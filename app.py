import os
import uuid
import json
import tempfile
from pathlib import Path
from flask import (
    Flask, request, session, redirect, url_for,
    render_template, send_file, jsonify, flash
)
from dotenv import load_dotenv

load_dotenv(override=True)

app = Flask(__name__)
app.secret_key = os.environ["SECRET_KEY"]
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB

UPLOAD_FOLDER = Path(tempfile.gettempdir()) / "rc_analyzer_uploads"
UPLOAD_FOLDER.mkdir(exist_ok=True)

APP_PASSWORD = os.environ["APP_PASSWORD"]


# ---------------------------------------------------------------------------
# Auth
# ---------------------------------------------------------------------------

ALLOWED_EMAIL_DOMAIN = "ringcentral.com"


@app.route("/login", methods=["GET", "POST"])
def login():
    error = None
    if request.method == "POST":
        email = request.form.get("email", "").strip().lower()
        password = request.form.get("password", "")
        if not email.endswith("@" + ALLOWED_EMAIL_DOMAIN):
            error = "Please sign in with your RingCentral email address."
        elif password == APP_PASSWORD:
            session["authed"] = True
            session["user_email"] = email
            return redirect(url_for("index"))
        else:
            error = "Incorrect email or password."
    return render_template("login.html", error=error)


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


def require_auth():
    if not session.get("authed"):
        return redirect(url_for("login"))
    return None


# ---------------------------------------------------------------------------
# Main app routes
# ---------------------------------------------------------------------------

@app.route("/")
def index():
    if not session.get("authed"):
        return redirect(url_for("login"))
    return render_template("index.html")


@app.route("/start", methods=["POST"])
def start():
    """Step 1: capture the customer URL + names, then crawl before anything else."""
    guard = require_auth()
    if guard:
        return guard

    customer = request.form.get("customer", "").strip()
    ae_name = request.form.get("ae_name", "").strip()
    company_url = request.form.get("company_url", "").strip()

    if not customer or not ae_name:
        flash("Please enter the customer name and your name.")
        return redirect(url_for("index"))

    run_id = uuid.uuid4().hex
    session["run_id"] = run_id
    session["customer"] = customer
    session["ae_name"] = ae_name
    session["company_url"] = company_url
    session["reporting_period"] = ""
    session["overrides"] = {}
    session["business_context"] = None
    session["messages"] = []
    session["filename"] = None

    if company_url:
        return redirect(url_for("discover", run_id=run_id))
    return redirect(url_for("upload_step", run_id=run_id))


@app.route("/discover/<run_id>")
def discover(run_id):
    guard = require_auth()
    if guard:
        return guard
    if session.get("run_id") != run_id:
        flash("Session mismatch — please start over.")
        return redirect(url_for("index"))
    return render_template(
        "discover.html",
        run_id=run_id,
        customer=session.get("customer"),
        company_url=session.get("company_url"),
    )


@app.route("/api/discover/<run_id>", methods=["POST"])
def api_discover(run_id):
    """Crawl the customer's website and return the business profile for confirmation."""
    guard = require_auth()
    if guard:
        return jsonify({"error": "Unauthorized"}), 401
    if session.get("run_id") != run_id:
        return jsonify({"error": "Session mismatch"}), 400

    company_url = session.get("company_url", "").strip()
    if not company_url:
        return jsonify({"available": False, "reason": "no_url"})

    biz = session.get("business_context")
    if biz is None:
        try:
            from business_context import build_business_context
            biz = build_business_context(company_url, session.get("customer", ""), [])
        except Exception as e:
            biz = {"available": False, "reason": f"error: {e}"}
        session["business_context"] = biz

    return jsonify(_business_summary_from(biz))


@app.route("/upload-step/<run_id>")
def upload_step(run_id):
    guard = require_auth()
    if guard:
        return guard
    if session.get("run_id") != run_id:
        flash("Session mismatch — please start over.")
        return redirect(url_for("index"))
    return render_template(
        "upload.html",
        run_id=run_id,
        customer=session.get("customer"),
        business=_business_summary_from(session.get("business_context")),
    )


@app.route("/upload/<run_id>", methods=["POST"])
def upload(run_id):
    guard = require_auth()
    if guard:
        return guard
    if session.get("run_id") != run_id:
        flash("Session mismatch — please start over.")
        return redirect(url_for("index"))

    file = request.files.get("report")
    if not file or not file.filename:
        flash("Please select a file.")
        return redirect(url_for("upload_step", run_id=run_id))
    if not file.filename.lower().endswith((".xlsx", ".xls")):
        flash("Only Excel files (.xlsx / .xls) are accepted.")
        return redirect(url_for("upload_step", run_id=run_id))

    upload_path = UPLOAD_FOLDER / f"{run_id}.xlsx"
    file.save(upload_path)
    session["filename"] = file.filename
    session["reporting_period"] = request.form.get("reporting_period", "").strip()

    return redirect(url_for("generate", run_id=run_id))


@app.route("/generate/<run_id>")
def generate(run_id):
    guard = require_auth()
    if guard:
        return guard

    if session.get("run_id") != run_id:
        flash("Session mismatch — please re-upload your file.")
        return redirect(url_for("index"))

    return render_template(
        "generate.html",
        run_id=run_id,
        filename=session.get("filename"),
        ae_name=session.get("ae_name"),
    )


@app.route("/api/process/<run_id>", methods=["POST"])
def api_process(run_id):
    """Run the data pipeline and generate the PPTX. Called via fetch from the browser."""
    guard = require_auth()
    if guard:
        return jsonify({"error": "Unauthorized"}), 401

    if session.get("run_id") != run_id:
        return jsonify({"error": "Session mismatch"}), 400

    upload_path = UPLOAD_FOLDER / f"{run_id}.xlsx"
    if not upload_path.exists():
        return jsonify({"error": "Uploaded file not found."}), 404

    try:
        pptx_path = _run_pipeline_and_build(run_id, upload_path, session.get("messages", []))
        session["pptx_path"] = str(pptx_path)
    except ValueError as e:
        return jsonify({"error": str(e)}), 400
    except Exception as e:
        return jsonify({"error": f"Processing error: {e}"}), 500

    return jsonify({
        "ok": True,
        "download_url": url_for("download", run_id=run_id),
        "business": _business_summary(),
    })


def _business_summary():
    """Lightweight business-context payload for the results UI."""
    return _business_summary_from(session.get("business_context"))


def _business_summary_from(biz):
    if not biz:
        return None
    if not biz.get("available"):
        return {"available": False, "reason": biz.get("reason", "")}
    return {
        "available": True,
        "summary": biz.get("summary", ""),
        "industry": biz.get("industry", ""),
        "lines_of_business": biz.get("lines_of_business", []),
        "predicted_call_reasons": [
            r.get("reason", "") for r in (biz.get("predicted_call_reasons") or [])
        ][:6],
        "suggested_avg_order_value": biz.get("suggested_avg_order_value"),
    }


def _run_pipeline_and_build(run_id, upload_path, messages):
    from pipeline import parse_sessions, build_result, distinct_queues
    from claude_client import classify_queues
    from deck import build_deck

    sdf = parse_sessions(upload_path)
    queues = distinct_queues(sdf)
    tiers = classify_queues(queues)
    result = build_result(sdf, tiers)

    override_period = session.get("reporting_period", "").strip()
    if override_period:
        result.reporting_period = override_period

    # Business context via Firecrawl (cached per run after first build)
    business_context = session.get("business_context")
    company_url = session.get("company_url", "").strip()
    if business_context is None and company_url:
        try:
            from business_context import build_business_context
            business_context = build_business_context(
                company_url, session.get("customer", ""), queues)
        except Exception as e:
            business_context = {"available": False, "reason": f"error: {e}"}
        session["business_context"] = business_context

    return build_deck(
        result=result,
        run_id=run_id,
        customer=session.get("customer", ""),
        ae_name=session.get("ae_name", ""),
        prior_instructions=messages,
        business_context=business_context,
        overrides=session.get("overrides", {}),
    )


@app.route("/download/<run_id>")
def download(run_id):
    guard = require_auth()
    if guard:
        return guard

    if session.get("run_id") != run_id:
        flash("Session mismatch.")
        return redirect(url_for("index"))

    pptx_path = session.get("pptx_path")
    if not pptx_path or not Path(pptx_path).exists():
        flash("No deck found — please generate one first.")
        return redirect(url_for("generate", run_id=run_id))

    filename = session.get("filename", "report").replace(".xlsx", "")
    return send_file(pptx_path, as_attachment=True, download_name=f"{filename}_AI_Business_Case.pptx")


@app.route("/api/refine/<run_id>", methods=["POST"])
def api_refine(run_id):
    """Accept a chat instruction, regenerate the deck."""
    guard = require_auth()
    if guard:
        return jsonify({"error": "Unauthorized"}), 401

    if session.get("run_id") != run_id:
        return jsonify({"error": "Session mismatch"}), 400

    body = request.get_json(silent=True) or {}
    instruction = (body.get("instruction") or "").strip()
    if not instruction:
        return jsonify({"error": "No instruction provided."}), 400

    messages = session.get("messages", [])
    messages.append({"role": "user", "content": instruction})
    session["messages"] = messages

    # Pull any structured facts the user supplied (order value, capture rate, etc.)
    applied = []
    try:
        from claude_client import extract_overrides
        ov = extract_overrides(instruction)
        overrides = session.get("overrides", {}) or {}
        for key, label in (("avg_order_value", "avg order value"),
                           ("capture_rate", "capture rate"),
                           ("air_rate_per_min", "AIR $/min")):
            val = ov.get(key)
            if val is not None:
                overrides[key] = val
                applied.append(f"{label} = {val}")
        session["overrides"] = overrides
    except Exception:
        pass

    upload_path = UPLOAD_FOLDER / f"{run_id}.xlsx"
    if not upload_path.exists():
        return jsonify({"error": "Uploaded file not found."}), 404

    try:
        pptx_path = _run_pipeline_and_build(run_id, upload_path, messages)
        session["pptx_path"] = str(pptx_path)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

    return jsonify({
        "ok": True,
        "download_url": url_for("download", run_id=run_id),
        "applied": applied,
    })


if __name__ == "__main__":
    app.run(debug=True)
