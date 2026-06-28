import os
import uuid
import json
import tempfile
from pathlib import Path
from flask import (
    Flask, request, session, redirect, url_for,
    render_template, send_file, jsonify, flash, Response, stream_with_context
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
    queues_file = request.files.get("queues_report")

    if not file or not file.filename:
        flash("Please select the Calls export.")
        return redirect(url_for("upload_step", run_id=run_id))
    if not file.filename.lower().endswith((".xlsx", ".xls")):
        flash("Only Excel files (.xlsx / .xls) are accepted for the Calls export.")
        return redirect(url_for("upload_step", run_id=run_id))

    if not queues_file or not queues_file.filename:
        flash("Please also upload the Queues report — it carries the abandoned-call data.")
        return redirect(url_for("upload_step", run_id=run_id))
    if not queues_file.filename.lower().endswith((".xlsx", ".xls")):
        flash("Only Excel files (.xlsx / .xls) are accepted for the Queues report.")
        return redirect(url_for("upload_step", run_id=run_id))

    # Optional 3rd/4th uploads: the CN Calls report (splits non-queue misses by
    # destination) and the Company Numbers report (per-number rollup + labels).
    cn_calls_file = request.files.get("cn_calls")
    company_numbers_file = request.files.get("company_numbers")
    for label, f in (("CN Calls report", cn_calls_file),
                     ("Company Numbers report", company_numbers_file)):
        if f and f.filename and not f.filename.lower().endswith((".xlsx", ".xls")):
            flash(f"Only Excel files (.xlsx / .xls) are accepted for the {label}.")
            return redirect(url_for("upload_step", run_id=run_id))

    upload_path = UPLOAD_FOLDER / f"{run_id}.xlsx"
    file.save(upload_path)
    queues_path = UPLOAD_FOLDER / f"{run_id}_queues.xlsx"
    queues_file.save(queues_path)

    if cn_calls_file and cn_calls_file.filename:
        cn_calls_file.save(UPLOAD_FOLDER / f"{run_id}_cncalls.xlsx")
        session["cn_calls_filename"] = cn_calls_file.filename
    if company_numbers_file and company_numbers_file.filename:
        company_numbers_file.save(UPLOAD_FOLDER / f"{run_id}_company.xlsx")
        session["company_numbers_filename"] = company_numbers_file.filename

    session["filename"] = file.filename
    session["queues_filename"] = queues_file.filename
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


@app.route("/api/process_stream/<run_id>")
def api_process_stream(run_id):
    """Run the pipeline while streaming real per-stage progress to the browser (SSE).

    The heavy work runs inside the response generator, so each ``yield`` flushes
    a progress event as that stage completes. We only READ the Flask session
    (writing mid-stream is impossible once headers are sent); the finished deck
    is found on the deterministic path ``<tmp>/rc_analyzer_decks/{run_id}.pptx``
    by the download route, so no session write is needed.
    """
    guard = require_auth()
    if guard:
        return guard
    if session.get("run_id") != run_id:
        return Response(
            "data: " + json.dumps({"error": "Session mismatch — please start over."}) + "\n\n",
            mimetype="text/event-stream",
        )

    upload_path = UPLOAD_FOLDER / f"{run_id}.xlsx"
    queues_path = UPLOAD_FOLDER / f"{run_id}_queues.xlsx"
    if not upload_path.exists() or not queues_path.exists():
        return Response(
            "data: " + json.dumps({"error": "Uploaded files not found — please re-upload."}) + "\n\n",
            mimetype="text/event-stream",
        )

    # Capture everything we need from the session up-front (no mid-stream reads).
    customer = session.get("customer", "")
    ae_name = session.get("ae_name", "")
    company_url = session.get("company_url", "").strip()
    business_context = session.get("business_context")
    reporting_period = session.get("reporting_period", "").strip()
    overrides = session.get("overrides", {}) or {}
    messages = session.get("messages", [])
    download_url = url_for("download", run_id=run_id)
    biz_summary_seed = business_context

    @stream_with_context
    def gen():
        def ev(**payload):
            return "data: " + json.dumps(payload) + "\n\n"

        try:
            from pipeline import (parse_sessions, build_result, distinct_queues,
                                  parse_queues_report, queues_report_queue_names)
            from claude_client import classify_queues
            from deck import build_deck

            yield ev(stage="parse", msg="Reading the call export…", pct=8)
            sdf = parse_sessions(upload_path)
            n_sessions = len(sdf)
            yield ev(stage="parse",
                     msg=f"De-duplicated call legs into {n_sessions:,} sessions.", pct=24)

            queues = distinct_queues(sdf)

            yield ev(stage="queues", msg="Reading the Queues report (abandoned calls)…", pct=32)
            queues_report = parse_queues_report(queues_path)
            queues = sorted(set(queues) | set(queues_report_queue_names(queues_report)))

            biz = business_context
            if biz is None and company_url:
                yield ev(stage="profile", msg="Profiling the business from its website…", pct=36)
                try:
                    from business_context import build_business_context
                    biz = build_business_context(company_url, customer, queues)
                except Exception as e:
                    biz = {"available": False, "reason": f"error: {e}"}

            yield ev(stage="classify",
                     msg=f"Classifying {len(queues)} call queues by revenue relevance…", pct=46)
            tiers = classify_queues(queues, business_context=biz)

            yield ev(stage="analyze", msg="Calculating the missed-call impact…", pct=62)
            result = build_result(sdf, tiers, queues_report=queues_report)
            if reporting_period:
                result.reporting_period = reporting_period
            _attach_call_destinations(result, run_id)

            yield ev(stage="deck", msg="Writing the slides with Claude…", pct=78)
            build_deck(
                result=result,
                run_id=run_id,
                customer=customer,
                ae_name=ae_name,
                prior_instructions=messages,
                business_context=biz,
                overrides=overrides,
            )

            yield ev(stage="done", msg="Deck ready.", pct=100,
                     download_url=download_url,
                     business=_business_summary_from(biz))
        except ValueError as e:
            yield ev(error=str(e))
        except Exception as e:
            yield ev(error=f"Processing error: {e}")

    return Response(gen(), mimetype="text/event-stream",
                    headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"})


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


def _attach_call_destinations(result, run_id):
    """Parse the optional CN Calls / Company Numbers uploads and attach the
    non-queue destination breakdown to the result. Silently no-ops if absent."""
    from pipeline import parse_call_destinations, parse_company_numbers

    company_path = UPLOAD_FOLDER / f"{run_id}_company.xlsx"
    cncalls_path = UPLOAD_FOLDER / f"{run_id}_cncalls.xlsx"
    labels = {}
    if company_path.exists():
        try:
            cn = parse_company_numbers(company_path)
            result.company_numbers = cn
            labels = cn.labels
        except Exception:
            pass
    if cncalls_path.exists():
        try:
            result.call_destinations = parse_call_destinations(cncalls_path, number_labels=labels)
        except Exception:
            pass
    return result


def _run_pipeline_and_build(run_id, upload_path, messages):
    from pipeline import (parse_sessions, build_result, distinct_queues,
                          parse_queues_report, queues_report_queue_names)
    from claude_client import classify_queues
    from deck import build_deck

    sdf = parse_sessions(upload_path)
    queues = distinct_queues(sdf)

    # Mandatory Queues report (2nd upload) — the only source of abandoned data.
    queues_path = UPLOAD_FOLDER / f"{run_id}_queues.xlsx"
    if not queues_path.exists():
        raise ValueError("The Queues report is missing — please re-upload both files.")
    queues_report = parse_queues_report(queues_path)
    # Tier both queue-name sets together so the abandoned table can be tiered.
    queues = sorted(set(queues) | set(queues_report_queue_names(queues_report)))

    # Business context via Firecrawl is crawled up-front (discover step) and
    # cached in the session. Build it here only as a fallback, BEFORE tiering,
    # so queue tiers can be reasoned from this customer's actual business.
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

    tiers = classify_queues(queues, business_context=business_context)
    result = build_result(sdf, tiers, queues_report=queues_report)
    _attach_call_destinations(result, run_id)

    override_period = session.get("reporting_period", "").strip()
    if override_period:
        result.reporting_period = override_period

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
        # SSE generation can't write the session mid-stream; fall back to the
        # deterministic deck path build_deck always writes.
        deterministic = Path(tempfile.gettempdir()) / "rc_analyzer_decks" / f"{run_id}.pptx"
        if deterministic.exists():
            pptx_path = str(deterministic)
        else:
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
