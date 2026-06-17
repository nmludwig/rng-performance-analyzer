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

load_dotenv()

app = Flask(__name__)
app.secret_key = os.environ["SECRET_KEY"]
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB

UPLOAD_FOLDER = Path(tempfile.gettempdir()) / "rc_analyzer_uploads"
UPLOAD_FOLDER.mkdir(exist_ok=True)

APP_PASSWORD = os.environ["APP_PASSWORD"]


# ---------------------------------------------------------------------------
# Auth
# ---------------------------------------------------------------------------

@app.route("/login", methods=["GET", "POST"])
def login():
    error = None
    if request.method == "POST":
        if request.form.get("password") == APP_PASSWORD:
            session["authed"] = True
            return redirect(url_for("index"))
        error = "Incorrect password."
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


@app.route("/upload", methods=["POST"])
def upload():
    guard = require_auth()
    if guard:
        return guard

    file = request.files.get("report")
    if not file or not file.filename:
        flash("Please select a file.")
        return redirect(url_for("index"))

    if not file.filename.lower().endswith((".xlsx", ".xls")):
        flash("Only Excel files (.xlsx / .xls) are accepted.")
        return redirect(url_for("index"))

    run_id = uuid.uuid4().hex
    upload_path = UPLOAD_FOLDER / f"{run_id}.xlsx"
    file.save(upload_path)

    # Store run metadata in session
    session["run_id"] = run_id
    session["filename"] = file.filename
    session["messages"] = []

    return redirect(url_for("configure", run_id=run_id))


@app.route("/configure/<run_id>", methods=["GET", "POST"])
def configure(run_id):
    guard = require_auth()
    if guard:
        return guard

    if session.get("run_id") != run_id:
        flash("Session mismatch — please re-upload your file.")
        return redirect(url_for("index"))

    if request.method == "POST":
        ae_name = request.form.get("ae_name", "").strip()
        avg_deal = request.form.get("avg_deal_value", "").strip()
        close_rate = request.form.get("close_rate", "").strip()

        try:
            avg_deal_val = float(avg_deal.replace(",", "").replace("$", ""))
            close_rate_val = float(close_rate.replace("%", "")) / 100
        except ValueError:
            flash("Please enter valid numbers for deal value and close rate.")
            return redirect(url_for("configure", run_id=run_id))

        session["ae_name"] = ae_name
        session["avg_deal_value"] = avg_deal_val
        session["close_rate"] = close_rate_val

        return redirect(url_for("generate", run_id=run_id))

    return render_template("configure.html", run_id=run_id, filename=session.get("filename"))


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

    from pipeline import parse_report
    from deck import build_deck

    try:
        results = parse_report(upload_path)
    except Exception as e:
        return jsonify({"error": f"Data pipeline error: {e}"}), 500

    try:
        pptx_path = build_deck(
            results=results,
            run_id=run_id,
            ae_name=session.get("ae_name", ""),
            avg_deal_value=session.get("avg_deal_value", 0),
            close_rate=session.get("close_rate", 0),
            prior_instructions=session.get("messages", []),
        )
        session["pptx_path"] = str(pptx_path)
    except Exception as e:
        return jsonify({"error": f"Deck generation error: {e}"}), 500

    return jsonify({"ok": True, "download_url": url_for("download", run_id=run_id)})


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

    upload_path = UPLOAD_FOLDER / f"{run_id}.xlsx"
    if not upload_path.exists():
        return jsonify({"error": "Uploaded file not found."}), 404

    from pipeline import parse_report
    from deck import build_deck

    try:
        results = parse_report(upload_path)
        pptx_path = build_deck(
            results=results,
            run_id=run_id,
            ae_name=session.get("ae_name", ""),
            avg_deal_value=session.get("avg_deal_value", 0),
            close_rate=session.get("close_rate", 0),
            prior_instructions=messages,
        )
        session["pptx_path"] = str(pptx_path)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

    return jsonify({"ok": True, "download_url": url_for("download", run_id=run_id)})


if __name__ == "__main__":
    app.run(debug=True)
