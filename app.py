import os
import time
import uuid
import json
import base64
import secrets
import tempfile
import urllib.parse
import urllib.request
from datetime import datetime, timezone
from pathlib import Path
from flask import (
    Flask, request, session, redirect, url_for,
    render_template, send_file, jsonify, flash, Response, stream_with_context
)
from dotenv import load_dotenv

from run_log import log_run, read_runs, log_idea, read_ideas

load_dotenv(override=True)

app = Flask(__name__)
app.secret_key = os.environ["SECRET_KEY"]
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB

APP_VERSION = "1.0.0"

UPLOAD_FOLDER = Path(tempfile.gettempdir()) / "rc_analyzer_uploads"
UPLOAD_FOLDER.mkdir(exist_ok=True)

# ---------------------------------------------------------------------------
# Auth — RingCentral corporate single sign-on (OAuth 2.0 authorization code)
# ---------------------------------------------------------------------------
#
# Employees log in with their real RingCentral credentials. We never pull any
# customer data with this login — it only proves the person is a RingCentral
# employee. After RingCentral authenticates them we read their profile email and
# require it to be an @ringcentral.com address.

RC_CLIENT_ID = os.environ.get("RC_CLIENT_ID", "")
RC_CLIENT_SECRET = os.environ.get("RC_CLIENT_SECRET", "")
RC_REDIRECT_URI = os.environ.get("RC_REDIRECT_URI", "")
RC_SERVER_URL = os.environ.get("RC_SERVER_URL", "https://platform.ringcentral.com").rstrip("/")

ALLOWED_EMAIL_DOMAIN = "ringcentral.com"


def _sso_configured():
    return bool(RC_CLIENT_ID and RC_CLIENT_SECRET and RC_REDIRECT_URI)


@app.route("/login")
def login():
    if session.get("authed"):
        return redirect(url_for("index"))
    error = request.args.get("error")
    if not _sso_configured():
        error = ("Single sign-on is not configured. Set RC_CLIENT_ID, "
                 "RC_CLIENT_SECRET and RC_REDIRECT_URI on the server.")
    return render_template("login.html", error=error, sso_ready=_sso_configured())


@app.route("/oauth/start")
def oauth_start():
    if not _sso_configured():
        return redirect(url_for("login"))
    state = secrets.token_urlsafe(24)
    session["oauth_state"] = state
    params = urllib.parse.urlencode({
        "response_type": "code",
        "client_id": RC_CLIENT_ID,
        "redirect_uri": RC_REDIRECT_URI,
        "state": state,
    })
    return redirect(f"{RC_SERVER_URL}/restapi/oauth/authorize?{params}")


@app.route("/oauth/callback")
def oauth_callback():
    if request.args.get("error"):
        return redirect(url_for("login", error="RingCentral sign-in was cancelled."))

    code = request.args.get("code")
    state = request.args.get("state")
    expected = session.pop("oauth_state", None)
    if not code or not state or not expected or state != expected:
        return redirect(url_for("login", error="Sign-in expired or invalid — please try again."))

    try:
        access_token = _rc_exchange_code(code)
        email = _rc_fetch_email(access_token)
    except Exception:
        return redirect(url_for("login", error="Could not complete RingCentral sign-in — please try again."))

    email = (email or "").strip().lower()
    if not email.endswith("@" + ALLOWED_EMAIL_DOMAIN):
        return redirect(url_for("login", error="This tool is restricted to RingCentral employees."))

    session["authed"] = True
    session["user_email"] = email
    _glip_event("🔓", "login", [
        ("User", email),
        ("Time (UTC)", _utc_now()),
        ("IP", _client_ip()),
    ])
    return redirect(url_for("index"))


def _rc_exchange_code(code):
    """Trade the authorization code for an access token (HTTP Basic client auth)."""
    data = urllib.parse.urlencode({
        "grant_type": "authorization_code",
        "code": code,
        "redirect_uri": RC_REDIRECT_URI,
    }).encode()
    basic = base64.b64encode(f"{RC_CLIENT_ID}:{RC_CLIENT_SECRET}".encode()).decode()
    req = urllib.request.Request(
        f"{RC_SERVER_URL}/restapi/oauth/token", data=data, method="POST",
        headers={
            "Authorization": f"Basic {basic}",
            "Content-Type": "application/x-www-form-urlencoded",
            "Accept": "application/json",
        },
    )
    with urllib.request.urlopen(req, timeout=15) as resp:
        payload = json.loads(resp.read().decode())
    return payload["access_token"]


def _rc_fetch_email(access_token):
    """Read the signed-in extension's contact email to enforce the employee check."""
    req = urllib.request.Request(
        f"{RC_SERVER_URL}/restapi/v1.0/account/~/extension/~",
        headers={"Authorization": f"Bearer {access_token}", "Accept": "application/json"},
    )
    with urllib.request.urlopen(req, timeout=15) as resp:
        payload = json.loads(resp.read().decode())
    return (payload.get("contact") or {}).get("email", "")


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


# Only these employees may view the admin run log. Comma-separated override via
# ADMIN_EMAILS in the environment; defaults to the tool owner.
ADMIN_EMAILS = {
    e.strip().lower()
    for e in os.environ.get("ADMIN_EMAILS", "matthew.ludwig@ringcentral.com").split(",")
    if e.strip()
}


# Optional RingCentral Glip incoming-webhook URL. Set GLIP_WEBHOOK_URL in the
# environment (.env) to a https://hooks.ringcentral.com/webhook/v2/... address
# to mirror idea submissions and app runs into a team chat. Never hardcode it.
GLIP_WEBHOOK_URL = os.environ.get("GLIP_WEBHOOK_URL", "").strip()


def notify_glip(text: str) -> None:
    """Post a short plain-text notification to the configured Glip webhook.
    Best-effort: any failure (no URL, network error, bad response) is swallowed
    so it can never disrupt the user-facing request."""
    if not GLIP_WEBHOOK_URL:
        return
    try:
        # ensure_ascii=True keeps the body pure ASCII (emoji become \uXXXX
        # escapes, which RingCentral renders correctly) so the request never
        # depends on the server's locale for encoding.
        payload = json.dumps(
            {"activity": "AIR Pro Performance Analyzer", "text": text},
            ensure_ascii=True).encode("ascii")
        req = urllib.request.Request(
            GLIP_WEBHOOK_URL, data=payload,
            headers={"Content-Type": "application/json"}, method="POST")
        resp = urllib.request.urlopen(req, timeout=5)
        resp.read()
    except Exception as e:
        try:
            app.logger.warning("notify_glip failed: %s: %s", type(e).__name__, e)
        except Exception:
            pass


def _client_ip() -> str:
    """Best-effort real client IP, honoring the nginx reverse proxy in front of
    gunicorn. X-Forwarded-For is a comma list (client, proxy1, ...); take the
    first entry. Falls back to remote_addr when there is no proxy header."""
    try:
        fwd = request.headers.get("X-Forwarded-For", "")
        if fwd:
            return fwd.split(",")[0].strip()
        return request.headers.get("X-Real-IP", "").strip() or (request.remote_addr or "")
    except Exception:
        return ""


def _glip_event(emoji: str, title: str, fields: list[tuple[str, str]],
                extra: str = "") -> None:
    """Post a richly-formatted event to Glip: a bold header line, then a set of
    labeled fields (User / Time (UTC) / IP / ...), matching the fuller format
    used by our other internal tools. Any field with an empty value is skipped.
    Best-effort — delegates delivery (and failure-swallowing) to notify_glip."""
    lines = [f"{emoji} **AIR Pro Performance Analyzer — {title}**"]
    for label, value in fields:
        value = (str(value) if value is not None else "").strip()
        if value:
            lines.append(f"- **{label}:** {value}")
    if extra:
        lines.append("")
        lines.append(extra)
    notify_glip("\n".join(lines))


def _utc_now() -> str:
    """Current time as an ISO-8601 UTC timestamp, e.g. 2026-07-20T17:37:50+00:00."""
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


@app.context_processor
def _inject_is_admin():
    return {"is_admin": session.get("user_email", "").strip().lower() in ADMIN_EMAILS,
            "app_version": APP_VERSION}


@app.route("/admin")
def admin():
    if not session.get("authed"):
        return redirect(url_for("login"))
    if session.get("user_email", "").strip().lower() not in ADMIN_EMAILS:
        return render_template("admin.html", runs=None, ideas=None, forbidden=True), 403
    return render_template("admin.html", runs=read_runs(), ideas=read_ideas(),
                           forbidden=False)


@app.route("/ideas", methods=["GET", "POST"])
def ideas():
    """Any signed-in employee can suggest a product improvement; submissions
    surface in the admin view. No customer data is involved."""
    if not session.get("authed"):
        return redirect(url_for("login"))
    if request.method == "POST":
        title = request.form.get("title", "").strip()
        detail = request.form.get("detail", "").strip()
        category = request.form.get("category", "").strip()
        if not title:
            flash("Please add a short summary of your idea.")
            return redirect(url_for("ideas"))
        submitter = session.get("user_email", "")
        log_idea({
            "title": title[:200],
            "detail": detail[:4000],
            "category": category[:60],
            "user_email": submitter,
        })
        _glip_event("💡", "new idea", [
            ("User", submitter or "a user"),
            ("Category", category[:60]),
            ("Idea", f"“{title[:200]}”"),
            ("Time (UTC)", _utc_now()),
            ("IP", _client_ip()),
        ], extra=(detail[:500] if detail else ""))
        flash("Thanks! Your idea was sent to the team.")
        return redirect(url_for("ideas", submitted=1))
    return render_template("ideas.html", submitted=request.args.get("submitted"))


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

    # Single-file model: the RingCentral Business Analytics "Call Records"
    # export now carries every outcome (answered / missed / voicemail /
    # abandoned) in one file, so there is no longer a second Queues upload.
    file = request.files.get("report")

    if not file or not file.filename:
        flash("Please select the Business Analytics Call Records export.")
        return redirect(url_for("upload_step", run_id=run_id))
    if not file.filename.lower().endswith((".xlsx", ".xls")):
        flash("Only Excel files (.xlsx / .xls) are accepted.")
        return redirect(url_for("upload_step", run_id=run_id))

    # Non-blocking filename-structure warning: the Business Analytics export
    # downloads as "Call_Records_…". A mismatched name usually means the wrong
    # report (e.g. an old Performance Report) was uploaded. Warn but proceed.
    if not file.filename.lower().startswith("call_records"):
        flash(
            f"Heads up: the file you uploaded (“{file.filename}”) doesn’t start "
            "with “Call_Records…”. Make sure it’s the Business Analytics Call "
            "Records export, not an older Performance Report.")

    upload_path = UPLOAD_FOLDER / f"{run_id}.xlsx"
    file.save(upload_path)

    session["filename"] = file.filename
    session["queues_filename"] = ""
    session["reporting_period"] = request.form.get("reporting_period", "").strip()

    # Average order value: the one ROI input not measured from the call logs. When
    # the AE enters it, store it as a supplied override so the deck treats it as the
    # customer's own figure (not an AI estimate). Blank → fall back to crawl estimate.
    aov_raw = (request.form.get("avg_order_value") or "").strip().replace(",", "")
    if aov_raw:
        try:
            aov_val = int(float(aov_raw))
            if aov_val > 0:
                overrides = session.get("overrides", {}) or {}
                overrides["avg_order_value"] = aov_val
                session["overrides"] = overrides
        except ValueError:
            pass

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
    if not upload_path.exists():
        return Response(
            "data: " + json.dumps({"error": "Uploaded file not found — please re-upload."}) + "\n\n",
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
    user_email = session.get("user_email", "")
    filename = session.get("filename", "")
    queues_filename = session.get("queues_filename", "")
    client_ip = _client_ip()

    @stream_with_context
    def gen():
        def ev(**payload):
            return "data: " + json.dumps(payload) + "\n\n"

        started = time.monotonic()
        run_meta = {
            "run_id": run_id,
            "user_email": user_email,
            "customer": customer,
            "ae_name": ae_name,
            "company_url": company_url,
            "reporting_period": reporting_period,
            "calls_file": filename,
            "queues_file": queues_filename,
            "refine": bool(messages),
        }
        _glip_event("▶️", "run started", [
            ("User", user_email or "a user"),
            ("Customer", customer or "a customer"),
            ("Mode", "refine" if messages else "new deck"),
            ("Calls file", filename),
            ("Queues file", queues_filename),
            ("Time (UTC)", _utc_now()),
            ("IP", client_ip),
        ])

        try:
            from pipeline import (parse_business_analytics, build_result,
                                  ba_queue_tiers)
            from deck import build_deck

            yield ev(stage="parse", msg="Reading the Business Analytics export…", pct=10)
            sdf = parse_business_analytics(upload_path)
            n_sessions = len(sdf)
            yield ev(stage="parse",
                     msg=f"De-duplicated call legs into {n_sessions:,} inbound calls.", pct=30)

            # No queue dimension in Business Analytics call records — all inbound
            # is one synthetic 'Direct line' bucket, so no queue classification.
            queues = [next(iter(ba_queue_tiers()))]
            tiers = ba_queue_tiers()
            queues_report = None

            biz = business_context
            if biz is None and company_url:
                yield ev(stage="profile", msg="Profiling the business from its website…", pct=42)
                try:
                    from business_context import build_business_context
                    biz = build_business_context(company_url, customer, queues)
                except Exception as e:
                    biz = {"available": False, "reason": f"error: {e}"}

            yield ev(stage="analyze", msg="Calculating the missed-call impact…", pct=62)
            result = build_result(sdf, tiers, queues_report=queues_report)
            if reporting_period:
                result.reporting_period = reporting_period

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

            _dur = round(time.monotonic() - started, 1)
            log_run({**run_meta, "status": "success",
                     "duration_s": _dur,
                     "sessions": int(n_sessions), "queues": len(queues)})
            _glip_event("✅", "deck ready", [
                ("User", user_email or "a user"),
                ("Customer", customer or "a customer"),
                ("Sessions", f"{n_sessions:,}"),
                ("Queues", str(len(queues))),
                ("Duration", f"{_dur}s"),
                ("Time (UTC)", _utc_now()),
                ("IP", client_ip),
            ])
            yield ev(stage="done", msg="Deck ready.", pct=100,
                     download_url=download_url,
                     business=_business_summary_from(biz))
        except ValueError as e:
            log_run({**run_meta, "status": "error",
                     "duration_s": round(time.monotonic() - started, 1),
                     "error": str(e)})
            _glip_event("❌", "run failed", [
                ("User", user_email or "a user"),
                ("Customer", customer or "a customer"),
                ("Error", str(e)),
                ("Time (UTC)", _utc_now()),
                ("IP", client_ip),
            ])
            yield ev(error=str(e))
        except Exception as e:
            log_run({**run_meta, "status": "error",
                     "duration_s": round(time.monotonic() - started, 1),
                     "error": f"Processing error: {e}"})
            _glip_event("❌", "run failed", [
                ("User", user_email or "a user"),
                ("Customer", customer or "a customer"),
                ("Error", f"Processing error: {e}"),
                ("Time (UTC)", _utc_now()),
                ("IP", client_ip),
            ])
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
        "aov_basis": biz.get("aov_basis", ""),
    }


def _run_pipeline_and_build(run_id, upload_path, messages):
    from pipeline import parse_business_analytics, build_result, ba_queue_tiers
    from deck import build_deck

    sdf = parse_business_analytics(upload_path)

    # Business Analytics call records carry no queue dimension — all inbound is
    # one synthetic 'Direct line' bucket, and the Result column carries every
    # outcome, so there is no second Queues report and no queue classification.
    queues = [next(iter(ba_queue_tiers()))]
    tiers = ba_queue_tiers()
    queues_report = None

    # Business context via Firecrawl is crawled up-front (discover step) and
    # cached in the session. Build it here only as a fallback.
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

    result = build_result(sdf, tiers, queues_report=queues_report)

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

    # The heavy regeneration is streamed by /api/process_stream so the browser
    # gets live per-stage progress. Here we only persist the instruction and any
    # extracted overrides to the session; the stream reads them back and rebuilds.
    return jsonify({
        "ok": True,
        "stream_url": url_for("api_process_stream", run_id=run_id),
        "applied": applied,
    })


if __name__ == "__main__":
    app.run(debug=True)
