# Gunicorn config — auto-loaded from the working directory even when the
# start command is just `gunicorn app:app --bind ...` (e.g. when Render's
# dashboard start command overrides render.yaml and drops the CLI flags).
#
# The streamed deck-generation endpoint (/api/process_stream) holds a sync
# worker for the whole job (Excel parse + Firecrawl + 2 Claude calls + PPTX
# build). With gunicorn's DEFAULT 30s sync timeout that worker is killed
# mid-stream, surfacing as "The connection to the server was lost." A long
# timeout keeps the streaming request alive for the full job.

import os

# Long enough for the entire generation to stream without the worker being
# reaped. Generation is typically 30-60s; 300s leaves generous headroom.
timeout = 300

# Honor Render's WEB_CONCURRENCY when set; otherwise default to 2 so a
# long-running stream on one worker doesn't block every other request.
workers = int(os.environ.get("WEB_CONCURRENCY", "2"))

worker_class = "sync"

# Don't recycle the worker mid-stream.
graceful_timeout = 300
keepalive = 5
