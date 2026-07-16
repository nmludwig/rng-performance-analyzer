"""Lightweight append-only log of deck-generation runs.

We deliberately store NO customer report data — only who ran a generation, when,
whether it succeeded, and a few operational details useful for an admin view.
Records are appended as JSON lines to a file next to the app; the file is
git-ignored and never leaves the server.
"""
import json
import threading
from datetime import datetime, timezone
from pathlib import Path

LOG_PATH = Path(__file__).resolve().parent / "run_log.jsonl"
_LOCK = threading.Lock()


def log_run(record: dict) -> None:
    """Append one run record. Never raises — logging must not break a run."""
    try:
        record = {"ts": datetime.now(timezone.utc).isoformat(timespec="seconds"), **record}
        line = json.dumps(record, ensure_ascii=False)
        with _LOCK:
            with open(LOG_PATH, "a", encoding="utf-8") as fh:
                fh.write(line + "\n")
    except Exception:
        pass


def read_runs(limit: int = 500) -> list[dict]:
    """Return the most recent runs, newest first."""
    if not LOG_PATH.exists():
        return []
    rows = []
    try:
        with open(LOG_PATH, "r", encoding="utf-8") as fh:
            for line in fh:
                line = line.strip()
                if not line:
                    continue
                try:
                    rows.append(json.loads(line))
                except Exception:
                    continue
    except Exception:
        return []
    rows.reverse()
    return rows[:limit]
