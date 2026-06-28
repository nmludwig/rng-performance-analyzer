"""
Firecrawl-powered business context.

Given a company URL, crawl the homepage (and a couple of obvious pages),
then ask Claude to build a business profile + predicted call reasons that the
deck uses to tailor the analysis. The whole feature degrades gracefully:
if there's no FIRECRAWL_API_KEY, no SDK, or the crawl fails, we return None
and the deck is built without business context.
"""

from __future__ import annotations
import os
import re


def _normalize_url(url: str) -> str:
    url = (url or "").strip()
    if not url:
        return ""
    if not re.match(r"^https?://", url, re.I):
        url = "https://" + url
    return url


def _scrape_markdown(app, url: str, timeout_ms: int = 25000) -> str | None:
    """Scrape one URL to markdown, tolerant of firecrawl-py v1 and v2 shapes.

    A hard ``timeout_ms`` is passed through so a slow/unresponsive target
    fails fast instead of blocking the (single) gunicorn worker for minutes —
    a hung scrape here previously starved Render's health check and cancelled
    deploys. The firecrawl SDK adds +5s to the request timeout internally.
    """
    # v1: app.scrape_url(url, params={"formats": ["markdown"]})
    # v2: app.scrape(url, formats=["markdown"])
    result = None
    for attempt in (
        lambda: app.scrape_url(url, params={"formats": ["markdown"]}, timeout=timeout_ms),
        lambda: app.scrape_url(url, formats=["markdown"], timeout=timeout_ms),
        lambda: app.scrape(url, formats=["markdown"], timeout=timeout_ms),
        lambda: app.scrape_url(url, params={"formats": ["markdown"]}),
        lambda: app.scrape_url(url, formats=["markdown"]),
        lambda: app.scrape(url, formats=["markdown"]),
        lambda: app.scrape_url(url),
    ):
        try:
            result = attempt()
            break
        except TypeError:
            continue
        except Exception:
            continue
    if result is None:
        return None
    # Unwrap a variety of return shapes
    if isinstance(result, dict):
        if "markdown" in result:
            return result["markdown"]
        data = result.get("data") or {}
        if isinstance(data, dict):
            return data.get("markdown")
        return None
    # Object with attributes (v2 pydantic models)
    md = getattr(result, "markdown", None)
    if md:
        return md
    data = getattr(result, "data", None)
    if data is not None:
        return getattr(data, "markdown", None)
    return None


def crawl_company(url: str, max_chars: int = 14000) -> dict:
    """Return {"ok": bool, "markdown": str|None, "url": str, "reason": str}."""
    url = _normalize_url(url)
    if not url:
        return {"ok": False, "markdown": None, "url": "", "reason": "no_url"}

    api_key = os.environ.get("FIRECRAWL_API_KEY")
    if not api_key:
        return {"ok": False, "markdown": None, "url": url, "reason": "no_key"}

    try:
        from firecrawl import FirecrawlApp as _Client
    except Exception:
        try:
            from firecrawl import Firecrawl as _Client
        except Exception:
            return {"ok": False, "markdown": None, "url": url, "reason": "no_sdk"}

    try:
        app = _Client(api_key=api_key)
    except Exception as e:
        return {"ok": False, "markdown": None, "url": url, "reason": f"init_error: {e}"}

    md = _scrape_markdown(app, url)
    if not md:
        return {"ok": False, "markdown": None, "url": url, "reason": "scrape_failed"}

    md = md[:max_chars]
    return {"ok": True, "markdown": md, "url": url, "reason": "ok"}


def build_business_context(url: str, customer: str, queue_names: list[str]) -> dict | None:
    """Crawl + profile. Returns a profile dict or None if unavailable."""
    crawl = crawl_company(url)
    if not crawl["ok"]:
        return {"available": False, "reason": crawl["reason"], "url": crawl["url"]}

    from claude_client import profile_business
    profile = profile_business(crawl["markdown"], customer, queue_names)
    profile["available"] = True
    profile["url"] = crawl["url"]
    return profile
