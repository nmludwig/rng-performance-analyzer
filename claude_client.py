"""
Server-side Claude API integration.
- Queue tiering: classify queue names into Tier A/B/C/D.
- Slide narrative text generation.
"""

from __future__ import annotations
import os
import json
import anthropic

_client: anthropic.Anthropic | None = None
MODEL = "claude-sonnet-4-6"


def get_client() -> anthropic.Anthropic:
    global _client
    if _client is None:
        _client = anthropic.Anthropic(api_key=os.environ["ANTHROPIC_API_KEY"])
    return _client


def _extract_json(text: str) -> dict | list:
    text = text.strip()
    if text.startswith("```"):
        text = text.split("```")[1]
        if text.lstrip().startswith("json"):
            text = text.lstrip()[4:]
    return json.loads(text)


TIERING_SYSTEM = """You classify RingCentral call-queue names into business tiers for a sales analysis.

Tiers:
- A = Named sales queue. The queue name explicitly indicates direct sales (contains "Sales", or is an obvious revenue line).
- B = Customer-facing product / retail counter. Product counters, retail, "Building Products", "Doors/Hardware", lumber, will-call — customer-facing but not explicitly "Sales".
- C = Branch main line or front desk. General "Main", "Front Desk", "Operator", "Main CQ", branch reception lines that a customer would call and that can carry sales intent.
- D = Back-office / internal. NOT customer-revenue facing: IT, Support Center, Infrastructure, Network, Accounting, Estimating, Shipping and Receiving, Credit, AP/AR, HR, Systems. These are EXCLUDED from the sales analysis.

Rules:
- Use only the queue name text to decide.
- When a name clearly contains "Sales" -> A.
- Shipping/Receiving, Credit, Accounting, IT, Support, Network, Infrastructure, Estimating, Systems -> D.
- Generic branch "Main"/"Front Desk"/"Operator" -> C.
- Respond with JSON only: an array of {"queue": <exact name>, "tier": "A|B|C|D", "classification": <short label>}.
- classification labels: A -> "Named 'Sales' queue", B -> "Customer-facing product / retail counter", C -> "Branch main line or front desk", D -> "Back-office / internal (excluded)".
- Include every queue exactly once, using the exact input string."""


def classify_queues(queue_names: list[str]) -> dict[str, dict]:
    """Return {queue_name: {"tier": "A".."D", "classification": "..."}}."""
    client = get_client()
    payload = "\n".join(f"- {q}" for q in queue_names)
    resp = client.messages.create(
        model=MODEL,
        max_tokens=4096,
        system=TIERING_SYSTEM,
        messages=[{"role": "user", "content": f"Classify these {len(queue_names)} queues:\n{payload}"}],
    )
    data = _extract_json(resp.content[0].text)
    result: dict[str, dict] = {}
    for item in data:
        q = item.get("queue", "").strip()
        if q:
            result[q] = {
                "tier": item.get("tier", "C").strip().upper()[:1] or "C",
                "classification": item.get("classification", "").strip(),
            }
    # Fallback for any queue Claude omitted
    for q in queue_names:
        if q not in result:
            result[q] = {"tier": "C", "classification": "Branch main line or front desk"}
    return result


NARRATIVE_SYSTEM = """You are a RingCentral sales engineer writing concise, factual slide copy for an AI Receptionist business case.

Rules:
- Use only the numbers provided; never invent data.
- Be precise and professional, no marketing fluff or superlatives.
- Be honest: if the data is weak for a queue, say so.
- Respond with JSON only, matching the requested schema.
- Keep bullets under 12 words and narrative under 3 sentences."""


def generate_narrative(context: dict, schema: dict, prior_instructions: list[dict] | None = None) -> dict:
    client = get_client()
    user = (
        f"Data context (JSON):\n{json.dumps(context, indent=2)}\n\n"
        f"Fill this schema (JSON only):\n{json.dumps(schema, indent=2)}"
    )
    messages = list(prior_instructions or []) + [{"role": "user", "content": user}]
    resp = client.messages.create(
        model=MODEL,
        max_tokens=1024,
        system=NARRATIVE_SYSTEM,
        messages=messages,
    )
    return _extract_json(resp.content[0].text)
