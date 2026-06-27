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
    text = text.strip()
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        # Tolerate prose before/after the JSON: grab the outermost array/object.
        starts = [i for i in (text.find("["), text.find("{")) if i != -1]
        ends = [i for i in (text.rfind("]"), text.rfind("}")) if i != -1]
        if starts and ends:
            return json.loads(text[min(starts): max(ends) + 1])
        raise


TIERING_SYSTEM = """You classify a business's inbound call-queue names into revenue-relevance tiers for a missed-call analysis. This must work for ANY industry — distribution, healthcare, professional services, hospitality, automotive, SaaS, home services, etc. Reason from what the queue name means for THIS business, not from any single industry's vocabulary.

Tiers (a gradient of how directly a call to this queue touches revenue):
- A = Direct sales / revenue line. The name explicitly signals selling, new business, or booking revenue (e.g. "Sales", "New Business", "Reservations", "New Patients", "Quotes").
- B = Customer-facing front line. Where existing or prospective customers reach the business to buy, order, schedule, or get service that drives revenue or retention (product/order desks, service/booking lines, account lines). Customer-facing but not explicitly "Sales".
- C = General main line / reception. A "Main", "Front Desk", "Operator", or main number a customer would call that can still carry sales or service intent.
- D = Back-office / internal. NOT customer-revenue facing: IT, internal support/help desk, accounting, AP/AR, credit, HR, logistics/shipping, facilities, systems. These are EXCLUDED from the analysis.

Rules:
- Decide from the queue name plus any business context provided. When unsure whether a queue is customer-facing, prefer C over D so it is not wrongly excluded.
- A name that clearly signals selling/booking new revenue -> A.
- Clearly internal/operational functions (IT, accounting, credit, HR, shipping/receiving, systems, internal help desk) -> D.
- Generic "Main"/"Front Desk"/"Operator"/main number -> C.
- Respond with JSON only: an array of {"queue": <exact name>, "tier": "A|B|C|D", "classification": <short neutral label>}.
- classification labels: A -> "Sales / revenue line", B -> "Customer-facing front line", C -> "Main line / reception", D -> "Back-office / internal (excluded)".
- Include every queue exactly once, using the exact input string."""


def _business_hint(business_context) -> str:
    """A short, neutral description of the business to steer queue tiering."""
    if not business_context or not isinstance(business_context, dict):
        return ""
    if not business_context.get("available", True):
        return ""
    parts = []
    industry = (business_context.get("industry") or "").strip()
    summary = (business_context.get("summary") or "").strip()
    lobs = business_context.get("lines_of_business") or []
    if industry:
        parts.append(f"Industry: {industry}")
    if summary:
        parts.append(f"What they do: {summary}")
    if lobs:
        parts.append("Lines of business: " + ", ".join(str(x) for x in lobs[:8]))
    if not parts:
        return ""
    return (
        "Business context for THIS customer (use it to judge whether each queue is "
        "customer/revenue-facing for this specific business):\n" + "\n".join(parts) + "\n\n"
    )


def classify_queues(queue_names: list[str], business_context=None) -> dict[str, dict]:
    """Return {queue_name: {"tier": "A".."D", "classification": "..."}}.

    When ``business_context`` (the crawled website profile) is supplied, it is
    fed to the model so tiering is reasoned from this customer's actual business
    rather than any hard-coded industry vocabulary.
    """
    client = get_client()
    payload = "\n".join(f"- {q}" for q in queue_names)
    user = (
        _business_hint(business_context)
        + f"Classify these {len(queue_names)} queues:\n{payload}"
    )
    result: dict[str, dict] = {}
    try:
        resp = client.messages.create(
            model=MODEL,
            max_tokens=4096,
            system=TIERING_SYSTEM,
            messages=[{"role": "user", "content": user}],
        )
        data = _extract_json(resp.content[0].text)
        for item in data:
            q = (item.get("queue") or "").strip()
            if q:
                result[q] = {
                    "tier": (item.get("tier", "C") or "C").strip().upper()[:1] or "C",
                    "classification": (item.get("classification") or "").strip(),
                }
    except Exception:
        # Never fail the whole run on a model/parse hiccup — fall back to a
        # neutral default tier for every queue (keeps them in the analysis).
        result = {}
    # Fallback for any queue Claude omitted or that failed to parse
    for q in queue_names:
        if q not in result:
            result[q] = {"tier": "C", "classification": "Main line / reception"}
    return result


NARRATIVE_SYSTEM = """You are a RingCentral sales engineer writing concise, factual slide copy for an AI Receptionist business case.

Rules:
- Use only the numbers provided; never invent data.
- Be precise and professional, no marketing fluff or superlatives.
- Be honest: if the data is weak for a queue, say so.
- Respond with JSON only, matching the requested schema.
- Keep bullets under 12 words and narrative under 3 sentences."""


BUSINESS_PROFILE_SYSTEM = """You are a RingCentral sales engineer profiling a prospect's business from their website, to tailor an inbound-call analysis.

You are given (1) markdown scraped from the company's website and (2) the names of their actual phone call queues. From these, infer how this business runs its phones and WHY customers call.

Return JSON only with this exact shape:
{
  "summary": "<=2 sentences on what the company does",
  "industry": "<short industry label>",
  "lines_of_business": ["<3-6 short phrases>"],
  "predicted_call_reasons": [
    {"reason": "<short label, e.g. 'Order status & tracking'>",
     "why": "<<=14 words on why this business gets these calls>",
     "tier": "A|B|C|D",
     "revenue_relevant": true}
  ],
  "suggested_avg_order_value": <integer USD, your best defensible estimate for a typical order/transaction, or null if unknowable>,
  "aov_basis": "<<=12 words explaining the order-value estimate, or empty>"
}

Rules:
- Ground every claim in the website text or the queue names; do not invent specific facts (no fake revenue, locations, or customer names).
- Predict 4-7 call reasons most likely for THIS business, whatever its industry. Map each to a tier: A=direct sales / new revenue, B=customer-facing front line (orders, scheduling, service that drives revenue/retention), C=general main line / reception, D=back-office / internal support.
- revenue_relevant=true when answering that call could win or retain revenue.
- suggested_avg_order_value: reason from the business type and typical transaction size (e.g. a distributor or B2B firm -> larger orders; a salon or cafe -> small; a SaaS or services firm -> a contract/booking value). Be conservative and defensible.
- JSON only, no prose, no code fences."""


def profile_business(markdown: str, customer: str, queue_names: list[str]) -> dict:
    client = get_client()
    queues = "\n".join(f"- {q}" for q in queue_names[:60])
    user = (
        f"Company name (as entered): {customer or 'unknown'}\n\n"
        f"Their call queue names:\n{queues}\n\n"
        f"Website content (markdown):\n{markdown}"
    )
    resp = client.messages.create(
        model=MODEL,
        max_tokens=1500,
        system=BUSINESS_PROFILE_SYSTEM,
        messages=[{"role": "user", "content": user}],
    )
    try:
        data = _extract_json(resp.content[0].text)
    except Exception:
        return {"summary": "", "industry": "", "lines_of_business": [],
                "predicted_call_reasons": [], "suggested_avg_order_value": None,
                "aov_basis": ""}
    return data if isinstance(data, dict) else {}


OVERRIDE_SYSTEM = """You extract structured analysis overrides from a user's free-text instruction in a call-analysis tool.

The user may state new business facts or modeling assumptions. Extract ONLY values they explicitly provide. Return JSON:
{
  "avg_order_value": <integer USD or null>,
  "capture_rate": <decimal 0-1 or null>,      // e.g. "8% capture" -> 0.08
  "air_rate_per_min": <decimal USD or null>,
  "notes": "<other instructions to honor verbatim, or empty>"
}

Rules:
- If the user only gives styling/wording instructions (not data), set all numeric fields null and put the instruction in notes.
- Never guess numbers the user didn't give. JSON only, no prose."""


def extract_overrides(instruction: str) -> dict:
    client = get_client()
    try:
        resp = client.messages.create(
            model=MODEL,
            max_tokens=400,
            system=OVERRIDE_SYSTEM,
            messages=[{"role": "user", "content": instruction}],
        )
        data = _extract_json(resp.content[0].text)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


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
