"""
Server-side Claude API integration.
Generates slide narrative text based on pipeline results and AE inputs.
"""

from __future__ import annotations
import os
import anthropic

_client: anthropic.Anthropic | None = None


def get_client() -> anthropic.Anthropic:
    global _client
    if _client is None:
        _client = anthropic.Anthropic(api_key=os.environ["ANTHROPIC_API_KEY"])
    return _client


SYSTEM_PROMPT = """You are an expert sales engineer helping build a business case presentation for RingCentral's AI receptionist product.

You receive structured call data (real customer data) and produce concise, professional slide narrative text.

Rules:
- Be specific — use the actual numbers provided, never invent data.
- Be honest — if the data doesn't support a strong AI receptionist case for a particular queue, say so.
- Avoid superlatives and marketing fluff. Write like a trusted advisor, not a pitch deck.
- All assumptions (deal value, close rate) must be explicitly labeled on slides.
- Use the low/mid/high recovery range (50%/70%/90%) for revenue opportunity — never a single figure.
- Keep slide text concise: bullets max 12 words each, narrative paragraphs max 3 sentences.
- Output valid JSON matching the schema requested in each prompt."""


def generate_slide_content(
    results_summary: dict,
    ae_name: str,
    avg_deal_value: float,
    close_rate: float,
    prior_instructions: list[dict],
    slide_schema: dict,
) -> dict:
    """
    Call Claude to generate slide text content.

    slide_schema: describes what fields to populate, e.g.
      {"headline": "str", "bullets": ["str"], "narrative": "str"}

    Returns a dict matching slide_schema.
    """
    client = get_client()

    user_content = f"""
Call data summary:
{_format_summary(results_summary)}

AE: {ae_name}
Average deal value: ${avg_deal_value:,.0f}
Close rate on answered inbound calls: {close_rate:.0%}

Slide schema to fill (respond with JSON only):
{slide_schema}
"""

    messages = list(prior_instructions) + [{"role": "user", "content": user_content}]

    response = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=1024,
        system=SYSTEM_PROMPT,
        messages=messages,
    )

    import json
    text = response.content[0].text.strip()
    # Strip markdown code fences if present
    if text.startswith("```"):
        text = text.split("```")[1]
        if text.startswith("json"):
            text = text[4:]
    return json.loads(text)


def _format_summary(s: dict) -> str:
    lines = []
    for k, v in s.items():
        if isinstance(v, dict):
            lines.append(f"{k}:")
            for kk, vv in v.items():
                lines.append(f"  {kk}: {vv}")
        else:
            lines.append(f"{k}: {v}")
    return "\n".join(lines)
