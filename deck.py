"""
PPTX generation using python-pptx.
Builds a business-case deck from pipeline results + Claude-generated narratives.
"""

from __future__ import annotations
import tempfile
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

from pipeline import PipelineResult
from claude_client import generate_slide_content

# RingCentral brand colors
RC_ORANGE = RGBColor(0xFF, 0x7A, 0x00)
RC_DARK = RGBColor(0x1A, 0x1A, 0x2E)
RC_LIGHT_GRAY = RGBColor(0xF5, 0xF5, 0xF5)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

SLIDE_W = Inches(13.33)
SLIDE_H = Inches(7.5)


def _add_text_box(slide, text, left, top, width, height,
                  font_size=18, bold=False, color=RC_DARK, align=PP_ALIGN.LEFT):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = color
    return txBox


def _title_slide(prs: Presentation, company_name: str, ae_name: str, date_str: str):
    layout = prs.slide_layouts[6]  # blank
    slide = prs.slides.add_slide(layout)

    # Background
    bg = slide.shapes.add_shape(
        1, 0, 0, SLIDE_W, SLIDE_H  # MSO_SHAPE_TYPE.RECTANGLE = 1
    )
    bg.fill.solid()
    bg.fill.fore_color.rgb = RC_DARK
    bg.line.fill.background()

    _add_text_box(slide, "AI Receptionist", Inches(1), Inches(1.5), Inches(11), Inches(1.2),
                  font_size=44, bold=True, color=RC_ORANGE, align=PP_ALIGN.LEFT)
    _add_text_box(slide, "Business Case Analysis", Inches(1), Inches(2.8), Inches(11), Inches(1),
                  font_size=32, bold=False, color=WHITE, align=PP_ALIGN.LEFT)
    _add_text_box(slide, company_name, Inches(1), Inches(4.0), Inches(11), Inches(0.8),
                  font_size=24, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
    _add_text_box(slide, f"Prepared by {ae_name}  ·  {date_str}",
                  Inches(1), Inches(4.9), Inches(11), Inches(0.6),
                  font_size=16, color=RGBColor(0xAA, 0xAA, 0xAA), align=PP_ALIGN.LEFT)


def _call_volume_slide(prs: Presentation, results: PipelineResult, content: dict):
    layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(layout)

    _add_text_box(slide, content.get("headline", "Your Call Volume at a Glance"),
                  Inches(0.5), Inches(0.3), Inches(12), Inches(0.8),
                  font_size=28, bold=True, color=RC_DARK)

    # Stat boxes
    stats = [
        ("Total Inbound Sessions", f"{results.total_inbound_sessions:,}", RC_DARK),
        ("Answered", f"{results.answered:,}\n({results.answer_rate:.0%})", RGBColor(0x22, 0x8B, 0x22)),
        ("Unanswered", f"{results.unanswered:,}\n({1-results.answer_rate:.0%})", RC_ORANGE),
    ]
    for i, (label, value, color) in enumerate(stats):
        x = Inches(0.5 + i * 4.3)
        box = slide.shapes.add_shape(1, x, Inches(1.3), Inches(3.8), Inches(2.2))
        box.fill.solid()
        box.fill.fore_color.rgb = RC_LIGHT_GRAY
        box.line.fill.background()
        _add_text_box(slide, label, x + Inches(0.15), Inches(1.4), Inches(3.5), Inches(0.5),
                      font_size=13, color=RGBColor(0x55, 0x55, 0x55))
        _add_text_box(slide, value, x + Inches(0.15), Inches(1.9), Inches(3.5), Inches(0.9),
                      font_size=28, bold=True, color=color)

    # Unanswered breakdown
    _add_text_box(slide, "Unanswered Breakdown",
                  Inches(0.5), Inches(3.7), Inches(12), Inches(0.5),
                  font_size=16, bold=True, color=RC_DARK)
    breakdown = [
        f"• Abandoned (caller hung up):  {results.abandoned:,}",
        f"• VM/Abandoned:  {results.vm_abandoned:,}",
        f"• VM/Missed:  {results.vm_missed:,}",
        f"• Missed/Other:  {results.missed:,}",
    ]
    _add_text_box(slide, "\n".join(breakdown),
                  Inches(0.5), Inches(4.2), Inches(6), Inches(2.5),
                  font_size=15, color=RC_DARK)

    narrative = content.get("narrative", "")
    if narrative:
        _add_text_box(slide, narrative,
                      Inches(6.5), Inches(4.0), Inches(6.3), Inches(3.0),
                      font_size=13, color=RGBColor(0x44, 0x44, 0x44))


def _queue_slide(prs: Presentation, results: PipelineResult, content: dict):
    layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(layout)

    _add_text_box(slide, content.get("headline", "Queue-by-Queue Breakdown"),
                  Inches(0.5), Inches(0.3), Inches(12), Inches(0.8),
                  font_size=28, bold=True, color=RC_DARK)

    sorted_queues = sorted(
        results.queue_stats.values(),
        key=lambda q: q.unanswered,
        reverse=True,
    )[:8]  # top 8 queues by unanswered volume

    # Simple table-style layout
    headers = ["Queue", "Total", "Answered", "Unanswered", "Answer Rate"]
    col_widths = [Inches(3.5), Inches(1.4), Inches(1.8), Inches(2.0), Inches(1.8)]
    col_starts = [Inches(0.3)]
    for w in col_widths[:-1]:
        col_starts.append(col_starts[-1] + w)

    row_h = Inches(0.45)
    header_y = Inches(1.2)

    for j, (header, x) in enumerate(zip(headers, col_starts)):
        _add_text_box(slide, header, x, header_y, col_widths[j], row_h,
                      font_size=12, bold=True, color=WHITE)
        # header bg
        hbg = slide.shapes.add_shape(1, x, header_y, col_widths[j], row_h)
        hbg.fill.solid()
        hbg.fill.fore_color.rgb = RC_DARK
        hbg.line.fill.background()
        # re-add text on top
        _add_text_box(slide, header, x, header_y + Inches(0.05), col_widths[j], row_h,
                      font_size=12, bold=True, color=WHITE)

    for i, qs in enumerate(sorted_queues):
        y = header_y + row_h * (i + 1)
        bg_color = RC_LIGHT_GRAY if i % 2 == 0 else WHITE
        row_vals = [
            qs.name[:40],
            f"{qs.total:,}",
            f"{qs.answered:,}",
            f"{qs.unanswered:,}",
            f"{qs.answer_rate:.0%}",
        ]
        for j, (val, x) in enumerate(zip(row_vals, col_starts)):
            rbg = slide.shapes.add_shape(1, x, y, col_widths[j], row_h)
            rbg.fill.solid()
            rbg.fill.fore_color.rgb = bg_color
            rbg.line.fill.background()
            text_color = RC_ORANGE if j == 3 and qs.unanswered > 0 else RC_DARK
            _add_text_box(slide, val, x + Inches(0.05), y + Inches(0.05), col_widths[j], row_h,
                          font_size=12, color=text_color)

    note = content.get("narrative", "")
    if note:
        _add_text_box(slide, note, Inches(0.3), Inches(6.5), Inches(12.5), Inches(0.8),
                      font_size=12, color=RGBColor(0x66, 0x66, 0x66))


def _roi_slide(prs: Presentation, results: PipelineResult,
               avg_deal_value: float, close_rate: float,
               ae_name: str, content: dict):
    layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(layout)

    _add_text_box(slide, content.get("headline", "Revenue Opportunity"),
                  Inches(0.5), Inches(0.3), Inches(12), Inches(0.8),
                  font_size=28, bold=True, color=RC_DARK)

    unanswered = results.unanswered
    recovery_rates = [("Conservative", 0.50), ("Moderate", 0.70), ("Aggressive", 0.90)]

    _add_text_box(slide,
                  f"Coverage Gap: {unanswered:,} unanswered inbound calls\n"
                  f"Assumptions entered by {ae_name}: "
                  f"${avg_deal_value:,.0f} avg deal value · {close_rate:.0%} close rate on answered calls",
                  Inches(0.5), Inches(1.1), Inches(12.3), Inches(0.9),
                  font_size=14, color=RGBColor(0x44, 0x44, 0x44))

    for i, (label, rate) in enumerate(recovery_rates):
        recovered = unanswered * rate
        revenue = recovered * close_rate * avg_deal_value
        x = Inches(0.5 + i * 4.2)

        box = slide.shapes.add_shape(1, x, Inches(2.2), Inches(3.9), Inches(2.6))
        box.fill.solid()
        box.fill.fore_color.rgb = RC_LIGHT_GRAY if i != 1 else RC_DARK
        box.line.fill.background()

        val_color = RC_ORANGE if i == 1 else RC_DARK
        label_color = WHITE if i == 1 else RGBColor(0x55, 0x55, 0x55)
        rev_color = RC_ORANGE if i == 1 else RC_DARK

        _add_text_box(slide, label, x + Inches(0.2), Inches(2.3), Inches(3.5), Inches(0.5),
                      font_size=14, bold=True, color=label_color)
        _add_text_box(slide, f"{rate:.0%} AI recovery rate",
                      x + Inches(0.2), Inches(2.8), Inches(3.5), Inches(0.4),
                      font_size=12, color=label_color)
        _add_text_box(slide, f"~{recovered:,.0f} calls recovered",
                      x + Inches(0.2), Inches(3.2), Inches(3.5), Inches(0.4),
                      font_size=12, color=label_color)
        _add_text_box(slide, f"${revenue:,.0f}",
                      x + Inches(0.2), Inches(3.7), Inches(3.5), Inches(0.8),
                      font_size=30, bold=True, color=rev_color)
        _add_text_box(slide, "potential revenue",
                      x + Inches(0.2), Inches(4.5), Inches(3.5), Inches(0.4),
                      font_size=12, color=label_color)

    disclaimer = (
        "All figures are estimates based on AE-entered assumptions. "
        "AI recovery rate reflects the share of currently-unanswered calls that an AI receptionist could handle. "
        "Not a RingCentral guarantee."
    )
    _add_text_box(slide, disclaimer, Inches(0.5), Inches(6.3), Inches(12.3), Inches(0.9),
                  font_size=11, color=RGBColor(0x88, 0x88, 0x88))

    if content.get("narrative"):
        _add_text_box(slide, content["narrative"],
                      Inches(0.5), Inches(5.2), Inches(12.3), Inches(1.0),
                      font_size=13, color=RC_DARK)


def _reconciliation_slide(prs: Presentation, results: PipelineResult):
    layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(layout)

    _add_text_box(slide, "Data Methodology & Validation",
                  Inches(0.5), Inches(0.3), Inches(12), Inches(0.8),
                  font_size=24, bold=True, color=RC_DARK)

    lines = [
        "• Unit of analysis: session (not call leg) — eliminates multi-ring inflation",
        "• Scope: inbound calls only; outbound and internal legs excluded",
        "• Outcome priority: Answered > VM/Abandoned > VM/Missed > Abandoned > Missed/Other",
        "• Park Off events are ignored (non-outcome marker)",
        "",
        f"Reconciliation: {results.reconciliation_note}",
        "",
        f"  Total inbound sessions: {results.total_inbound_sessions:,}",
        f"  Answered:               {results.answered:,}",
        f"  VM/Abandoned:           {results.vm_abandoned:,}",
        f"  VM/Missed:              {results.vm_missed:,}",
        f"  Abandoned:              {results.abandoned:,}",
        f"  Missed/Other:           {results.missed:,}",
    ]

    status_color = RGBColor(0x22, 0x8B, 0x22) if results.reconciliation_ok else RGBColor(0xCC, 0x00, 0x00)
    _add_text_box(slide, "\n".join(lines),
                  Inches(0.5), Inches(1.2), Inches(12), Inches(5.5),
                  font_size=14, color=RC_DARK)

    # Rewrite reconciliation line with status color
    # (simplification — just note it inline)


def build_deck(
    results: PipelineResult,
    run_id: str,
    ae_name: str,
    avg_deal_value: float,
    close_rate: float,
    prior_instructions: list[dict],
) -> Path:
    import datetime

    out_dir = Path(tempfile.gettempdir()) / "rc_analyzer_decks"
    out_dir.mkdir(exist_ok=True)
    out_path = out_dir / f"{run_id}.pptx"

    # Build summary for Claude
    top_queues = sorted(results.queue_stats.values(), key=lambda q: q.unanswered, reverse=True)[:5]
    results_summary = {
        "total_inbound_sessions": results.total_inbound_sessions,
        "answered": results.answered,
        "answer_rate": f"{results.answer_rate:.1%}",
        "unanswered": results.unanswered,
        "vm_abandoned": results.vm_abandoned,
        "vm_missed": results.vm_missed,
        "abandoned": results.abandoned,
        "missed": results.missed,
        "top_queues_by_unanswered": {
            q.name: {"total": q.total, "unanswered": q.unanswered, "answer_rate": f"{q.answer_rate:.0%}"}
            for q in top_queues
        },
    }

    # Generate slide narratives via Claude
    volume_content = generate_slide_content(
        results_summary=results_summary,
        ae_name=ae_name,
        avg_deal_value=avg_deal_value,
        close_rate=close_rate,
        prior_instructions=prior_instructions,
        slide_schema={"headline": "str (max 10 words)", "narrative": "str (max 3 sentences, insight-focused)"},
    )

    queue_content = generate_slide_content(
        results_summary=results_summary,
        ae_name=ae_name,
        avg_deal_value=avg_deal_value,
        close_rate=close_rate,
        prior_instructions=prior_instructions,
        slide_schema={
            "headline": "str (max 10 words, focus on queue patterns)",
            "narrative": "str (1-2 sentences, note which queues need attention most)",
        },
    )

    roi_content = generate_slide_content(
        results_summary=results_summary,
        ae_name=ae_name,
        avg_deal_value=avg_deal_value,
        close_rate=close_rate,
        prior_instructions=prior_instructions,
        slide_schema={
            "headline": "str (max 10 words, quantified if possible)",
            "narrative": "str (2-3 sentences connecting unanswered calls to revenue, no invented numbers)",
        },
    )

    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    date_str = datetime.date.today().strftime("%B %d, %Y")
    company_name = "Prospect"  # TODO: could be derived from filename or entered on configure screen

    _title_slide(prs, company_name, ae_name, date_str)
    _call_volume_slide(prs, results, volume_content)
    _queue_slide(prs, results, queue_content)
    _roi_slide(prs, results, avg_deal_value, close_rate, ae_name, roi_content)
    _reconciliation_slide(prs, results)

    prs.save(out_path)
    return out_path
