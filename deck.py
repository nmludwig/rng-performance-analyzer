"""
PPTX generation — RingCentral AI Receptionist business-case deck.
Matches the FBM 4-slide design. Slides 1, 2, and 4 are built (slide 3,
the AIR capability mapping, is deferred).
"""

from __future__ import annotations
import tempfile
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION

from pipeline import PipelineResult
from claude_client import generate_narrative

# Palette
RC_BLUE = RGBColor(0x06, 0x2E, 0x5C)
RC_NAVY = RGBColor(0x1A, 0x2B, 0x4A)
RC_ORANGE = RGBColor(0xFF, 0x7A, 0x00)
RC_RED = RGBColor(0xC0, 0x2A, 0x2A)
RC_TEAL = RGBColor(0x0A, 0x8A, 0x8A)
RC_GOLD = RGBColor(0xC8, 0x8A, 0x00)
RC_PURPLE = RGBColor(0x5B, 0x3E, 0x96)
DARK = RGBColor(0x20, 0x20, 0x28)
GRAY = RGBColor(0x66, 0x66, 0x66)
LIGHT = RGBColor(0xF4, 0xF5, 0xF8)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
CARD_BORDER = RGBColor(0xE2, 0xE4, 0xEA)
ROW_ALT = RGBColor(0xF6, 0xF1, 0xEE)

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)


def _text(slide, text, left, top, width, height, *, size=14, bold=False,
          color=DARK, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
          italic=False, font="Calibri", wrap=True, line_spacing=None):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    tf.margin_left = Pt(2); tf.margin_right = Pt(2)
    tf.margin_top = Pt(1); tf.margin_bottom = Pt(1)
    p = tf.paragraphs[0]
    p.alignment = align
    if line_spacing:
        p.line_spacing = line_spacing
    r = p.add_run(); r.text = text
    r.font.size = Pt(size); r.font.bold = bold; r.font.italic = italic
    r.font.color.rgb = color; r.font.name = font
    return tb


def _rich(slide, segments, left, top, width, height, *, anchor=MSO_ANCHOR.TOP,
          align=PP_ALIGN.LEFT, line_spacing=1.0, space_after=4):
    """segments: list of paragraphs, each a list of (text, {opts}) runs."""
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = Pt(2); tf.margin_right = Pt(2)
    tf.margin_top = Pt(1); tf.margin_bottom = Pt(1)
    for i, para in enumerate(segments):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = line_spacing
        p.space_after = Pt(space_after)
        for txt, opts in para:
            r = p.add_run(); r.text = txt
            r.font.size = Pt(opts.get("size", 10))
            r.font.bold = opts.get("bold", False)
            r.font.italic = opts.get("italic", False)
            r.font.color.rgb = opts.get("color", DARK)
            r.font.name = opts.get("font", "Calibri")
    return tb


def _rect(slide, left, top, width, height, fill, *, line=None, line_w=None, radius=False):
    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE,
        left, top, width, height)
    if fill is None:
        shape.fill.background()
    else:
        shape.fill.solid(); shape.fill.fore_color.rgb = fill
    if line is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = line
        shape.line.width = line_w or Pt(0.75)
    shape.shadow.inherit = False
    return shape


def _circle(slide, left, top, dia, fill, text=None, *, text_color=WHITE, size=12):
    c = slide.shapes.add_shape(MSO_SHAPE.OVAL, left, top, dia, dia)
    c.fill.solid(); c.fill.fore_color.rgb = fill
    c.line.fill.background()
    c.shadow.inherit = False
    if text is not None:
        tf = c.text_frame
        tf.word_wrap = False
        p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
        r = p.add_run(); r.text = text
        r.font.size = Pt(size); r.font.bold = True
        r.font.color.rgb = text_color; r.font.name = "Calibri"
    return c


def _logo(slide):
    _text(slide, "Ring", Inches(0.45), Inches(0.18), Inches(0.9), Inches(0.4),
          size=22, bold=True, color=RC_BLUE, font="Arial")
    _text(slide, "Central", Inches(1.25), Inches(0.18), Inches(1.4), Inches(0.4),
          size=22, bold=True, color=RC_ORANGE, font="Arial")


def _footer(slide, page):
    _text(slide, str(page), Inches(0.45), Inches(7.12), Inches(0.5), Inches(0.3),
          size=9, color=GRAY)
    _text(slide, "Confidential", Inches(5.9), Inches(7.12), Inches(1.5), Inches(0.3),
          size=9, color=GRAY, align=PP_ALIGN.CENTER)
    _text(slide, "©2026 RingCentral", Inches(11.0), Inches(7.12), Inches(1.9), Inches(0.3),
          size=9, color=GRAY, align=PP_ALIGN.RIGHT)


def _title_block(slide, title, subtitle):
    _text(slide, title, Inches(0.5), Inches(0.62), Inches(12.3), Inches(0.7),
          size=30, bold=True, color=DARK, font="Arial")
    _text(slide, subtitle, Inches(0.52), Inches(1.32), Inches(12.3), Inches(0.4),
          size=12, italic=True, color=GRAY)


# ---------------------------------------------------------------------------
# Slide 1 — methodology credibility
# ---------------------------------------------------------------------------

def _slide1(prs, r: PipelineResult, ctx, narr):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _logo(s)
    _title_block(s, narr.get("title", "How we know these are genuine missed calls — not spam, not noise"),
                 narr.get("subtitle", ""))

    cards = [
        ("1", RC_BLUE, "Session ID deduplication", narr.get("card1", [])),
        ("2", RC_TEAL, "SPAM / solicitation filter", narr.get("card2", [])),
        ("3", RC_RED, "Business-hours concentration", narr.get("card3", [])),
        ("4", RC_GOLD, "Repeat caller evidence", narr.get("card4", [])),
    ]
    card_w, card_h = Inches(6.05), Inches(1.95)
    xs = [Inches(0.5), Inches(6.78)]
    ys = [Inches(1.85), Inches(3.92)]
    for i, (num, color, heading, paras) in enumerate(cards):
        x = xs[i % 2]; y = ys[i // 2]
        _rect(s, x, y, card_w, card_h, WHITE, line=CARD_BORDER, radius=True)
        _circle(s, x + Inches(0.18), y + Inches(0.16), Inches(0.42), color, num, size=14)
        _text(s, heading, x + Inches(0.75), y + Inches(0.16), card_w - Inches(0.9), Inches(0.4),
              size=15, bold=True, color=color, font="Arial")
        segs = []
        for label, body in paras:
            segs.append([(label + " ", {"bold": True, "size": 9.5, "color": DARK}),
                         (body, {"size": 9.5, "color": GRAY})])
        _rich(s, segs, x + Inches(0.22), y + Inches(0.62), card_w - Inches(0.45), card_h - Inches(0.7),
              line_spacing=1.0, space_after=3)

    # bottom stat strip
    strip_y = Inches(6.05)
    _rect(s, Inches(0.5), strip_y, Inches(12.33), Inches(0.92), RC_NAVY, radius=True)
    stats = [
        (f"{r.universe_sessions:,}", "legitimate inbound\nsessions analyzed"),
        (f"{r.spam_sessions_removed:,}", f"spam sessions\nremoved ({_pct(r.spam_sessions_removed, r.inbound_sessions)})"),
        (f"{r.phantom_legs_removed:,}", "phantom legs removed\nby Session ID method"),
        (f"{r.total_missed:,}", f"genuine missed\ncalls ({_pct(r.total_missed, r.universe_sessions)})"),
        (f"{round(r.business_hours_miss_pct*100)}%", "during staffed\nbusiness hours"),
        (f"{r.repeat_callers:,}", "repeat callers\nnever got through"),
    ]
    seg_w = Inches(12.33) / 6
    for i, (val, lbl) in enumerate(stats):
        sx = Inches(0.5) + seg_w * i
        _text(s, val, sx, strip_y + Inches(0.10), seg_w, Inches(0.42),
              size=22, bold=True, color=WHITE, align=PP_ALIGN.CENTER, font="Arial")
        _text(s, lbl, sx, strip_y + Inches(0.52), seg_w, Inches(0.36),
              size=8, color=RGBColor(0xC8,0xD0,0xE0), align=PP_ALIGN.CENTER)
        if i:
            _rect(s, sx, strip_y + Inches(0.18), Pt(1), Inches(0.56), RGBColor(0x33,0x44,0x66))
    _footer(s, 1)


# ---------------------------------------------------------------------------
# Slide 2 — missed-call summary
# ---------------------------------------------------------------------------

def _slide2(prs, r: PipelineResult, ctx, narr, sales_queue_calls):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _logo(s)
    _title_block(s, narr.get("title", f"{r.total_missed:,} genuine missed calls across sales-taking queues"),
                 narr.get("subtitle", ""))

    # Left big-stat card
    lx, ly, lw, lh = Inches(0.5), Inches(1.95), Inches(3.25), Inches(4.7)
    _rect(s, lx, ly, lw, lh, WHITE, line=CARD_BORDER, radius=True)
    _text(s, "GENUINE MISSED\nCALLS — A+B+C QUEUES", lx + Inches(0.25), ly + Inches(0.55),
          lw - Inches(0.5), Inches(0.7), size=12, bold=True, color=GRAY, font="Arial")
    _text(s, f"{r.total_missed:,}", lx + Inches(0.18), ly + Inches(1.15), lw - Inches(0.3), Inches(1.2),
          size=72, bold=True, color=RC_RED, font="Arial")
    _rich(s, [[(f"{r.miss_rate*100:.1f}%", {"bold": True, "size": 13, "color": DARK}),
               (f" of {r.universe_sessions:,} queue sessions went unanswered",
                {"size": 13, "color": DARK})]],
          lx + Inches(0.25), ly + Inches(2.35), lw - Inches(0.45), Inches(0.7))
    _rect(s, lx + Inches(0.25), ly + Inches(3.05), lw - Inches(0.5), Pt(1), CARD_BORDER)
    _text(s, f"≈ {round(r.misses_per_day)} missed calls every day",
          lx + Inches(0.25), ly + Inches(3.15), lw - Inches(0.5), Inches(0.4),
          size=13, bold=True, italic=True, color=RC_ORANGE, align=PP_ALIGN.CENTER)
    _rich(s, [[(f"{sales_queue_calls:,}", {"bold": True, "size": 13, "color": RC_RED}),
               (" of these are confirmed\nsales-queue calls", {"size": 12, "color": DARK})]],
          lx + Inches(0.25), ly + Inches(3.7), lw - Inches(0.5), Inches(0.8),
          align=PP_ALIGN.CENTER)

    # Middle: chart + breakdown
    mx = Inches(4.0)
    _text(s, f"{round(r.business_hours_miss_pct*100)}% of misses during staffed hours (M–F, 7a–6p)",
          mx, Inches(1.95), Inches(4.6), Inches(0.35), size=12, bold=True, color=DARK, font="Arial")
    _hourly_chart(s, r, mx, Inches(2.35), Inches(4.55), Inches(1.85))

    _text(s, "How they were missed:", mx, Inches(4.35), Inches(4.6), Inches(0.3),
          size=12, bold=True, color=DARK, font="Arial")
    bd = [
        (r.abandoned, "Abandoned while ringing", RC_RED, RGBColor(0xFC,0xEC,0xEC)),
        (r.voicemail_total, "Went to voicemail (left msg)", RC_GOLD, RGBColor(0xFB,0xF3,0xE0)),
        (r.missed, "Rang out — no answer", RC_PURPLE, RGBColor(0xF0,0xEC,0xF7)),
    ]
    by = Inches(4.7)
    for val, lbl, col, bg in bd:
        _rect(s, mx, by, Inches(4.55), Inches(0.55), bg, radius=True)
        pct = _pct(val, r.total_missed)
        _rich(s, [[(f"{val:,}  ", {"bold": True, "size": 15, "color": col}),
                   (f"{pct}  ", {"bold": True, "size": 11, "color": col}),
                   (lbl, {"size": 11, "color": DARK})]],
              mx + Inches(0.2), by, Inches(4.3), Inches(0.55), anchor=MSO_ANCHOR.MIDDLE)
        by += Inches(0.68)
    _text(s, f"{r.repeat_callers} callers tried 2+ times and never got through — proof of buying intent, not spam",
          mx, by + Inches(0.02), Inches(4.6), Inches(0.5), size=10, italic=True, color=RC_RED)

    # Right: worst-hit queues table
    rx = Inches(8.85)
    _text(s, "Worst-hit queues", rx, Inches(1.95), Inches(4.0), Inches(0.35),
          size=14, bold=True, color=DARK, font="Arial")
    worst = sorted(r.queue_stats.values(), key=lambda q: q.total_missed, reverse=True)[:10]
    _worst_table(s, worst, rx, Inches(2.35), Inches(3.95))
    _footer(s, 2)


def _hourly_chart(s, r, x, y, w, h):
    cd = CategoryChartData()
    hours = sorted(r.hourly_missed.keys())
    labels = [_hour_label(hh) for hh in hours]
    cd.categories = labels
    cd.add_series("Missed", [r.hourly_missed[hh] for hh in hours])
    gframe = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, x, y, w, h, cd)
    chart = gframe.chart
    chart.has_legend = False
    chart.has_title = False
    plot = chart.plots[0]
    plot.gap_width = 40
    plot.series[0].format.fill.solid()
    plot.series[0].format.fill.fore_color.rgb = RC_RED
    cat = chart.category_axis
    cat.tick_labels.font.size = Pt(8)
    cat.format.line.color.rgb = CARD_BORDER
    val = chart.value_axis
    val.visible = False
    val.has_major_gridlines = False
    try:
        val.major_gridlines.format.line.fill.background()
    except Exception:
        pass
    return gframe


def _worst_table(s, queues, x, y, w):
    rows = len(queues) + 1
    tbl_shape = s.shapes.add_table(rows, 4, x, y, w, Inches(0.32) * rows)
    tbl = tbl_shape.table
    tbl.columns[0].width = Inches(2.5)
    tbl.columns[1].width = Inches(0.45)
    tbl.columns[2].width = Inches(0.55)
    tbl.columns[3].width = Inches(0.45)
    headers = ["Queue", "T", "Miss", "%"]
    for j, htext in enumerate(headers):
        c = tbl.cell(0, j)
        c.fill.solid(); c.fill.fore_color.rgb = RC_NAVY
        _cell(c, htext, WHITE, bold=True, size=9, align=PP_ALIGN.CENTER if j else PP_ALIGN.LEFT)
    for i, q in enumerate(queues, 1):
        bg = ROW_ALT if i % 2 else WHITE
        vals = [q.name[:26], q.tier, f"{q.total_missed}", f"{round(q.miss_rate*100)}%"]
        for j, v in enumerate(vals):
            c = tbl.cell(i, j)
            c.fill.solid(); c.fill.fore_color.rgb = bg
            col = RC_RED if (j == 3 and q.miss_rate >= 0.7) else DARK
            _cell(c, v, col, bold=(j == 3 and q.miss_rate >= 0.7), size=8.5,
                  align=PP_ALIGN.CENTER if j else PP_ALIGN.LEFT)


# ---------------------------------------------------------------------------
# Slide 3 — AIR opportunity signals (abandoned + sub-60s answered)
# ---------------------------------------------------------------------------

def _slide3(prs, r: PipelineResult, ctx, narr):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _logo(s)
    _title_block(s, narr.get("title", "Where AI Receptionist captures revenue today"),
                 narr.get("subtitle", ""))

    abandon_rate = r.abandoned_total / r.universe_sessions if r.universe_sessions else 0

    # Two stacked stat cards on the left
    lx, lw = Inches(0.5), Inches(4.15)
    # Card A — abandoned callers
    ay, ah = Inches(1.95), Inches(2.25)
    _rect(s, lx, ay, lw, ah, WHITE, line=CARD_BORDER, radius=True)
    _text(s, "CALLERS WHO WAITED,\nTHEN HUNG UP", lx + Inches(0.25), ay + Inches(0.22),
          lw - Inches(0.5), Inches(0.6), size=12, bold=True, color=GRAY, font="Arial")
    _text(s, f"{r.abandoned_total:,}", lx + Inches(0.2), ay + Inches(0.78), lw - Inches(0.3), Inches(1.0),
          size=58, bold=True, color=RC_RED, font="Arial")
    _rich(s, [[(f"{abandon_rate*100:.1f}%", {"bold": True, "size": 12, "color": DARK}),
               (" of inbound queue calls were abandoned in queue", {"size": 12, "color": GRAY})]],
          lx + Inches(0.25), ay + Inches(1.78), lw - Inches(0.5), Inches(0.4))

    # Card B — answered under 60s
    by, bh = Inches(4.4), Inches(2.25)
    _rect(s, lx, by, lw, bh, WHITE, line=CARD_BORDER, radius=True)
    _text(s, "ANSWERED CALLS UNDER\n60 SECONDS", lx + Inches(0.25), by + Inches(0.22),
          lw - Inches(0.5), Inches(0.6), size=12, bold=True, color=GRAY, font="Arial")
    _text(s, f"{r.answered_under_60:,}", lx + Inches(0.2), by + Inches(0.78), lw - Inches(0.3), Inches(1.0),
          size=58, bold=True, color=RC_ORANGE, font="Arial")
    _rich(s, [[(f"{r.under_60_pct*100:.0f}%", {"bold": True, "size": 12, "color": DARK}),
               (" of answered calls — short, routine calls AIR can handle", {"size": 12, "color": GRAY})]],
          lx + Inches(0.25), by + Inches(1.78), lw - Inches(0.5), Inches(0.4))

    # Right — most-abandoned queues table
    rx = Inches(5.0)
    _text(s, "Where callers give up — most-abandoned queues", rx, Inches(1.95),
          Inches(7.8), Inches(0.35), size=14, bold=True, color=DARK, font="Arial")
    top_ab = sorted((q for q in r.queue_stats.values() if q.abandoned_total > 0),
                    key=lambda q: q.abandoned_total, reverse=True)[:12]
    _abandon_table(s, top_ab, rx, Inches(2.4), Inches(7.85))

    # Takeaway strip
    _rect(s, Inches(0.5), Inches(6.85), Inches(12.33), Pt(0.5), CARD_BORDER)
    _text(s, "AI Receptionist answers instantly — recovering abandoned callers and "
             "deflecting short, routine calls so staff focus on revenue conversations.",
          Inches(0.5), Inches(6.92), Inches(12.3), Inches(0.4),
          size=11, italic=True, color=RC_BLUE, align=PP_ALIGN.CENTER)
    _footer(s, 3)


def _abandon_table(s, queues, x, y, w):
    rows = len(queues) + 1
    row_in = 0.34
    tbl_shape = s.shapes.add_table(rows, 5, x, y, w, Inches(row_in * rows))
    tbl = tbl_shape.table
    tbl.first_row = False
    tbl.horz_banding = False
    widths = [4.05, 0.7, 1.2, 0.9, 1.0]
    for j, ww in enumerate(widths):
        tbl.columns[j].width = Inches(ww)
    for i in range(rows):
        tbl.rows[i].height = Inches(row_in)
    headers = ["Queue", "Tier", "Abandoned", "Ab. %", "Ans. <60s"]
    for j, htext in enumerate(headers):
        c = tbl.cell(0, j)
        c.fill.solid(); c.fill.fore_color.rgb = RC_NAVY
        _cell(c, htext, WHITE, bold=True, size=9,
              align=PP_ALIGN.LEFT if j == 0 else PP_ALIGN.CENTER)
    tier_col = {"A": RC_RED, "B": RC_GOLD, "C": RC_BLUE}
    for i, q in enumerate(queues, 1):
        bg = ROW_ALT if i % 2 else WHITE
        vals = [q.name[:40], q.tier, f"{q.abandoned_total:,}",
                f"{round(q.abandon_rate*100)}%", f"{q.answered_under_60:,}"]
        for j, v in enumerate(vals):
            c = tbl.cell(i, j)
            c.fill.solid(); c.fill.fore_color.rgb = bg
            if j == 1:
                col = tier_col.get(q.tier, DARK); bold = True
            elif j == 3:
                col = RC_RED if q.abandon_rate >= 0.5 else DARK
                bold = q.abandon_rate >= 0.5
            else:
                col = DARK; bold = False
            _cell(c, v, col, bold=bold, size=9,
                  align=PP_ALIGN.LEFT if j == 0 else PP_ALIGN.CENTER)


# ---------------------------------------------------------------------------
# Slide 4 — queue-level table
# ---------------------------------------------------------------------------

def _slide4(prs, r: PipelineResult, ctx, narr):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _logo(s)
    _title_block(s, narr.get("title", "Queue-level missed call analysis — sales-taking queues (Tier A+B+C)"),
                 narr.get("subtitle", ""))

    # Tier summary cards
    tier_meta = [
        ("A", "Named 'Sales'", RC_RED, RGBColor(0xFC,0xEC,0xEC)),
        ("B", "Retail / Product Counter", RC_GOLD, RGBColor(0xFB,0xF3,0xE0)),
        ("C", "Branch Front Desk / Main", RC_BLUE, RGBColor(0xE9,0xEE,0xF6)),
    ]
    cw = Inches(4.05)
    cxs = [Inches(0.5), Inches(4.68), Inches(8.86)]
    for (tier, label, col, bg), cx in zip(tier_meta, cxs):
        qs = [q for q in r.queue_stats.values() if q.tier == tier]
        nq = len(qs); inb = sum(q.inbound for q in qs); mis = sum(q.total_missed for q in qs)
        mr = mis / inb if inb else 0
        _rect(s, cx, Inches(1.85), cw, Inches(0.95), bg, line=col, line_w=Pt(1), radius=True)
        _circle(s, cx + Inches(0.18), Inches(2.02), Inches(0.42), col, tier, size=15)
        _text(s, label, cx + Inches(0.72), Inches(2.0), cw - Inches(0.85), Inches(0.35),
              size=13, bold=True, color=col, font="Arial")
        _rich(s, [[(f"{nq} ", {"bold": True, "size": 12, "color": DARK}),
                   ("queues   ", {"size": 11, "color": GRAY}),
                   (f"{inb:,} ", {"bold": True, "size": 12, "color": DARK}),
                   ("inbound", {"size": 11, "color": GRAY})]],
              cx + Inches(0.72), Inches(2.32), cw - Inches(0.8), Inches(0.25))
        _rich(s, [[(f"{mis:,} ", {"bold": True, "size": 12, "color": col}),
                   ("missed   ", {"size": 11, "color": GRAY}),
                   (f"{mr*100:.1f}%", {"bold": True, "size": 12, "color": col})]],
              cx + Inches(0.72), Inches(2.55), cw - Inches(0.8), Inches(0.25))

    # Full queue table
    order = {"A": 0, "B": 1, "C": 2}
    queues = sorted(r.queue_stats.values(),
                    key=lambda q: (order.get(q.tier, 9), -q.total_missed))
    _full_table(s, queues, Inches(0.5), Inches(2.92), Inches(12.33))
    _footer(s, 4)


def _full_table(s, queues, x, y, w):
    rows = len(queues) + 1
    row_in = 0.135
    tbl_shape = s.shapes.add_table(rows, 7, x, y, w, Inches(row_in * rows))
    tbl = tbl_shape.table
    tbl.first_row = False
    tbl.horz_banding = False
    widths = [4.6, 0.7, 3.6, 1.1, 1.0, 0.85, 0.85]
    for j, ww in enumerate(widths):
        tbl.columns[j].width = Inches(ww)
    for i in range(rows):
        tbl.rows[i].height = Inches(row_in)
    headers = ["Queue Name", "Tier", "Classification", "Inbound", "Missed", "Miss %", "Ans %"]
    for j, htext in enumerate(headers):
        c = tbl.cell(0, j)
        c.fill.solid(); c.fill.fore_color.rgb = RC_NAVY
        _cell(c, htext, WHITE, bold=True, size=8.5, pad=0,
              align=PP_ALIGN.LEFT if j in (0, 2) else PP_ALIGN.CENTER)
    tier_col = {"A": RC_RED, "B": RC_GOLD, "C": RC_BLUE}
    for i, q in enumerate(queues, 1):
        bg = ROW_ALT if i % 2 else WHITE
        vals = [q.name[:46], q.tier, q.classification[:38],
                f"{q.inbound:,}", f"{q.total_missed:,}",
                f"{round(q.miss_rate*100)}%", f"{round(q.answer_rate*100)}%"]
        for j, v in enumerate(vals):
            c = tbl.cell(i, j)
            c.fill.solid(); c.fill.fore_color.rgb = bg
            if j == 1:
                col = tier_col.get(q.tier, DARK); bold = True
            elif j == 5:
                col = RC_RED if q.miss_rate >= 0.7 else DARK
                bold = q.miss_rate >= 0.7
            else:
                col = DARK; bold = False
            _cell(c, v, col, bold=bold, size=7.5, pad=0,
                  align=PP_ALIGN.LEFT if j in (0, 2) else PP_ALIGN.CENTER)


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------

def _cell(cell, text, color, *, bold=False, size=9, align=PP_ALIGN.LEFT, pad=1):
    cell.margin_left = Pt(4); cell.margin_right = Pt(4)
    cell.margin_top = Pt(pad); cell.margin_bottom = Pt(pad)
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf = cell.text_frame; tf.word_wrap = False
    p = tf.paragraphs[0]; p.alignment = align
    r = p.add_run(); r.text = text
    r.font.size = Pt(size); r.font.bold = bold
    r.font.color.rgb = color; r.font.name = "Calibri"


def _pct(part, whole):
    return f"{round(100*part/whole)}%" if whole else "0%"


def _hour_label(h):
    suffix = "a" if h < 12 else "p"
    hh = h if h <= 12 else h - 12
    return f"{hh}{suffix}"


# ---------------------------------------------------------------------------
# orchestration
# ---------------------------------------------------------------------------

def build_deck(result: PipelineResult, run_id: str, customer: str, ae_name: str,
               prior_instructions: list[dict] | None = None) -> Path:
    out_dir = Path(tempfile.gettempdir()) / "rc_analyzer_decks"
    out_dir.mkdir(exist_ok=True)
    out_path = out_dir / f"{run_id}.pptx"

    sales_queue_calls = sum(q.total_missed for q in result.queue_stats.values() if q.tier == "A")
    top_queues = sorted(result.queue_stats.values(), key=lambda q: q.total_missed, reverse=True)[:5]

    ctx = {
        "customer": customer,
        "reporting_period": result.reporting_period,
        "universe_sessions": result.universe_sessions,
        "total_missed": result.total_missed,
        "miss_rate_pct": round(result.miss_rate * 100, 1),
        "answer_rate_pct": round(result.answer_rate * 100, 1),
        "abandoned": result.abandoned,
        "voicemail": result.voicemail_total,
        "rang_out": result.missed,
        "phantom_legs_removed": result.phantom_legs_removed,
        "spam_sessions_removed": result.spam_sessions_removed,
        "raw_inbound_legs": result.raw_inbound_legs,
        "inbound_sessions": result.inbound_sessions,
        "business_hours_miss_pct": round(result.business_hours_miss_pct * 100),
        "repeat_callers": result.repeat_callers,
        "misses_per_day": round(result.misses_per_day),
        "abandoned_total": result.abandoned_total,
        "answered_under_60": result.answered_under_60,
        "under_60_pct": round(result.under_60_pct * 100),
        "sales_queue_missed": sales_queue_calls,
        "num_queues": len(result.queue_stats),
        "top_missed_queues": {q.name: q.total_missed for q in top_queues},
    }

    narr1 = _narr1(ctx, prior_instructions)
    narr2 = _narr_titles(ctx, prior_instructions, "slide2")
    narr3 = {"title": "Where AI Receptionist captures revenue today",
             "subtitle": f"Abandoned-in-queue callers and short routine calls · Tier A+B+C · {result.reporting_period}"}
    narr4 = {"title": "Queue-level missed call analysis (Tier A+B+C)",
             "subtitle": f"Session-deduplicated · spam-filtered · {result.reporting_period} · back-office (Tier D) excluded · {len(result.queue_stats)} queues shown"}

    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    _slide1(prs, result, ctx, narr1)
    _slide2(prs, result, ctx, narr2, sales_queue_calls)
    _slide3(prs, result, ctx, narr3)
    _slide4(prs, result, ctx, narr4)

    prs.save(out_path)
    return out_path


def _narr1(ctx, prior):
    schema = {
        "title": "str — slide title, credibility framing about genuine missed calls",
        "subtitle": "str — source line: queues, tier, period, sessions analyzed",
        "card1": [["The problem:", "..."], ["Our fix:", "..."], ["Proof:", "..."]],
        "card2": [["The rule:", "..."], ["Result:", "..."], ["What it means:", "..."]],
        "card3": [["The test:", "..."], ["Result:", "..."], ["What it means:", "..."]],
        "card4": [["The test:", "..."], ["What we found:", "..."], ["What it means:", "..."]],
    }
    try:
        out = generate_narrative(ctx, schema, prior)
        # normalize card paras to list of [label, body]
        for k in ("card1", "card2", "card3", "card4"):
            out[k] = [(p[0], p[1]) for p in out.get(k, [])]
        return out
    except Exception:
        return _fallback1(ctx)


def _narr_titles(ctx, prior, which):
    schema = {"title": "str — slide headline", "subtitle": "str — methodology source line"}
    try:
        return generate_narrative({**ctx, "slide": which}, schema, prior)
    except Exception:
        if which == "slide2":
            return {"title": f"{ctx['total_missed']:,} genuine missed calls across sales-taking queues",
                    "subtitle": f"Tier A+B+C · {ctx['reporting_period']} · session-deduplicated · spam-filtered · back-office excluded"}
        return {"title": "Queue-level missed call analysis — sales-taking queues (Tier A+B+C)",
                "subtitle": f"Session-deduplicated · spam-filtered · {ctx['reporting_period']} · back-office (Tier D) excluded"}


def _fallback1(ctx):
    return {
        "title": "How we know these are genuine missed calls — not spam, not noise",
        "subtitle": f"RingCentral Performance Reports · Tier A+B+C queues · {ctx['reporting_period']} · {ctx['universe_sessions']:,} legitimate inbound sessions analyzed",
        "card1": [("The problem:", "Standard reports count one row per agent ring leg — a call routed to many agents shows as many phantom missed calls."),
                  ("Our fix:", f"Every leg was grouped by Session ID — one call, one count. {ctx['phantom_legs_removed']} phantom legs removed."),
                  ("Proof:", f"{ctx['raw_inbound_legs']:,} raw legs → {ctx['inbound_sessions']:,} unique sessions. Counts are callers, not ring events.")],
        "card2": [("The rule:", "Any session whose longest leg lasted 5 seconds or less was classified as spam and excluded."),
                  ("Result:", f"Only {ctx['spam_sessions_removed']} sessions removed — every caller counted waited more than 5 seconds."),
                  ("What it means:", "These are real people trying to reach the business, not robocalls.")],
        "card3": [("The test:", "If misses were after-hours, they would cluster outside business hours — a staffing issue, not an AI opportunity."),
                  ("Result:", f"{ctx['business_hours_miss_pct']}% of misses landed Monday–Friday, 7am–6pm, during fully staffed hours."),
                  ("What it means:", "These are answerable calls. People were at their desks; calls simply weren't picked up.")],
        "card4": [("The test:", "Spam dials every number once; real buyers call back because they need something."),
                  ("What we found:", f"{ctx['repeat_callers']} callers tried 2+ times and never got through."),
                  ("What it means:", "Repeat unanswered callers are the strongest evidence real buyers are being lost.")],
    }
