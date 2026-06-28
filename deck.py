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

# ---------------------------------------------------------------------------
# RingCentral 2026 corporate template (rc-presentation-template skill)
# ---------------------------------------------------------------------------
# Mandatory brand font. PowerPoint renders this when Inter Tight is installed;
# the .pptx names it regardless so it's correct on a branded machine.
FONT = "Inter Tight"

# Brand assets bundled in the repo (logos + pre-rendered gradient backgrounds).
_BRAND = Path(__file__).resolve().parent / "assets" / "brand"
BG_LIGHT = str(_BRAND / "bg_light.png")   # content / stats / tables
BG_WARM = str(_BRAND / "bg_warm.png")     # covers / dividers / closing
LOGO_COLOR = str(_BRAND / "logo_color.png")  # on light gradient
LOGO_WHITE = str(_BRAND / "logo_white.png")  # on warm gradient

# Brand colors. Accent orange is used ONLY as an accent (callouts, active nodes).
RC_ORANGE = RGBColor(0xFF, 0x88, 0x00)   # FF8800 brand accent
RC_NAVY = RGBColor(0x1B, 0x2A, 0x4A)     # 1B2A4A table-header fill
RC_BLUE = RGBColor(0x06, 0x2E, 0x5C)
RC_RED = RGBColor(0xC0, 0x2A, 0x2A)
RC_TEAL = RGBColor(0x0A, 0x8A, 0x8A)
RC_GOLD = RGBColor(0xC8, 0x8A, 0x00)
RC_PURPLE = RGBColor(0x5B, 0x3E, 0x96)
DARK = RGBColor(0x1A, 0x1A, 0x1A)        # 1A1A1A body/title on light slides
GRAY = RGBColor(0x88, 0x88, 0x88)        # muted footer / captions
LIGHT = RGBColor(0xF4, 0xF5, 0xF8)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
CARD_BORDER = RGBColor(0xE2, 0xE4, 0xEA)
ROW_ALT = RGBColor(0xF8, 0xF4, 0xF0)     # alternating table row

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)


def _bg(slide, *, warm=False):
    """Full-bleed RingCentral gradient as the backmost shape on the slide."""
    pic = slide.shapes.add_picture(BG_WARM if warm else BG_LIGHT,
                                   0, 0, SLIDE_W, SLIDE_H)
    # Send the picture to the very back so all content renders on top of it.
    spTree = slide.shapes._spTree
    spTree.remove(pic._element)
    spTree.insert(2, pic._element)
    return pic


def _text(slide, text, left, top, width, height, *, size=14, bold=False,
          color=DARK, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
          italic=False, font=FONT, wrap=True, line_spacing=None):
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
            r.font.name = opts.get("font", FONT)
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
        r.font.color.rgb = text_color; r.font.name = FONT
    return c


def _logo(slide, *, warm=False):
    # Official RingCentral wordmark: white on warm gradient, color on light.
    slide.shapes.add_picture(LOGO_WHITE if warm else LOGO_COLOR,
                             Inches(0.3), Inches(0.22), height=Inches(0.3))


def _footer(slide, page=None, *, warm=False):
    # Page number is stamped in a final pass (see _stamp_page_numbers); the
    # optional `page` arg is ignored so adding/removing slides never desyncs.
    col = WHITE if warm else GRAY
    _text(slide, "Confidential", Inches(5.9), Inches(7.12), Inches(1.5), Inches(0.3),
          size=8, color=col, align=PP_ALIGN.CENTER)
    _text(slide, "©2026 RingCentral", Inches(11.0), Inches(7.12), Inches(1.9), Inches(0.3),
          size=8, color=col, align=PP_ALIGN.RIGHT)


# Slides whose background is the warm gradient (white text), set during build.
_WARM_SLIDES = set()


def _stamp_page_numbers(prs):
    for i, slide in enumerate(prs.slides, 1):
        col = WHITE if i in _WARM_SLIDES else GRAY
        _text(slide, str(i), Inches(0.45), Inches(7.12), Inches(0.5), Inches(0.3),
              size=8, color=col)


def _title_block(slide, title, subtitle, *, warm=False):
    # Long titles wrap to a second line; shrink the size and push the subtitle
    # down so the wrapped title never collides with it (a wider fallback font
    # makes this worse, so size off character count, not measured width).
    long = len(title) > 52
    _text(slide, title, Inches(0.5), Inches(0.5), Inches(12.3), Inches(0.95),
          size=22 if long else 28, bold=True, color=WHITE if warm else DARK, font=FONT)
    _text(slide, subtitle, Inches(0.52), Inches(1.55) if long else Inches(1.32),
          Inches(12.3), Inches(0.4),
          size=12, italic=True, color=WHITE if warm else GRAY, font=FONT)


# ---------------------------------------------------------------------------
# Slide 1 — methodology credibility
# ---------------------------------------------------------------------------

# ---------------------------------------------------------------------------
# Slide 2 — missed-call summary
# ---------------------------------------------------------------------------

def _slide2(prs, r: PipelineResult, ctx, narr, sales_queue_calls):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _bg(s)
    _logo(s)
    _title_block(s, narr.get("title", f"{r.total_missed:,} genuine missed calls across customer-facing queues"),
                 narr.get("subtitle", ""))

    # Left big-stat card
    lx, ly, lw, lh = Inches(0.5), Inches(1.95), Inches(3.25), Inches(4.7)
    _rect(s, lx, ly, lw, lh, WHITE, line=CARD_BORDER, radius=True)
    _text(s, "GENUINE MISSED\nCALLS — CUSTOMER QUEUES", lx + Inches(0.25), ly + Inches(0.55),
          lw - Inches(0.5), Inches(0.7), size=12, bold=True, color=GRAY, font=FONT)
    # Size the headline number to the digit count so 5- and 6-figure totals
    # both fit on ONE line in the ~2.95in card (72pt overflowed and wrapped).
    _big_n = f"{r.total_missed:,}"
    # Tier the headline size to digit-count so it fits on ONE line even when a
    # WIDER fallback font is substituted (Inter Tight may be absent on the
    # rendering machine). Sizes chosen to fit the ~2.95in card with margin.
    _nlen = len(_big_n)
    _big_size = 66 if _nlen <= 4 else 54 if _nlen == 5 else 46 if _nlen == 6 else 38
    _text(s, _big_n, lx + Inches(0.1), ly + Inches(1.3), lw - Inches(0.2), Inches(1.2),
          size=_big_size, bold=True, color=RC_RED, font=FONT, wrap=False,
          align=PP_ALIGN.CENTER)
    _rich(s, [[(f"{r.miss_rate*100:.1f}%", {"bold": True, "size": 13, "color": DARK}),
               (f" of {r.universe_sessions:,} inbound sessions went unanswered",
                {"size": 13, "color": DARK})]],
          lx + Inches(0.25), ly + Inches(2.35), lw - Inches(0.45), Inches(0.7))
    _rect(s, lx + Inches(0.25), ly + Inches(3.05), lw - Inches(0.5), Pt(1), CARD_BORDER)
    _text(s, f"≈ {round(r.misses_per_day)} missed calls every day",
          lx + Inches(0.25), ly + Inches(3.15), lw - Inches(0.5), Inches(0.4),
          size=13, bold=True, italic=True, color=RC_ORANGE, align=PP_ALIGN.CENTER)
    _rich(s, [[(f"{sales_queue_calls:,}", {"bold": True, "size": 13, "color": RC_RED}),
               (" are revenue-line calls\n(sales · orders · bookings)", {"size": 12, "color": DARK})]],
          lx + Inches(0.25), ly + Inches(3.7), lw - Inches(0.5), Inches(0.8),
          align=PP_ALIGN.CENTER)

    # Middle: chart + breakdown
    mx = Inches(4.0)
    _text(s, f"{round(r.business_hours_miss_pct*100)}% of misses hit during staffed hours — peak overflow, "
             f"when every agent is already on a call",
          mx, Inches(1.95), Inches(4.6), Inches(0.5), size=11, bold=True, color=DARK, font=FONT, line_spacing=1.0)
    _hourly_chart(s, r, mx, Inches(2.35), Inches(4.55), Inches(1.85))

    _text(s, "How they were missed:", mx, Inches(4.35), Inches(4.6), Inches(0.3),
          size=12, bold=True, color=DARK, font=FONT)
    # The two dispositions of a genuine missed call — these sum to total_missed
    # exactly (rang-out + voicemail). "Abandoned in queue" is a DIFFERENT measure
    # (from the Queues report), not a third slice, so it's shown separately below
    # rather than inside this breakdown (which would push the total past 100%).
    qr = r.queues_report
    abandoned_n = qr.abandoned if (qr and qr.abandoned) else r.abandoned
    bd = [
        (r.missed, "Rang out — no answer", RC_PURPLE, RGBColor(0xF0,0xEC,0xF7)),
        (r.voicemail_total, "Went to voicemail (left msg)", RC_GOLD, RGBColor(0xFB,0xF3,0xE0)),
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
    if abandoned_n:
        _text(s, f"Separately, the Queues report shows {abandoned_n:,} callers abandoned in queue "
                 f"before reaching anyone — a distinct measure, not part of the split above.",
              mx, by + Inches(0.02), Inches(4.6), Inches(0.55), size=9.5, italic=True, color=RC_RED)
    else:
        _text(s, f"{r.repeat_callers} callers tried 2+ times and never got through — intent, not spam.",
              mx, by + Inches(0.02), Inches(4.6), Inches(0.5), size=10, italic=True, color=RC_RED)

    # Right: worst-hit queues table. Prefer the Queues report (real per-queue
    # abandoned); the Calls export collapses all abandoned calls into "Unknown",
    # which hides which sales/customer-facing queues actually lost the calls.
    rx = Inches(8.85)
    if qr and qr.queues:
        _text(s, "Worst-hit queues — by abandoned", rx, Inches(1.95), Inches(4.0), Inches(0.35),
              size=14, bold=True, color=DARK, font=FONT)
        worst = sorted((q for q in qr.queues if (q.tier or "C") != "D"),
                       key=lambda q: q.abandoned, reverse=True)[:10]
        _worst_table_qr(s, worst, rx, Inches(2.35), Inches(3.95))
        if any(getattr(q, "answered", None) == 0 and q.inbound > 0 for q in worst):
            _text(s, "† unstaffed queue — 0 agents ever answered; a routing/staffing fix (see slide 5), not a data error.",
                  rx, Inches(5.95), Inches(4.0), Inches(0.6), size=8, italic=True, color=GRAY)
    else:
        _text(s, "Worst-hit queues", rx, Inches(1.95), Inches(4.0), Inches(0.35),
              size=14, bold=True, color=DARK, font=FONT)
        worst = sorted(r.queue_stats.values(), key=lambda q: q.total_missed, reverse=True)[:10]
        _worst_table(s, worst, rx, Inches(2.35), Inches(3.95))
    _footer(s, 2)


def _worst_table_qr(s, queues, x, y, w):
    """Worst-hit table from the Queues report (real abandoned per queue)."""
    rows = len(queues) + 1
    tbl_shape = s.shapes.add_table(rows, 4, x, y, w, Inches(0.32) * rows)
    tbl = tbl_shape.table
    tbl.columns[0].width = Inches(2.5)
    tbl.columns[1].width = Inches(0.45)
    tbl.columns[2].width = Inches(0.55)
    tbl.columns[3].width = Inches(0.45)
    headers = ["Queue", "Rev", "Aband", "%"]
    for j, htext in enumerate(headers):
        c = tbl.cell(0, j)
        c.fill.solid(); c.fill.fore_color.rgb = RC_NAVY
        _cell(c, htext, WHITE, bold=True, size=9, align=PP_ALIGN.CENTER if j else PP_ALIGN.LEFT)
    for i, q in enumerate(queues, 1):
        bg = ROW_ALT if i % 2 else WHITE
        # "Rev" column: ✓ marks a revenue-line queue (sales/orders/bookings).
        is_rev = (q.tier or "C") in ("A", "B")
        # A 100% abandon rate on a queue with zero answered isn't a data error — it's
        # an UNSTAFFED queue (0 agents). Mark it with † so it reads as the config
        # finding it is (see slide 5), not a broken number.
        unstaffed = (getattr(q, "answered", None) == 0 and q.inbound > 0)
        pct_txt = f"{round(q.abandon_rate*100)}%" + ("†" if unstaffed else "")
        vals = [q.name[:26], "✓" if is_rev else "", f"{q.abandoned}", pct_txt]
        for j, v in enumerate(vals):
            c = tbl.cell(i, j)
            c.fill.solid(); c.fill.fore_color.rgb = bg
            if j == 1:
                col = RC_ORANGE
            elif j == 3 and q.abandon_rate >= 0.7:
                col = RC_RED
            else:
                col = DARK
            _cell(c, v, col, bold=(j == 1) or (j == 3 and q.abandon_rate >= 0.7), size=8.5,
                  align=PP_ALIGN.CENTER if j else PP_ALIGN.LEFT)


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
    headers = ["Queue", "Rev", "Miss", "%"]
    for j, htext in enumerate(headers):
        c = tbl.cell(0, j)
        c.fill.solid(); c.fill.fore_color.rgb = RC_NAVY
        _cell(c, htext, WHITE, bold=True, size=9, align=PP_ALIGN.CENTER if j else PP_ALIGN.LEFT)
    for i, q in enumerate(queues, 1):
        bg = ROW_ALT if i % 2 else WHITE
        is_rev = (q.tier or "C") in ("A", "B")
        vals = [q.name[:26], "✓" if is_rev else "", f"{q.total_missed}", f"{round(q.miss_rate*100)}%"]
        for j, v in enumerate(vals):
            c = tbl.cell(i, j)
            c.fill.solid(); c.fill.fore_color.rgb = bg
            if j == 1:
                col = RC_ORANGE
            elif j == 3 and q.miss_rate >= 0.7:
                col = RC_RED
            else:
                col = DARK
            _cell(c, v, col, bold=(j == 1) or (j == 3 and q.miss_rate >= 0.7), size=8.5,
                  align=PP_ALIGN.CENTER if j else PP_ALIGN.LEFT)


# ---------------------------------------------------------------------------
# Slide 3 — miss rate by hour of day (reference deck's key slide)
# ---------------------------------------------------------------------------

def _slide_hourly(prs, r: PipelineResult, ctx, narr):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _bg(s)
    _logo(s)
    _title_block(s, narr.get("title", "Calls slip away after hours, on weekends — and even midday"),
                 narr.get("subtitle", ""))

    # Big miss-rate column chart (left)
    _miss_rate_chart(s, r, Inches(0.5), Inches(2.05), Inches(8.0), Inches(4.55))

    # Three callout cards (right)
    cx, cw = Inches(8.85), Inches(3.98)
    cards = [
        (_range_or_pct(r.after_hours_miss_lo, r.after_hours_miss_hi, r.after_hours_miss_rate),
         "of after-hours calls (6pm–6am) are missed"),
        (_range_or_pct(r.weekend_miss_lo, r.weekend_miss_hi, r.weekend_miss_rate),
         "missed on Saturdays and Sundays"),
        (f"~{round(r.midday_miss_rate*100)}%",
         "missed even during peak midday hours"),
    ]
    cy = Inches(2.05); ch = Inches(1.32); gap = Inches(0.2)
    for big, label in cards:
        _rect(s, cx, cy, cw, ch, LIGHT, radius=True)
        _text(s, big, cx + Inches(0.3), cy + Inches(0.16), cw - Inches(0.6), Inches(0.6),
              size=30, bold=True, color=RC_ORANGE, font=FONT)
        _text(s, label, cx + Inches(0.3), cy + Inches(0.78), cw - Inches(0.6), Inches(0.45),
              size=12, color=GRAY)
        cy = Emu(int(cy) + int(ch) + int(gap))

    _text(s, "AIR answers instantly — every hour, every day.",
          cx, cy + Inches(0.05), cw, Inches(0.4),
          size=13, bold=True, italic=True, color=RC_BLUE)
    _footer(s, 3)


def _miss_rate_chart(s, r, x, y, w, h):
    cd = CategoryChartData()
    cd.categories = [_hour_label24(hh) for hh in range(24)]
    cd.add_series("Miss %", [round(r.hourly_miss_rate.get(hh, 0) * 100) for hh in range(24)])
    gframe = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, x, y, w, h, cd)
    chart = gframe.chart
    chart.has_legend = False
    chart.has_title = True
    chart.chart_title.text_frame.text = "Miss %"
    chart.chart_title.text_frame.paragraphs[0].runs[0].font.size = Pt(13)
    chart.chart_title.text_frame.paragraphs[0].runs[0].font.bold = True
    plot = chart.plots[0]
    plot.gap_width = 35
    plot.series[0].format.fill.solid()
    plot.series[0].format.fill.fore_color.rgb = RC_ORANGE
    cat = chart.category_axis
    cat.tick_labels.font.size = Pt(9)
    cat.format.line.color.rgb = CARD_BORDER
    val = chart.value_axis
    val.minimum_scale = 0
    val.maximum_scale = 100
    val.tick_labels.font.size = Pt(9)
    val.has_major_gridlines = True
    try:
        val.major_gridlines.format.line.color.rgb = RGBColor(0xEE, 0xEE, 0xEE)
    except Exception:
        pass
    return gframe


# ---------------------------------------------------------------------------
# Slide 4 — AIR opportunity signals (abandoned + sub-60s answered)
# ---------------------------------------------------------------------------

def _slide3(prs, r: PipelineResult, ctx, narr):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _bg(s)
    _logo(s)
    _title_block(s, narr.get("title", "Where AI Receptionist captures revenue today"),
                 narr.get("subtitle", ""))

    # Real abandoned data comes from the Queues report (2nd upload). The Calls
    # export doesn't carry an "Abandoned" disposition, so prefer the Queues
    # report; fall back to the Calls-derived figure only if it's absent.
    qr = r.queues_report
    if qr and qr.inbound:
        abandoned_n = qr.abandoned
        abandon_rate = qr.abandon_rate
    else:
        abandoned_n = r.abandoned_total
        abandon_rate = r.abandoned_total / r.universe_sessions if r.universe_sessions else 0

    # Two stacked stat cards on the left
    lx, lw = Inches(0.5), Inches(4.15)
    # Card A — abandoned callers
    ay, ah = Inches(1.95), Inches(2.25)
    _rect(s, lx, ay, lw, ah, WHITE, line=CARD_BORDER, radius=True)
    _text(s, "CALLERS WHO WAITED,\nTHEN HUNG UP", lx + Inches(0.25), ay + Inches(0.22),
          lw - Inches(0.5), Inches(0.6), size=12, bold=True, color=GRAY, font=FONT)
    _text(s, f"{abandoned_n:,}", lx + Inches(0.2), ay + Inches(0.78), lw - Inches(0.3), Inches(1.0),
          size=58, bold=True, color=RC_RED, font=FONT)
    _rich(s, [[(f"{abandon_rate*100:.1f}%", {"bold": True, "size": 12, "color": DARK}),
               (" of queue calls were abandoned while waiting", {"size": 12, "color": GRAY})]],
          lx + Inches(0.25), ay + Inches(1.78), lw - Inches(0.5), Inches(0.4))

    # Card B — answered under 60s
    by, bh = Inches(4.4), Inches(2.25)
    _rect(s, lx, by, lw, bh, WHITE, line=CARD_BORDER, radius=True)
    _text(s, "ANSWERED CALLS UNDER\n60 SECONDS", lx + Inches(0.25), by + Inches(0.22),
          lw - Inches(0.5), Inches(0.6), size=12, bold=True, color=GRAY, font=FONT)
    _text(s, f"{r.answered_under_60:,}", lx + Inches(0.2), by + Inches(0.78), lw - Inches(0.3), Inches(1.0),
          size=58, bold=True, color=RC_ORANGE, font=FONT)
    _rich(s, [[(f"{r.under_60_pct*100:.0f}%", {"bold": True, "size": 12, "color": DARK}),
               (" of answered calls ran under 60s — an indicator of routine volume an AI receptionist could handle", {"size": 11, "color": GRAY})]],
          lx + Inches(0.25), by + Inches(1.72), lw - Inches(0.5), Inches(0.5))

    # Right — most-abandoned queues table (from the Queues report when present)
    rx = Inches(5.0)
    _text(s, "Where callers give up — most-abandoned queues", rx, Inches(1.95),
          Inches(7.8), Inches(0.35), size=14, bold=True, color=DARK, font=FONT)

    if qr and qr.queues:
        top_ab = sorted((q for q in qr.queues if q.abandoned > 0 and q.tier != "D"),
                        key=lambda q: q.abandoned, reverse=True)[:11]
        _abandon_table_qr(s, top_ab, rx, Inches(2.4), Inches(7.85))
        if any(getattr(q, "answered", None) == 0 and q.inbound > 0 for q in top_ab):
            _text(s, "† unstaffed queue — 0 agents ever answered (a routing/staffing fix, see slide 5), not a data error.",
                  rx, Inches(6.05), Inches(7.85), Inches(0.3), size=8, italic=True, color=GRAY)
        # Wait-time / SLA context strip
        bits = []
        if qr.avg_wait:
            bits.append(("Avg. wait ", qr.avg_wait))
        if qr.longest_wait:
            bits.append(("Longest wait ", qr.longest_wait))
        if qr.sla_pct:
            bits.append(("Service level ", f"{qr.sla_pct*100:.0f}%"))
        if bits:
            segs = []
            for i, (label, val) in enumerate(bits):
                if i:
                    segs.append(("    ·    ", {"size": 11, "color": CARD_BORDER}))
                segs.append((label, {"size": 11, "color": GRAY}))
                segs.append((val, {"size": 11, "bold": True, "color": RC_RED}))
            _rich(s, [segs], rx, Inches(6.35), Inches(7.85), Inches(0.35))
    else:
        top_ab = sorted((q for q in r.queue_stats.values() if q.abandoned_total > 0),
                        key=lambda q: q.abandoned_total, reverse=True)[:12]
        _abandon_table(s, top_ab, rx, Inches(2.4), Inches(7.85))

    # Takeaway strip
    _rect(s, Inches(0.5), Inches(6.85), Inches(12.33), Pt(0.5), CARD_BORDER)
    _text(s, "AI Receptionist answers instantly — recovering abandoned callers and taking many "
             "short, routine calls so staff focus on revenue conversations.",
          Inches(0.5), Inches(6.92), Inches(12.3), Inches(0.4),
          size=11, italic=True, color=RC_BLUE, align=PP_ALIGN.CENTER)
    _footer(s, 4)


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
    headers = ["Queue", "Rev", "Abandoned", "Ab. %", "Ans. <60s"]
    for j, htext in enumerate(headers):
        c = tbl.cell(0, j)
        c.fill.solid(); c.fill.fore_color.rgb = RC_NAVY
        _cell(c, htext, WHITE, bold=True, size=9,
              align=PP_ALIGN.LEFT if j == 0 else PP_ALIGN.CENTER)
    for i, q in enumerate(queues, 1):
        bg = ROW_ALT if i % 2 else WHITE
        is_rev = (q.tier or "C") in ("A", "B")
        vals = [q.name[:40], "✓" if is_rev else "", f"{q.abandoned_total:,}",
                f"{round(q.abandon_rate*100)}%", f"{q.answered_under_60:,}"]
        for j, v in enumerate(vals):
            c = tbl.cell(i, j)
            c.fill.solid(); c.fill.fore_color.rgb = bg
            if j == 1:
                col = RC_ORANGE; bold = True
            elif j == 3:
                col = RC_RED if q.abandon_rate >= 0.5 else DARK
                bold = q.abandon_rate >= 0.5
            else:
                col = DARK; bold = False
            _cell(c, v, col, bold=bold, size=9,
                  align=PP_ALIGN.LEFT if j == 0 else PP_ALIGN.CENTER)


def _abandon_table_qr(s, queues, x, y, w):
    """Most-abandoned-queues table sourced from the Queues report (real abandoned)."""
    rows = len(queues) + 1
    row_in = 0.34
    tbl_shape = s.shapes.add_table(rows, 5, x, y, w, Inches(row_in * rows))
    tbl = tbl_shape.table
    tbl.first_row = False
    tbl.horz_banding = False
    widths = [3.75, 0.7, 1.2, 1.2, 1.0]
    for j, ww in enumerate(widths):
        tbl.columns[j].width = Inches(ww)
    for i in range(rows):
        tbl.rows[i].height = Inches(row_in)
    headers = ["Queue", "Rev", "Inbound", "Abandoned", "Ab. %"]
    for j, htext in enumerate(headers):
        c = tbl.cell(0, j)
        c.fill.solid(); c.fill.fore_color.rgb = RC_NAVY
        _cell(c, htext, WHITE, bold=True, size=9,
              align=PP_ALIGN.LEFT if j == 0 else PP_ALIGN.CENTER)
    for i, q in enumerate(queues, 1):
        bg = ROW_ALT if i % 2 else WHITE
        is_rev = (q.tier or "C") in ("A", "B")
        unstaffed = (getattr(q, "answered", None) == 0 and q.inbound > 0)
        pct_txt = f"{round(q.abandon_rate*100)}%" + ("†" if unstaffed else "")
        vals = [q.name[:38], "✓" if is_rev else "", f"{q.inbound:,}",
                f"{q.abandoned:,}", pct_txt]
        for j, v in enumerate(vals):
            c = tbl.cell(i, j)
            c.fill.solid(); c.fill.fore_color.rgb = bg
            if j == 1:
                col = RC_ORANGE; bold = True
            elif j == 4:
                col = RC_RED if q.abandon_rate >= 0.5 else DARK
                bold = q.abandon_rate >= 0.5
            else:
                col = DARK; bold = False
            _cell(c, v, col, bold=bold, size=9,
                  align=PP_ALIGN.LEFT if j == 0 else PP_ALIGN.CENTER)


# ---------------------------------------------------------------------------
# Slide 4 — queue-level table
# ---------------------------------------------------------------------------

# ---------------------------------------------------------------------------
# Slide — config-fixable vs. structural gap (license-justification framing)
# ---------------------------------------------------------------------------

def _missed_time_split(r: PipelineResult):
    """Split genuine missed calls into after-hours (structural) vs business-hours.

    Sourced from the Calls export (the only per-call timestamped feed), spam
    excluded so it reconciles to the headline universe. Returns None if the
    timestamp/in-business-hours signal isn't available.
    """
    df = r.sessions_df
    if df is None or "in_business_hours" not in df.columns:
        return None
    missed_vals = ["missed", "abandoned", "vm/missed", "voicemail"]
    m = df[df["outcome"].astype(str).str.strip().str.lower().isin(missed_vals)]
    if "is_spam" in m.columns:
        m = m[~m["is_spam"].astype(bool)]
    total = len(m)
    after = int((~m["in_business_hours"].astype(bool)).sum())
    return {"total": total, "after": after, "business": total - after}


_DEST_COLOR = {
    "Ring group / front desk": RC_ORANGE,
    "Direct to a person's extension": RC_BLUE,
    "IVR / auto-attendant menu": RC_TEAL,
    "Main line / auto-receptionist": RC_GOLD,
}


def _unstaffed_queues(r: PipelineResult):
    """Queues with inbound calls but zero ever answered — config/staffing fixes."""
    qr = r.queues_report
    if not (qr and qr.queues):
        return []
    return [q for q in qr.queues
            if q.answered == 0 and q.inbound > 0 and (q.tier or "C") != "D"]


def _slide_config_vs_air(prs, r: PipelineResult, ctx, narr):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _bg(s)
    _logo(s)
    _title_block(s, narr.get("title", "What queue config can fix — and what needs always-on coverage"),
                 narr.get("subtitle",
                          "Separating calls recoverable by routing/staffing from the structural "
                          "gap that no configuration can answer"))

    total = r.total_missed
    unstaffed = _unstaffed_queues(r)
    config_fixable = sum(q.abandoned for q in unstaffed)
    structural = max(total - config_fixable, 0)
    split = _missed_time_split(r)
    after_hours = split["after"] if split else 0

    # Three flow cards: total -> config-fixable -> structural (AIR)
    cy, ch = Inches(2.05), Inches(2.75)
    cards = [
        (Inches(0.5), Inches(3.55), WHITE, CARD_BORDER, DARK, GRAY,
         "TOTAL GENUINE MISSED", f"{total:,}", "customer-facing queues · spam-filtered · de-duplicated",
         "Every inbound call that never reached a person."),
        (Inches(4.78), Inches(3.55), RGBColor(0xFB,0xF3,0xE0), RC_GOLD, RC_GOLD, GRAY,
         "CONFIG-FIXABLE (PRO SERVICES)", f"{config_fixable:,}",
         f"{len(unstaffed)} unstaffed queue{'s' if len(unstaffed)!=1 else ''} · 0 agents ever answered",
         "Routing/staffing fix — no new licenses required."),
        (Inches(9.06), Inches(3.77), RGBColor(0xFC,0xEC,0xEC), RC_RED, RC_RED, GRAY,
         "STRUCTURAL GAP — NEEDS ALWAYS-ON COVERAGE", f"{structural:,}",
         "Calls arriving when staff are busy, after hours, or at peak",
         "No routing rule answers a call when no human is free."),
    ]
    for x, w, bg, border, numcol, subcol, head, big, sub, foot in cards:
        _rect(s, x, cy, w, ch, bg, line=border, line_w=Pt(1.25), radius=True)
        _text(s, head, x + Inches(0.22), cy + Inches(0.2), w - Inches(0.4), Inches(0.5),
              size=11, bold=True, color=numcol, font=FONT)
        _text(s, big, x + Inches(0.18), cy + Inches(0.72), w - Inches(0.3), Inches(1.0),
              size=46, bold=True, color=numcol, font=FONT, wrap=False)
        _text(s, sub, x + Inches(0.22), cy + Inches(1.82), w - Inches(0.4), Inches(0.5),
              size=10.5, bold=True, color=DARK, font=FONT)
        _text(s, foot, x + Inches(0.22), cy + Inches(2.25), w - Inches(0.4), Inches(0.45),
              size=10, italic=True, color=subcol)

    # Minus / equals connectors between the cards
    _text(s, "−", Inches(4.18), cy + Inches(0.95), Inches(0.6), Inches(0.8),
          size=34, bold=True, color=GRAY, align=PP_ALIGN.CENTER)
    _text(s, "=", Inches(8.46), cy + Inches(0.95), Inches(0.6), Inches(0.8),
          size=34, bold=True, color=GRAY, align=PP_ALIGN.CENTER)

    # Unstaffed-queue detail line under the middle card
    if unstaffed:
        names = " · ".join(f"{q.name} ({q.abandoned:,})" for q in unstaffed[:3])
        _text(s, f"Unstaffed: {names}", Inches(4.78), cy + ch + Inches(0.05),
              Inches(4.0), Inches(0.5), size=8.5, italic=True, color=RC_GOLD)

    # Bottom emphasis bar — the after-hours floor config literally cannot touch
    by = Inches(5.95)
    _rect(s, Inches(0.5), by, Inches(12.33), Inches(0.85), RGBColor(0xFC,0xEC,0xEC),
          line=RC_RED, line_w=Pt(1), radius=True)
    if split and after_hours:
        ah_pct = round(after_hours / split["total"] * 100) if split["total"] else 0
        _rich(s, [[(f"{after_hours:,} ", {"bold": True, "size": 16, "color": RC_RED}),
                   (f"of the missed calls ({ah_pct}%) arrive entirely after hours — when no one is "
                    "scheduled. ", {"size": 12, "color": DARK}),
                   ("Configuration cannot recover a single one; only always-on coverage can.",
                    {"size": 12, "bold": True, "color": RC_RED})]],
              Inches(0.75), by, Inches(11.8), Inches(0.85), anchor=MSO_ANCHOR.MIDDLE)
    else:
        _rich(s, [[("Even with perfectly tuned routing, the structural gap above remains — "
                    "it is answered only by always-on coverage, not by reconfiguring queues.",
                    {"size": 12, "bold": True, "color": RC_RED})]],
              Inches(0.75), by, Inches(11.8), Inches(0.85), anchor=MSO_ANCHOR.MIDDLE)
    _footer(s)


# ---------------------------------------------------------------------------
# Slide — predicted call reasons (Firecrawl business context)
# ---------------------------------------------------------------------------

_TIER_BADGE = {
    "A": (RC_ORANGE, RGBColor(0xFF, 0xF1, 0xE3), "Revenue-line"),
    "B": (RC_ORANGE, RGBColor(0xFF, 0xF1, 0xE3), "Revenue-line"),
    "C": (RC_BLUE, RGBColor(0xE9, 0xEE, 0xF6), "General"),
    "D": (GRAY, LIGHT, "Internal"),
}


def _slide_call_reasons(prs, r: PipelineResult, ctx, narr):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _bg(s)
    _logo(s)
    biz = ctx.get("business") or {}
    summary = (biz.get("summary") or "").strip()
    url = biz.get("url", "")
    customer = ctx.get("customer", "this business")
    reasons_present = bool(biz.get("predicted_call_reasons"))

    # This is the OPENER: ground the deck in who they are and who's calling them,
    # then land the headline missed-call number. With no business context (no
    # Firecrawl crawl), fall back to a clean, number-led opener so the slide
    # always renders.
    if not (summary or reasons_present):
        _title_block(s, f"{r.total_missed:,} missed calls — and who was trying to reach {customer}",
                     f"RingCentral Performance Reports · {r.reporting_period} · "
                     f"session-deduplicated · spam-filtered · customer-facing lines only")
        _footer(s)
        return

    _title_block(s, narr.get("title", f"Who {customer} is — and who's calling them"),
                 narr.get("subtitle",
                          f"{r.total_missed:,} missed calls in {r.reporting_period} · "
                          f"these are the callers a business like this hears from every day"))

    # Business profile band — stack the three pieces vertically (industry,
    # lines of business, one-line summary), each capped to a single non-wrapping
    # row. The old side-by-side layout collided when the summary wrapped over the
    # lines-of-business list.
    by = Inches(1.8)
    _rect(s, Inches(0.5), by, Inches(12.33), Inches(1.05), LIGHT, line=CARD_BORDER, radius=True)
    industry = (biz.get("industry") or "").strip()
    lobs = biz.get("lines_of_business") or []
    head = industry if industry else "Business profile"
    _text(s, head, Inches(0.8), by + Inches(0.12), Inches(11.6), Inches(0.32),
          size=13, bold=True, color=RC_BLUE, font=FONT, wrap=False)
    if lobs:
        lob_line = " · ".join(str(x) for x in lobs[:6])
        if len(lob_line) > 140:
            lob_line = lob_line[:138].rstrip(" ·") + "…"
        _text(s, lob_line, Inches(0.8), by + Inches(0.46), Inches(11.7), Inches(0.3),
              size=9.5, color=GRAY, font=FONT, wrap=False)
    if summary:
        s_line = summary if len(summary) <= 150 else summary[:148].rstrip() + "…"
        _text(s, s_line, Inches(0.8), by + Inches(0.74), Inches(11.7), Inches(0.28),
              size=9.5, italic=True, color=DARK, font=FONT, wrap=False)

    # Predicted call-reason cards (up to 6)
    reasons = (biz.get("predicted_call_reasons") or [])[:6]
    cw, chh = Inches(4.05), Inches(1.55)
    xs = [Inches(0.5), Inches(4.68), Inches(8.86)]
    ys = [Inches(3.1), Inches(4.85)]
    for idx, item in enumerate(reasons):
        cx = xs[idx % 3]; cy = ys[idx // 3]
        tier = str(item.get("tier", "C")).upper()[:1]
        col, bg, tlabel = _TIER_BADGE.get(tier, _TIER_BADGE["C"])
        _rect(s, cx, cy, cw, chh, WHITE, line=CARD_BORDER, radius=True)
        _rect(s, cx, cy, Inches(0.12), chh, col)
        _text(s, str(item.get("reason", ""))[:48], cx + Inches(0.32), cy + Inches(0.18),
              cw - Inches(1.6), Inches(0.55), size=14, bold=True, color=RC_BLUE, font=FONT)
        # tier badge
        bw = Inches(1.15)
        _rect(s, cx + cw - bw - Inches(0.18), cy + Inches(0.2), bw, Inches(0.34), bg, radius=True)
        _text(s, tlabel, cx + cw - bw - Inches(0.18), cy + Inches(0.225), bw, Inches(0.3),
              size=9, bold=True, color=col, align=PP_ALIGN.CENTER, font=FONT)
        _text(s, str(item.get("why", ""))[:120], cx + Inches(0.32), cy + Inches(0.8),
              cw - Inches(0.55), Inches(0.7), size=11, color=GRAY, line_spacing=1.05)
        # Only badge revenue-relevant on genuine revenue-line cards (tier A/B), so the
        # marker stays meaningful instead of appearing on General/Internal reasons too.
        if item.get("revenue_relevant") and tier in ("A", "B"):
            _text(s, "● revenue-relevant", cx + Inches(0.32), cy + chh - Inches(0.32),
                  cw - Inches(0.55), Inches(0.28), size=9, bold=True, color=RC_ORANGE)

    _text(s, "Caller types inferred from public website content — illustrative, not a claim about your call logs. "
             "The point: these are routine, answerable calls — the same kind that show up later in your RingCentral "
             "data as abandoned and sub-60-second calls.",
          Inches(0.5), Inches(6.7), Inches(12.33), Inches(0.5),
          size=10, italic=True, color=GRAY, align=PP_ALIGN.CENTER)
    _footer(s)


# ---------------------------------------------------------------------------
# ROI model (assumptions are clearly labelled on the slides)
# ---------------------------------------------------------------------------

# AIR usage pricing
AIR_RATE_PER_MIN = 0.20
# Staffing-equivalent productivity (calls handled per agent per year)
CALLS_PER_FTE_EFFICIENT = 41_000   # busy, well-run inbound desk -> fewer FTEs
CALLS_PER_FTE_LEAN = 23_000        # realistic everyday throughput -> more FTEs
COST_PER_FTE = 50_000              # fully-loaded annual cost per agent
# Revenue assumptions
AVG_ORDER_VALUE = 500              # $ per recovered order (editable assumption)
# Deliberately conservative capture rates — the share of revenue-line missed calls
# an AI receptionist converts into a booked order. Kept low so the model is a floor,
# not a hero number a CFO dismisses.
CAPTURE_RATES = [0.02, 0.03, 0.05]
CONSERVATIVE_CAPTURE = 0.02        # the cell we lead with / highlight


def _roi_model(r: PipelineResult, air_rate: float = AIR_RATE_PER_MIN,
               rev_missed_month: float | None = None) -> dict:
    days = r.days_in_period or 30
    per_day_missed = r.total_missed / days
    missed_per_year = per_day_missed * 365
    missed_per_month = per_day_missed * 30.4
    inbound_per_month = (r.universe_sessions / days) * 30.4

    # Conservative opportunity pool: ONLY revenue-line missed calls (sales/orders/
    # bookings) — never the full missed-call count. If we weren't handed a pool,
    # fall back to the observed revenue-line monthly figure, else the total.
    if rev_missed_month and rev_missed_month > 0:
        rev_missed_per_month = float(rev_missed_month)
    else:
        rev_missed_per_month = missed_per_month
    rev_missed_per_year = rev_missed_per_month * 12

    # AIR fields every inbound call; minutes = calls x avg talk minutes
    air_minutes_month = inbound_per_month * r.avg_answered_minutes
    air_cost_month = air_minutes_month * air_rate
    air_cost_year = air_cost_month * 12

    # Staffing to answer every inbound call across all hours (the AIR equivalent).
    # Sized on full inbound volume, not just the missed subset, so the comparison
    # is like-for-like ("answer every call, 24/7").
    inbound_per_year = (r.universe_sessions / days) * 365
    fte_lo = inbound_per_year / CALLS_PER_FTE_EFFICIENT
    fte_hi = inbound_per_year / CALLS_PER_FTE_LEAN
    hire_lo = fte_lo * COST_PER_FTE
    hire_hi = fte_hi * COST_PER_FTE

    return {
        "missed_per_year": missed_per_year,
        "missed_per_month": missed_per_month,
        "rev_missed_per_year": rev_missed_per_year,
        "rev_missed_per_month": rev_missed_per_month,
        "inbound_per_month": inbound_per_month,
        "air_minutes_month": air_minutes_month,
        "air_cost_month": air_cost_month,
        "air_cost_year": air_cost_year,
        "fte_lo": fte_lo,
        "fte_hi": fte_hi,
        "hire_lo": hire_lo,
        "hire_hi": hire_hi,
        "air_rate": air_rate,
    }


def _money(v) -> str:
    v = float(v)
    if abs(v) >= 1_000_000:
        return f"${v/1_000_000:.1f}M"
    if abs(v) >= 1_000:
        return f"${round(v/1_000)}K"
    return f"${round(v)}"


# ---------------------------------------------------------------------------
# Slide 6 — AIR capability mapping
# ---------------------------------------------------------------------------

# ---------------------------------------------------------------------------
# Slide 7 — far cheaper than hiring
# ---------------------------------------------------------------------------

# ---------------------------------------------------------------------------
# Slide 8 — recovered revenue
# ---------------------------------------------------------------------------

def _slide_revenue(prs, r: PipelineResult, ctx, narr):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _bg(s)
    _logo(s)
    m = ctx["roi"]
    aov = ctx.get("aov", AVG_ORDER_VALUE)
    cap_hi = ctx.get("capture_override") or CONSERVATIVE_CAPTURE
    rev_year = m["rev_missed_per_year"]
    rev_month = m["rev_missed_per_month"]
    _title_block(s, narr.get("title", "Sizing the revenue opportunity — conservatively"),
                 narr.get("subtitle",
                          f"Revenue-line missed calls only (sales · orders · bookings) · "
                          f"{ctx.get('reporting_period','')} run-rate · illustrative, using your own order value"))

    # Funnel strip — pool is the REVENUE-LINE missed calls, never total volume.
    fy = Inches(1.95)
    steps = [
        (f"{round(rev_month):,}", "revenue-line missed / month", RC_RED),
        (f"{round(rev_year):,}", "revenue-line missed / year", RC_GOLD),
        ("× capture %", "booked as orders", RC_BLUE),
        ("= recovered $", "added revenue", RC_ORANGE),
    ]
    sw = Inches(3.0); sx = Inches(0.5)
    for big, lab, col in steps:
        _rect(s, sx, fy, sw, Inches(0.95), WHITE, line=CARD_BORDER, radius=True)
        _text(s, big, sx + Inches(0.2), fy + Inches(0.12), sw - Inches(0.4), Inches(0.5),
              size=24, bold=True, color=col, font=FONT)
        _text(s, lab, sx + Inches(0.2), fy + Inches(0.62), sw - Inches(0.4), Inches(0.3),
              size=11, color=GRAY)
        sx = Emu(int(sx) + int(sw) + int(Inches(0.11)))

    # Recovered-revenue table by capture rate (revenue-line pool only)
    _text(s, "Annual recovered revenue — revenue-line missed calls only, by capture rate",
          Inches(0.5), Inches(3.25), Inches(12.3), Inches(0.35), size=14, bold=True, color=DARK, font=FONT)
    _revenue_table(s, m, aov, cap_hi, Inches(0.5), Inches(3.7), Inches(12.33))

    # Headline takeaway — lead with the CONSERVATIVE cell, framed as a floor.
    rec_lo = cap_hi * rev_year * aov
    _rect(s, Inches(0.5), Inches(6.05), Inches(12.33), Inches(0.62), RC_BLUE, radius=True)
    _rich(s, [[
        (f"Even at a conservative {round(cap_hi*100)}% capture and a ${aov:,} order, that's ", {"size": 13, "color": RGBColor(0xCD,0xD9,0xEA)}),
        (f"{_money(rec_lo)}/year recovered ", {"bold": True, "size": 16, "color": WHITE, "font": FONT}),
        ("— set your own order value to size it precisely.", {"size": 13, "color": RGBColor(0xCD,0xD9,0xEA)}),
    ]], Inches(0.7), Inches(6.16), Inches(12.0), Inches(0.4), anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
    # Explicit caveat line so no exec can say it was hidden.
    _text(s, f"Illustrative. Counts only missed calls to revenue-line queues — not total call volume. "
             f"{ctx.get('reporting_period','')} run-rate ×12; assumes your average order value. Validate against your CRM.",
          Inches(0.5), Inches(6.75), Inches(12.33), Inches(0.4),
          size=9, italic=True, color=GRAY, align=PP_ALIGN.CENTER)
    _footer(s, 8)


def _revenue_table(s, m, aov, cap_hi, x, y, w):
    aovs = [max(50, round(aov / 2)), aov, aov * 2]
    rows = len(CAPTURE_RATES) + 1
    cols = len(aovs) + 1
    row_in = 0.55
    tbl_shape = s.shapes.add_table(rows, cols, x, y, w, Inches(row_in * rows))
    tbl = tbl_shape.table
    tbl.first_row = False
    tbl.horz_banding = False
    tbl.columns[0].width = Inches(3.33)
    for j in range(1, cols):
        tbl.columns[j].width = Inches(3.0)
    for i in range(rows):
        tbl.rows[i].height = Inches(row_in)
    # header
    _cell(tbl.cell(0, 0), "Capture rate  ╲  Avg order value", WHITE, bold=True, size=11)
    tbl.cell(0, 0).fill.solid(); tbl.cell(0, 0).fill.fore_color.rgb = RC_NAVY
    for j, col_aov in enumerate(aovs, 1):
        c = tbl.cell(0, j)
        c.fill.solid(); c.fill.fore_color.rgb = RC_NAVY
        _cell(c, f"${col_aov:,}/order", WHITE, bold=True, size=11, align=PP_ALIGN.CENTER)
    for i, cap in enumerate(CAPTURE_RATES, 1):
        c0 = tbl.cell(i, 0)
        c0.fill.solid(); c0.fill.fore_color.rgb = ROW_ALT if i % 2 else WHITE
        _cell(c0, f"{round(cap*100)}% of missed calls", DARK, bold=True, size=12)
        for j, col_aov in enumerate(aovs, 1):
            rec = cap * m["rev_missed_per_year"] * col_aov
            c = tbl.cell(i, j)
            highlight = (abs(cap - cap_hi) < 1e-9 and col_aov == aov)
            c.fill.solid()
            c.fill.fore_color.rgb = RGBColor(0xFF, 0xF1, 0xE3) if highlight else (ROW_ALT if i % 2 else WHITE)
            _cell(c, _money(rec), RC_ORANGE if highlight else DARK,
                  bold=highlight, size=13 if highlight else 12, align=PP_ALIGN.CENTER)


# ---------------------------------------------------------------------------
# Slide 9 — rollout investment
# ---------------------------------------------------------------------------

# ---------------------------------------------------------------------------
# Slide 10 — recommendation & next steps
# ---------------------------------------------------------------------------

def _slide_next(prs, r: PipelineResult, ctx, narr):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _bg(s, warm=True)
    _WARM_SLIDES.add(len(prs.slides))   # white footer + page number on this slide
    _logo(s, warm=True)
    m = ctx["roi"]
    _title_block(s, narr.get("title", "Recommendation & next steps"),
                 narr.get("subtitle", "Turn the missed-call gap into recovered revenue — starting with a focused pilot"),
                 warm=True)

    steps = [
        ("Confirm the numbers", "Review this analysis with your operations team and validate the queues, volumes, and order value against your own reporting."),
        ("Pilot on the worst gap", "Stand up AIR on the highest-miss queues and the after-hours window first — fastest, most visible recovery."),
        ("Measure recovered calls", "Track answered-vs-missed and captured leads for 30–60 days against the baseline in this deck."),
        ("Roll out companywide", "Expand AIR across every location once the pilot proves capture — with free implementation and 4 months free."),
    ]
    y = Inches(2.05); rh = Inches(1.1); gap = Inches(0.12)
    for i, (title, body) in enumerate(steps, 1):
        _rect(s, Inches(0.5), y, Inches(12.33), rh, WHITE, line=CARD_BORDER, radius=True)
        _circle(s, Inches(0.78), y + Inches(0.3), Inches(0.5), RC_ORANGE, str(i), size=18)
        _text(s, title, Inches(1.6), y + Inches(0.16), Inches(11.0), Inches(0.4),
              size=16, bold=True, color=RC_BLUE, font=FONT)
        _text(s, body, Inches(1.6), y + Inches(0.56), Inches(10.9), Inches(0.5),
              size=12, color=GRAY, line_spacing=1.05)
        y = Emu(int(y) + int(rh) + int(gap))

    cap_hi = ctx.get("capture_override") or CONSERVATIVE_CAPTURE
    rec_lo = cap_hi * m["rev_missed_per_year"] * ctx.get("aov", AVG_ORDER_VALUE)
    _text(s, f"Even at a conservative {round(cap_hi*100)}% capture on revenue-line missed calls, "
             f"that's ~{_money(rec_lo)}/year in recovered revenue — validated against your own order value.",
          Inches(0.5), Inches(6.5), Inches(12.33), Inches(0.4),
          size=13, bold=True, italic=True, color=WHITE, align=PP_ALIGN.CENTER, font=FONT)
    _footer(s, 10, warm=True)


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
    r.font.color.rgb = color; r.font.name = FONT


def _pct(part, whole):
    return f"{round(100*part/whole)}%" if whole else "0%"


def _hour_label(h):
    suffix = "a" if h < 12 else "p"
    hh = h if h <= 12 else h - 12
    return f"{hh}{suffix}"


def _hour_label24(h):
    if h == 0:
        return "12a"
    if h < 12:
        return f"{h}a"
    if h == 12:
        return "12p"
    return f"{h - 12}p"


def _range_or_pct(lo, hi, agg):
    """Show a "X–Y%" range when both endpoints are meaningful and the spread is
    wide; otherwise fall back to the volume-weighted aggregate "~Z%"."""
    lo_p, hi_p, agg_p = round(lo * 100), round(hi * 100), round(agg * 100)
    if hi_p - lo_p >= 8 and lo_p > 0:
        return f"{lo_p}–{hi_p}%"
    return f"~{agg_p}%"


# ---------------------------------------------------------------------------
# orchestration
# ---------------------------------------------------------------------------

def build_deck(result: PipelineResult, run_id: str, customer: str, ae_name: str,
               prior_instructions: list[dict] | None = None,
               business_context: dict | None = None,
               overrides: dict | None = None) -> Path:
    out_dir = Path(tempfile.gettempdir()) / "rc_analyzer_decks"
    out_dir.mkdir(exist_ok=True)
    out_path = out_dir / f"{run_id}.pptx"
    overrides = overrides or {}

    # Revenue-line = queues with clear buy/order/booking or customer-service intent
    # (the old A + B tiers). These are the missed calls that most directly cost money.
    #
    # Tier labels live on the Queues report (the authoritative per-queue feed), NOT
    # on queue_stats (built from the Calls export, whose Queue column is mostly
    # blank — so tiering it yields ~0). Source the revenue-line pool from the
    # Queues report: missed = inbound − answered for A/B queues. Fall back to
    # queue_stats only when no Queues report is present.
    _qr0 = result.queues_report
    if _qr0 and _qr0.queues:
        sales_queue_calls = sum(max((q.inbound or 0) - (q.answered or 0), 0)
                                for q in _qr0.queues if (q.tier or "C") in ("A", "B"))
    else:
        sales_queue_calls = sum(q.total_missed for q in result.queue_stats.values()
                                if (q.tier or "C") in ("A", "B"))
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

    # Non-queue traffic: calls that never entered a managed ACD call queue land in
    # the "Unknown" bucket (direct extension dials, ring/hunt groups, IVR). When this
    # dominates, the story is "your misses never reach a queue at all" — the single
    # strongest AIR argument: routing/SLA/overflow staffing cannot catch a call that
    # never enters a queue; only an always-on AI receptionist answers every line.
    # Real abandoned-in-queue data from the Queues report (2nd upload)
    qr = result.queues_report
    if qr and qr.inbound:
        ctx["queue_inbound"] = qr.inbound
        ctx["queue_abandoned"] = qr.abandoned
        ctx["queue_abandon_rate_pct"] = round(qr.abandon_rate * 100, 1)
        ctx["queue_longest_wait"] = qr.longest_wait
        ctx["queue_sla_pct"] = round(qr.sla_pct * 100)

    # How many misses actually entered a managed call queue vs. never reached one.
    #
    # We can't trust the Calls export's "Queue" column for this — it's blank on the
    # vast majority of rows, so the "Unknown" bucket balloons and falsely implies
    # ~100% of misses are unqueued. The Queues report is the authoritative source
    # for queue activity, so reconcile against it: queue misses = inbound − answered
    # for customer-facing queues (tier != 'D'; back-office queues carry no customer
    # and are out of scope). Everything else in total_missed never reached a queue.
    if result.total_missed:
        queue_missed = None
        if qr and qr.queues:
            tiered = [q for q in qr.queues if (getattr(q, "tier", "") or "").upper() != "D"]
            if tiered:
                queue_missed = sum(max((q.inbound or 0) - (q.answered or 0), 0) for q in tiered)
        if queue_missed is None and qr and qr.inbound:
            queue_missed = max(qr.inbound - qr.answered, 0)
        if queue_missed is None:
            # No Queues report at all — fall back to the (unreliable) Unknown bucket.
            _unq = result.queue_stats.get("Unknown")
            unqueued_missed = _unq.total_missed if _unq is not None else result.total_missed
        else:
            unqueued_missed = max(result.total_missed - queue_missed, 0)
            ctx["queue_missed"] = queue_missed
        ctx["unqueued_missed"] = unqueued_missed
        ctx["unqueued_share_pct"] = round(100 * unqueued_missed / result.total_missed)

    # Business context + modeling overrides
    business = business_context if (business_context and business_context.get("available")) else None
    ctx["business"] = business

    def _num(v):
        try:
            return float(v) if v is not None else None
        except (TypeError, ValueError):
            return None

    aov = _num(overrides.get("avg_order_value"))
    if aov is None and business:
        aov = _num(business.get("suggested_avg_order_value"))
    ctx["aov"] = int(aov) if aov and aov > 0 else AVG_ORDER_VALUE

    air_rate = _num(overrides.get("air_rate_per_min"))
    ctx["air_rate"] = air_rate if air_rate and air_rate > 0 else AIR_RATE_PER_MIN

    cap = _num(overrides.get("capture_rate"))
    ctx["capture_override"] = cap if cap and 0 < cap < 1 else None

    # Scale the observed revenue-line missed pool to a 30.4-day month so it lines
    # up with the model's monthly basis (the report period may not be exactly 30 days).
    _rev_missed_month = sales_queue_calls * (30.4 / (result.days_in_period or 30))
    ctx["roi"] = _roi_model(result, ctx["air_rate"], rev_missed_month=_rev_missed_month)
    ctx["sales_queue_missed"] = sales_queue_calls

    # Narrative titles/subtitles for the slides we keep. The deck was deliberately
    # simplified to a tight, defensible sequence (see the build order below):
    # every slide after the business-context opener is verifiable straight from
    # the RingCentral reports — nothing modeled except the one caveated slide.
    narr2 = _narr_titles(ctx, prior_instructions, "slide2")
    narr_hourly = {"title": "Calls slip away after hours, on weekends — and even midday",
                   "subtitle": f"Inbound miss rate by hour of day · {result.reporting_period} · when the business closes or gets busy, calls go unanswered"}
    if qr and qr.inbound:
        narr3_sub = (f"{qr.abandoned:,} of {qr.inbound:,} queue callers abandoned "
                     f"({qr.abandon_rate*100:.0f}%) · short routine calls · {result.reporting_period}")
    else:
        narr3_sub = f"Abandoned-in-queue callers and short routine calls · customer-facing queues · {result.reporting_period}"
    narr3 = {"title": "Where AI Receptionist captures revenue today",
             "subtitle": narr3_sub}

    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    _WARM_SLIDES.clear()   # reset per-build (module-level set)

    # Simplified, defensible 6-slide flow:
    #   1. Who they are + who's calling them (business context up front)
    #   2. Missed calls by queue/company — the headline number + breakdown
    #   3. Hourly trend + after-hours / weekend / midday miss rates
    #   4. Answerable-call signals: abandons, answered <60s, wait/SLA
    #   5. What queue config can fix vs. the structural gap only AIR closes —
    #      the answer to the "just help us configure our queues" objection
    #   6. Illustrative opportunity (single, clearly-caveated money slide)
    #   7. Recommendation & next steps
    _slide_call_reasons(prs, result, ctx, {})   # opener — robust to missing business context
    _slide2(prs, result, ctx, narr2, sales_queue_calls)
    _slide_hourly(prs, result, ctx, narr_hourly)
    _slide3(prs, result, ctx, narr3)
    _slide_config_vs_air(prs, result, ctx, {})
    _slide_revenue(prs, result, ctx, {})
    _slide_next(prs, result, ctx, {})

    _stamp_page_numbers(prs)
    prs.save(out_path)
    return out_path


def _narr_titles(ctx, prior, which):
    schema = {"title": "str — slide headline", "subtitle": "str — methodology source line"}
    try:
        return generate_narrative({**ctx, "slide": which}, schema, prior)
    except Exception:
        if which == "slide2":
            return {"title": f"{ctx['total_missed']:,} genuine missed calls across customer-facing queues",
                    "subtitle": f"Customer-facing queues · {ctx['reporting_period']} · session-deduplicated · spam-filtered · internal/back-office excluded"}
        return {"title": "Where customers are getting missed — by queue",
                "subtitle": f"Session-deduplicated · spam-filtered · {ctx['reporting_period']} · customer-facing queues only · ranked by calls lost"}

