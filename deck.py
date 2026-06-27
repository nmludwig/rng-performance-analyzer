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

def _slide1(prs, r: PipelineResult, ctx, narr):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _bg(s)
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
              size=15, bold=True, color=color, font=FONT)
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
              size=22, bold=True, color=WHITE, align=PP_ALIGN.CENTER, font=FONT)
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
    _bg(s)
    _logo(s)
    _title_block(s, narr.get("title", f"{r.total_missed:,} genuine missed calls across sales-taking queues"),
                 narr.get("subtitle", ""))

    # Left big-stat card
    lx, ly, lw, lh = Inches(0.5), Inches(1.95), Inches(3.25), Inches(4.7)
    _rect(s, lx, ly, lw, lh, WHITE, line=CARD_BORDER, radius=True)
    _text(s, "GENUINE MISSED\nCALLS — A+B+C QUEUES", lx + Inches(0.25), ly + Inches(0.55),
          lw - Inches(0.5), Inches(0.7), size=12, bold=True, color=GRAY, font=FONT)
    # Size the headline number to the digit count so 5- and 6-figure totals
    # both fit on ONE line in the ~2.95in card (72pt overflowed and wrapped).
    _big_n = f"{r.total_missed:,}"
    # Tier the headline size to digit-count so it fits on ONE line even when a
    # WIDER fallback font is substituted (Inter Tight may be absent on the
    # rendering machine). Sizes chosen to fit the ~2.95in card with margin.
    _nlen = len(_big_n)
    _big_size = 72 if _nlen <= 5 else 50 if _nlen == 6 else 42
    _text(s, _big_n, lx + Inches(0.1), ly + Inches(1.3), lw - Inches(0.2), Inches(1.2),
          size=_big_size, bold=True, color=RC_RED, font=FONT, wrap=False,
          align=PP_ALIGN.CENTER)
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
          mx, Inches(1.95), Inches(4.6), Inches(0.35), size=12, bold=True, color=DARK, font=FONT)
    _hourly_chart(s, r, mx, Inches(2.35), Inches(4.55), Inches(1.85))

    _text(s, "How they were missed:", mx, Inches(4.35), Inches(4.6), Inches(0.3),
          size=12, bold=True, color=DARK, font=FONT)
    # Real abandoned count comes from the Queues report (2nd upload); the Calls
    # export has no "Abandoned" disposition so r.abandoned is always 0 there.
    qr = r.queues_report
    abandoned_n = qr.abandoned if (qr and qr.abandoned) else r.abandoned
    bd = [
        (abandoned_n, "Abandoned in queue (left waiting)", RC_RED, RGBColor(0xFC,0xEC,0xEC)),
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
    headers = ["Queue", "T", "Aband", "%"]
    for j, htext in enumerate(headers):
        c = tbl.cell(0, j)
        c.fill.solid(); c.fill.fore_color.rgb = RC_NAVY
        _cell(c, htext, WHITE, bold=True, size=9, align=PP_ALIGN.CENTER if j else PP_ALIGN.LEFT)
    for i, q in enumerate(queues, 1):
        bg = ROW_ALT if i % 2 else WHITE
        vals = [q.name[:26], q.tier or "C", f"{q.abandoned}", f"{round(q.abandon_rate*100)}%"]
        for j, v in enumerate(vals):
            c = tbl.cell(i, j)
            c.fill.solid(); c.fill.fore_color.rgb = bg
            col = RC_RED if (j == 3 and q.abandon_rate >= 0.7) else DARK
            _cell(c, v, col, bold=(j == 3 and q.abandon_rate >= 0.7), size=8.5,
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
               (" of answered calls — short, routine calls AIR can handle", {"size": 12, "color": GRAY})]],
          lx + Inches(0.25), by + Inches(1.78), lw - Inches(0.5), Inches(0.4))

    # Right — most-abandoned queues table (from the Queues report when present)
    rx = Inches(5.0)
    _text(s, "Where callers give up — most-abandoned queues", rx, Inches(1.95),
          Inches(7.8), Inches(0.35), size=14, bold=True, color=DARK, font=FONT)

    if qr and qr.queues:
        top_ab = sorted((q for q in qr.queues if q.abandoned > 0 and q.tier != "D"),
                        key=lambda q: q.abandoned, reverse=True)[:11]
        _abandon_table_qr(s, top_ab, rx, Inches(2.4), Inches(7.85))
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
    _text(s, "AI Receptionist answers instantly — recovering abandoned callers and "
             "deflecting short, routine calls so staff focus on revenue conversations.",
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
    headers = ["Queue", "Tier", "Inbound", "Abandoned", "Ab. %"]
    for j, htext in enumerate(headers):
        c = tbl.cell(0, j)
        c.fill.solid(); c.fill.fore_color.rgb = RC_NAVY
        _cell(c, htext, WHITE, bold=True, size=9,
              align=PP_ALIGN.LEFT if j == 0 else PP_ALIGN.CENTER)
    tier_col = {"A": RC_RED, "B": RC_GOLD, "C": RC_BLUE}
    for i, q in enumerate(queues, 1):
        bg = ROW_ALT if i % 2 else WHITE
        vals = [q.name[:38], q.tier or "C", f"{q.inbound:,}",
                f"{q.abandoned:,}", f"{round(q.abandon_rate*100)}%"]
        for j, v in enumerate(vals):
            c = tbl.cell(i, j)
            c.fill.solid(); c.fill.fore_color.rgb = bg
            if j == 1:
                col = tier_col.get(q.tier, DARK); bold = True
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

def _slide4(prs, r: PipelineResult, ctx, narr):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _bg(s)
    _logo(s)

    # The Calls export only stamps a queue name on ANSWERED legs — abandoned /
    # missed calls arrive with a blank Queue and collapse into "Unknown", which
    # makes every sales/customer-facing queue look like it had 0 missed calls.
    # The Queues report (2nd upload) is the only source with real per-queue
    # abandoned counts, so we build this slide from it whenever it's present.
    qr = r.queues_report
    use_qr = bool(qr and qr.queues)

    _title_block(s, narr.get("title", "Queue-level missed call analysis — sales-taking queues (Tier A+B+C)"),
                 narr.get("subtitle",
                          "Per-queue abandoned calls from the RingCentral Queues report — "
                          "the only source that attributes missed calls to the queue they were waiting in."
                          if use_qr else ""))

    # Tier summary cards
    tier_meta = [
        ("A", "Sales / revenue line", RC_RED, RGBColor(0xFC,0xEC,0xEC)),
        ("B", "Customer-facing front line", RC_GOLD, RGBColor(0xFB,0xF3,0xE0)),
        ("C", "Main line / reception", RC_BLUE, RGBColor(0xE9,0xEE,0xF6)),
    ]
    cw = Inches(4.05)
    cxs = [Inches(0.5), Inches(4.68), Inches(8.86)]
    for (tier, label, col, bg), cx in zip(tier_meta, cxs):
        if use_qr:
            qs = [q for q in qr.queues if (q.tier or "C") == tier]
            inb = sum(q.inbound for q in qs); mis = sum(q.abandoned for q in qs)
        else:
            qs = [q for q in r.queue_stats.values() if q.tier == tier]
            inb = sum(q.inbound for q in qs); mis = sum(q.total_missed for q in qs)
        nq = len(qs)
        mr = mis / inb if inb else 0
        _rect(s, cx, Inches(1.85), cw, Inches(0.95), bg, line=col, line_w=Pt(1), radius=True)
        _circle(s, cx + Inches(0.18), Inches(2.02), Inches(0.42), col, tier, size=15)
        _text(s, label, cx + Inches(0.72), Inches(2.0), cw - Inches(0.85), Inches(0.35),
              size=13, bold=True, color=col, font=FONT)
        _rich(s, [[(f"{nq} ", {"bold": True, "size": 12, "color": DARK}),
                   ("queues   ", {"size": 11, "color": GRAY}),
                   (f"{inb:,} ", {"bold": True, "size": 12, "color": DARK}),
                   ("inbound", {"size": 11, "color": GRAY})]],
              cx + Inches(0.72), Inches(2.32), cw - Inches(0.8), Inches(0.25))
        _rich(s, [[(f"{mis:,} ", {"bold": True, "size": 12, "color": col}),
                   ("abandoned   " if use_qr else "missed   ", {"size": 11, "color": GRAY}),
                   (f"{mr*100:.1f}%", {"bold": True, "size": 12, "color": col})]],
              cx + Inches(0.72), Inches(2.55), cw - Inches(0.8), Inches(0.25))

    # Full queue table
    order = {"A": 0, "B": 1, "C": 2, "D": 3}
    if use_qr:
        queues = sorted((q for q in qr.queues if (q.tier or "C") != "D"),
                        key=lambda q: (order.get(q.tier or "C", 9), -q.abandoned))
        _full_table_qr(s, queues, Inches(0.5), Inches(2.92), Inches(12.33))
    else:
        queues = sorted(r.queue_stats.values(),
                        key=lambda q: (order.get(q.tier, 9), -q.total_missed))
        _full_table(s, queues, Inches(0.5), Inches(2.92), Inches(12.33))
    _footer(s, 5)


def _full_table_qr(s, queues, x, y, w):
    """Queue-level table sourced from the Queues report (real per-queue abandoned)."""
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
    headers = ["Queue Name", "Tier", "Classification", "Inbound", "Answered", "Abandoned", "Ab. %"]
    for j, htext in enumerate(headers):
        c = tbl.cell(0, j)
        c.fill.solid(); c.fill.fore_color.rgb = RC_NAVY
        _cell(c, htext, WHITE, bold=True, size=8.5, pad=0,
              align=PP_ALIGN.LEFT if j in (0, 2) else PP_ALIGN.CENTER)
    tier_col = {"A": RC_RED, "B": RC_GOLD, "C": RC_BLUE}
    for i, q in enumerate(queues, 1):
        bg = ROW_ALT if i % 2 else WHITE
        tier = q.tier or "C"
        # A queue with inbound calls but 0 answered is unstaffed/misconfigured —
        # that's a routing/staffing fix (Professional Services), NOT a license gap.
        unstaffed = q.answered == 0 and q.inbound > 0
        classification = (q.classification or "")[:38]
        if unstaffed:
            classification = "UNSTAFFED — 0 agents answered"
        vals = [q.name[:46], tier, classification,
                f"{q.inbound:,}", f"{q.answered:,}", f"{q.abandoned:,}",
                f"{round(q.abandon_rate*100)}%"]
        for j, v in enumerate(vals):
            c = tbl.cell(i, j)
            c.fill.solid(); c.fill.fore_color.rgb = bg
            if j == 1:
                col = tier_col.get(tier, DARK); bold = True
            elif j == 2 and unstaffed:
                col = RC_GOLD; bold = True
            elif j == 6:
                col = RC_RED if q.abandon_rate >= 0.5 else DARK
                bold = q.abandon_rate >= 0.5
            else:
                col = DARK; bold = False
            _cell(c, v, col, bold=bold, size=7.5, pad=0,
                  align=PP_ALIGN.LEFT if j in (0, 2) else PP_ALIGN.CENTER)


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
    _title_block(s, narr.get("title", "What queue config can fix — and what only AIR can"),
                 narr.get("subtitle",
                          "Separating calls recoverable by routing/staffing from the structural "
                          "gap that needs always-on coverage"))

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
         "TOTAL GENUINE MISSED", f"{total:,}", "A+B+C queues · spam-filtered · de-duplicated",
         "Every inbound call that never reached a person."),
        (Inches(4.78), Inches(3.55), RGBColor(0xFB,0xF3,0xE0), RC_GOLD, RC_GOLD, GRAY,
         "CONFIG-FIXABLE (PRO SERVICES)", f"{config_fixable:,}",
         f"{len(unstaffed)} unstaffed queue{'s' if len(unstaffed)!=1 else ''} · 0 agents ever answered",
         "Routing/staffing fix — no new licenses required."),
        (Inches(9.06), Inches(3.77), RGBColor(0xFC,0xEC,0xEC), RC_RED, RC_RED, GRAY,
         "STRUCTURAL GAP — ONLY AIR CLOSES", f"{structural:,}",
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
    "A": (RC_RED, RGBColor(0xFC, 0xEC, 0xEC), "Sales"),
    "B": (RC_GOLD, RGBColor(0xFB, 0xF3, 0xE0), "Customer-facing"),
    "C": (RC_BLUE, RGBColor(0xE9, 0xEE, 0xF6), "Main line"),
    "D": (GRAY, LIGHT, "Back-office"),
}


def _slide_call_reasons(prs, r: PipelineResult, ctx, narr):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _bg(s)
    _logo(s)
    biz = ctx.get("business") or {}
    summary = (biz.get("summary") or "").strip()
    url = biz.get("url", "")
    _title_block(s, narr.get("title", "What your callers are likely calling about"),
                 narr.get("subtitle", f"Predicted from {url} · mapped to your queues — so coverage gaps mean lost business"))

    # Business profile band
    by = Inches(1.85)
    _rect(s, Inches(0.5), by, Inches(12.33), Inches(1.0), LIGHT, line=CARD_BORDER, radius=True)
    industry = (biz.get("industry") or "").strip()
    lobs = biz.get("lines_of_business") or []
    head = industry if industry else "Business profile"
    _text(s, head, Inches(0.8), by + Inches(0.16), Inches(4.0), Inches(0.35),
          size=14, bold=True, color=RC_BLUE, font=FONT)
    if lobs:
        _text(s, " · ".join(str(x) for x in lobs[:6]), Inches(0.8), by + Inches(0.56), Inches(11.4), Inches(0.35),
              size=11, color=GRAY)
    if summary:
        _text(s, summary, Inches(4.9), by + Inches(0.16), Inches(7.7), Inches(0.36),
              size=11, italic=True, color=DARK, line_spacing=1.0)

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
        if item.get("revenue_relevant"):
            _text(s, "● revenue-relevant", cx + Inches(0.32), cy + chh - Inches(0.32),
                  cw - Inches(0.55), Inches(0.28), size=9, bold=True, color=RC_ORANGE)

    _text(s, "Predicted from public website content — confirm with the customer. Every missed call above is a "
             "missed revenue or retention moment AIR would have answered.",
          Inches(0.5), Inches(6.75), Inches(12.33), Inches(0.5),
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
CAPTURE_RATES = [0.02, 0.05, 0.10]


def _roi_model(r: PipelineResult, air_rate: float = AIR_RATE_PER_MIN) -> dict:
    days = r.days_in_period or 30
    per_day_missed = r.total_missed / days
    missed_per_year = per_day_missed * 365
    missed_per_month = per_day_missed * 30.4
    inbound_per_month = (r.universe_sessions / days) * 30.4

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

def _slide_capabilities(prs, r: PipelineResult, ctx, narr):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _bg(s)
    _logo(s)
    _title_block(s, narr.get("title", "AI Receptionist answers every call"),
                 narr.get("subtitle", "Always-on coverage for the exact gaps in your data — no ring-out, no voicemail dead-ends"))

    cards = [
        ("Answers 24/7", "Every call picked up instantly — after hours, weekends, and holidays, the windows where your miss rate is highest."),
        ("Zero hold, zero abandon", "No queue wait means callers stop hanging up — the abandoned-in-queue callers are captured, not lost."),
        ("Handles routine calls", "Hours, locations, order status, basic questions — the short, sub-60-second calls — resolved without a human."),
        ("Qualifies & routes", "Understands intent, captures lead details, and warm-routes revenue conversations to the right queue or rep."),
        ("Captures every lead", "Name, number, and reason logged on every interaction — even when no one is available to take it live."),
        ("Scales with no hiring", "Add unlimited concurrent calls across every location without recruiting, training, or overtime."),
    ]
    cw, chh = Inches(4.05), Inches(1.95)
    gx, gy = Inches(0.5), Inches(0.55)
    xs = [Inches(0.5), Inches(4.68), Inches(8.86)]
    ys = [Inches(2.0), Inches(4.15)]
    for idx, (title, body) in enumerate(cards):
        cx = xs[idx % 3]; cy = ys[idx // 3]
        _rect(s, cx, cy, cw, chh, WHITE, line=CARD_BORDER, radius=True)
        _rect(s, cx, cy, Inches(0.12), chh, RC_ORANGE)
        _text(s, title, cx + Inches(0.32), cy + Inches(0.2), cw - Inches(0.5), Inches(0.4),
              size=16, bold=True, color=RC_BLUE, font=FONT)
        _text(s, body, cx + Inches(0.32), cy + Inches(0.72), cw - Inches(0.55), Inches(1.1),
              size=11.5, color=GRAY, line_spacing=1.1)
    _footer(s, 6)


# ---------------------------------------------------------------------------
# Slide 7 — far cheaper than hiring
# ---------------------------------------------------------------------------

def _slide_cost(prs, r: PipelineResult, ctx, narr):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _bg(s)
    _logo(s)
    m = ctx["roi"]
    _title_block(s, narr.get("title", "Far cheaper than hiring to close the gap"),
                 narr.get("subtitle", "To answer every inbound call across all hours by hand vs. letting AIR field every call"))

    # Left card — hire staff
    lx, lw = Inches(0.5), Inches(5.9)
    ly, lh = Inches(2.05), Inches(4.1)
    _rect(s, lx, ly, lw, lh, RGBColor(0xFC, 0xEC, 0xEC), line=RC_RED, line_w=Pt(1.25), radius=True)
    _text(s, "OPTION A · HIRE STAFF", lx + Inches(0.35), ly + Inches(0.28), lw - Inches(0.7), Inches(0.4),
          size=13, bold=True, color=RC_RED, font=FONT)
    _text(s, f"{_money(m['hire_lo'])}–{_money(m['hire_hi'])}",
          lx + Inches(0.3), ly + Inches(0.78), lw - Inches(0.6), Inches(1.0),
          size=46, bold=True, color=RC_RED, font=FONT)
    _text(s, "per year, fully loaded", lx + Inches(0.35), ly + Inches(1.78), lw - Inches(0.7), Inches(0.35),
          size=13, color=GRAY)
    _rich(s, [
        [(f"{max(1, round(m['fte_lo']))}–{max(2, round(m['fte_hi']))} full-time agents", {"bold": True, "size": 13, "color": DARK}),
         ("  for round-the-clock coverage", {"size": 12, "color": GRAY})],
        [("• Recruiting, training, turnover, overtime", {"size": 12, "color": GRAY})],
        [("• Still no coverage at 2am or on Sundays", {"size": 12, "color": GRAY})],
        [("• Capacity is fixed — spikes still ring out", {"size": 12, "color": GRAY})],
    ], lx + Inches(0.35), ly + Inches(2.35), lw - Inches(0.7), Inches(1.6), line_spacing=1.15, space_after=7)

    # Right card — AIR
    rx = Inches(6.9)
    _rect(s, rx, ly, lw, lh, RGBColor(0xE9, 0xEE, 0xF6), line=RC_BLUE, line_w=Pt(1.25), radius=True)
    _text(s, "OPTION B · AI RECEPTIONIST", rx + Inches(0.35), ly + Inches(0.28), lw - Inches(0.7), Inches(0.4),
          size=13, bold=True, color=RC_BLUE, font=FONT)
    _text(s, f"{_money(m['air_cost_year'])}",
          rx + Inches(0.3), ly + Inches(0.78), lw - Inches(0.6), Inches(1.0),
          size=46, bold=True, color=RC_BLUE, font=FONT)
    _text(s, f"per year (~{_money(m['air_cost_month'])}/mo usage)", rx + Inches(0.35), ly + Inches(1.78), lw - Inches(0.7), Inches(0.35),
          size=13, color=GRAY)
    _rich(s, [
        [("Answers 100% of calls", {"bold": True, "size": 13, "color": DARK}),
         (" — every hour, every day", {"size": 12, "color": GRAY})],
        [("• No recruiting, training, or turnover", {"size": 12, "color": GRAY})],
        [("• Full nights / weekends / holidays coverage", {"size": 12, "color": GRAY})],
        [("• Unlimited concurrent calls, every location", {"size": 12, "color": GRAY})],
    ], rx + Inches(0.35), ly + Inches(2.35), lw - Inches(0.7), Inches(1.6), line_spacing=1.15, space_after=7)

    _text(s, f"Assumes ~{round(m['air_minutes_month']):,} AIR minutes/mo at ${m['air_rate']:.2f}/min "
             f"and ~${COST_PER_FTE/1000:.0f}K fully-loaded per agent. Edit assumptions to fit the account.",
          Inches(0.5), Inches(6.45), Inches(12.33), Inches(0.4),
          size=10, italic=True, color=GRAY, align=PP_ALIGN.CENTER)
    _footer(s, 7)


# ---------------------------------------------------------------------------
# Slide 8 — recovered revenue
# ---------------------------------------------------------------------------

def _slide_revenue(prs, r: PipelineResult, ctx, narr):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _bg(s)
    _logo(s)
    m = ctx["roi"]
    aov = ctx.get("aov", AVG_ORDER_VALUE)
    cap_hi = ctx.get("capture_override") or 0.05
    _title_block(s, narr.get("title", "The ROI: missed calls become recovered orders"),
                 narr.get("subtitle", f"{round(m['missed_per_year']):,} revenue-relevant missed calls/year (spam already removed) · ${aov:,} avg order value"))

    # Funnel strip
    fy = Inches(1.95)
    steps = [
        (f"{round(m['missed_per_year']):,}", "missed calls / year", RC_RED),
        (f"{round(m['missed_per_month']):,}", "missed calls / month", RC_GOLD),
        ("× capture %", "convert to orders", RC_BLUE),
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

    # Recovered-revenue table by capture rate
    _text(s, "Annual recovered revenue by capture rate", Inches(0.5), Inches(3.25),
          Inches(12.3), Inches(0.35), size=14, bold=True, color=DARK, font=FONT)
    _revenue_table(s, m, aov, cap_hi, Inches(0.5), Inches(3.7), Inches(12.33))

    # Headline takeaway (capture rate)
    rec5 = cap_hi * m["missed_per_year"] * aov
    _rect(s, Inches(0.5), Inches(6.25), Inches(12.33), Inches(0.7), RC_BLUE, radius=True)
    _rich(s, [[
        (f"{_money(rec5)} recovered/year ", {"bold": True, "size": 17, "color": WHITE, "font": FONT}),
        (f"at {round(cap_hi*100)}% capture and ${aov:,}/order — against ", {"size": 14, "color": RGBColor(0xCD,0xD9,0xEA)}),
        (f"{_money(m['air_cost_year'])} ", {"bold": True, "size": 17, "color": RC_ORANGE, "font": FONT}),
        ("of AIR investment.", {"size": 14, "color": RGBColor(0xCD,0xD9,0xEA)}),
    ]], Inches(0.7), Inches(6.4), Inches(12.0), Inches(0.45), anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
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
            rec = cap * m["missed_per_year"] * col_aov
            c = tbl.cell(i, j)
            highlight = (abs(cap - cap_hi) < 1e-9 and col_aov == aov)
            c.fill.solid()
            c.fill.fore_color.rgb = RGBColor(0xFF, 0xF1, 0xE3) if highlight else (ROW_ALT if i % 2 else WHITE)
            _cell(c, _money(rec), RC_ORANGE if highlight else DARK,
                  bold=highlight, size=13 if highlight else 12, align=PP_ALIGN.CENTER)


# ---------------------------------------------------------------------------
# Slide 9 — rollout investment
# ---------------------------------------------------------------------------

def _slide_investment(prs, r: PipelineResult, ctx, narr):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _bg(s)
    _logo(s)
    m = ctx["roi"]
    _title_block(s, narr.get("title", "The investment to roll out companywide"),
                 narr.get("subtitle", "Simple usage-based pricing — pay for minutes AIR actually handles, nothing to install"))

    tiles = [
        (f"${m['air_rate']:.2f}", "per AIR minute", RC_BLUE),
        (f"{round(m['air_minutes_month']):,}", "AIR minutes / month", RC_BLUE),
        (f"{_money(m['air_cost_month'])}", "per month", RC_ORANGE),
        (f"{_money(m['air_cost_year'])}", "per year", RC_ORANGE),
    ]
    tw = Inches(3.0); tx = Inches(0.5); ty = Inches(2.15)
    for big, lab, col in tiles:
        _rect(s, tx, ty, tw, Inches(1.75), WHITE, line=CARD_BORDER, radius=True)
        _text(s, big, tx + Inches(0.2), ty + Inches(0.35), tw - Inches(0.4), Inches(0.75),
              size=38, bold=True, color=col, font=FONT, align=PP_ALIGN.CENTER)
        _text(s, lab, tx + Inches(0.2), ty + Inches(1.2), tw - Inches(0.4), Inches(0.4),
              size=13, color=GRAY, align=PP_ALIGN.CENTER)
        tx = Emu(int(tx) + int(tw) + int(Inches(0.11)))

    # Included benefits band
    by = Inches(4.35)
    _rect(s, Inches(0.5), by, Inches(12.33), Inches(1.7), RGBColor(0xE9, 0xEE, 0xF6), line=RC_BLUE, line_w=Pt(1), radius=True)
    _text(s, "Included at no extra cost", Inches(0.8), by + Inches(0.22), Inches(11.7), Inches(0.4),
          size=15, bold=True, color=RC_BLUE, font=FONT)
    _rich(s, [
        [("✓  Free implementation & configuration", {"bold": True, "size": 13, "color": DARK})],
        [("✓  First 4 months of usage free", {"bold": True, "size": 13, "color": DARK})],
    ], Inches(0.8), by + Inches(0.72), Inches(5.8), Inches(0.9), line_spacing=1.2, space_after=6)
    _rich(s, [
        [("✓  No hardware, no on-prem install", {"bold": True, "size": 13, "color": DARK})],
        [("✓  Scales to every location, no hiring", {"bold": True, "size": 13, "color": DARK})],
    ], Inches(6.9), by + Inches(0.72), Inches(5.8), Inches(0.9), line_spacing=1.2, space_after=6)

    _text(s, f"Usage estimate derived from {round(m['inbound_per_month']):,} inbound calls/mo × "
             f"~{r.avg_answered_minutes:.1f} min avg talk time. Final pricing per signed order.",
          Inches(0.5), Inches(6.3), Inches(12.33), Inches(0.4),
          size=10, italic=True, color=GRAY, align=PP_ALIGN.CENTER)
    _footer(s, 9)


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

    cap_hi = ctx.get("capture_override") or 0.05
    rec5 = cap_hi * m["missed_per_year"] * ctx.get("aov", AVG_ORDER_VALUE)
    _text(s, f"The opportunity: ~{_money(rec5)}/year in recovered revenue for ~{_money(m['air_cost_year'])} in AIR investment.",
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

    # Real abandoned-in-queue data from the Queues report (2nd upload)
    qr = result.queues_report
    if qr and qr.inbound:
        ctx["queue_inbound"] = qr.inbound
        ctx["queue_abandoned"] = qr.abandoned
        ctx["queue_abandon_rate_pct"] = round(qr.abandon_rate * 100, 1)
        ctx["queue_longest_wait"] = qr.longest_wait
        ctx["queue_sla_pct"] = round(qr.sla_pct * 100)

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

    ctx["roi"] = _roi_model(result, ctx["air_rate"])

    narr1 = _narr1(ctx, prior_instructions)
    narr2 = _narr_titles(ctx, prior_instructions, "slide2")
    narr_hourly = {"title": "Calls slip away after hours, on weekends — and even midday",
                   "subtitle": f"Inbound miss rate by hour of day · {result.reporting_period} · when the business closes or gets busy, calls go unanswered"}
    if qr and qr.inbound:
        narr3_sub = (f"{qr.abandoned:,} of {qr.inbound:,} queue callers abandoned "
                     f"({qr.abandon_rate*100:.0f}%) · short routine calls · {result.reporting_period}")
    else:
        narr3_sub = f"Abandoned-in-queue callers and short routine calls · Tier A+B+C · {result.reporting_period}"
    narr3 = {"title": "Where AI Receptionist captures revenue today",
             "subtitle": narr3_sub}
    narr4 = {"title": "Queue-level missed call analysis (Tier A+B+C)",
             "subtitle": f"Session-deduplicated · spam-filtered · {result.reporting_period} · back-office (Tier D) excluded · {len(result.queue_stats)} queues shown"}

    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    _WARM_SLIDES.clear()   # reset per-build (module-level set)

    _slide1(prs, result, ctx, narr1)
    _slide2(prs, result, ctx, narr2, sales_queue_calls)
    _slide_hourly(prs, result, ctx, narr_hourly)
    _slide3(prs, result, ctx, narr3)
    _slide4(prs, result, ctx, narr4)
    _slide_config_vs_air(prs, result, ctx, {})
    if business and business.get("predicted_call_reasons"):
        _slide_call_reasons(prs, result, ctx, {})
    _slide_capabilities(prs, result, ctx, {})
    _slide_cost(prs, result, ctx, {})
    _slide_revenue(prs, result, ctx, {})
    _slide_investment(prs, result, ctx, {})
    _slide_next(prs, result, ctx, {})

    _stamp_page_numbers(prs)
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
