"""
IFRS 18 Comprehensive Presentation Generator
Creates a professional PowerPoint presentation covering all key aspects of IFRS 18
(Presentation and Disclosure in Financial Statements).
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.dml import MSO_THEME_COLOR

# ── Color Palette ──────────────────────────────────────────────────────────────
DARK_NAVY    = RGBColor(0x1F, 0x38, 0x64)   # deep navy – primary header/bg
ROYAL_BLUE   = RGBColor(0x1F, 0x5C, 0x99)   # section headings
BRIGHT_BLUE  = RGBColor(0x2E, 0x75, 0xB6)   # accent boxes
LIGHT_BLUE   = RGBColor(0xDE, 0xEA, 0xF1)   # subtle card backgrounds
SKY_BLUE     = RGBColor(0x9D, 0xC3, 0xE6)   # divider lines / icons
WHITE        = RGBColor(0xFF, 0xFF, 0xFF)
DARK_GRAY    = RGBColor(0x40, 0x40, 0x40)   # body text
MID_GRAY     = RGBColor(0x7F, 0x7F, 0x7F)   # captions / meta
LIGHT_GRAY   = RGBColor(0xF2, 0xF2, 0xF2)   # slide background tint
GOLD         = RGBColor(0xED, 0x7D, 0x31)   # call-out / emphasis
GREEN        = RGBColor(0x70, 0xAD, 0x47)   # positive / new feature
RED_ORANGE   = RGBColor(0xC0, 0x50, 0x20)   # warning / old standard

# ── Slide size (16:9 widescreen) ───────────────────────────────────────────────
SLIDE_W = Inches(13.33)
SLIDE_H = Inches(7.5)


# ══════════════════════════════════════════════════════════════════════════════
# Helper utilities
# ══════════════════════════════════════════════════════════════════════════════

def new_prs() -> Presentation:
    prs = Presentation()
    prs.slide_width  = SLIDE_W
    prs.slide_height = SLIDE_H
    return prs


def blank_slide(prs: Presentation):
    """Add a slide with the blank layout."""
    blank_layout = prs.slide_layouts[6]
    return prs.slides.add_slide(blank_layout)


def fill_shape(shape, color: RGBColor):
    shape.fill.solid()
    shape.fill.fore_color.rgb = color


def no_border(shape):
    shape.line.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    shape.line.width = Pt(0)


def add_rect(slide, left, top, width, height, color: RGBColor, line_color=None, line_width=Pt(0)):
    shape = slide.shapes.add_shape(
        1,  # MSO_SHAPE_TYPE.RECTANGLE
        left, top, width, height
    )
    fill_shape(shape, color)
    if line_color:
        shape.line.color.rgb = line_color
        shape.line.width = line_width
    else:
        shape.line.fill.background()
    return shape


def add_textbox(slide, text, left, top, width, height,
                font_size=Pt(12), bold=False, color=DARK_GRAY,
                align=PP_ALIGN.LEFT, italic=False, wrap=True):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = font_size
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    return txBox


def add_text_to_shape(shape, text, font_size=Pt(12), bold=False,
                      color=WHITE, align=PP_ALIGN.CENTER,
                      italic=False, vertical_anchor=None):
    from pptx.enum.text import MSO_ANCHOR
    tf = shape.text_frame
    tf.word_wrap = True
    if vertical_anchor:
        tf.vertical_anchor = vertical_anchor
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = font_size
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    return shape


def add_slide_header(slide, title: str, subtitle: str = ""):
    """Standard slide header bar."""
    # top colour band
    add_rect(slide, 0, 0, SLIDE_W, Inches(1.25), DARK_NAVY)
    # accent line
    add_rect(slide, 0, Inches(1.25), SLIDE_W, Inches(0.06), GOLD)

    add_textbox(slide, title,
                Inches(0.4), Inches(0.18), Inches(12.5), Inches(0.7),
                font_size=Pt(28), bold=True, color=WHITE, align=PP_ALIGN.LEFT)
    if subtitle:
        add_textbox(slide, subtitle,
                    Inches(0.4), Inches(0.82), Inches(12.5), Inches(0.38),
                    font_size=Pt(13), bold=False, color=SKY_BLUE, align=PP_ALIGN.LEFT)


def add_footer(slide, text="IFRS 18 | Presentation and Disclosure in Financial Statements"):
    add_rect(slide, 0, Inches(7.1), SLIDE_W, Inches(0.4), DARK_NAVY)
    add_textbox(slide, text,
                Inches(0.3), Inches(7.12), Inches(12.0), Inches(0.35),
                font_size=Pt(9), color=SKY_BLUE, align=PP_ALIGN.LEFT)


def bullet_paragraph(tf, text: str, level: int = 0,
                     font_size=Pt(13), bold=False,
                     color=DARK_GRAY, space_before=Pt(4)):
    """Append a bullet paragraph to an existing text frame."""
    from pptx.util import Pt as _Pt
    p = tf.add_paragraph()
    p.level = level
    p.space_before = space_before
    p.alignment = PP_ALIGN.LEFT
    run = p.add_run()
    run.text = text
    run.font.size = font_size
    run.font.bold = bold
    run.font.color.rgb = color


def info_card(slide, left, top, width, height,
              title: str, body_lines: list,
              header_color=BRIGHT_BLUE, body_color=LIGHT_BLUE,
              title_size=Pt(14), body_size=Pt(11.5)):
    """Draws a titled card with a header strip and body text block."""
    # header
    hdr = add_rect(slide, left, top, width, Inches(0.45), header_color)
    add_text_to_shape(hdr, title, font_size=title_size, bold=True,
                      color=WHITE, align=PP_ALIGN.LEFT)
    hdr.text_frame.paragraphs[0].runs[0]  # already set
    # shift text left padding
    hdr.text_frame.margin_left  = Inches(0.15)
    hdr.text_frame.margin_top   = Inches(0.04)

    # body
    body = add_rect(slide, left, top + Inches(0.45),
                    width, height - Inches(0.45), body_color)
    body.line.color.rgb = SKY_BLUE
    body.line.width = Pt(0.75)
    tf = body.text_frame
    tf.word_wrap = True
    tf.margin_left  = Inches(0.12)
    tf.margin_top   = Inches(0.08)
    tf.margin_right = Inches(0.08)

    first = True
    for line in body_lines:
        if first:
            p = tf.paragraphs[0]
            first = False
        else:
            p = tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.space_before = Pt(3)
        run = p.add_run()
        run.text = line
        run.font.size = body_size
        run.font.color.rgb = DARK_GRAY


def number_circle(slide, left, top, diameter, number: str,
                  bg_color=BRIGHT_BLUE, text_color=WHITE, font_size=Pt(20)):
    """Draws a circular numbered badge."""
    circle = slide.shapes.add_shape(
        9,  # OVAL
        left, top, diameter, diameter
    )
    fill_shape(circle, bg_color)
    circle.line.fill.background()
    add_text_to_shape(circle, number,
                      font_size=font_size, bold=True, color=text_color)


def arrow_shape(slide, left, top, width, height, color=GOLD):
    """Right-pointing chevron arrow."""
    from pptx.util import Emu
    arrow = slide.shapes.add_shape(
        13,  # RIGHT_ARROW
        left, top, width, height
    )
    fill_shape(arrow, color)
    arrow.line.fill.background()
    return arrow


# ══════════════════════════════════════════════════════════════════════════════
# Individual slide builders
# ══════════════════════════════════════════════════════════════════════════════

def slide_01_title(prs):
    """Cover / Title slide."""
    slide = blank_slide(prs)

    # Background – full dark navy
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, DARK_NAVY)

    # Decorative accent rectangle – left band
    add_rect(slide, 0, 0, Inches(0.22), SLIDE_H, GOLD)

    # Geometric circles (decorative)
    for n, (cx, cy, sz, alpha) in enumerate([
        (Inches(11.5), Inches(-0.5), Inches(3.8), ROYAL_BLUE),
        (Inches(12.2), Inches(1.2),  Inches(2.0), BRIGHT_BLUE),
        (Inches(10.8), Inches(5.8),  Inches(2.5), RGBColor(0x18, 0x2D, 0x50)),
    ]):
        c = slide.shapes.add_shape(9, cx, cy, sz, sz)
        fill_shape(c, alpha)
        c.line.fill.background()

    # Main title
    add_textbox(slide, "IFRS 18",
                Inches(0.6), Inches(1.4), Inches(9), Inches(1.6),
                font_size=Pt(72), bold=True, color=WHITE, align=PP_ALIGN.LEFT)

    # Gold divider line
    add_rect(slide, Inches(0.6), Inches(3.1), Inches(5.5), Inches(0.07), GOLD)

    # Subtitle
    add_textbox(slide,
                "Presentation and Disclosure\nin Financial Statements",
                Inches(0.6), Inches(3.25), Inches(9.5), Inches(1.3),
                font_size=Pt(26), bold=False, color=SKY_BLUE, align=PP_ALIGN.LEFT)

    # Tag-line / descriptor
    add_textbox(slide,
                "A Comprehensive Guide to the New Standard — Replacing IAS 1",
                Inches(0.6), Inches(4.65), Inches(9.5), Inches(0.55),
                font_size=Pt(14), bold=False, color=RGBColor(0xCC, 0xCC, 0xCC),
                align=PP_ALIGN.LEFT)

    # Effective-date pill
    pill = add_rect(slide, Inches(0.6), Inches(5.45), Inches(3.4), Inches(0.55),
                    GOLD)
    add_text_to_shape(pill, "Effective:  1 January 2027",
                      font_size=Pt(13), bold=True, color=WHITE)

    # Bottom caption
    add_textbox(slide,
                "IASB — Issued May 2024  |  Supersedes IAS 1 (2007)",
                Inches(0.6), Inches(6.3), Inches(10), Inches(0.4),
                font_size=Pt(11), color=MID_GRAY, align=PP_ALIGN.LEFT)

    add_footer(slide)


def slide_02_agenda(prs):
    """Table of Contents / Agenda."""
    slide = blank_slide(prs)
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, LIGHT_GRAY)
    add_slide_header(slide, "Agenda", "What we will cover today")

    topics = [
        ("01", "Overview — What is IFRS 18?",           BRIGHT_BLUE),
        ("02", "Why IFRS 18?  Problems with IAS 1",      ROYAL_BLUE),
        ("03", "New Statement of Profit or Loss",         BRIGHT_BLUE),
        ("04", "Three Income Categories",                 ROYAL_BLUE),
        ("05", "Required Subtotals",                      BRIGHT_BLUE),
        ("06", "Management-defined Performance Measures", ROYAL_BLUE),
        ("07", "MPM Disclosure Requirements",             BRIGHT_BLUE),
        ("08", "Aggregation & Disaggregation Principles", ROYAL_BLUE),
        ("09", "Changes to Statement of Cash Flows",      BRIGHT_BLUE),
        ("10", "Effective Date & Transition",             ROYAL_BLUE),
        ("11", "Key Takeaways",                           BRIGHT_BLUE),
    ]

    col_w = Inches(5.8)
    col_gap = Inches(0.3)
    left1 = Inches(0.4)
    left2 = left1 + col_w + col_gap
    top_start = Inches(1.45)
    row_h = Inches(0.51)
    row_gap = Inches(0.04)

    for i, (num, label, color) in enumerate(topics):
        col = 0 if i < 6 else 1
        row = i if i < 6 else i - 6
        left = left1 if col == 0 else left2
        top = top_start + row * (row_h + row_gap)

        # number badge
        badge = add_rect(slide, left, top, Inches(0.48), row_h, color)
        add_text_to_shape(badge, num, font_size=Pt(13), bold=True, color=WHITE)

        # label box
        lbl = add_rect(slide, left + Inches(0.5), top,
                       col_w - Inches(0.52), row_h, WHITE)
        lbl.line.color.rgb = RGBColor(0xCC, 0xCC, 0xCC)
        lbl.line.width = Pt(0.5)
        add_text_to_shape(lbl, label,
                          font_size=Pt(12.5), bold=False,
                          color=DARK_GRAY, align=PP_ALIGN.LEFT)
        lbl.text_frame.margin_left = Inches(0.12)

    add_footer(slide)


def slide_03_overview(prs):
    """What is IFRS 18?"""
    slide = blank_slide(prs)
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, LIGHT_GRAY)
    add_slide_header(slide, "Overview — What is IFRS 18?",
                     "A new era for financial statement presentation")

    # Central graphic – large circle
    main_circle = slide.shapes.add_shape(9,
                                         Inches(4.9), Inches(1.5),
                                         Inches(3.5), Inches(3.5))
    fill_shape(main_circle, DARK_NAVY)
    main_circle.line.color.rgb = GOLD
    main_circle.line.width = Pt(3)
    add_text_to_shape(main_circle, "IFRS 18",
                      font_size=Pt(34), bold=True, color=WHITE)

    # Surrounding fact cards
    facts = [
        (Inches(0.3),  Inches(1.5),  Inches(4.3), Inches(1.6),
         "Issued",
         "May 2024 by the IASB\nReplaces IAS 1 (2007)\nNew standard for all IFRS reporters",
         BRIGHT_BLUE),

        (Inches(8.8),  Inches(1.5),  Inches(4.3), Inches(1.6),
         "Objective",
         "Improve comparability & transparency\nof financial statements\nacross entities and industries",
         ROYAL_BLUE),

        (Inches(0.3),  Inches(3.3),  Inches(4.3), Inches(1.6),
         "Scope",
         "All entities preparing financial\nstatements under IFRS\n(all industries, all sizes)",
         ROYAL_BLUE),

        (Inches(8.8),  Inches(3.3),  Inches(4.3), Inches(1.6),
         "Effective Date",
         "Annual periods beginning\non or after 1 January 2027\nEarly adoption permitted",
         BRIGHT_BLUE),

        (Inches(0.3),  Inches(5.1),  Inches(4.3), Inches(1.6),
         "Key Change 1",
         "Structured P&L with three\nmandatory income categories:\nOperating · Investing · Financing",
         DARK_NAVY),

        (Inches(8.8),  Inches(5.1),  Inches(4.3), Inches(1.6),
         "Key Change 2",
         "New rules for Management-defined\nPerformance Measures (MPMs)\nwith reconciliation to IFRS line items",
         DARK_NAVY),
    ]

    for (l, t, w, h, ttl, body, col) in facts:
        info_card(slide, l, t, w, h, ttl,
                  body.split("\n"), header_color=col)

    # Connector arrows (decorative)
    for ax, ay in [(Inches(4.65), Inches(2.2)), (Inches(4.65), Inches(4.0)),
                   (Inches(8.45), Inches(2.2)), (Inches(8.45), Inches(4.0))]:
        arr = slide.shapes.add_shape(13, ax, ay, Inches(0.35), Inches(0.28))
        fill_shape(arr, GOLD)
        arr.line.fill.background()

    add_footer(slide)


def slide_04_why_ifrs18(prs):
    """Why IFRS 18? — Problems with IAS 1 (comparison layout)."""
    slide = blank_slide(prs)
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, LIGHT_GRAY)
    add_slide_header(slide, "Why IFRS 18?",
                     "Limitations of IAS 1 that IFRS 18 addresses")

    # Two-column comparison
    # Left: IAS 1 Issues  |  Right: IFRS 18 Solutions
    mid = Inches(6.67)

    add_rect(slide, Inches(0.3), Inches(1.4), Inches(6.0), Inches(0.5),
             RED_ORANGE)
    lbl_l = slide.shapes[-1]
    add_text_to_shape(lbl_l, "⚠  Limitations of IAS 1",
                      font_size=Pt(15), bold=True, color=WHITE)

    add_rect(slide, Inches(6.7), Inches(1.4), Inches(6.0), Inches(0.5), GREEN)
    lbl_r = slide.shapes[-1]
    add_text_to_shape(lbl_r, "✔  How IFRS 18 Solves Them",
                      font_size=Pt(15), bold=True, color=WHITE)

    problems = [
        "No mandatory categories in P&L\n→ entities classify items inconsistently",
        "No required subtotals\n→ operating profit varies across companies",
        "Unrestricted use of non-GAAP measures\n→ reduces comparability",
        "Vague aggregation guidance\n→ obscures material information",
        "Inconsistent P&L and cash flow\npresentation for investments",
    ]
    solutions = [
        "Three mandatory income categories:\nOperating, Investing & Financing",
        "Required subtotals including\nOperating Profit and Profit before financing",
        "MPM framework with mandatory\nreconciliation and prominence rules",
        "Clear principles: material items\nmust be separately presented",
        "Integral vs. non-integral associates\nalign P&L and cash flow presentation",
    ]

    for i, (prob, sol) in enumerate(zip(problems, solutions)):
        top = Inches(2.05) + i * Inches(0.95)
        h   = Inches(0.88)

        # problem card
        pc = add_rect(slide, Inches(0.3), top, Inches(6.0), h, WHITE)
        pc.line.color.rgb = RED_ORANGE
        pc.line.width = Pt(1)
        tf = pc.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0.12)
        tf.margin_top  = Inches(0.06)
        p0 = tf.paragraphs[0]
        p0.alignment = PP_ALIGN.LEFT
        lines = prob.split("\n")
        r0 = p0.add_run()
        r0.text = lines[0]
        r0.font.size = Pt(11.5)
        r0.font.bold = True
        r0.font.color.rgb = RED_ORANGE
        if len(lines) > 1:
            p1 = tf.add_paragraph()
            p1.alignment = PP_ALIGN.LEFT
            r1 = p1.add_run()
            r1.text = lines[1]
            r1.font.size = Pt(11)
            r1.font.color.rgb = DARK_GRAY

        # solution card
        sc = add_rect(slide, Inches(6.7), top, Inches(6.0), h, WHITE)
        sc.line.color.rgb = GREEN
        sc.line.width = Pt(1)
        tf2 = sc.text_frame
        tf2.word_wrap = True
        tf2.margin_left = Inches(0.12)
        tf2.margin_top  = Inches(0.06)
        p2 = tf2.paragraphs[0]
        p2.alignment = PP_ALIGN.LEFT
        lines2 = sol.split("\n")
        r2 = p2.add_run()
        r2.text = lines2[0]
        r2.font.size = Pt(11.5)
        r2.font.bold = True
        r2.font.color.rgb = GREEN
        if len(lines2) > 1:
            p3 = tf2.add_paragraph()
            p3.alignment = PP_ALIGN.LEFT
            r3 = p3.add_run()
            r3.text = lines2[1]
            r3.font.size = Pt(11)
            r3.font.color.rgb = DARK_GRAY

        # arrow between
        arr = slide.shapes.add_shape(13,
                                     Inches(6.35), top + Inches(0.3),
                                     Inches(0.3), Inches(0.28))
        fill_shape(arr, GOLD)
        arr.line.fill.background()

    add_footer(slide)


def slide_05_new_pl_structure(prs):
    """New Statement of Profit or Loss — overview."""
    slide = blank_slide(prs)
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, LIGHT_GRAY)
    add_slide_header(slide, "New Statement of Profit or Loss",
                     "A structured, comparable presentation of income and expenses")

    # Central funnel-style layout showing categories flowing to profit

    # Top: Revenue box
    rev = add_rect(slide, Inches(4.4), Inches(1.5), Inches(4.5), Inches(0.65),
                   DARK_NAVY)
    add_text_to_shape(rev, "Revenue", font_size=Pt(16), bold=True, color=WHITE)

    # Three category columns
    cat_data = [
        (Inches(0.3),  "OPERATING\nCATEGORY",    BRIGHT_BLUE,
         "+ Revenue\n+ Other operating income\n– Cost of sales\n– Operating expenses\n– D&A"),
        (Inches(4.75), "INVESTING\nCATEGORY",     ROYAL_BLUE,
         "+ Income from non-integral\n   associates & JVs\n+ Returns on surplus cash\n– Related costs"),
        (Inches(9.2),  "FINANCING\nCATEGORY",     DARK_NAVY,
         "+ Gains on financial liabilities\n– Interest expense on debt\n– Other financing costs\n– FX on borrowings"),
    ]

    for (lft, ttl, col, items) in cat_data:
        # category header
        hdr = add_rect(slide, lft, Inches(2.35), Inches(3.9), Inches(0.7), col)
        add_text_to_shape(hdr, ttl, font_size=Pt(14), bold=True, color=WHITE)

        # items box
        itm = add_rect(slide, lft, Inches(3.1), Inches(3.9), Inches(2.3), WHITE)
        itm.line.color.rgb = col
        itm.line.width = Pt(1.5)
        tf = itm.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0.12)
        tf.margin_top  = Inches(0.1)
        first = True
        for line in items.strip().split("\n"):
            if first:
                p = tf.paragraphs[0]; first = False
            else:
                p = tf.add_paragraph()
            p.alignment = PP_ALIGN.LEFT
            p.space_before = Pt(2)
            run = p.add_run()
            run.text = line
            run.font.size = Pt(11.5)
            run.font.color.rgb = DARK_GRAY

    # Subtotals row
    subtotals = [
        (Inches(0.3),  "Operating Profit",             BRIGHT_BLUE),
        (Inches(4.75), "+ Investing Income / Expense",  ROYAL_BLUE),
        (Inches(9.2),  "+ Financing Income / Expense",  DARK_NAVY),
    ]
    for (lft, lbl, col) in subtotals:
        s = add_rect(slide, lft, Inches(5.55), Inches(3.9), Inches(0.55), col)
        add_text_to_shape(s, lbl, font_size=Pt(12), bold=True, color=WHITE)

    # Final Profit box
    profit = add_rect(slide, Inches(3.5), Inches(6.25), Inches(6.3), Inches(0.65), GOLD)
    add_text_to_shape(profit, "PROFIT BEFORE TAX",
                      font_size=Pt(17), bold=True, color=WHITE)

    # Downward arrows
    for ax in [Inches(2.1), Inches(6.6), Inches(11.1)]:
        arr = slide.shapes.add_shape(15, ax, Inches(5.25), Inches(0.3), Inches(0.28))
        fill_shape(arr, MID_GRAY)
        arr.line.fill.background()

    add_footer(slide)


def slide_06_operating(prs):
    """Operating Category — deep-dive."""
    slide = blank_slide(prs)
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, LIGHT_GRAY)
    add_slide_header(slide, "The Operating Category",
                     "Income and expenses from the entity's main business activities")

    # Left: definition block
    defn = add_rect(slide, Inches(0.3), Inches(1.5), Inches(4.9), Inches(5.4),
                    BRIGHT_BLUE)
    add_text_to_shape(defn, "", font_size=Pt(12), color=WHITE)  # placeholder
    tf = defn.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.2)
    tf.margin_top  = Inches(0.15)

    lines_def = [
        ("Definition", True, Pt(17), WHITE),
        ("", False, Pt(6), WHITE),
        ("The RESIDUAL category — it captures all", False, Pt(12), WHITE),
        ("income and expenses NOT classified in the", False, Pt(12), WHITE),
        ("Investing or Financing categories.", False, Pt(12), WHITE),
        ("", False, Pt(6), WHITE),
        ("Always includes:", True, Pt(13), GOLD),
        ("  •  Revenue from contracts with customers", False, Pt(12), WHITE),
        ("  •  Cost of sales", False, Pt(12), WHITE),
        ("  •  Selling, general & administrative costs", False, Pt(12), WHITE),
        ("  •  Depreciation & amortisation", False, Pt(12), WHITE),
        ("  •  Impairment of operating assets", False, Pt(12), WHITE),
        ("  •  Share of profit of integral associates", False, Pt(12), WHITE),
        ("  •  FX on operating items", False, Pt(12), WHITE),
        ("", False, Pt(6), WHITE),
        ("May also include:", True, Pt(13), GOLD),
        ("  •  Research & development costs", False, Pt(12), WHITE),
        ("  •  Government grants", False, Pt(12), WHITE),
        ("  •  Gains/losses on disposal of PP&E", False, Pt(12), WHITE),
    ]
    first = True
    for (txt, bold, sz, col) in lines_def:
        if first:
            p = tf.paragraphs[0]; first = False
        else:
            p = tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.space_before = Pt(1)
        run = p.add_run()
        run.text = txt
        run.font.size = sz
        run.font.bold = bold
        run.font.color.rgb = col

    # Right: Key point cards
    right_items = [
        ("Residual Category",
         "Operating is defined by exclusion — if an item is not Investing or Financing, it belongs here. This ensures all items are classified."),
        ("Integral vs. Non-Integral\nAssociates & JVs",
         "Associates integral to the entity's main activities → Operating.\nAll others → Investing category."),
        ("Income Tax",
         "Income tax expense is presented separately AFTER all three category subtotals, consistent with IAS 12."),
        ("Why It Matters",
         "Standardising the Operating category makes 'Operating Profit' comparable across companies and industries for the first time."),
    ]
    top_r = Inches(1.5)
    for (ttl, body) in right_items:
        info_card(slide, Inches(5.4), top_r, Inches(7.6), Inches(1.28),
                  ttl, [body], header_color=BRIGHT_BLUE)
        top_r += Inches(1.35)

    add_footer(slide)


def slide_07_investing(prs):
    """Investing Category — deep-dive."""
    slide = blank_slide(prs)
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, LIGHT_GRAY)
    add_slide_header(slide, "The Investing Category",
                     "Returns from assets not related to the main business activities")

    # Definition strip
    defn_strip = add_rect(slide, Inches(0.3), Inches(1.5), SLIDE_W - Inches(0.6),
                          Inches(0.85), ROYAL_BLUE)
    add_text_to_shape(
        defn_strip,
        "Investing category = income and expenses from assets that generate returns INDEPENDENTLY of the entity's main activities",
        font_size=Pt(13), bold=False, color=WHITE, align=PP_ALIGN.LEFT
    )
    defn_strip.text_frame.margin_left = Inches(0.2)

    # What's IN vs NOT IN
    in_items = [
        "Income from non-integral associates / JVs",
        "Dividends from equity investments (FVOCI)",
        "Interest from surplus cash & short-term investments",
        "Gains/losses on non-integral investment disposals",
        "FX gains/losses on investing assets",
        "Impairment of investing-category assets",
    ]
    out_items = [
        "Returns from integral associates / JVs  → Operating",
        "Interest on operating trade receivables  → Operating",
        "Interest on pension liability  → Financing",
        "Interest expense on borrowings  → Financing",
        "Revenue from sale of goods/services  → Operating",
    ]

    info_card(slide, Inches(0.3), Inches(2.5), Inches(6.1), Inches(4.0),
              "✔  Classified as INVESTING",
              ["• " + x for x in in_items],
              header_color=ROYAL_BLUE, body_size=Pt(12))

    info_card(slide, Inches(6.6), Inches(2.5), Inches(6.4), Inches(4.0),
              "✘  NOT Classified as Investing",
              ["• " + x for x in out_items],
              header_color=RED_ORANGE, body_size=Pt(12))

    # Bottom note
    note = add_rect(slide, Inches(0.3), Inches(6.6), SLIDE_W - Inches(0.6),
                    Inches(0.58), LIGHT_BLUE)
    note.line.color.rgb = ROYAL_BLUE
    note.line.width = Pt(0.75)
    add_text_to_shape(
        note,
        "Key Principle: An asset is in the Investing category when its returns do not depend on the volume or success of the entity's main operating activities.",
        font_size=Pt(11.5), bold=False, color=DARK_NAVY, align=PP_ALIGN.LEFT
    )
    note.text_frame.margin_left = Inches(0.15)

    add_footer(slide)


def slide_08_financing(prs):
    """Financing Category — deep-dive."""
    slide = blank_slide(prs)
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, LIGHT_GRAY)
    add_slide_header(slide, "The Financing Category",
                     "Costs and returns directly linked to how the entity is funded")

    # Definition strip
    defn_strip = add_rect(slide, Inches(0.3), Inches(1.5), SLIDE_W - Inches(0.6),
                          Inches(0.85), DARK_NAVY)
    add_text_to_shape(
        defn_strip,
        "Financing category = income and expenses arising from liabilities and equity instruments that represent the entity's financing structure",
        font_size=Pt(13), bold=False, color=WHITE, align=PP_ALIGN.LEFT
    )
    defn_strip.text_frame.margin_left = Inches(0.2)

    # Three-column cards
    cols = [
        ("Financing\nLiabilities",
         DARK_NAVY,
         [
             "• Interest expense on bonds & loans",
             "• Amortisation of debt issue costs",
             "• Fair value changes on financial liabilities",
             "• Gains/losses on extinguishment of debt",
             "• FX on foreign-currency borrowings",
         ]),
        ("Equity-Related\nItems",
         BRIGHT_BLUE,
         [
             "• Interest on lease liabilities (IFRS 16)",
             "• Discount unwinding on long-term provisions",
             "• Net interest on pension obligations (IAS 19)",
             "• FX on equity-denominated instruments",
         ]),
        ("What is Excluded",
         RED_ORANGE,
         [
             "✘  Interest on trade payables  → Operating",
             "✘  Bank charges on operations  → Operating",
             "✘  Dividends paid  → Equity (not P&L)",
             "✘  Investment returns  → Investing",
         ]),
    ]

    col_w = Inches(4.0)
    for i, (ttl, col, items) in enumerate(cols):
        lft = Inches(0.3) + i * (col_w + Inches(0.3))
        info_card(slide, lft, Inches(2.5), col_w, Inches(4.0),
                  ttl, items, header_color=col, body_size=Pt(12))

    # Bottom note
    note = add_rect(slide, Inches(0.3), Inches(6.6), SLIDE_W - Inches(0.6),
                    Inches(0.58), LIGHT_BLUE)
    note.line.color.rgb = DARK_NAVY
    note.line.width = Pt(0.75)
    add_text_to_shape(
        note,
        "Key Rule: Entities that classify interest in the Investing or Financing category in the cash flow statement MUST use the SAME classification in the P&L.",
        font_size=Pt(11.5), bold=False, color=DARK_NAVY, align=PP_ALIGN.LEFT
    )
    note.text_frame.margin_left = Inches(0.15)

    add_footer(slide)


def slide_09_subtotals(prs):
    """Required Subtotals."""
    slide = blank_slide(prs)
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, LIGHT_GRAY)
    add_slide_header(slide, "Required Subtotals in the P&L",
                     "IFRS 18 mandates specific line items and subtotals for the first time")

    # Waterfall-style visual
    items = [
        ("Revenue",                                 BRIGHT_BLUE,   True,  False),
        ("+ Other Operating Income / (Expenses)",   BRIGHT_BLUE,   False, False),
        ("= OPERATING PROFIT  ★",                   GOLD,          True,  True),
        ("+ Investing Income / (Expenses)",          ROYAL_BLUE,   False, False),
        ("= PROFIT BEFORE FINANCING & TAX",         DARK_NAVY,     True,  True),
        ("+ Financing Income / (Expenses)",          DARK_NAVY,    False, False),
        ("= PROFIT BEFORE TAX  ★",                  GOLD,          True,  True),
        ("– Income Tax Expense",                    MID_GRAY,      False, False),
        ("= PROFIT FOR THE PERIOD  ★",              GREEN,         True,  True),
    ]

    bar_w = Inches(8.0)
    indent = Inches(0.0)
    top = Inches(1.5)
    row_h = Inches(0.54)
    gap   = Inches(0.03)

    for (label, col, bold, is_total) in items:
        lft = Inches(0.5) + (Inches(0.4) if not is_total else 0)
        w   = bar_w - (Inches(0.4) if not is_total else 0)
        bar = add_rect(slide, lft, top, w, row_h, col)
        add_text_to_shape(bar, label,
                          font_size=Pt(13 if is_total else 12),
                          bold=bold, color=WHITE, align=PP_ALIGN.LEFT)
        bar.text_frame.margin_left = Inches(0.2)
        top += row_h + gap

    # Right: annotations
    ann_items = [
        (Inches(1.55), "★  Mandatory subtotals\n      required by IFRS 18"),
        (Inches(3.17), "Entities may present additional\nsubtotals if useful (but cannot\nremove mandatory ones)"),
        (Inches(5.72), "★  New: 'Profit before financing\n      and income tax' subtotal"),
        (Inches(6.78), "Consistent with IAS 12\npresentation requirements"),
    ]
    for (t, txt) in ann_items:
        ann = add_rect(slide, Inches(9.0), t, Inches(4.0), Inches(0.85),
                       LIGHT_BLUE)
        ann.line.color.rgb = BRIGHT_BLUE
        ann.line.width = Pt(0.75)
        add_text_to_shape(ann, txt, font_size=Pt(11), bold=False,
                          color=DARK_NAVY, align=PP_ALIGN.LEFT)
        ann.text_frame.margin_left = Inches(0.1)
        ann.text_frame.margin_top  = Inches(0.06)

    add_footer(slide)


def slide_10_mpm_intro(prs):
    """Management-defined Performance Measures (MPMs) — introduction."""
    slide = blank_slide(prs)
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, LIGHT_GRAY)
    add_slide_header(slide, "Management-defined Performance Measures (MPMs)",
                     "New rules for non-GAAP and alternative performance measures")

    # Large quote / definition
    quote_box = add_rect(slide, Inches(0.3), Inches(1.5), SLIDE_W - Inches(0.6),
                         Inches(1.3), DARK_NAVY)
    add_text_to_shape(
        quote_box,
        "\"An MPM is a subtotal of income and expenses that is not specified in IFRS Standards,\n"
        "is used or referenced in public communications outside the financial statements,\n"
        "and communicates management's view of an aspect of financial performance.\"",
        font_size=Pt(13.5), italic=True, color=SKY_BLUE, align=PP_ALIGN.CENTER
    )

    # Four info cards
    cards = [
        ("What Qualifies as an MPM?",
         BRIGHT_BLUE,
         [
             "• A subtotal of income & expenses (not just ratios)",
             "• Used in earnings releases, investor presentations,\n  management commentary, or analyst calls",
             "• Purports to communicate management's view of\n  an aspect of financial performance",
             "• NOT specified / required by IFRS Standards",
         ]),
        ("Common Examples",
         ROYAL_BLUE,
         [
             "• Adjusted EBITDA / Underlying EBITDA",
             "• Adjusted Operating Profit",
             "• Core Earnings / Recurring Earnings",
             "• Adjusted EPS (earnings per share)",
             "• Free Cash Flow subtotals (if in P&L context)",
         ]),
        ("What Does NOT Qualify?",
         RED_ORANGE,
         [
             "• Ratios: profit margins, ROCE, EPS",
             "• Individual line items not labelled a subtotal",
             "• Items required by an IFRS Standard",
             "• Subtotals used only internally (never in public comms)",
         ]),
        ("Why This Matters",
         GREEN,
         [
             "• Provides a level playing field for APMs globally",
             "• Makes adjustments to IFRS profit transparent",
             "• Reconciliation required to nearest IFRS subtotal",
             "• Prominence rules prevent misleading presentation",
         ]),
    ]

    col_w = Inches(6.0)
    for i, (ttl, col, items) in enumerate(cards):
        lft = Inches(0.3) + (i % 2) * (col_w + Inches(0.45))
        top = Inches(2.95) + (i // 2) * Inches(1.9)
        info_card(slide, lft, top, col_w, Inches(1.82),
                  ttl, items, header_color=col, body_size=Pt(11.5))

    add_footer(slide)


def slide_11_mpm_disclosure(prs):
    """MPM Disclosure Requirements."""
    slide = blank_slide(prs)
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, LIGHT_GRAY)
    add_slide_header(slide, "MPM Disclosure Requirements",
                     "What entities must disclose for every MPM they use")

    # Central process-flow: 5 boxes connected by arrows
    steps = [
        ("1\nLabel",
         "Use a descriptive label\nthat conveys the nature\nof the MPM",
         BRIGHT_BLUE),
        ("2\nExplain",
         "Explain why the MPM\nprovides useful info about\nfinancial performance",
         ROYAL_BLUE),
        ("3\nReconcile",
         "Reconcile to the nearest\nmandatory IFRS subtotal\nline-by-line",
         DARK_NAVY),
        ("4\nTax &\nMinorities",
         "Show the tax effect and\nnon-controlling interest\neffect of each adjustment",
         BRIGHT_BLUE),
        ("5\nConsistency",
         "Apply the same definition\nconsistently year-over-year;\ndisclose changes",
         ROYAL_BLUE),
    ]

    box_w = Inches(2.25)
    box_h = Inches(2.5)
    top_b = Inches(1.7)
    gap   = Inches(0.2)

    for i, (num, desc, col) in enumerate(steps):
        lft = Inches(0.3) + i * (box_w + gap)

        # number circle
        circ = slide.shapes.add_shape(9,
                                      lft + Inches(0.7), top_b,
                                      Inches(0.85), Inches(0.85))
        fill_shape(circ, col)
        circ.line.fill.background()
        add_text_to_shape(circ, num.split("\n")[0],
                          font_size=Pt(22), bold=True, color=WHITE)

        # label
        add_textbox(slide, num.split("\n")[1] if "\n" in num else "",
                    lft, top_b + Inches(0.55), box_w, Inches(0.45),
                    font_size=Pt(14), bold=True, color=col, align=PP_ALIGN.CENTER)

        # description card
        dc = add_rect(slide, lft, top_b + Inches(1.1), box_w, Inches(1.38), WHITE)
        dc.line.color.rgb = col
        dc.line.width = Pt(1.5)
        tf = dc.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0.1)
        tf.margin_top  = Inches(0.08)
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        run.text = desc
        run.font.size = Pt(11.5)
        run.font.color.rgb = DARK_GRAY

        # connector arrow (except after last)
        if i < len(steps) - 1:
            arr = slide.shapes.add_shape(13,
                                         lft + box_w + Inches(0.03),
                                         top_b + Inches(0.6),
                                         Inches(0.15), Inches(0.28))
            fill_shape(arr, GOLD)
            arr.line.fill.background()

    # Prominence Rule highlight
    prom = add_rect(slide, Inches(0.3), Inches(4.5), SLIDE_W - Inches(0.6),
                    Inches(0.85), DARK_NAVY)
    add_text_to_shape(
        prom,
        "⚡  Prominence Rule:  MPMs must NOT be given more prominence than the nearest comparable mandatory IFRS subtotal\n"
        "in ANY communication — including earnings releases, investor presentations, and management commentary.",
        font_size=Pt(12), bold=False, color=WHITE, align=PP_ALIGN.LEFT
    )
    prom.text_frame.margin_left  = Inches(0.2)
    prom.text_frame.margin_top   = Inches(0.1)
    prom.text_frame.word_wrap    = True

    # Reconciliation example table
    add_textbox(slide, "Example MPM Reconciliation",
                Inches(0.3), Inches(5.5), Inches(4.5), Inches(0.4),
                font_size=Pt(13), bold=True, color=DARK_NAVY)

    tbl_data = [
        ("Operating Profit (IFRS)",              "2,500"),
        ("Add: Restructuring costs",             "  150"),
        ("Add: Impairment of goodwill",          "  200"),
        ("Less: Gain on disposal",               " (80)"),
        ("= Adjusted Operating Profit (MPM)",    "2,770"),
    ]
    for j, (item, amt) in enumerate(tbl_data):
        is_total = j == len(tbl_data) - 1
        bg = BRIGHT_BLUE if is_total else (LIGHT_BLUE if j % 2 == 0 else WHITE)
        fc = WHITE if is_total else DARK_GRAY
        row = add_rect(slide, Inches(0.3), Inches(5.95) + j * Inches(0.24),
                       Inches(4.8), Inches(0.23), bg)
        row.line.color.rgb = SKY_BLUE
        row.line.width = Pt(0.5)
        add_text_to_shape(row, item,
                          font_size=Pt(10), bold=is_total, color=fc,
                          align=PP_ALIGN.LEFT)
        row.text_frame.margin_left = Inches(0.1)

        amt_row = add_rect(slide, Inches(5.15), Inches(5.95) + j * Inches(0.24),
                           Inches(0.85), Inches(0.23), bg)
        amt_row.line.color.rgb = SKY_BLUE
        amt_row.line.width = Pt(0.5)
        add_text_to_shape(amt_row, amt,
                          font_size=Pt(10), bold=is_total, color=fc,
                          align=PP_ALIGN.RIGHT)

    add_footer(slide)


def slide_12_aggregation(prs):
    """Aggregation and Disaggregation Principles."""
    slide = blank_slide(prs)
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, LIGHT_GRAY)
    add_slide_header(slide, "Aggregation & Disaggregation Principles",
                     "When to combine — and when to separate — line items")

    # Two-sided layout
    # Left: Aggregation
    info_card(slide, Inches(0.3), Inches(1.5), Inches(6.0), Inches(5.5),
              "AGGREGATION — When to Combine",
              [
                  "Items may be aggregated only when they have:",
                  "",
                  "  • Similar nature AND",
                  "  • Similar function in the entity's activities",
                  "",
                  "Examples of items that CAN be aggregated:",
                  "  • Multiple revenue streams with same nature",
                  "  • Similar operating costs (same function)",
                  "",
                  "Why it matters:",
                  "  • Prevents over-disclosure / clutter",
                  "  • Reduces information overload for users",
                  "  • Must not obscure material information",
              ],
              header_color=BRIGHT_BLUE, body_size=Pt(12))

    # Right: Disaggregation
    info_card(slide, Inches(6.6), Inches(1.5), Inches(6.4), Inches(5.5),
              "DISAGGREGATION — When to Separate",
              [
                  "Items MUST be disaggregated when they have:",
                  "",
                  "  • Different nature OR different function, AND",
                  "  • Material individually",
                  "",
                  "Examples requiring separate presentation:",
                  "  • Revenue vs. other operating income",
                  "  • Finance income vs. finance costs",
                  "  • Restructuring costs (if material)",
                  "  • Write-downs (vs. routine depreciation)",
                  "",
                  "Minimum line items required:",
                  "  • Revenue (separately from other income)",
                  "  • Share of P&L of associates (if applicable)",
                  "  • Income tax expense",
              ],
              header_color=ROYAL_BLUE, body_size=Pt(12))

    # Bottom: key principle pill
    pill = add_rect(slide, Inches(0.3), Inches(7.1), SLIDE_W - Inches(0.6),
                    Inches(0.48), DARK_NAVY)
    add_text_to_shape(
        pill,
        "Core Principle:  Information must be presented at a level that gives users a faithful and complete picture — neither too detailed nor too condensed.",
        font_size=Pt(12), bold=False, color=WHITE, align=PP_ALIGN.LEFT
    )
    pill.text_frame.margin_left = Inches(0.2)

    add_footer(slide)


def slide_13_cashflows(prs):
    """Changes to the Statement of Cash Flows."""
    slide = blank_slide(prs)
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, LIGHT_GRAY)
    add_slide_header(slide, "Changes to the Statement of Cash Flows",
                     "Improved consistency between P&L categories and cash flow classification")

    # Summary comparison: IAS 7 → IFRS 18
    add_textbox(slide, "Under IAS 7 (Old)",
                Inches(0.3), Inches(1.55), Inches(4.0), Inches(0.4),
                font_size=Pt(14), bold=True, color=RED_ORANGE)
    add_textbox(slide, "Under IFRS 18 (New)",
                Inches(7.0), Inches(1.55), Inches(4.0), Inches(0.4),
                font_size=Pt(14), bold=True, color=GREEN)

    changes = [
        (
            "Interest Paid",
            "Choice: Operating OR Financing",
            "MUST match P&L category\n(Operating or Financing — entity's policy)",
            BRIGHT_BLUE,
        ),
        (
            "Interest Received",
            "Choice: Operating OR Investing",
            "MUST match P&L category\n(Operating or Investing — entity's policy)",
            ROYAL_BLUE,
        ),
        (
            "Dividends Received",
            "Choice: Operating OR Investing",
            "Non-integral: Investing\nIntegral associates: Operating",
            DARK_NAVY,
        ),
        (
            "Dividends Paid",
            "Choice: Operating OR Financing",
            "Financing (equity distribution)\nor Operating if policy consistent with P&L",
            BRIGHT_BLUE,
        ),
        (
            "Taxes Paid",
            "Operating (general rule)\nor allocation across activities",
            "Consistent allocation required;\nenforced link to related P&L items",
            ROYAL_BLUE,
        ),
    ]

    for i, (item, old, new, col) in enumerate(changes):
        top = Inches(2.05) + i * Inches(0.97)
        h   = Inches(0.9)

        # item label
        lbl = add_rect(slide, Inches(0.3), top, Inches(2.5), h, col)
        add_text_to_shape(lbl, item, font_size=Pt(12), bold=True, color=WHITE)

        # old rule
        old_box = add_rect(slide, Inches(2.85), top, Inches(4.2), h, WHITE)
        old_box.line.color.rgb = RED_ORANGE
        old_box.line.width = Pt(1)
        tf1 = old_box.text_frame
        tf1.word_wrap = True
        tf1.margin_left = Inches(0.1)
        tf1.margin_top  = Inches(0.06)
        p1 = tf1.paragraphs[0]
        p1.alignment = PP_ALIGN.LEFT
        r1 = p1.add_run()
        r1.text = old
        r1.font.size = Pt(11.5)
        r1.font.color.rgb = DARK_GRAY

        # arrow
        arr = slide.shapes.add_shape(13,
                                     Inches(7.1), top + Inches(0.3),
                                     Inches(0.3), Inches(0.28))
        fill_shape(arr, GOLD)
        arr.line.fill.background()

        # new rule
        new_box = add_rect(slide, Inches(7.5), top, Inches(5.5), h, WHITE)
        new_box.line.color.rgb = GREEN
        new_box.line.width = Pt(1)
        tf2 = new_box.text_frame
        tf2.word_wrap = True
        tf2.margin_left = Inches(0.1)
        tf2.margin_top  = Inches(0.06)
        p2 = tf2.paragraphs[0]
        p2.alignment = PP_ALIGN.LEFT
        r2 = p2.add_run()
        r2.text = new
        r2.font.size = Pt(11.5)
        r2.font.bold = True
        r2.font.color.rgb = GREEN
        for extra_line in new.split("\n")[1:]:
            p_extra = tf2.add_paragraph()
            p_extra.alignment = PP_ALIGN.LEFT
            r_extra = p_extra.add_run()
            r_extra.text = extra_line
            r_extra.font.size = Pt(11.5)
            r_extra.font.color.rgb = DARK_GRAY

    add_footer(slide)


def slide_14_effective_date(prs):
    """Effective Date & Transition."""
    slide = blank_slide(prs)
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, LIGHT_GRAY)
    add_slide_header(slide, "Effective Date & Transition Requirements",
                     "How and when to adopt IFRS 18")

    # Timeline bar
    milestones = [
        ("May 2024",         "IFRS 18\nIssued",              BRIGHT_BLUE),
        ("2025 – 2026",      "Preparation\n& Planning",       ROYAL_BLUE),
        ("1 Jan 2027",       "Effective\nDate",               GOLD),
        ("2027 Reports",     "First IFRS 18\nStatements",     GREEN),
        ("Ongoing",          "Continuous\nApplication",       DARK_NAVY),
    ]

    tl_top  = Inches(1.6)
    tl_left = Inches(0.5)
    tl_w    = SLIDE_W - Inches(1.0)
    gap     = tl_w / (len(milestones) - 1)

    # Line
    add_rect(slide, tl_left, tl_top + Inches(0.45), tl_w, Inches(0.08), SKY_BLUE)

    for i, (date, label, col) in enumerate(milestones):
        cx = tl_left + i * gap
        # circle
        c = slide.shapes.add_shape(9, cx - Inches(0.35), tl_top, Inches(0.7), Inches(0.7))
        fill_shape(c, col)
        c.line.color.rgb = WHITE
        c.line.width = Pt(2)

        # date label above
        add_textbox(slide, date,
                    cx - Inches(0.9), tl_top - Inches(0.5), Inches(1.8), Inches(0.45),
                    font_size=Pt(10), bold=True, color=col, align=PP_ALIGN.CENTER)

        # event label below
        add_textbox(slide, label,
                    cx - Inches(0.9), tl_top + Inches(0.8), Inches(1.8), Inches(0.6),
                    font_size=Pt(10), bold=False, color=DARK_GRAY, align=PP_ALIGN.CENTER)

    # Transition details – two columns
    left_items = [
        ("Retrospective Application",
         "IFRS 18 is applied RETROSPECTIVELY — comparative information\nmust be restated to comply with the new standard."),
        ("Early Adoption",
         "Early adoption is permitted for annual periods beginning\nbefore 1 January 2027, provided it is disclosed."),
        ("Transition Relief — MPMs",
         "Entities need NOT provide comparative MPM disclosures\nin the year of initial adoption of IFRS 18."),
    ]
    right_items = [
        ("Practical Expedients",
         "The comparative period cash flow statement does NOT\nneed to be restated on first-time application."),
        ("Disclosure on Adoption",
         "Must disclose: nature of changes, reason for early adoption (if applicable), and impact on financial statements."),
        ("IAS 8 Applicability",
         "Standard transition requirements under IAS 8 apply to changes\nin accounting policies resulting from IFRS 18 adoption."),
    ]

    for i, (ttl, body) in enumerate(left_items):
        info_card(slide, Inches(0.3), Inches(3.0) + i * Inches(1.35),
                  Inches(6.1), Inches(1.25),
                  ttl, [body], header_color=BRIGHT_BLUE, body_size=Pt(11.5))

    for i, (ttl, body) in enumerate(right_items):
        info_card(slide, Inches(6.6), Inches(3.0) + i * Inches(1.35),
                  Inches(6.4), Inches(1.25),
                  ttl, [body], header_color=ROYAL_BLUE, body_size=Pt(11.5))

    add_footer(slide)


def slide_15_key_takeaways(prs):
    """Key Takeaways / Summary."""
    slide = blank_slide(prs)
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, DARK_NAVY)
    add_rect(slide, 0, 0, Inches(0.22), SLIDE_H, GOLD)

    # Title
    add_textbox(slide, "Key Takeaways",
                Inches(0.55), Inches(0.3), Inches(12), Inches(0.75),
                font_size=Pt(36), bold=True, color=WHITE, align=PP_ALIGN.LEFT)
    add_rect(slide, Inches(0.55), Inches(1.05), Inches(5.0), Inches(0.06), GOLD)

    takeaways = [
        (
            "01",
            "Structured P&L",
            "IFRS 18 introduces three mandatory income categories — Operating, Investing, and Financing — eliminating inconsistent classification practices.",
            BRIGHT_BLUE,
        ),
        (
            "02",
            "Required Subtotals",
            "Operating Profit and Profit before Financing & Tax are now mandatory subtotals, giving investors a comparable baseline for analysis.",
            GOLD,
        ),
        (
            "03",
            "MPM Transparency",
            "Management-defined Performance Measures (MPMs) must be labelled, explained, reconciled to IFRS subtotals, and never given undue prominence.",
            BRIGHT_BLUE,
        ),
        (
            "04",
            "Cash Flow Consistency",
            "Cash flow classifications for interest and dividends must align with the entity's P&L category choices — no more arbitrary choices.",
            GOLD,
        ),
        (
            "05",
            "Better Aggregation Rules",
            "Clear principles ensure that material items are separately disclosed and similar items are appropriately grouped for readability.",
            BRIGHT_BLUE,
        ),
        (
            "06",
            "Effective 1 Jan 2027",
            "Applies retrospectively from annual periods beginning on or after 1 January 2027. Early adoption is permitted — start planning now.",
            GOLD,
        ),
    ]

    col_w = Inches(5.9)
    row_h = Inches(1.6)
    for i, (num, title, body, col) in enumerate(takeaways):
        col_idx = i % 2
        row_idx = i // 2
        lft = Inches(0.4) + col_idx * (col_w + Inches(0.4))
        top = Inches(1.3) + row_idx * (row_h + Inches(0.1))

        # number badge
        badge = slide.shapes.add_shape(9, lft, top + Inches(0.4),
                                        Inches(0.6), Inches(0.6))
        fill_shape(badge, col)
        badge.line.fill.background()
        add_text_to_shape(badge, num, font_size=Pt(14), bold=True, color=WHITE)

        # title
        add_textbox(slide, title,
                    lft + Inches(0.75), top + Inches(0.38),
                    Inches(4.8), Inches(0.42),
                    font_size=Pt(15), bold=True, color=col, align=PP_ALIGN.LEFT)

        # body
        add_textbox(slide, body,
                    lft + Inches(0.75), top + Inches(0.82),
                    Inches(4.8), Inches(0.75),
                    font_size=Pt(11), bold=False, color=RGBColor(0xCC, 0xCC, 0xCC),
                    align=PP_ALIGN.LEFT)

    add_footer(slide)


def slide_16_thankyou(prs):
    """Thank You / Questions slide."""
    slide = blank_slide(prs)

    # Full dark navy background
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, DARK_NAVY)
    add_rect(slide, 0, 0, Inches(0.22), SLIDE_H, GOLD)

    # Large decorative circle
    c1 = slide.shapes.add_shape(9, Inches(7.5), Inches(-1.0), Inches(7.0), Inches(7.0))
    fill_shape(c1, ROYAL_BLUE)
    c1.line.fill.background()

    c2 = slide.shapes.add_shape(9, Inches(9.0), Inches(2.5), Inches(4.5), Inches(4.5))
    fill_shape(c2, BRIGHT_BLUE)
    c2.line.fill.background()

    # Thank You text
    add_textbox(slide, "Thank You",
                Inches(0.6), Inches(1.8), Inches(7), Inches(1.6),
                font_size=Pt(60), bold=True, color=WHITE, align=PP_ALIGN.LEFT)

    add_rect(slide, Inches(0.6), Inches(3.5), Inches(4.5), Inches(0.06), GOLD)

    add_textbox(slide, "Questions & Discussion",
                Inches(0.6), Inches(3.7), Inches(7.5), Inches(0.6),
                font_size=Pt(20), bold=False, color=SKY_BLUE, align=PP_ALIGN.LEFT)

    add_textbox(slide,
                "IFRS 18 — Presentation and Disclosure in Financial Statements\n"
                "Effective for annual periods beginning on or after 1 January 2027\n"
                "Issued by the International Accounting Standards Board (IASB), May 2024",
                Inches(0.6), Inches(4.5), Inches(8.5), Inches(1.2),
                font_size=Pt(12), bold=False, color=MID_GRAY, align=PP_ALIGN.LEFT)

    # Bottom disclaimer
    disc = add_rect(slide, Inches(0.0), Inches(6.75), SLIDE_W, Inches(0.75),
                    RGBColor(0x12, 0x22, 0x3D))
    add_text_to_shape(disc,
                      "This presentation is prepared for educational and informational purposes. "
                      "Please refer to the full IFRS 18 standard for authoritative guidance.",
                      font_size=Pt(9), color=MID_GRAY, align=PP_ALIGN.CENTER)


# ══════════════════════════════════════════════════════════════════════════════
# Main builder
# ══════════════════════════════════════════════════════════════════════════════

def build_presentation(output_path: str = "IFRS_18_Comprehensive_Presentation.pptx"):
    prs = new_prs()

    slide_01_title(prs)
    slide_02_agenda(prs)
    slide_03_overview(prs)
    slide_04_why_ifrs18(prs)
    slide_05_new_pl_structure(prs)
    slide_06_operating(prs)
    slide_07_investing(prs)
    slide_08_financing(prs)
    slide_09_subtotals(prs)
    slide_10_mpm_intro(prs)
    slide_11_mpm_disclosure(prs)
    slide_12_aggregation(prs)
    slide_13_cashflows(prs)
    slide_14_effective_date(prs)
    slide_15_key_takeaways(prs)
    slide_16_thankyou(prs)

    prs.save(output_path)
    print(f"✅  Presentation saved → {output_path}  ({len(prs.slides)} slides)")


if __name__ == "__main__":
    build_presentation()
