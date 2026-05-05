import collections
import collections.abc
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

def create_ifrs18_presentation():
    prs = Presentation()

    # Define Corporate Colors (Big 4 Style: Dark Blue, Light Gray, Navy)
    NAVY = RGBColor(0, 32, 96)
    GRAY = RGBColor(242, 242, 242)
    DARK_GRAY = RGBColor(89, 89, 89)
    ACCENT_BLUE = RGBColor(0, 112, 192)

    def add_footer(slide, slide_number):
        # Footer Note
        txBox = slide.shapes.add_textbox(Inches(0.5), Inches(7.1), Inches(4), Inches(0.3))
        tf = txBox.text_frame
        p = tf.paragraphs[0]
        p.text = "Confidential - For Internal Use Only | IFRS 18 Executive Briefing"
        p.font.size = Pt(8)
        p.font.color.rgb = DARK_GRAY

        # Page Number
        txBoxNum = slide.shapes.add_textbox(Inches(9), Inches(7.1), Inches(0.5), Inches(0.3))
        tfNum = txBoxNum.text_frame
        pNum = tfNum.paragraphs[0]
        pNum.text = str(slide_number)
        pNum.font.size = Pt(8)
        pNum.font.color.rgb = DARK_GRAY
        pNum.alignment = PP_ALIGN.RIGHT

    # --- Slide 1: Title Slide ---
    slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(slide_layout)
    title = slide.shapes.title
    subtitle = slide.placeholders[1]

    title.text = "IFRS 18: Presentation and Disclosure in Financial Statements"
    subtitle.text = "Enhancing Comparability and Transparency in Financial Reporting\nMay 5, 2026 | Executive Briefing"
    
    # Styling Title
    title.text_frame.paragraphs[0].font.color.rgb = NAVY
    title.text_frame.paragraphs[0].font.bold = True
    
    add_footer(slide, 1)

    # --- Slide 2: Executive Summary ---
    slide_layout = prs.slide_layouts[1]
    slide = prs.slides.add_slide(slide_layout)
    slide.shapes.title.text = "Executive Summary: Why IFRS 18 Matters"
    
    body = slide.shapes.placeholders[1]
    tf = body.text_frame
    tf.text = "IFRS 18 replaces IAS 1 to drive consistency in how financial performance is reported."
    
    points = [
        "Structure: New P&L categories (Operating, Investing, Financing) to eliminate 'non-GAAP' inconsistencies.",
        "Management Performance Measures (MPMs): Now brought into audited financial statements for the first time.",
        "Aggregation/Disaggregation: Strict new requirements to prevent 'obscuring' material data through generic labeling."
    ]
    
    for point in points:
        p = tf.add_paragraph()
        p.text = point
        p.level = 1

    add_footer(slide, 2)

    # --- Slide 3: New Structure of the Statement of Profit or Loss ---
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "The New Mandatory P&L Structure"
    
    content = slide.shapes.placeholders[1].text_frame
    content.text = "Total income and expenses are now allocated into three defined categories:"
    
    cats = [
        "Operating Category: The 'residual' bucket—includes all items not classified elsewhere.",
        "Investing Category: Income/expenses from assets that generate returns independently (e.g., JVs, Associates).",
        "Financing Category: Costs related to raising capital and interest on liabilities."
    ]
    
    for cat in cats:
        p = content.add_paragraph()
        p.text = cat
        p.level = 1
    
    # Visual Box for 'Operating Profit'
    left = Inches(1)
    top = Inches(5)
    width = Inches(8)
    height = Inches(1)
    shape = slide.shapes.add_shape(1, left, top, width, height) # 1 is Rectangle
    shape.fill.solid()
    shape.fill.foreground_color.rgb = NAVY
    shape.text = "MANDATORY SUBTOTAL: OPERATING PROFIT"
    
    add_footer(slide, 3)

    # --- Slide 4: Management Performance Measures (MPMs) ---
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Management Performance Measures (MPMs)"
    
    content = slide.shapes.placeholders[1].text_frame
    content.text = "Bringing non-GAAP measures into the audited financial statements."
    
    items = [
        "Definition: Subtotals of income/expenses used in public communications outside financial statements.",
        "Audit Requirement: MPMs must be disclosed in a single, dedicated note within the financial statements.",
        "Mandatory Reconciliation: Companies must reconcile the MPM to the most comparable IFRS subtotal."
    ]
    
    for item in items:
        p = content.add_paragraph()
        p.text = item
        p.level = 1

    add_footer(slide, 4)

    # --- Slide 5: Aggregation and Disaggregation ---
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Aggregation & Disaggregation Principles"
    
    content = slide.shapes.placeholders[1].text_frame
    content.text = "Reducing the 'noise' in financial reporting."
    
    rules = [
        "Elimination of generic labels: Avoid 'Other' unless the items are truly immaterial.",
        "Materiality Focus: If an item has different characteristics, it must be presented separately.",
        "Location of disclosure: Primary statements must provide the most relevant view; notes provide the detail."
    ]
    
    for rule in rules:
        p = content.add_paragraph()
        p.text = rule
        p.level = 1

    add_footer(slide, 5)

    # --- Slide 6: Implementation Roadmap ---
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Implementation Roadmap (2024–2027)"
    
    content = slide.shapes.placeholders[1].text_frame
    
    phases = [
        "2024 - 2025: Impact Assessment & Gap Analysis (Systems, Data, MPMs).",
        "2026: Parallel Reporting & Restatement of Comparatives.",
        "Jan 1, 2027: Effective Date (First IFRS 18 Compliant Annual Report)."
    ]
    
    for phase in phases:
        p = content.add_paragraph()
        p.text = phase
        p.level = 0
        p.space_before = Pt(12)

    add_footer(slide, 6)

    # Save
    prs.save('IFRS_18_Executive_Presentation.pptx')

if __name__ == "__main__":
    create_ifrs18_presentation()
