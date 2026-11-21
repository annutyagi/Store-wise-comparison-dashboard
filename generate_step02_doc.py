import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

def create_hmw_doc(filename):
    os.makedirs(os.path.dirname(filename), exist_ok=True)
    doc = docx.Document()
    
    # Page Margins
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)
        
    # Color Palette - Professional Corporate Blue
    COLOR_PRIMARY = RGBColor(26, 54, 93)     # #1A365D - Deep Navy
    COLOR_SECONDARY = RGBColor(43, 108, 176) # #2B6CB0 - Classic Slate Blue
    COLOR_TEXT = RGBColor(45, 55, 72)        # #2D3748 - Dark Charcoal
    COLOR_MUTED = RGBColor(113, 128, 150)    # #718096 - Muted Grey
    COLOR_PRIORITY = RGBColor(197, 48, 48)   # #C53030 - Dark Crimson
    
    # Normal Style
    style_normal = doc.styles['Normal']
    style_normal.font.name = 'Calibri'
    style_normal.font.size = Pt(11)
    style_normal.font.color.rgb = COLOR_TEXT
    style_normal.paragraph_format.line_spacing = 1.15
    style_normal.paragraph_format.space_after = Pt(6)

    def set_cell_background(cell, hex_color):
        shading_xml = f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>'
        cell._tc.get_or_add_tcPr().append(parse_xml(shading_xml))

    def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
        tcPr = cell._tc.get_or_add_tcPr()
        tcMar = OxmlElement('w:tcMar')
        for margin, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
            node = OxmlElement(f'w:{margin}')
            node.set(qn('w:w'), str(val))
            node.set(qn('w:type'), 'dxa')
            tcMar.append(node)
        tcPr.append(tcMar)

    def add_custom_heading(text, level, space_before=14, space_after=6):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.bold = True
        if level == 1:
            run.font.size = Pt(16)
            run.font.color.rgb = COLOR_PRIMARY
        elif level == 2:
            run.font.size = Pt(13)
            run.font.color.rgb = COLOR_SECONDARY
        elif level == 3:
            run.font.size = Pt(11.5)
            run.font.color.rgb = COLOR_PRIMARY
        return p

    def style_table(table, col_widths, col_alignments):
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        # Header formatting
        for idx, cell in enumerate(table.rows[0].cells):
            set_cell_background(cell, '1A365D')
            set_cell_margins(cell, top=120, bottom=120, left=150, right=150)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            for p in cell.paragraphs:
                p.alignment = col_alignments[idx]
                for run in p.runs:
                    run.font.bold = True
                    run.font.color.rgb = RGBColor(255, 255, 255)
                    run.font.size = Pt(10)
        
        # Body rows formatting
        for r_idx, row in enumerate(table.rows[1:], start=1):
            bg_color = 'F7FAFC' if r_idx % 2 == 1 else 'FFFFFF'
            for idx, cell in enumerate(row.cells):
                set_cell_background(cell, bg_color)
                set_cell_margins(cell, top=100, bottom=100, left=150, right=150)
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                for p in cell.paragraphs:
                    p.alignment = col_alignments[idx]
                    for run in p.runs:
                        run.font.size = Pt(9.5)
                        run.font.color.rgb = COLOR_TEXT
                        
        # Widths
        for row in table.rows:
            for idx, width in enumerate(col_widths):
                row.cells[idx].width = Inches(width)

    # Document Header / Banner
    title_p = doc.add_paragraph()
    title_p.paragraph_format.space_before = Pt(0)
    title_p.paragraph_format.space_after = Pt(2)
    t_run = title_p.add_run("STORE WISE SALES COMPARISON DASHBOARD")
    t_run.font.size = Pt(22)
    t_run.font.bold = True
    t_run.font.color.rgb = COLOR_PRIMARY

    sub_p = doc.add_paragraph()
    sub_p.paragraph_format.space_after = Pt(14)
    s_run = sub_p.add_run("Step 2: How Might We (HMW) Design-Thinking Specification")
    s_run.font.size = Pt(12)
    s_run.font.color.rgb = COLOR_MUTED

    # Metadata Table
    meta_table = doc.add_table(rows=4, cols=2)
    meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta_data = [
        [("Learner Name:", "Anu Tyagi"), ("Roll Number:", "2500320140025")],
        [("Batch ID:", "ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01"), ("Institution:", "ABES Engineering College")],
        [("Project Mentor:", "Mr. Dinesh"), ("Document Type:", "Design Thinking & Problem Reframing")],
        [("Upstream Artifact:", "docs/01_Project_Understanding.docx"), ("Target Taxonomy:", "HMW-01 to HMW-12")]
    ]
    for r_idx, row_vals in enumerate(meta_data):
        row = meta_table.rows[r_idx]
        for c_idx, (k, v) in enumerate(row_vals):
            cell = row.cells[c_idx]
            set_cell_background(cell, 'EDF2F7')
            set_cell_margins(cell, top=70, bottom=70, left=120, right=120)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            r_k = p.add_run(f"{k} ")
            r_k.bold = True
            r_k.font.size = Pt(9.5)
            r_k.font.color.rgb = COLOR_PRIMARY
            r_v = p.add_run(v)
            r_v.font.size = Pt(9.5)
            r_v.font.color.rgb = COLOR_TEXT
            cell.width = Inches(3.25)

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # Section 1: Introduction to Design Thinking Reframing
    add_custom_heading("1. Design-Thinking Reframing Methodology", level=1)
    p = doc.add_paragraph(
        "Design thinking reframing bridges business frustrations into actionable opportunity statements. "
        "Rather than jumping prematurely to technological features (e.g., 'building a Streamlit dashboard'), "
        "the 'How Might We' (HMW) paradigm articulates the human, organizational, and operational needs of each "
        "stakeholder identified in Step 1. Every HMW statement is calibrated to be specific, human-centric, "
        "open-ended, and strictly tied to established Operational Questions (OQ-01 through OQ-10)."
    )

    # Section 2: Stakeholder Pain Points
    add_custom_heading("2. Stakeholder Pain Point Analysis", level=1)
    p = doc.add_paragraph(
        "Detailed empathy mapping across all primary and secondary stakeholders reveals critical operational bottlenecks:"
    )

    stakeholder_pain_points = [
        ("SH-01: Regional Sales Director (Executive Sponsor)", [
            "Pain Point PP-01.1: Struggles to benchmark store revenue, gross margins, and sales velocity across disparate geographic territories without wading through bulky monthly spreadsheet dumps.",
            "Pain Point PP-01.2: Lacks proactive early-warning indicators for outlets falling behind monthly sales targets until financial closing, preventing timely commercial interventions.",
            "Pain Point PP-01.3: Inability to correlate store square footage with sales productivity to assess real estate efficiency and capital allocation."
        ]),
        ("SH-02: Store Operations Manager (Store Lead)", [
            "Pain Point PP-02.1: Unable to accurately identify peak trading hours and customer footfall rhythms, leading to understaffed checkout counters and customer churn during rushes.",
            "Pain Point PP-02.2: Cannot easily differentiate weekday vs. weekend revenue velocity and basket sizes to optimize weekend inventory staging.",
            "Pain Point PP-02.3: Blindsided by localized spikes in order cancellations and product returns, causing unmonitored net revenue shrinkage."
        ]),
        ("SH-03: Inventory & Category Merchandiser (Merchandising Lead)", [
            "Pain Point PP-03.1: Lacks visibility into which product categories dominate in specific demographic tiers (Urban vs. Suburban vs. Rural), resulting in mismatched store inventory.",
            "Pain Point PP-03.2: Unable to determine whether promotional discount markdowns stimulate profitable unit volume or simply erode gross merchandise margins.",
            "Pain Point PP-03.3: Lacks store-level SKU performance indicators required to re-allocate slow-moving inventory to high-demand outlets."
        ]),
        ("SH-04: Lead Business Data Analyst (Technical BA / QA)", [
            "Pain Point PP-04.1: Spends substantial hours writing repetitive data cleaning scripts and generating ad-hoc static charts instead of conducting diagnostic analytics.",
            "Pain Point PP-04.2: Lacks automated exploratory tools to visualize non-parametric retail distributions (skewness, median, IQR) for basket values across stores.",
            "Pain Point PP-04.3: Faces difficulty validating transaction anomalies and correlations between tender payment types and ticket sizes."
        ]),
        ("SH-05: Academic Project Mentor - Mr. Dinesh (Evaluator)", [
            "Pain Point PP-05.1: Needs transparent, verifiable traceability linking academic curriculum concepts (EDA, SQL, testing) to realistic business questions.",
            "Pain Point PP-05.2: Requires assurance of code quality, database normalization, and automated verification rather than purely cosmetic front-end visuals."
        ]),
        ("SH-06: Retail End Customer (Consumer - Indirect)", [
            "Pain Point PP-06.1: Experiences long checkout queue delays during unanticipated peak shopping hours.",
            "Pain Point PP-06.2: Encounters out-of-stock essentials in popular categories due to poor store-level inventory replenishment."
        ])
    ]

    for sh_title, pps in stakeholder_pain_points:
        add_custom_heading(sh_title, level=2, space_before=10, space_after=3)
        for pp in pps:
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Inches(0.2)
            p.paragraph_format.space_after = Pt(2)
            parts = pp.split(":", 1)
            r_b = p.add_run(parts[0] + ":")
            r_b.bold = True
            r_b.font.color.rgb = COLOR_PRIMARY
            r_t = p.add_run(parts[1])
            r_t.font.color.rgb = COLOR_TEXT

    # Section 3: Comprehensive HMW Statements Table
    add_custom_heading("3. Structured 'How Might We' (HMW) Statements", level=1)
    p = doc.add_paragraph(
        "The following 12 HMW statements transform identified stakeholder pain points into specific analytical opportunities. "
        "Every statement links directly to upstream Stakeholder IDs (SH) and Operational Questions (OQ)."
    )

    hmw_table = doc.add_table(rows=13, cols=5)
    hmw_headers = ["HMW ID", "How Might We Statement", "Primary Stakeholder", "Core Pain Point Addressed", "Linked OQ"]
    for i, h in enumerate(hmw_headers):
        hmw_table.rows[0].cells[i].paragraphs[0].text = h

    hmw_data = [
        ["HMW-01", "How might we enable regional directors to evaluate and rank store-level revenue and profit margins simultaneously across multiple territories?", "SH-01 (Regional Director)", "PP-01.1 (Static reporting friction)", "OQ-01"],
        ["HMW-02", "How might we help retail executives benchmark floor-space productivity (sales/sq.ft.) and average transaction value across stores of varying physical sizes?", "SH-01 (Regional Director)", "PP-01.3 (Space efficiency blindness)", "OQ-02"],
        ["HMW-03", "How might we assist category merchandisers in identifying which product categories drive the highest revenue across Urban, Suburban, and Rural store tiers?", "SH-03 (Merchandiser)", "PP-03.1 (Demographic inventory mismatch)", "OQ-03"],
        ["HMW-04", "How might we provide commercial leadership with dynamic month-over-month sales velocity trends across regional store clusters before the fiscal month closes?", "SH-01 (Regional Director)", "PP-01.1 (Delayed trend visibility)", "OQ-04"],
        ["HMW-05", "How might we empower store managers to optimize checkout staffing and shift allocation by visualizing peak hourly customer transaction windows?", "SH-02 (Store Manager)", "PP-02.1 (Unpredicted peak rush hours)", "OQ-10"],
        ["HMW-06", "How might we help store operations teams tailor inventory staging and weekend operational readiness by distinguishing weekday versus weekend demand patterns?", "SH-02 (Store Manager)", "PP-02.2 (Weekend demand volatility)", "OQ-05"],
        ["HMW-07", "How might we alert store leadership and category leads to localized spikes in product returns and order cancellations to curb net revenue leakage?", "SH-02 (Store Manager), SH-03", "PP-02.3 (Unmonitored return spikes)", "OQ-06"],
        ["HMW-08", "How might we enable merchandisers to determine whether promotional discounting genuinely expands sales volume or merely erodes product profit margins?", "SH-03 (Merchandiser)", "PP-03.2 (Discount elasticity uncertainty)", "OQ-07"],
        ["HMW-09", "How might we provide proactive alerts and variance tracking for stores falling behind budgeted monthly sales quotas so leadership can intervene rapidly?", "SH-01 (Regional Director)", "PP-01.2 (Late target shortfall awareness)", "OQ-08"],
        ["HMW-10", "How might we help operations and finance leads analyze customer payment preferences (UPI, Card, Cash) and their correlation with overall basket size?", "SH-04 (Analyst), SH-02", "PP-04.3 (Tender type basket opacity)", "OQ-09"],
        ["HMW-11", "How might we eliminate repetitive manual data wrangling for analysts while providing rigorous statistical distribution metrics (IQR, skewness, median)?", "SH-04 (Lead Analyst)", "PP-04.1, PP-04.2 (Manual ETL & ad-hoc stats)", "OQ-02, OQ-09"],
        ["HMW-12", "How might we assure academic mentors and evaluators of end-to-end data integrity, architectural rigor, and statistical validity across the analytics pipeline?", "SH-05 (Academic Mentor)", "PP-05.1, PP-05.2 (Academic & technical rigor)", "OQ-01 to OQ-10"]
    ]

    for r_idx, row_vals in enumerate(hmw_data, start=1):
        for c_idx, val in enumerate(row_vals):
            hmw_table.rows[r_idx].cells[c_idx].paragraphs[0].text = val
    style_table(hmw_table, [0.8, 2.4, 1.2, 1.4, 0.7], [WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER])

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # Section 4: Top 5 Priority HMW Statements
    add_custom_heading("4. Top 5 Priority HMW Statements & Justification", level=1)
    p = doc.add_paragraph(
        "To establish maximum business and technical impact for the initial product increments, the top 5 HMW statements "
        "are designated as core architectural priorities:"
    )

    prio_table = doc.add_table(rows=6, cols=4)
    prio_headers = ["Priority Rank", "HMW ID", "Core Reframed Challenge", "Strategic Justification / Value Rationale"]
    for i, h in enumerate(prio_headers):
        prio_table.rows[0].cells[i].paragraphs[0].text = h

    prio_data = [
        ["PRIORITY 1", "HMW-01", "Comparative Multi-Store Revenue & Margin Benchmarking", "Fundamental core requirement; without normalized multi-store sales and profitability ranking, no macro-level commercial decisions or resource allocations can occur."],
        ["PRIORITY 2", "HMW-09", "Target vs. Actual Sales Quota Variance Tracking", "Directly addresses revenue leakage by providing actionable early-warning alerts for underperforming outlets before monthly financial close."],
        ["PRIORITY 3", "HMW-05", "Peak Hourly Trading Windows & Footfall Optimization", "Delivers immediate daily operational ROI at the store level by synchronizing staff scheduling with empirical transaction rushes."],
        ["PRIORITY 4", "HMW-03", "Demographic Store-Tier Category Revenue Alignment", "Maximizes inventory turnover and gross margins by eliminating category mismatches across Urban, Suburban, and Rural store footprints."],
        ["PRIORITY 5", "HMW-08", "Discount Elasticity vs. Gross Margin Preservation", "Protects bottom-line enterprise profitability by statistically uncovering whether markdown strategies create accretive volume or margin destruction."]
    ]

    for r_idx, row_vals in enumerate(prio_data, start=1):
        for c_idx, val in enumerate(row_vals):
            prio_table.rows[r_idx].cells[c_idx].paragraphs[0].text = val
    style_table(prio_table, [1.0, 0.8, 2.0, 2.7], [WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT])

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # Section 5: OQ Coverage Verification
    add_custom_heading("5. Operational Question (OQ) Coverage Audit", level=1)
    p = doc.add_paragraph(
        "Formal audit verifying that 100% of Operational Questions (OQ-01 through OQ-10) defined in Step 1 "
        "are thoroughly addressed by the HMW framework:"
    )

    cov_table = doc.add_table(rows=11, cols=4)
    cov_headers = ["OQ ID", "Operational Question Summary", "Addressing HMW Statement(s)", "Coverage Status"]
    for i, h in enumerate(cov_headers):
        cov_table.rows[0].cells[i].paragraphs[0].text = h

    cov_data = [
        ["OQ-01", "Highest/lowest gross sales revenue & profit margin across stores", "HMW-01 (Priority 1), HMW-12", "100% Fully Covered"],
        ["OQ-02", "Store comparison in average transaction value & sales velocity / sq.ft.", "HMW-02, HMW-11", "100% Fully Covered"],
        ["OQ-03", "Category revenue contribution across store tiers (Urban/Suburban/Rural)", "HMW-03 (Priority 4)", "100% Fully Covered"],
        ["OQ-04", "Month-over-month (MoM) sales trends across regional clusters", "HMW-04", "100% Fully Covered"],
        ["OQ-05", "Distribution of sales performance across weekdays vs. weekends", "HMW-06", "100% Fully Covered"],
        ["OQ-06", "Stores with highest return/cancellation rates and net revenue impact", "HMW-07", "100% Fully Covered"],
        ["OQ-07", "Discount percentage correlation with unit volume and gross margin", "HMW-08 (Priority 5)", "100% Fully Covered"],
        ["OQ-08", "Underperforming stores consistently missing monthly targets", "HMW-09 (Priority 2)", "100% Fully Covered"],
        ["OQ-09", "Payment method distribution across stores and basket size correlation", "HMW-10, HMW-11", "100% Fully Covered"],
        ["OQ-10", "Peak trading hours and temporal traffic distribution patterns", "HMW-05 (Priority 3)", "100% Fully Covered"]
    ]

    for r_idx, row_vals in enumerate(cov_data, start=1):
        for c_idx, val in enumerate(row_vals):
            cov_table.rows[r_idx].cells[c_idx].paragraphs[0].text = val
    style_table(cov_table, [0.8, 2.7, 1.8, 1.2], [WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER])

    doc.save(filename)
    print(f"HMW Document successfully created at: {filename}")

if __name__ == "__main__":
    import sys
    target = sys.argv[1] if len(sys.argv) > 1 else "docs/02_HMW_Statements.docx"
    create_hmw_doc(target)
