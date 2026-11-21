import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

def create_needs_analysis_doc(filename):
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
    
    # Style Configurations
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
    s_run = sub_p.add_run("Step 3: Needs Analysis & Business Requirements (BR) Specification")
    s_run.font.size = Pt(12)
    s_run.font.color.rgb = COLOR_MUTED

    # Metadata Box Table
    meta_table = doc.add_table(rows=4, cols=2)
    meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta_data = [
        [("Learner Name:", "Anu Tyagi"), ("Roll Number:", "2500320140025")],
        [("Batch ID:", "ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01"), ("Institution:", "ABES Engineering College")],
        [("Project Mentor:", "Mr. Dinesh"), ("Document Type:", "Business Requirements Analysis (BR)")],
        [("Upstream Artifacts:", "Step 1 (Understanding) & Step 2 (HMW)"), ("Requirement IDs:", "BR-01 through BR-14")]
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

    # Section 1: Methodology & Principles
    add_custom_heading("1. Business Needs Analysis Methodology", level=1)
    p = doc.add_paragraph(
        "A rigorous Needs Analysis translates design-thinking opportunities (HMW statements) into concrete Business Requirements (BR). "
        "In strict adherence to professional business analysis best practices (BABOK / IEEE 830), this document articulates WHAT capabilities "
        "the business requires and the measurable BENEFITS realized, deliberately abstracting away technology implementation details "
        "(zero mention of Streamlit, Python libraries, or SQL dialect syntax)."
    )
    p_syn = doc.add_paragraph(
        "Standard Canonical Syntax: 'The business needs to <capability> so that <benefit>.'"
    )
    p_syn.runs[0].italic = True
    p_syn.runs[0].bold = True

    # Section 2: Detailed Business Requirements Table
    add_custom_heading("2. Business Requirements Matrix", level=1)
    p = doc.add_paragraph(
        "The following matrix specifies 14 distinct Business Requirements categorized using MoSCoW prioritization "
        "(64.3% Must Have, ensuring solid alignment with the ~60% guideline). Each requirement is paired with a concrete, testable measurement method."
    )

    br_table = doc.add_table(rows=15, cols=6)
    br_headers = ["BR ID", "Requirement Statement (Capability & Business Benefit)", "Primary Stakeholder", "Source HMW", "MoSCoW Priority", "Measurement & Verification Method"]
    for i, h in enumerate(br_headers):
        br_table.rows[0].cells[i].paragraphs[0].text = h

    br_data = [
        [
            "BR-01",
            "The business needs to benchmark and rank store-level gross sales revenue and profit margins across geographic territories so that regional directors can identify top performers and margin-diluting stores.",
            "SH-01 (Regional Director)",
            "HMW-01",
            "Must",
            "Verified by comparing multi-store ranked revenue and margin calculations against verified reference test totals; 100% mathematical accuracy."
        ],
        [
            "BR-02",
            "The business needs to calculate and compare sales velocity per square foot and average transaction value across all retail store sizes so that executives can evaluate retail real-estate productivity and store expansion ROI.",
            "SH-01 (Regional Director)",
            "HMW-02",
            "Must",
            "Verified by formula audit (Total Revenue / Total Sq.Ft. and Total Revenue / Total Orders) matching across all store size cohorts."
        ],
        [
            "BR-03",
            "The business needs to analyze product category sales contribution segmented by store demographic tiers (Urban, Suburban, Rural) so that merchandisers can eliminate stock-outs and align inventory stocking to regional demographics.",
            "SH-03 (Merchandiser)",
            "HMW-03",
            "Must",
            "Verified by category revenue percentage breakdown summing exactly to 100% across each tier with zero unclassified categories."
        ],
        [
            "BR-04",
            "The business needs to track month-over-month (MoM) sales velocity trends across regional store clusters so that commercial leadership can detect seasonal demand shifts and take proactive corrective actions prior to month-end.",
            "SH-01 (Regional Director)",
            "HMW-04",
            "Must",
            "Verified by calculating percentage variance between consecutive calendar months with automated reconciliation against fiscal records."
        ],
        [
            "BR-05",
            "The business needs to identify hourly customer transaction volumes and peak sales trading windows for each store so that store managers can optimize staff shift scheduling and reduce customer checkout wait times.",
            "SH-02 (Store Manager)",
            "HMW-05",
            "Must",
            "Verified by hourly volume distribution test matching timestamp aggregation for 24 distinct hourly buckets (00:00 to 23:00)."
        ],
        [
            "BR-06",
            "The business needs to separate and contrast weekday versus weekend sales volumes and basket sizes so that store operations teams can optimize weekend inventory staging and logistics schedules.",
            "SH-02 (Store Manager)",
            "HMW-06",
            "Should",
            "Verified by calendar day-type classification test verifying that Monday-Friday and Saturday-Sunday aggregates partition correctly."
        ],
        [
            "BR-07",
            "The business needs to monitor product return and order cancellation rates across stores and product categories so that management can identify localized quality issues, reduce return rates, and prevent net revenue leakage.",
            "SH-02 (Store Manager), SH-03",
            "HMW-07",
            "Must",
            "Verified by net sales reconciliation: Net Sales = Gross Sales - Returns - Cancellations across all store records."
        ],
        [
            "BR-08",
            "The business needs to evaluate the relationship between discount percentages, unit sales volume, and gross profit margins across product lines so that merchandisers can design margin-protective pricing and promotional campaigns.",
            "SH-03 (Merchandiser)",
            "HMW-08",
            "Must",
            "Verified by correlation coefficient testing and margin variance verification across distinct discount tiers (0%, 5%, 10%, 15%, 20%+)."
        ],
        [
            "BR-09",
            "The business needs to measure actual store sales against budgeted monthly revenue quotas and flag variance thresholds exceeding +/- 5% so that regional management can execute timely sales interventions.",
            "SH-01 (Regional Director)",
            "HMW-09",
            "Must",
            "Verified by variance rule assertion: Variance = Actual - Target; automated trigger test when absolute variance exceeds 5%."
        ],
        [
            "BR-10",
            "The business needs to aggregate customer payment method adoption (UPI, Credit Card, Cash) and analyze their relationship with basket size across store locations so that finance and store managers can minimize payment friction and transaction charges.",
            "SH-04 (Analyst), SH-02",
            "HMW-10",
            "Should",
            "Verified by tender split reconciliation (UPI + Card + Cash = Total Transactions) and basket size distribution verification."
        ],
        [
            "BR-11",
            "The business needs to automate statistical distribution analysis (mean, median, interquartile range, skewness, outlier detection) for transaction values so that business analysts can instantly evaluate customer spending patterns without manual calculation errors.",
            "SH-04 (Lead Analyst)",
            "HMW-11",
            "Should",
            "Verified by parametric and non-parametric statistical unit test suite confirming mathematical alignment with standard formulas."
        ],
        [
            "BR-12",
            "The business needs to enforce data validation rules, relational referential integrity, and reconciliation controls across all transactional records so that stakeholders and evaluators have guaranteed auditability and trust in the analytical results.",
            "SH-05 (Academic Mentor)",
            "HMW-12",
            "Must",
            "Verified by automated data integrity test suite proving zero orphaned records, zero null primary keys, and exact mathematical row reconciliations."
        ],
        [
            "BR-13",
            "The business needs to export aggregated comparative store performance summaries into tabular spreadsheet formats so that business users can share findings with external executive partners.",
            "SH-01 (Regional Director), SH-04",
            "HMW-01, HMW-11",
            "Could",
            "Verified by exporting tabular data extracts and validating column structure, data types, and row completeness against on-screen values."
        ],
        [
            "BR-14",
            "The business needs to provide multi-currency international sales conversion and real-time live currency exchange rate updates for cross-border operations.",
            "SH-01 (Regional Director)",
            "HMW-01",
            "Won't",
            "Formally deferred: Project is strictly scoped to single-currency domestic retail (INR / ₹) as per Assumption AS-02."
        ]
    ]

    for r_idx, row_vals in enumerate(br_data, start=1):
        for c_idx, val in enumerate(row_vals):
            br_table.rows[r_idx].cells[c_idx].paragraphs[0].text = val
    style_table(br_table, [0.7, 2.5, 1.1, 0.7, 0.7, 1.5], [WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT])

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # Section 3: Data Entities and Attributes Required per Requirement
    add_custom_heading("3. Data Entities and Attributes Required per Requirement", level=1)
    p = doc.add_paragraph(
        "To guarantee that every business capability can be implemented deterministically, the data attributes required by each "
        "business requirement are formally specified below:"
    )

    data_table = doc.add_table(rows=15, cols=3)
    data_headers = ["BR ID", "Required Business Entities", "Mandatory Data Attributes"]
    for i, h in enumerate(data_headers):
        data_table.rows[0].cells[i].paragraphs[0].text = h

    data_mapping = [
        ["BR-01", "Stores, Orders, Order_Items, Products", "store_id, store_name, region_name, total_amount, cost_price, selling_price, order_date"],
        ["BR-02", "Stores, Orders", "store_id, store_name, store_size_sqft, order_id, total_amount"],
        ["BR-03", "Stores, Products, Order_Items, Orders", "store_tier (Urban/Suburban/Rural), category_name, sub_category, quantity, item_subtotal"],
        ["BR-04", "Orders, Stores", "order_date (Year-Month), region_name, store_id, total_amount"],
        ["BR-05", "Orders, Stores", "store_id, order_time (Hour: 00-23), order_id, total_amount"],
        ["BR-06", "Orders, Stores", "order_date, day_of_week, is_weekend_flag, store_id, total_amount"],
        ["BR-07", "Orders, Order_Items, Products", "order_status (Completed/Returned/Cancelled), return_amount, refund_reason, store_id"],
        ["BR-08", "Order_Items, Products", "discount_percentage, unit_price, cost_price, quantity, gross_profit_margin, store_id"],
        ["BR-09", "Stores, Sales_Targets, Orders", "store_id, target_month_year, target_revenue_inr, actual_revenue_inr, variance_amount"],
        ["BR-10", "Orders, Payments", "payment_method (UPI/Credit Card/Debit Card/Cash), order_id, total_amount, store_id"],
        ["BR-11", "Orders, Stores", "order_id, total_amount, store_id, transaction_timestamp (for statistical series)"],
        ["BR-12", "All Relational Tables", "primary_keys, foreign_key_constraints, not_null_constraints, audit_timestamps"],
        ["BR-13", "Aggregated Query Results", "aggregated_kpi_columns, store_metadata, date_range_filter_parameters"],
        ["BR-14", "Currency Exchange Rates", "exchange_rate_inr_usd, conversion_timestamp (Out-of-Scope / Deferred)"]
    ]

    for r_idx, row_vals in enumerate(data_mapping, start=1):
        for c_idx, val in enumerate(row_vals):
            data_table.rows[r_idx].cells[c_idx].paragraphs[0].text = val
    style_table(data_table, [0.8, 2.2, 3.5], [WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT])

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # Section 4: Gap Analysis
    add_custom_heading("4. Gap Analysis (Bidirectional Coverage Audit)", level=1)
    p = doc.add_paragraph(
        "A rigorous bidirectional gap analysis ensures complete alignment between design-thinking problem statements (HMWs) "
        "and business requirements (BRs), verifying that no requirement exists without an upstream need and no need is ignored."
    )

    add_custom_heading("4.1 Forward Check: HMW Statements with No Business Requirement", level=2)
    p_fwd = doc.add_paragraph(
        "Audit Result: ZERO GAPS. Every HMW statement from HMW-01 through HMW-12 maps directly to at least one Business Requirement:\n"
        "• HMW-01 -> BR-01, BR-13, BR-14\n"
        "• HMW-02 -> BR-02\n"
        "• HMW-03 -> BR-03\n"
        "• HMW-04 -> BR-04\n"
        "• HMW-05 -> BR-05\n"
        "• HMW-06 -> BR-06\n"
        "• HMW-07 -> BR-07\n"
        "• HMW-08 -> BR-08\n"
        "• HMW-09 -> BR-09\n"
        "• HMW-10 -> BR-10\n"
        "• HMW-11 -> BR-11, BR-13\n"
        "• HMW-12 -> BR-12"
    )

    add_custom_heading("4.2 Backward Check: Business Requirements with No Source HMW", level=2)
    p_bwd = doc.add_paragraph(
        "Audit Result: ZERO ORPHANED REQUIREMENTS. Every Business Requirement originates from an approved HMW statement. "
        "No requirements have been artificially introduced outside stakeholder needs."
    )

    add_custom_heading("4.3 Testability & Verification Audit", level=2)
    p_test = doc.add_paragraph(
        "Audit Result: 100% TESTABLE. Every requirement possesses a deterministic quantitative measurement method "
        "(reconciliation balancing, statistical parameter matching, or formula verification), ensuring seamless progression "
        "to Functional Requirements (Step 6) and Automated Test Cases (Step 13)."
    )

    doc.save(filename)
    print(f"Needs Analysis Document successfully created at: {filename}")

if __name__ == "__main__":
    import sys
    target = sys.argv[1] if len(sys.argv) > 1 else "docs/03_Needs_Analysis.docx"
    create_needs_analysis_doc(target)
