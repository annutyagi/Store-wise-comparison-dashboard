import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

def create_user_stories_doc(filename):
    os.makedirs(os.path.dirname(filename), exist_ok=True)
    doc = docx.Document()
    
    # Page Margins
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)
        
    COLOR_PRIMARY = RGBColor(26, 54, 93)     # #1A365D - Deep Navy
    COLOR_SECONDARY = RGBColor(43, 108, 176) # #2B6CB0 - Classic Slate Blue
    COLOR_TEXT = RGBColor(45, 55, 72)        # #2D3748 - Dark Charcoal
    COLOR_MUTED = RGBColor(113, 128, 150)    # #718096 - Muted Grey
    
    style_normal = doc.styles['Normal']
    style_normal.font.name = 'Calibri'
    style_normal.font.size = Pt(10.5)
    style_normal.font.color.rgb = COLOR_TEXT
    style_normal.paragraph_format.line_spacing = 1.15
    style_normal.paragraph_format.space_after = Pt(5)

    def set_cell_background(cell, hex_color):
        shading_xml = f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>'
        cell._tc.get_or_add_tcPr().append(parse_xml(shading_xml))

    def set_cell_margins(cell, top=100, bottom=100, left=130, right=130):
        tcPr = cell._tc.get_or_add_tcPr()
        tcMar = OxmlElement('w:tcMar')
        for margin, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
            node = OxmlElement(f'w:{margin}')
            node.set(qn('w:w'), str(val))
            node.set(qn('w:type'), 'dxa')
            tcMar.append(node)
        tcPr.append(tcMar)

    def add_custom_heading(text, level, space_before=14, space_after=5):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.bold = True
        if level == 1:
            run.font.size = Pt(15)
            run.font.color.rgb = COLOR_PRIMARY
        elif level == 2:
            run.font.size = Pt(12)
            run.font.color.rgb = COLOR_SECONDARY
        elif level == 3:
            run.font.size = Pt(10.5)
            run.font.color.rgb = COLOR_PRIMARY
        return p

    def style_table(table, col_widths, col_alignments):
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        for idx, cell in enumerate(table.rows[0].cells):
            set_cell_background(cell, '1A365D')
            set_cell_margins(cell, top=110, bottom=110, left=130, right=130)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            for p in cell.paragraphs:
                p.alignment = col_alignments[idx]
                for run in p.runs:
                    run.font.bold = True
                    run.font.color.rgb = RGBColor(255, 255, 255)
                    run.font.size = Pt(9.0)
        
        for r_idx, row in enumerate(table.rows[1:], start=1):
            bg_color = 'F7FAFC' if r_idx % 2 == 1 else 'FFFFFF'
            for idx, cell in enumerate(row.cells):
                set_cell_background(cell, bg_color)
                set_cell_margins(cell, top=80, bottom=80, left=130, right=130)
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                for p in cell.paragraphs:
                    p.alignment = col_alignments[idx]
                    for run in p.runs:
                        run.font.size = Pt(8.5)
                        run.font.color.rgb = COLOR_TEXT
                        
        for row in table.rows:
            for idx, width in enumerate(col_widths):
                row.cells[idx].width = Inches(width)

    # Title & Banner
    title_p = doc.add_paragraph()
    title_p.paragraph_format.space_before = Pt(0)
    title_p.paragraph_format.space_after = Pt(2)
    t_run = title_p.add_run("STORE WISE SALES COMPARISON DASHBOARD")
    t_run.font.size = Pt(22)
    t_run.font.bold = True
    t_run.font.color.rgb = COLOR_PRIMARY

    sub_p = doc.add_paragraph()
    sub_p.paragraph_format.space_after = Pt(14)
    s_run = sub_p.add_run("Step 11: Agile User Stories & Acceptance Criteria Specification")
    s_run.font.size = Pt(12)
    s_run.font.color.rgb = COLOR_MUTED

    # Metadata Block
    meta_table = doc.add_table(rows=4, cols=2)
    meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta_data = [
        [("Learner Name:", "Anu Tyagi"), ("Roll Number:", "2500320140025")],
        [("Batch ID:", "ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01"), ("Institution:", "ABES Engineering College")],
        [("Project Mentor:", "Mr. Dinesh"), ("Document Type:", "Agile User Stories & Sprint Plan")],
        [("Total User Stories:", "16 Stories across 5 Epics"), ("Total Story Points:", "44 Points (Fibonacci Sizing)")]
    ]
    for r_idx, row_vals in enumerate(meta_data):
        row = meta_table.rows[r_idx]
        for c_idx, (k, v) in enumerate(row_vals):
            cell = row.cells[c_idx]
            set_cell_background(cell, 'EDF2F7')
            set_cell_margins(cell, top=60, bottom=60, left=110, right=110)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            r_k = p.add_run(f"{k} ")
            r_k.bold = True
            r_k.font.size = Pt(9.0)
            r_k.font.color.rgb = COLOR_PRIMARY
            r_v = p.add_run(v)
            r_v.font.size = Pt(9.0)
            r_v.font.color.rgb = COLOR_TEXT
            cell.width = Inches(3.25)

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # Section 1: Executive Summary & Agile Framework
    add_custom_heading("1. Agile User Stories Framework", level=1)
    doc.add_paragraph(
        "User stories decompose high-level business requirements (BR) and technical features (F) into granular, human-centered "
        "work increments. Each story follows the canonical Agile formulation: 'As a <persona>, I want <goal> so that <benefit>' "
        "and is paired with precise Given-When-Then (Gherkin syntax) acceptance criteria. Every story is estimated using Fibonacci "
        "points (1, 2, 3, 5) ensuring that each increment is buildable in under half a day."
    )

    # Section 2: Comprehensive User Stories Table
    add_custom_heading("2. Master User Stories Matrix", level=1)
    
    us_table = doc.add_table(rows=17, cols=7)
    us_headers = ["US ID", "Primary Persona", "User Story Formulation (Goal & Benefit)", "Feature ID", "Linked BR", "Priority", "Points"]
    for i, h in enumerate(us_headers):
        us_table.rows[0].cells[i].paragraphs[0].text = h

    us_data = [
        ["US-01", "SH-04 (Lead Analyst)", "As a Lead Analyst, I want to load normalized retail tables via cached read-only SQLite connections so that queries run in under 50ms without locks.", "F-06", "BR-12", "Must", "3"],
        ["US-02", "SH-05 (Academic Mentor)", "As an Academic Mentor, I want automated validation checks to verify primary key uniqueness and foreign key integrity so that analytics are audit-proof.", "F-06", "BR-12", "Must", "2"],
        ["US-03", "SH-01 (Regional Director)", "As a Regional Director, I want an executive summary banner with 4 Tailwind KPI cards (Revenue, Orders, ATV, Margin) so that I can gauge business health in seconds.", "F-01", "BR-01", "Must", "3"],
        ["US-04", "SH-01 (Regional Director)", "As a Regional Director, I want global sidebar filters (Date range, Region, Store tier) so that all dashboard visuals slice dynamically to my selected cohort.", "F-01, F-03", "BR-01, BR-04", "Must", "3"],
        ["US-05", "SH-01 (Regional Director)", "As a Regional Director, I want to rank stores by gross revenue and profit margin on a dual-axis chart so that I can isolate top performers from margin-diluting outlets.", "F-01", "BR-01", "Must", "3"],
        ["US-06", "SH-01 (Regional Director)", "As a Regional Director, I want a scatter plot of sales velocity per sq.ft. against average transaction value so that I can evaluate floor-space capital productivity.", "F-01", "BR-02", "Must", "3"],
        ["US-07", "SH-03 (Merchandiser)", "As a Category Merchandiser, I want a 100% stacked bar chart of category revenue share across store tiers so that I can optimize inventory allocation by demographics.", "F-02", "BR-03", "Must", "3"],
        ["US-08", "SH-03 (Merchandiser)", "As a Category Merchandiser, I want correlation and OLS regression analysis between discount % and gross margin so that I can prevent margin-destroying promotions.", "F-02", "BR-08", "Must", "5"],
        ["US-09", "SH-01 (Regional Director)", "As a Regional Director, I want to track month-over-month (MoM) regional sales velocity trends so that I can detect seasonal trajectory shifts before month-end.", "F-03", "BR-04", "Must", "3"],
        ["US-10", "SH-02 (Store Manager)", "As a Store Manager, I want a 24-hour trading rush heatmap across stores so that I can optimize cashier staffing ahead of peak customer footfall windows.", "F-03", "BR-05", "Must", "3"],
        ["US-11", "SH-02 (Store Manager)", "As a Store Manager, I want to contrast weekday vs. weekend sales volumes, basket size, and item count so that I can tailor weekend logistics replenishment.", "F-03", "BR-06", "Should", "2"],
        ["US-12", "SH-02 (Store Manager)", "As a Store Manager, I want to audit returned and canceled orders alongside a Pareto defect chart so that I can mitigate net revenue leakage.", "F-04", "BR-07", "Must", "3"],
        ["US-13", "SH-01 (Regional Director)", "As a Regional Director, I want monthly target sales variance tracking with amber/red alerts for variances exceeding +/- 5% so that I can intervene with lagging stores.", "F-04", "BR-09", "Must", "3"],
        ["US-14", "SH-04 (Lead Analyst)", "As a Lead Analyst, I want to inspect payment tender adoption and basket value distributions using boxplots so that I can evaluate payment fee friction and basket elasticity.", "F-05", "BR-10", "Should", "3"],
        ["US-15", "SH-04 (Lead Analyst)", "As a Lead Analyst, I want non-parametric distribution statistics (median, IQR, skewness) and outlier detection so that I can accurately model skewed ticket sizes.", "F-05", "BR-11", "Should", "3"],
        ["US-16", "SH-04 (Lead Analyst)", "As a Lead Analyst, I want a one-click CSV export of the currently filtered comparative summary table so that I can share analytical findings with external retail stakeholders.", "F-01", "BR-13", "Could", "2"]
    ]
    for r_idx, row_vals in enumerate(us_data, start=1):
        for c_idx, val in enumerate(row_vals):
            us_table.rows[r_idx].cells[c_idx].paragraphs[0].text = val
    style_table(us_table, [0.7, 1.2, 2.5, 0.7, 0.7, 0.7, 0.5], [WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER])

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # Section 3: Acceptance Criteria per User Story (Gherkin Syntax)
    add_custom_heading("3. Detailed Acceptance Criteria (Given / When / Then)", level=1)
    
    acceptance_specs = [
        ("US-01: SQLite Caching & Fast Ingestion", [
            "GIVEN the SQLite database 'data/app.db' is deployed,",
            "WHEN 'src/data_loader.py:load_table()' is invoked,",
            "THEN data is returned as a cached DataFrame in < 50ms with PRAGMA query_only=ON preventing disk mutation locks."
        ]),
        ("US-02: Referential Integrity & Schema Assertions", [
            "GIVEN transactional data is loaded into memory,",
            "WHEN 'src/validator.py:validate_schema()' runs automated checks,",
            "THEN zero orphan foreign keys, zero null primary keys, and 100% price consistency (selling >= cost) are asserted."
        ]),
        ("US-03: Executive Overview KPI Cards", [
            "GIVEN the application root landing page ('app.py') is open,",
            "WHEN default date and region filters are active,",
            "THEN 4 responsive Tailwind KPI cards display Gross Revenue, Completed Orders, ATV, and Margin with zero NaN displays."
        ]),
        ("US-04: Global Multi-Dimensional Filtering", [
            "GIVEN the Streamlit sidebar filter controls are rendered,",
            "WHEN the user modifies the Region filter from 'All' to 'North',",
            "THEN all dashboard charts and metrics re-render within 1.0 second displaying only North stores."
        ]),
        ("US-05: Ranked Store Performance Dual-Axis Chart", [
            "GIVEN the Store Comparison page is active,",
            "WHEN comparative store sales are queried,",
            "THEN a Plotly dual-axis chart ranks stores in descending revenue order alongside a gross profit margin line."
        ]),
        ("US-06: Space Productivity & ATV Scatter Matrix", [
            "GIVEN store area metadata ('store_size_sqft'),",
            "WHEN space productivity is selected,",
            "THEN a scatter plot visualizes floor space vs sales velocity with quadrant threshold lines isolating space-inefficient outlets."
        ]),
        ("US-07: Demographic Tier Category Revenue Mix", [
            "GIVEN product category sales transactions,",
            "WHEN category drilldown is opened,",
            "THEN a 100% stacked horizontal bar chart shows category revenue distribution summing exactly to 100% per demographic tier."
        ]),
        ("US-08: Discount Elasticity & OLS Regression", [
            "GIVEN item-level discount percentages and margin records,",
            "WHEN elasticity analysis is viewed,",
            "THEN a scatter plot with OLS linear regression displays Pearson 'r' and flags discounts exceeding 20% that cause margin destruction."
        ]),
        ("US-09: Month-over-Month Regional Velocity Trends", [
            "GIVEN 24 months of historical order time series,",
            "WHEN temporal trends page is selected,",
            "THEN a multi-line trend chart displays regional revenue curves with automated month-over-month percentage growth rates."
        ]),
        ("US-10: 24-Hour Trading Rush Heatmap", [
            "GIVEN transactional checkout timestamps,",
            "WHEN hourly rush distribution is viewed,",
            "THEN a Store-by-Hour intensity heatmap highlights rush trading hours (17:00 to 21:00) with custom rupee hover tooltips."
        ]),
        ("US-11: Weekday versus Weekend Purchasing Split", [
            "GIVEN calendar date classifications,",
            "WHEN day-type comparison is selected,",
            "THEN comparative metric badges and Welch's t-test p-value (< 0.05) prove that weekend ATV is significantly higher than weekday."
        ]),
        ("US-12: Net Revenue Leakage & Return Audit", [
            "GIVEN order status classifications,",
            "WHEN target variance and returns view is loaded,",
            "THEN net revenue reconciles exactly as Gross Sales - Returns - Cancellations alongside a Pareto category return defect chart."
        ]),
        ("US-13: Monthly Target Quota Variance Alerting", [
            "GIVEN monthly target quotas from 'dim_sales_targets',",
            "WHEN actual monthly sales are evaluated,",
            "THEN a bullet graph renders and visual amber/red warning badges trigger for stores deviating by more than +/- 5% from target."
        ]),
        ("US-14: Payment Tender Split & Basket Value Boxplots", [
            "GIVEN customer tender method records (UPI, Card, Cash),",
            "WHEN payment dynamics are viewed,",
            "THEN a donut chart displays tender share and Tukey boxplots show median and IQR ticket sizes across payment types."
        ]),
        ("US-15: Non-Parametric Distribution & Outlier Diagnostics", [
            "GIVEN retail transaction amounts,",
            "WHEN statistical EDA view is loaded,",
            "THEN summary tables display Mean, Median, StdDev, IQR, and Skewness with Tukey boxplot whiskers isolating bulk corporate outliers."
        ]),
        ("US-16: Filtered Tabular CSV Data Export", [
            "GIVEN an active filtered multi-store summary table,",
            "WHEN the user clicks the 'Download CSV' button,",
            "THEN a clean CSV extract downloads matching exactly on-screen row counts and formatted currency values."
        ])
    ]

    for title, ac_lines in acceptance_specs:
        add_custom_heading(title, level=2, space_before=8, space_after=2)
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.2)
        p.paragraph_format.space_after = Pt(2)
        for line in ac_lines:
            p.add_run(f"• {line}\n")

    # Section 4: Epics Grouping & 2-Week Sprint Schedule
    add_custom_heading("4. Epics Structure & 2-Week Sprint Execution Schedule", level=1)
    doc.add_paragraph(
        "User stories are organized into 5 logical engineering Epics and scheduled across a 10-day (2-week) agile sprint cycle:"
    )

    sprint_table = doc.add_table(rows=11, cols=5)
    sprint_headers = ["Sprint Day", "Scheduled User Stories", "Epic Focus", "Story Points", "Deliverable Milestone"]
    for i, h in enumerate(sprint_headers):
        sprint_table.rows[0].cells[i].paragraphs[0].text = h

    sprint_schedule = [
        ["Day 1", "US-01, US-02", "Epic 1: Data Architecture", "5 pts", "SQLite Data Loader, Schema Validation & Integrity Assertions"],
        ["Day 2", "US-03, US-04", "Epic 2: Dashboard Foundation", "6 pts", "Executive Landing Page, Tailwind KPI Cards, Global Sidebar Filters"],
        ["Day 3", "US-05", "Epic 3: Store Benchmarking", "3 pts", "Dual-Axis Ranked Store Revenue & Profit Margin Visualizer"],
        ["Day 4", "US-06", "Epic 3: Store Benchmarking", "3 pts", "Space Productivity Velocity vs Sq.Ft. Scatter Matrix"],
        ["Day 5", "US-07", "Epic 3: Merchandising Mix", "3 pts", "Demographic Store-Tier 100% Stacked Category Revenue Mix"],
        ["Day 6", "US-08", "Epic 3: Merchandising Mix", "5 pts", "Discount Elasticity Scatter Plot & OLS Linear Regression Engine"],
        ["Day 7", "US-09, US-11", "Epic 4: Temporal Insights", "5 pts", "MoM Regional Trajectory & Weekday vs. Weekend Welch's t-test"],
        ["Day 8", "US-10, US-12", "Epic 4: Temporal & Leakage", "6 pts", "24-Hour Trading Rush Heatmap & Net Revenue Leakage Pareto Audit"],
        ["Day 9", "US-13, US-14", "Epic 4: Target & Tender", "6 pts", "Target Quota Variance Dynamic Alerts & Payment Tender Boxplots"],
        ["Day 10", "US-15, US-16", "Epic 4 & 5: Stats & Export", "5 pts", "Statistical Distribution EDA (IQR/Skewness) & CSV Export Utility"]
    ]
    for r_idx, row_vals in enumerate(sprint_schedule, start=1):
        for c_idx, val in enumerate(row_vals):
            sprint_table.rows[r_idx].cells[c_idx].paragraphs[0].text = val
    style_table(sprint_table, [1.0, 1.4, 1.8, 0.8, 1.5], [WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT])

    doc.save(filename)
    print(f"User Stories Document successfully created at: {filename}")

if __name__ == "__main__":
    import sys
    target = sys.argv[1] if len(sys.argv) > 1 else "docs/11_User_Stories.docx"
    create_user_stories_doc(target)
