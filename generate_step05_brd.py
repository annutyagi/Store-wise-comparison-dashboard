import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

def create_brd_doc(filename):
    os.makedirs(os.path.dirname(filename), exist_ok=True)
    doc = docx.Document()
    
    # Page Margins
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)
        
    # Color Palette - Professional Corporate Navy & Slate
    COLOR_PRIMARY = RGBColor(26, 54, 93)     # #1A365D - Deep Navy
    COLOR_SECONDARY = RGBColor(43, 108, 176) # #2B6CB0 - Classic Slate Blue
    COLOR_TEXT = RGBColor(45, 55, 72)        # #2D3748 - Dark Charcoal
    COLOR_MUTED = RGBColor(113, 128, 150)    # #718096 - Muted Grey
    
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
            run.font.size = Pt(15)
            run.font.color.rgb = COLOR_PRIMARY
        elif level == 2:
            run.font.size = Pt(12.5)
            run.font.color.rgb = COLOR_SECONDARY
        elif level == 3:
            run.font.size = Pt(11)
            run.font.color.rgb = COLOR_PRIMARY
        return p

    def style_table(table, col_widths, col_alignments):
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        # Header formatting
        for idx, cell in enumerate(table.rows[0].cells):
            set_cell_background(cell, '1A365D')
            set_cell_margins(cell, top=110, bottom=110, left=140, right=140)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            for p in cell.paragraphs:
                p.alignment = col_alignments[idx]
                for run in p.runs:
                    run.font.bold = True
                    run.font.color.rgb = RGBColor(255, 255, 255)
                    run.font.size = Pt(9.5)
        
        # Body rows formatting
        for r_idx, row in enumerate(table.rows[1:], start=1):
            bg_color = 'F7FAFC' if r_idx % 2 == 1 else 'FFFFFF'
            for idx, cell in enumerate(row.cells):
                set_cell_background(cell, bg_color)
                set_cell_margins(cell, top=90, bottom=90, left=140, right=140)
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                for p in cell.paragraphs:
                    p.alignment = col_alignments[idx]
                    for run in p.runs:
                        run.font.size = Pt(9.0)
                        run.font.color.rgb = COLOR_TEXT
                        
        # Column widths
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
    s_run = sub_p.add_run("Step 5: Business Requirements Document (BRD) | Version 1.0 (Baseline)")
    s_run.font.size = Pt(12)
    s_run.font.color.rgb = COLOR_MUTED

    # Section 1: Document Control Table
    add_custom_heading("1. Document Control & Revision History", level=1)
    doc_ctrl_table = doc.add_table(rows=5, cols=4)
    doc_ctrl_headers = ["Version", "Release Date", "Author / Role", "Reviewer / Approver"]
    for i, h in enumerate(doc_ctrl_headers):
        doc_ctrl_table.rows[0].cells[i].paragraphs[0].text = h
    
    doc_ctrl_data = [
        ["v0.1", "10-Nov-2026", "Anu Tyagi (Lead Business Analyst)", "Mr. Dinesh (Initial Outline)"],
        ["v0.2", "14-Nov-2026", "Anu Tyagi (Lead Business Analyst)", "Mr. Dinesh (HMW & Needs Alignment)"],
        ["v0.3", "18-Nov-2026", "Anu Tyagi (Lead Business Analyst)", "Mr. Dinesh (Traceability Verification)"],
        ["v1.0", "20-Nov-2026", "Anu Tyagi (Lead Business Analyst)", "Mr. Dinesh (Formal Baseline Sign-Off)"]
    ]
    for r_idx, row_vals in enumerate(doc_ctrl_data, start=1):
        for c_idx, val in enumerate(row_vals):
            doc_ctrl_table.rows[r_idx].cells[c_idx].paragraphs[0].text = val
    style_table(doc_ctrl_table, [1.0, 1.4, 2.1, 2.0], [WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT])

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # Section 2: Executive Summary
    add_custom_heading("2. Executive Summary", level=1)
    doc.add_paragraph(
        "Modern multi-outlet retail enterprises operate in highly competitive, data-intensive environments where revenue growth, "
        "gross margins, and customer satisfaction depend heavily on operational agility. Despite widespread adoption of digital "
        "point-of-sale terminals, retail leadership and store managers continue to suffer from analytical latency—relying on static, "
        "disconnected end-of-month spreadsheets that obscure real-time store-level sales velocity, regional variances, and merchandising bottlenecks. "
        "This Business Requirements Document (BRD) defines the formal business capabilities, governance principles, operational rules, "
        "and acceptance benchmarks required for the Store Wise Sales Comparison Dashboard."
    )
    doc.add_paragraph(
        "By synthesizing data across 10+ retail outlets spanning Urban, Suburban, and Rural formats, the platform establishes a centralized "
        "decision-support system that accelerates commercial decision-making, protects gross merchandise margins, optimizes workforce scheduling, "
        "and exposes operational revenue leakage."
    )

    # Section 3: Business Problem & Opportunity
    add_custom_heading("3. Business Problem and Opportunity Statement", level=1)
    
    add_custom_heading("3.1 The Business Problem", level=2)
    doc.add_paragraph(
        "Enterprise retail executives lack a unified, cross-store comparison platform. Store operations managers work with isolated daily totals, "
        "unable to identify whether local sales slowdowns stem from macroeconomic trends, footfall drops during specific trading hours, or category-level "
        "merchandising flaws. Merchandisers lack empirical feedback on discount elasticity, frequently approving promotional markdowns that expand "
        "volume but destroy store profitability. Furthermore, delayed monthly target reconciliation causes commercial teams to discover revenue deficits "
        "only after fiscal periods close, completely eliminating the opportunity for intra-month corrective action."
    )

    add_custom_heading("3.2 The Strategic Business Opportunity", level=2)
    doc.add_paragraph(
        "Deploying an interactive, responsive comparative dashboard presents a substantial business opportunity to:\n"
        "1. Compress analytical reporting turnaround from 15 business days to under 2 seconds.\n"
        "2. Protect profit margins by statistically uncovering discount thresholds that cause margin destruction.\n"
        "3. Synchronize store-level staffing schedules with empirical hourly transaction distributions, eliminating peak-hour checkout bottlenecks.\n"
        "4. Maximize floor-space capital productivity by continuously monitoring sales velocity per square foot across store formats."
    )

    # Section 4: Objectives and Success Metrics
    add_custom_heading("4. Business Objectives and Success Metrics", level=1)
    doc.add_paragraph(
        "Project execution is anchored to five SMART business objectives and governed by explicit commercial Key Performance Indicators (KPIs):"
    )

    kpi_table = doc.add_table(rows=6, cols=4)
    kpi_headers = ["Metric / KPI Name", "Baseline / Legacy State", "Target Benchmark", "Strategic Business Objective"]
    for i, h in enumerate(kpi_headers):
        kpi_table.rows[0].cells[i].paragraphs[0].text = h

    kpi_data = [
        ["Report Generation Time", "10-15 business days (Manual Excel)", "< 2.0 seconds (Interactive)", "OBJ-01 (Multi-Store Benchmarking)"],
        ["Target Variance Alerting Lag", "30 days (Post-month reconciliation)", "Intra-month (Continuous tracking)", "OBJ-04 (Target Realization Engine)"],
        ["Peak-Hour Staffing Efficiency", "Fixed static shifts (Unmatched to traffic)", "Dynamic alignment to hourly traffic", "OBJ-03 (Temporal Analysis)"],
        ["Category Margin Protection", "Ad-hoc manual margin calculations", "Automated discount elasticity view", "OBJ-02 (Category Drilldown)"],
        ["Data Quality & Reconciliation", "Frequent spreadsheet formula discrepancies", "100% automated integrity validation", "OBJ-05 (Automated Health Scoring)"]
    ]
    for r_idx, row_vals in enumerate(kpi_data, start=1):
        for c_idx, val in enumerate(row_vals):
            kpi_table.rows[r_idx].cells[c_idx].paragraphs[0].text = val
    style_table(kpi_table, [1.6, 1.8, 1.6, 1.5], [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT])

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # Section 5: Stakeholders RACI Table
    add_custom_heading("5. Stakeholder Governance — RACI Matrix", level=1)
    doc.add_paragraph(
        "Governance roles are defined using the RACI framework (Responsible, Accountable, Consulted, Informed) across key project delivery phases:"
    )

    raci_table = doc.add_table(rows=7, cols=7)
    raci_headers = ["Project Lifecycle Phase", "SH-01 (Reg. Dir)", "SH-02 (Store Mgr)", "SH-03 (Merch)", "SH-04 (Lead BA)", "SH-05 (Mentor)", "SH-06 (Customer)"]
    for i, h in enumerate(raci_headers):
        raci_table.rows[0].cells[i].paragraphs[0].text = h

    raci_data = [
        ["Requirements Elicitation & BRD", "A", "C", "C", "R", "C", "I"],
        ["KPI & Formula Specification", "C", "C", "C", "R", "A", "I"],
        ["Data Architecture & Modeling", "I", "I", "I", "R", "A", "I"],
        ["UI Layout & Wireframe Review", "C", "C", "C", "R", "A", "I"],
        ["User Acceptance Testing (UAT)", "A", "R", "R", "R", "C", "I"],
        ["Academic Sign-Off & Viva Defense", "I", "I", "I", "R", "A", "I"]
    ]
    for r_idx, row_vals in enumerate(raci_data, start=1):
        for c_idx, val in enumerate(row_vals):
            raci_table.rows[r_idx].cells[c_idx].paragraphs[0].text = val
    style_table(raci_table, [2.3, 0.7, 0.7, 0.7, 0.7, 0.7, 0.7], [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER])

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # Section 6: Scope In / Out
    add_custom_heading("6. Project Scope Boundaries", level=1)
    doc.add_paragraph(
        "To ensure disciplined execution and adhere strictly to academic and zero-budget resource boundaries, "
        "the system scope is partitioned as follows:"
    )
    
    scope_table = doc.add_table(rows=6, cols=2)
    scope_headers = ["In-Scope Capabilities & Boundaries", "Out-of-Scope Exclusions & Deferrals"]
    for i, h in enumerate(scope_headers):
        scope_table.rows[0].cells[i].paragraphs[0].text = h
    
    scope_data = [
        ["Multi-store comparative sales, margin, and order volume analytics across 10+ outlets.", "Live Point-of-Sale (POS) terminal hardware telemetry or cashier cash-drawer synchronization."],
        ["Category and demographic tier drill-down (Urban, Suburban, Rural).", "Enterprise ERP write-back integration (direct automated updates into SAP S/4HANA or Oracle)."],
        ["Temporal trend analysis: Month-over-month (MoM), weekday vs. weekend, and hourly sales heatmaps.", "Sub-second live financial tick data streaming engines (e.g., Apache Kafka / Apache Flink)."],
        ["Target variance monitoring with dynamic +/- 5% threshold alerting.", "Deep Learning predictive forecasting / LSTM neural network architectures (reserved for Phase 2)."],
        ["Descriptive statistical distributions (mean, median, IQR, skewness) and cloud deployment.", "Multi-tenant enterprise Single Sign-On (SSO / SAML 2.0 / Okta) and multi-currency foreign exchange."]
    ]
    for r_idx, row_vals in enumerate(scope_data, start=1):
        for c_idx, val in enumerate(row_vals):
            scope_table.rows[r_idx].cells[c_idx].paragraphs[0].text = val
    style_table(scope_table, [3.25, 3.25], [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT])

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # Section 7: Detailed Business Requirements (BR-01 to BR-14)
    add_custom_heading("7. Detailed Business Requirements Specification", level=1)
    doc.add_paragraph(
        "The 14 baseline business requirements are formally cataloged below with description, business rationale, "
        "MoSCoW priority, upstream HMW linkage, and precise acceptance criteria:"
    )

    br_specs = [
        ("BR-01", "Multi-Store Revenue and Margin Benchmarking", "Must", "HMW-01",
         "The business needs to benchmark and rank store-level gross sales revenue and profit margins across geographic territories so that regional directors can identify top performers and margin-diluting stores.",
         "Without cross-store ranking, leadership cannot evaluate regional retail health or allocate promotional budgets effectively.",
         "GIVEN a selected date range, WHEN the user accesses the comparative overview, THEN all stores must be ranked in descending order by gross sales revenue and profit margin percentage with zero mathematical discrepancy."),
        
        ("BR-02", "Space Productivity and Average Transaction Value (ATV)", "Must", "HMW-02",
         "The business needs to calculate and compare sales velocity per square foot and average transaction value across all retail store sizes so that executives can evaluate retail real-estate productivity and store expansion ROI.",
         "Physical store footprints vary from 1,200 to 8,000 sq.ft.; absolute revenue obscures true real-estate efficiency.",
         "GIVEN store square footage data, WHEN the space productivity metric is generated, THEN sales velocity must equal total net sales divided by store square footage, and ATV must equal total net sales divided by completed orders."),
        
        ("BR-03", "Demographic Store-Tier Category Revenue Alignment", "Must", "HMW-03",
         "The business needs to analyze product category sales contribution segmented by store demographic tiers (Urban, Suburban, Rural) so that merchandisers can eliminate stock-outs and align inventory stocking to regional demographics.",
         "Urban stores exhibit vastly different category affinities compared to rural outlets; inventory misallocation causes localized markdowns.",
         "GIVEN store classifications (Urban, Suburban, Rural), WHEN category contribution is computed, THEN category sales percentages must sum to 100% within each tier with zero unallocated categories."),
        
        ("BR-04", "Month-over-Month (MoM) Regional Velocity Tracking", "Must", "HMW-04",
         "The business needs to track month-over-month (MoM) sales velocity trends across regional store clusters so that commercial leadership can detect seasonal demand shifts and take proactive corrective actions prior to month-end.",
         "Monthly revenue trends provide early warning of demand contraction before quarterly financial statements are finalized.",
         "GIVEN transactional time series, WHEN consecutive calendar months are evaluated, THEN MoM percentage variance must be computed accurately across each regional cluster."),
        
        ("BR-05", "Hourly Transaction Density and Peak Trading Windows", "Must", "HMW-05",
         "The business needs to identify hourly customer transaction volumes and peak sales trading windows for each store so that store managers can optimize staff shift scheduling and reduce customer checkout wait times.",
         "Frontline staff scheduling without empirical rush-hour data results in checkout queue bottlenecks and lost sales.",
         "GIVEN timestamped sales logs, WHEN temporal aggregation is performed, THEN transaction volume and sales must be partitioned into 24 distinct hourly buckets (00:00 to 23:00) per store."),
        
        ("BR-06", "Weekday versus Weekend Demand Partitioning", "Should", "HMW-06",
         "The business needs to separate and contrast weekday versus weekend sales volumes and basket sizes so that store operations teams can optimize weekend inventory staging and logistics schedules.",
         "Weekend footfall surges require distinct supply chain replenishment buffers compared to weekday operational baselines.",
         "GIVEN calendar date attributes, WHEN weekday and weekend filters are evaluated, THEN orders must correctly partition into Monday-Friday vs Saturday-Sunday with distinct ATV calculations."),
        
        ("BR-07", "Net Revenue Leakage and Return Rate Monitoring", "Must", "HMW-07",
         "The business needs to monitor product return and order cancellation rates across stores and product categories so that management can identify localized quality issues, reduce return rates, and prevent net revenue leakage.",
         "High gross sales often mask severe net revenue erosion caused by product returns and canceled orders.",
         "GIVEN order status classifications, WHEN net sales are reconciled, THEN Net Sales must exactly equal Gross Sales minus Returns minus Cancellations."),
        
        ("BR-08", "Discount Elasticity and Margin Preservation", "Must", "HMW-08",
         "The business needs to evaluate the relationship between discount percentages, unit sales volume, and gross profit margins across product lines so that merchandisers can design margin-protective pricing and promotional campaigns.",
         "Deep discounting may artificially inflate unit volumes while destroying overall store gross profitability.",
         "GIVEN product discount records, WHEN discount tiers are analyzed, THEN correlation between discount rate, volume, and profit margin must be calculated with statistical rigor."),
        
        ("BR-09", "Target Revenue Quota Variance Alerting", "Must", "HMW-09",
         "The business needs to measure actual store sales against budgeted monthly revenue quotas and flag variance thresholds exceeding +/- 5% so that regional management can execute timely sales interventions.",
         "Static monthly budget spreadsheets delay intervention until after underperformance has caused fiscal damage.",
         "GIVEN monthly target quotas, WHEN actual monthly revenue is compared against budget, THEN any variance exceeding +/- 5% must trigger an automated visual alert indicator."),
        
        ("BR-10", "Payment Tender Split and Basket Size Correlation", "Should", "HMW-10",
         "The business needs to aggregate customer payment method adoption (UPI, Credit Card, Cash) and analyze their relationship with basket size across store locations so that finance and store managers can minimize payment friction and transaction charges.",
         "Payment tender fees and customer checkout speeds vary significantly between UPI, cards, and cash.",
         "GIVEN tender transaction records, WHEN payment methods are evaluated, THEN transaction count and basket sizes must be segmented by payment type with zero orphaned tenders."),
        
        ("BR-11", "Automated Statistical Distribution Analytics", "Should", "HMW-11",
         "The business needs to automate statistical distribution analysis (mean, median, interquartile range, skewness, outlier detection) for transaction values so that business analysts can instantly evaluate customer spending patterns without manual calculation errors.",
         "Retail transaction distributions are heavily skewed; relying solely on arithmetic mean distorts financial projections.",
         "GIVEN transaction values, WHEN statistical analysis is triggered, THEN mean, median, IQR, standard deviation, and skewness must be calculated in adherence to standard statistical formulas."),
        
        ("BR-12", "Data Integrity, Referential Integrity, and Audit Reconciliation", "Must", "HMW-12",
         "The business needs to enforce data validation rules, relational referential integrity, and reconciliation controls across all transactional records so that stakeholders and evaluators have guaranteed auditability and trust in the analytical results.",
         "Inconsistent data or orphaned records invalidate executive reporting and academic compliance.",
         "GIVEN relational datasets, WHEN automated integrity assertions execute, THEN zero orphan keys, zero null primary keys, and exact aggregate balance reconciliation must be proven."),
        
        ("BR-13", "Aggregated Tabular Data Export", "Could", "HMW-01, HMW-11",
         "The business needs to export aggregated comparative store performance summaries into tabular spreadsheet formats so that business users can share findings with external executive partners.",
         "External stakeholders often require portable tabular extracts for executive boardroom briefings.",
         "GIVEN filtered on-screen analytics, WHEN data export is invoked, THEN the downloaded CSV/Excel file must exactly match on-screen filtered records and totals."),
        
        ("BR-14", "Multi-Currency Cross-Border Conversion", "Won't", "HMW-01",
         "The business needs to provide multi-currency international sales conversion and real-time live currency exchange rate updates for cross-border operations.",
         "Deferred: Current system boundaries are strictly domestic retail denominated in INR (₹) as per Assumption AS-02.",
         "DEFERRED FOR PHASE 2: Out of scope for current baseline release.")
    ]

    for br_id, title, prio, hmw, desc, rat, ac in br_specs:
        add_custom_heading(f"{br_id}: {title} (Priority: {prio})", level=2, space_before=10, space_after=2)
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.2)
        p.paragraph_format.space_after = Pt(2)
        
        r_hmw = p.add_run(f"Source HMW: {hmw} | Priority: {prio}\n")
        r_hmw.bold = True
        r_hmw.font.color.rgb = COLOR_SECONDARY
        
        r_d_b = p.add_run("Requirement Description: ")
        r_d_b.bold = True
        p.add_run(f"{desc}\n")
        
        r_r_b = p.add_run("Business Rationale: ")
        r_r_b.bold = True
        p.add_run(f"{rat}\n")
        
        r_ac_b = p.add_run("Acceptance Criteria: ")
        r_ac_b.bold = True
        p.add_run(f"{ac}")

    # Section 8: Business Rules
    add_custom_heading("8. Business Rules Catalog", level=1)
    doc.add_paragraph(
        "The following operational business rules govern all calculations, validations, and logic across the analytical solution:"
    )

    rules_table = doc.add_table(rows=7, cols=3)
    rules_headers = ["Rule ID", "Business Rule Definition & Logical Formula", "Governing Principle & Enforcement"]
    for i, h in enumerate(rules_headers):
        rules_table.rows[0].cells[i].paragraphs[0].text = h

    rules_data = [
        ["BURL-01", "Net Sales Revenue = Gross Sales Revenue - (Returned Value + Cancelled Value)", "Enforces true revenue recognition; gross revenue cannot be reported in isolation."],
        ["BURL-02", "Gross Margin % = ((Net Selling Price - Unit Cost Price) / Net Selling Price) * 100", "Guarantees accurate store-level gross profitability benchmarking."],
        ["BURL-03", "Average Transaction Value (ATV) = Total Net Sales / Total Count of Completed Transactions", "Excludes canceled orders and returns from basket size calculation."],
        ["BURL-04", "Sales Velocity per Sq.Ft. = Total Annualized Net Sales / Store Physical Square Footage", "Normalizes sales productivity across small (kiosk) and large (hypermarket) store footprints."],
        ["BURL-05", "Target Quota Variance % = ((Actual Revenue - Target Revenue) / Target Revenue) * 100; Alert if ABS(Variance) >= 5%", "Automated early warning threshold for underperforming retail units."],
        ["BURL-06", "Valid Promotional Discount Range: 0.0% <= Discount % <= 50.0%", "Protects brand equity and flags extreme pricing anomalies for manual audit review."]
    ]
    for r_idx, row_vals in enumerate(rules_data, start=1):
        for c_idx, val in enumerate(row_vals):
            rules_table.rows[r_idx].cells[c_idx].paragraphs[0].text = val
    style_table(rules_table, [1.0, 3.2, 2.3], [WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT])

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # Section 9: Data Requirements
    add_custom_heading("9. Data Requirements & Governance", level=1)
    doc.add_paragraph(
        "To ensure robust data pipelines, data source characteristics, refresh frequencies, and retention policies are defined below:"
    )

    data_req_table = doc.add_table(rows=5, cols=4)
    data_req_headers = ["Entity Domain", "Source Description", "Refresh / Update Frequency", "Data Quality & Integrity Rules"]
    for i, h in enumerate(data_req_headers):
        data_req_table.rows[0].cells[i].paragraphs[0].text = h

    data_req_data = [
        ["Store Master", "Simulated enterprise retail master table (10+ stores across 3 regions)", "Static / Quarterly update", "Unique store_id, valid region_id, positive non-zero square footage."],
        ["Product Catalog", "Simulated retail catalog (100+ SKUs across 5 primary categories)", "Monthly update", "Unique product_id, non-negative cost and selling prices, valid category code."],
        ["Sales Transactions", "Synthetic POS transactional engine (10,000+ orders spanning 24 months)", "Batch append / Instant local query", "Unique order_id, valid foreign keys to Store and Product, ISO timestamps."],
        ["Sales Targets", "Monthly store revenue budget table generated per fiscal period", "Annual / Quarterly budget cycle", "Exact match to active store_ids, target revenue > 0 in INR (₹)."]
    ]
    for r_idx, row_vals in enumerate(data_req_data, start=1):
        for c_idx, val in enumerate(row_vals):
            data_req_table.rows[r_idx].cells[c_idx].paragraphs[0].text = val
    style_table(data_req_table, [1.4, 1.8, 1.5, 1.8], [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT])

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # Section 10: Assumptions, Constraints, Dependencies
    add_custom_heading("10. Assumptions, Constraints, and Technical Dependencies", level=1)
    
    add_custom_heading("10.1 Assumptions", level=2)
    doc.add_paragraph(
        "• AS-01: Synthetic dataset accurately reflects realistic retail purchasing distributions, seasonality, and basket dynamics.\n"
        "• AS-02: All transactions, revenue, and margins are denominated uniformly in Indian Rupees (INR / ₹).\n"
        "• AS-03: End users operate standard modern web browsers (Chrome, Edge, Firefox, Safari) at >= 1280x720 desktop resolution.\n"
        "• AS-04: Relational SQLite database delivers sub-second analytical query performance when indexed properly.\n"
        "• AS-05: System usage is predominantly analytical read-only; no manual in-app transactional data mutations."
    )

    add_custom_heading("10.2 Constraints", level=2)
    doc.add_paragraph(
        "• CO-01: Zero cloud budget constraint: All tools, libraries, and hosting must operate on free tiers.\n"
        "• CO-02: Streamlit Community Cloud enforces a strict 1 GB RAM allocation limit.\n"
        "• CO-03: Project delivery governed by the strict academic timetable of ABES MCA 2nd Year semester.\n"
        "• CO-04: Strict compliance with privacy policies: Zero real-world PII or proprietary customer data.\n"
        "• CO-05: Single developer architecture engineered solely by student data analyst Anu Tyagi."
    )

    add_custom_heading("10.3 Technical Dependencies", level=2)
    doc.add_paragraph(
        "• DEP-01: Python 3.10+ execution runtime environment.\n"
        "• DEP-02: SQLite 3.35+ relational database engine supporting window analytical functions.\n"
        "• DEP-03: Web browser client compatibility with modern HTML5 and WebGL for interactive Plotly charts.\n"
        "• DEP-04: Streamlit Community Cloud hosting availability and GitHub repository synchronization."
    )

    # Section 11: Risks and Mitigation Matrix
    add_custom_heading("11. Risk Management & Mitigation Matrix", level=1)
    doc.add_paragraph(
        "Proactive identification of technical and project risks ensures robust execution and business continuity:"
    )

    risk_table = doc.add_table(rows=6, cols=5)
    risk_headers = ["Risk ID", "Risk Description", "Likelihood", "Impact", "Proactive Mitigation Strategy"]
    for i, h in enumerate(risk_headers):
        risk_table.rows[0].cells[i].paragraphs[0].text = h

    risk_data = [
        ["RSK-01", "Memory exhaustion on Streamlit Community Cloud due to large in-memory dataframes.", "Medium", "High", "Implement indexed SQL-level aggregation so only aggregated records are loaded into Python memory; utilize @st.cache_data."],
        ["RSK-02", "Unrealistic synthetic data distorting statistical distribution insights.", "Medium", "Medium", "Calibrate synthetic data generator with real-world empirical retail parameters (Pareto 80/20 category distributions, log-normal basket values)."],
        ["RSK-03", "Scope creep during UI dashboard design exceeding academic timeline.", "High", "High", "Enforce strict MoSCoW prioritization and require RTM traceability before introducing any new visual module."],
        ["RSK-04", "Browser rendering lag when plotting high-density hourly heatmaps.", "Low", "Medium", "Pre-aggregate hourly transaction density at database query level before passing JSON payloads to Plotly."],
        ["RSK-05", "Non-compliance with ABES academic project reporting formatting standards.", "Low", "High", "Establish early structural alignment with ABES official report template and submit iterative drafts to Mentor Mr. Dinesh."]
    ]
    for r_idx, row_vals in enumerate(risk_data, start=1):
        for c_idx, val in enumerate(row_vals):
            risk_table.rows[r_idx].cells[c_idx].paragraphs[0].text = val
    style_table(risk_table, [0.8, 2.0, 0.8, 0.8, 2.1], [WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT])

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # Section 12: Glossary
    add_custom_heading("12. Glossary of Terms & Acronyms", level=1)
    glossary_items = [
        ("ATV (Average Transaction Value): ", "The average monetary expenditure per customer checkout, calculated as Total Net Revenue divided by Total Order Count."),
        ("B-Tree Index: ", "A self-balancing relational database index structure enabling logarithmic time complexity for analytical search and range queries."),
        ("Gross Margin: ", "The percentage of sales revenue retained after subtracting the direct cost of goods sold (COGS)."),
        ("IQR (Interquartile Range): ", "A robust measure of statistical dispersion equal to the difference between the 75th and 25th percentiles (Q3 - Q1)."),
        ("MoM (Month-over-Month): ", "A comparative metric tracking percentage growth or contraction in sales velocity relative to the immediately preceding calendar month."),
        ("MoSCoW Prioritization: ", "A requirements management framework categorizing items into Must Have, Should Have, Could Have, and Won't Have."),
        ("RACI: ", "A governance matrix identifying who is Responsible, Accountable, Consulted, and Informed for each project activity."),
        ("RTM (Requirements Traceability Matrix): ", "A cross-referencing artifact ensuring continuous bidirectional mapping between business requirements and technical deliverables."),
        ("Sales Velocity per Sq.Ft.: ", "A spatial productivity metric measuring net revenue generated per square foot of physical store retail selling space."),
        ("Skewness: ", "A statistical measure of asymmetry in a distribution curve, indicating whether spending clusters around lower values with a long high-value tail.")
    ]
    for term, definition in glossary_items:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.2)
        p.paragraph_format.space_after = Pt(2)
        r_t = p.add_run(term)
        r_t.bold = True
        r_t.font.color.rgb = COLOR_PRIMARY
        r_d = p.add_run(definition)
        r_d.font.color.rgb = COLOR_TEXT

    # Section 13: Sign-Off Block
    add_custom_heading("13. Formal Project Document Sign-Off", level=1)
    doc.add_paragraph(
        "By signing below, the author, academic mentor, and stakeholder representatives confirm that this Business Requirements "
        "Document (BRD v1.0) accurately encapsulates project goals, operational rules, and acceptance criteria:"
    )

    sign_table = doc.add_table(rows=4, cols=4)
    sign_headers = ["Signatory Role", "Name & Title", "Signature", "Approval Date"]
    for i, h in enumerate(sign_headers):
        sign_table.rows[0].cells[i].paragraphs[0].text = h

    sign_data = [
        ["Lead Business Analyst (Author)", "Anu Tyagi (MCA 2nd Year, ABES)", "_______________________", "20-Nov-2026"],
        ["Academic Project Mentor (Approver)", "Mr. Dinesh", "_______________________", "20-Nov-2026"],
        ["Executive Sponsor Representative", "Regional Sales Director (SH-01 Proxy)", "_______________________", "20-Nov-2026"]
    ]
    for r_idx, row_vals in enumerate(sign_data, start=1):
        for c_idx, val in enumerate(row_vals):
            sign_table.rows[r_idx].cells[c_idx].paragraphs[0].text = val
    style_table(sign_table, [2.0, 2.2, 1.3, 1.0], [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER])

    doc.save(filename)
    print(f"BRD Document successfully created at: {filename}")

if __name__ == "__main__":
    import sys
    target = sys.argv[1] if len(sys.argv) > 1 else "docs/05_BRD.docx"
    create_brd_doc(target)
