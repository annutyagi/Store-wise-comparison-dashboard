import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

def create_project_understanding_doc(filename):
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

    def add_custom_heading(text, level, space_before=12, space_after=6):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.bold = True
        if level == 1:
            run.font.size = Pt(16)
            run.font.color.rgb = COLOR_PRIMARY
            # Add bottom accent bar or divider line simulation
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
                        
        # Apply explicit column widths
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
    s_run = sub_p.add_run("Step 1: Project Understanding Document | Academic & Industrial Specification")
    s_run.font.size = Pt(12)
    s_run.font.color.rgb = COLOR_MUTED

    # Metadata Box Table
    meta_table = doc.add_table(rows=4, cols=2)
    meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta_data = [
        [("Learner Name:", "Anu Tyagi"), ("Roll Number:", "2500320140025")],
        [("Batch ID:", "ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01"), ("Institution:", "ABES Engineering College")],
        [("Project Mentor:", "Mr. Dinesh"), ("Domain:", "Data Analytics")],
        [("Target Stack:", "Python, Streamlit, Plotly, SQLite"), ("Deployment:", "Streamlit Community Cloud")]
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

    # Section 1 & 2: Executive Overview
    add_custom_heading("1. Executive Project Overview", level=1)
    p = doc.add_paragraph(
        "The retail enterprise environment demands continuous, granular evaluation of multi-location sales velocity, "
        "operational margins, and consumer basket dynamics. Modern retail enterprises operating multiple store outlets "
        "across geographic regions face significant operational friction due to disparate reporting, static end-of-month "
        "spreadsheets, and delayed diagnostic visibility into store-level anomalies. This project establishes an interactive, "
        "data-driven Store Wise Sales Comparison Dashboard utilizing Python, Streamlit, Plotly, and SQLite to empower retail "
        "decision-makers with near-real-time comparative sales performance, category drill-downs, and statistical health metrics."
    )

    # Section 3: Problem Statement
    add_custom_heading("2. Problem Statement", level=1)
    p_box = doc.add_table(rows=1, cols=1)
    p_box.alignment = WD_TABLE_ALIGNMENT.CENTER
    c = p_box.rows[0].cells[0]
    set_cell_background(c, 'EBF8FF')
    set_cell_margins(c, top=140, bottom=140, left=180, right=180)
    c.width = Inches(6.5)
    cp = c.paragraphs[0]
    cp.paragraph_format.space_after = Pt(0)
    cr_bold = cp.add_run("Formal Problem Statement: ")
    cr_bold.bold = True
    cr_bold.font.color.rgb = COLOR_PRIMARY
    cr = cp.add_run(
        "Retail store managers and regional executives lack a unified, interactive analytical platform to evaluate store-level "
        "sales velocity, benchmark multi-store performance across product categories and regions, and detect operational revenue "
        "leakage across time periods."
    )
    cr.font.color.rgb = COLOR_TEXT
    cr.italic = True

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # Section 4: Intended Audience / Users
    add_custom_heading("3. Intended Audience & User Personas", level=1)
    p = doc.add_paragraph(
        "The application is engineered for distinct operational tiers within a retail enterprise hierarchy. "
        "Each user archetype interacts with the platform at varying cadences to drive data-backed operational interventions:"
    )

    users_table = doc.add_table(rows=5, cols=4)
    users_headers = ["User Role / Persona", "Access Cadence", "Primary Workflow / Modules", "Key Information Need"]
    for i, h in enumerate(users_headers):
        users_table.rows[0].cells[i].paragraphs[0].text = h
    
    users_data = [
        ["Regional Sales Director", "Weekly & Monthly", "Executive Overview, Multi-Store Benchmarking, Regional Comparison", "High-level gross revenue, regional quota achievement, outlier store identification"],
        ["Store Operations Manager", "Daily (Opening & Closing)", "Store Performance Detail, Hourly Sales Heatmap, Payment Mix", "Daily store revenue target tracking, peak transaction hours, basket size variance"],
        ["Inventory & Category Merchandiser", "Weekly & Bi-Weekly", "Category & SKU Breakdown, Product Affinity, Return Analysis", "Category contribution margins, slow-moving lines, product return rate per outlet"],
        ["Business / Data Analyst", "Ad-Hoc & Weekly", "Statistical Exploratory Data Analysis (EDA), Correlation, Data Export", "Inter-store variance, IQR/skewness of basket value, discount vs. sales elasticity"]
    ]
    for r_idx, row_vals in enumerate(users_data, start=1):
        for c_idx, val in enumerate(row_vals):
            users_table.rows[r_idx].cells[c_idx].paragraphs[0].text = val
    style_table(users_table, [1.5, 1.2, 2.0, 1.8], [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT])

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # Section 5: Stakeholders Table
    add_custom_heading("4. Stakeholder Identification & Matrix", level=1)
    p = doc.add_paragraph(
        "Comprehensive stakeholder analysis ensures alignment across strategic, operational, analytical, and academic dimensions. "
        "All stakeholder IDs (SH-01 through SH-06) remain locked throughout the project lifecycle."
    )

    sh_table = doc.add_table(rows=7, cols=6)
    sh_headers = ["SH ID", "Stakeholder", "Role Description", "Core Interest", "Power / Influence", "Key Decision Enabled"]
    for i, h in enumerate(sh_headers):
        sh_table.rows[0].cells[i].paragraphs[0].text = h

    sh_data = [
        ["SH-01", "Regional Sales Director", "Executive Sponsor / Regional Business Head", "Regional revenue maximization, store profitability benchmarks", "High (H)", "Store budget allocation, regional marketing spends, underperforming store interventions"],
        ["SH-02", "Store Operations Manager", "Operational End User / Retail Store Lead", "Daily target attainment, staff scheduling, transaction velocity", "High (H)", "Shift staffing adjustments based on peak hours, local stock replenishment triggers"],
        ["SH-03", "Inventory & Category Manager", "Merchandising Lead / Supply Chain Partner", "Product category revenue share, stock velocity, markdown efficacy", "Medium (M)", "Store-specific category shelf-space allocation, promotional markdown approvals"],
        ["SH-04", "Lead Business Data Analyst", "Technical Business Analyst / QA Lead", "Metric definition accuracy, analytical query latency, EDA depth", "Medium (M)", "KPI standard definition, statistical anomaly detection methodology validation"],
        ["SH-05", "Academic Project Mentor (Mr. Dinesh)", "Project Evaluator / Academic Advisor", "Engineering rigor, methodological correctness, curriculum compliance", "High (H)", "Academic milestone acceptance, project viva grading, code quality sign-off"],
        ["SH-06", "Retail End Customer", "Consumer Purchasing Entity (Indirect)", "Checkout efficiency, product availability, transparent pricing", "Low (L)", "Indirect impact: faster checkout during peak hours, optimized local store inventory"]
    ]
    for r_idx, row_vals in enumerate(sh_data, start=1):
        for c_idx, val in enumerate(row_vals):
            sh_table.rows[r_idx].cells[c_idx].paragraphs[0].text = val
    style_table(sh_table, [0.7, 1.2, 1.3, 1.2, 0.8, 1.3], [WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT])

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # Section 6: Operational Questions (OQ-01 to OQ-10) with Data Feasibility Mapping
    add_custom_heading("5. Operational Questions & Data Feasibility Mapping", level=1)
    p = doc.add_paragraph(
        "To ensure business utility, the platform directly answers 10 rigorously formulated Operational Questions (OQ-01 through OQ-10). "
        "A rigorous feasibility check confirms that every operational question is fully answered by deterministic fields in the underlying retail schema."
    )

    oq_table = doc.add_table(rows=11, cols=4)
    oq_headers = ["OQ ID", "Operational Question", "Required Underlying Data Fields", "Feasibility Verification & Visual Output"]
    for i, h in enumerate(oq_headers):
        oq_table.rows[0].cells[i].paragraphs[0].text = h

    oq_data = [
        ["OQ-01", "Which retail stores generate the highest and lowest gross sales revenue and profit margin in a selected timeframe?", "store_id, sales_amount, cost_amount, profit, transaction_date", "Feasible: Aggregated sum of sales & margin; rendered via Ranked Bar Chart & KPI cards."],
        ["OQ-02", "How do individual stores compare in average transaction value (ATV) and sales velocity per square foot?", "store_id, transaction_id, sales_amount, store_size_sqft", "Feasible: ATV = sum(sales)/count(transactions); Velocity = sum(sales)/sqft; Scatter Plot."],
        ["OQ-03", "Which product categories drive the highest revenue contribution across different store tiers (Urban, Suburban, Rural)?", "store_tier, category_name, sales_amount, store_id", "Feasible: Two-way grouping on store_tier and category_name; Stacked 100% Bar Chart."],
        ["OQ-04", "How do monthly sales trends vary month-over-month (MoM) across regional store clusters?", "region_name, transaction_date (YYYY-MM), sales_amount", "Feasible: Date truncation to month, percentage delta calculation; Multi-line Trend Chart."],
        ["OQ-05", "What is the distribution of sales performance across weekdays versus weekends for each store?", "store_id, transaction_date, day_of_week_flag, sales_amount", "Feasible: Date parsing into weekday/weekend binary classification; Grouped Bar Chart."],
        ["OQ-06", "Which stores exhibit the highest return and cancellation rates, and what is the net revenue impact?", "order_status, return_amount, net_sales, store_id", "Feasible: Filter on status in ('Returned', 'Cancelled'); Return % metric card & Pareto Chart."],
        ["OQ-07", "How does discount percentage correlate with total unit volume and gross margin across product lines per store?", "discount_pct, quantity, gross_margin, store_id, product_id", "Feasible: Pearson/Spearman correlation and scatter plot with trendline."],
        ["OQ-08", "Which top underperforming stores consistently fail to achieve their assigned monthly revenue targets?", "store_id, actual_sales, target_sales, variance_amount, month_year", "Feasible: Variance calculation (Actual - Target) and percentage attainment; Bullet Graph."],
        ["OQ-09", "How do customer payment methods (UPI, Credit Card, Cash) distribute across stores, and does payment type correlate with basket size?", "payment_method, basket_amount, store_id, transaction_id", "Feasible: Distribution breakdown by method; Box plots of basket amount by payment type."],
        ["OQ-10", "What are the peak trading hours and temporal traffic distribution patterns across stores?", "store_id, transaction_time (Hour: 0-23), transaction_count, sales_amount", "Feasible: Hour extraction from timestamp; Hourly Heatmap Matrix (Store vs Hour)."]
    ]
    for r_idx, row_vals in enumerate(oq_data, start=1):
        for c_idx, val in enumerate(row_vals):
            oq_table.rows[r_idx].cells[c_idx].paragraphs[0].text = val
    style_table(oq_table, [0.7, 2.3, 1.8, 1.7], [WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT])

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # Section 7: SMART Objectives
    add_custom_heading("6. Project SMART Objectives", level=1)
    p = doc.add_paragraph(
        "To maintain engineering discipline, project milestones are governed by five SMART (Specific, Measurable, "
        "Achievable, Relevant, Time-bound) objectives:"
    )

    obj_table = doc.add_table(rows=6, cols=3)
    obj_headers = ["Objective ID", "SMART Objective Definition", "Verification Metric & Alignment"]
    for i, h in enumerate(obj_headers):
        obj_table.rows[0].cells[i].paragraphs[0].text = h

    obj_data = [
        ["OBJ-01", "Multi-Store Comparative Analytics: Engineer an interactive benchmarking module allowing simultaneous comparison of 5 to 15 stores across 4 core sales KPIs with query latency < 1.5 seconds.", "Latency verified via Streamlit execution profiler; covers OQ-01, OQ-02."],
        ["OBJ-02", "Category & Merchandising Drilldown: Develop hierarchical category-to-product sales visualizations enabling regional merchandisers to isolate top/bottom 15% revenue lines within 3 clicks.", "Workflow verified via UI usability test; covers OQ-03, OQ-07."],
        ["OBJ-03", "Temporal & Seasonality Modeling: Implement automated time-series aggregation calculating MoM, weekday-vs-weekend, and hourly sales patterns across a 24-month synthetic retail dataset.", "Validated against 100% of calendar dates in dataset; covers OQ-04, OQ-05, OQ-10."],
        ["OBJ-04", "Variance & Target Realization Engine: Construct dynamic target variance tracking that flags stores deviating by more than +/- 5% from budgeted sales targets.", "Validated with automated SQLite variance calculation assertions; covers OQ-08."],
        ["OBJ-05", "Statistical EDA & Automated Health Scoring: Deliver descriptive statistical summaries (mean, median, standard deviation, interquartile range, skewness) for transaction distributions.", "Validated via pytest numerical accuracy assertions against NumPy/Pandas; covers OQ-09."]
    ]
    for r_idx, row_vals in enumerate(obj_data, start=1):
        for c_idx, val in enumerate(row_vals):
            obj_table.rows[r_idx].cells[c_idx].paragraphs[0].text = val
    style_table(obj_table, [0.8, 3.7, 2.0], [WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT])

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # Section 8: Scope
    add_custom_heading("7. Project Scope Definition", level=1)
    p = doc.add_paragraph(
        "Clear scoping prevents feature creep and establishes realistic technical boundaries tailored for an academic capstone "
        "and industrial prototype:"
    )

    scope_table = doc.add_table(rows=6, cols=2)
    scope_headers = ["In-Scope Functional & Technical Boundaries", "Out-of-Scope Elements (Explicit Exclusions)"]
    for i, h in enumerate(scope_headers):
        scope_table.rows[0].cells[i].paragraphs[0].text = h

    scope_data = [
        ["Interactive multi-store sales comparison dashboard built in Python (Streamlit framework).", "Live Point-of-Sale (POS) hardware IoT telemetry or cash drawer synchronization."],
        ["Synthetic relational SQLite retail database modeled to 3NF standards (Stores, Products, Orders, Order_Items, Targets).", "Enterprise ERP automated write-back (e.g., direct write access to live SAP S/4HANA or Oracle Retail)."],
        ["Dynamic Plotly visualization suite (Ranked bars, multi-line trends, hourly heatmaps, scatter distributions, KPI cards).", "High-frequency real-time stock-trading style sub-second streaming engines (e.g., Apache Kafka / Flink)."],
        ["Statistical Exploratory Data Analysis (EDA) module covering parametric and non-parametric retail distributions.", "Advanced Machine Learning predictive sales forecasting / Deep Learning LSTM pipelines (reserved for v2.0)."],
        ["Automated pytest suite (>85% code coverage) and public cloud deployment on Streamlit Community Cloud.", "Multi-tenant enterprise Role-Based Access Control (RBAC) with corporate Single Sign-On (SSO / SAML 2.0 / Okta)."]
    ]
    for r_idx, row_vals in enumerate(scope_data, start=1):
        for c_idx, val in enumerate(row_vals):
            scope_table.rows[r_idx].cells[c_idx].paragraphs[0].text = val
    style_table(scope_table, [3.25, 3.25], [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT])

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # Section 9: Assumptions & Constraints
    add_custom_heading("8. Assumptions and Constraints", level=1)
    
    add_custom_heading("8.1 Project Assumptions", level=2)
    assumptions_p = [
        ("AS-01: Synthetic Data Realism: ", "The generated dataset accurately replicates real-world retail purchasing behavior, including seasonal spikes (Q4 holiday surge), weekend footfall increases, and authentic discount distributions."),
        ("AS-02: Operational Single Currency: ", "All financial metrics, revenue sums, discounts, and target quotas are denominated uniformly in Indian Rupees (INR / ₹) with zero multi-currency conversion overhead."),
        ("AS-03: Modern Client Environment: ", "End users access the web dashboard via standard modern HTML5-compliant web browsers (Chrome, Edge, Firefox, Safari) with JavaScript enabled at a minimum 1280x720 desktop resolution."),
        ("AS-04: Relational SQLite Sufficiency: ", "SQLite is fully capable of delivering sub-second analytical query performance on normalized retail datasets containing 10,000 to 50,000 transaction rows when proper B-tree indexing is applied."),
        ("AS-05: Non-destructive Analytical Usage: ", "The dashboard operates primarily as a read-heavy analytical reporting system, ensuring zero accidental mutation of raw transactional records during analytical sessions.")
    ]
    for tag, desc in assumptions_p:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.25)
        p.paragraph_format.space_after = Pt(3)
        r_b = p.add_run(tag)
        r_b.bold = True
        r_b.font.color.rgb = COLOR_PRIMARY
        r_d = p.add_run(desc)
        r_d.font.color.rgb = COLOR_TEXT

    add_custom_heading("8.2 Technical and Academic Constraints", level=2)
    constraints_p = [
        ("CO-01: Zero Budget Infrastructure: ", "The project must execute entirely on zero-cost tiers, utilizing local open-source Python libraries, local SQLite storage, and free-tier Streamlit Community Cloud hosting."),
        ("CO-02: Streamlit Cloud Memory Limits: ", "Public cloud deployment on Streamlit Community Cloud imposes a hard constraint of 1 GB RAM allocation, necessitating aggressive in-memory caching (`@st.cache_data`) and query optimization."),
        ("CO-03: Strict Academic Timeline: ", "All deliverables, testing suites, academic project reports (ABES template), and presentation slides must be finalized within the 2nd year MCA curriculum semester milestone schedule."),
        ("CO-04: No Production Data Exposure: ", "In adherence to strict academic compliance and privacy guidelines, zero proprietary or PII (Personally Identifiable Information) customer data will be collected, stored, or processed."),
        ("CO-05: Single Developer Architecture: ", "System architecture, ETL script design, UI wireframing, testing, and deployment are engineered and maintained by a single student data analyst (Anu Tyagi).")
    ]
    for tag, desc in constraints_p:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.25)
        p.paragraph_format.space_after = Pt(3)
        r_b = p.add_run(tag)
        r_b.bold = True
        r_b.font.color.rgb = COLOR_SECONDARY
        r_d = p.add_run(desc)
        r_d.font.color.rgb = COLOR_TEXT

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # Section 10: Success Criteria
    add_custom_heading("9. Measurable Success Criteria", level=1)
    p = doc.add_paragraph(
        "Project validation is measured against unambiguous quantitative benchmarks established across technical, "
        "analytical, and academic dimensions:"
    )

    sc_table = doc.add_table(rows=7, cols=4)
    sc_headers = ["Criteria ID", "Success Dimension", "Target Measurement / Threshold", "Evaluation Method"]
    for i, h in enumerate(sc_headers):
        sc_table.rows[0].cells[i].paragraphs[0].text = h

    sc_data = [
        ["SC-01", "Query & UI Rendering Latency", "Dashboard loads initial view in < 2.0s; filter updates render in < 1.0s.", "Browser DevTools Performance audit and Streamlit execution profiler."],
        ["SC-02", "Operational Question Coverage", "100% coverage (10 of 10 OQs answered directly with dedicated visuals).", "Traceability Matrix (RTM) cross-verification in Step 4."],
        ["SC-03", "Data Quality & Normalization", "Zero duplicate primary keys, zero null foreign keys in transaction logs.", "Pytest database integrity test suite running automated assertions."],
        ["SC-04", "Automated Test Coverage", ">= 85% branch and statement test coverage across ETL, queries, and calculations.", "Pytest and pytest-cov automated coverage reports."],
        ["SC-05", "Cloud Deployment Health", "99.0% application uptime on Streamlit Community Cloud with public URL access.", "Live cloud verification and accessibility testing across devices."],
        ["SC-06", "Academic Milestone Approval", "Full compliance with ABES project template, passing mentor review with zero defects.", "Final project documentation sign-off by Mentor Mr. Dinesh."]
    ]
    for r_idx, row_vals in enumerate(sc_data, start=1):
        for c_idx, val in enumerate(row_vals):
            sc_table.rows[r_idx].cells[c_idx].paragraphs[0].text = val
    style_table(sc_table, [0.7, 1.8, 2.3, 1.7], [WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT])

    doc.save(filename)
    print(f"Document successfully created at: {filename}")

if __name__ == "__main__":
    import sys
    target_path = sys.argv[1] if len(sys.argv) > 1 else "docs/01_Project_Understanding.docx"
    create_project_understanding_doc(target_path)
