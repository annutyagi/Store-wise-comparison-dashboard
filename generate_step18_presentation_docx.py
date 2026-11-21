import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

def set_cell_background(cell, hex_color):
    tcPr = cell._element.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=80, bottom=80, left=100, right=100):
    tcPr = cell._element.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

def create_presentation_docx(filename):
    os.makedirs(os.path.dirname(filename), exist_ok=True)
    doc = docx.Document()
    
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)
        
    COLOR_PRIMARY = RGBColor(26, 54, 93)     # #1A365D Corporate Navy
    COLOR_SECONDARY = RGBColor(43, 108, 176) # #2B6CB0 Blue
    COLOR_DARK = RGBColor(45, 55, 72)        # #2D3748 Charcoal
    COLOR_MUTED = RGBColor(113, 128, 150)    # #718096 Slate Grey
    HEX_HEADER = "1A365D"
    HEX_ZEBRA = "F7FAFC"
    HEX_BANNER = "EBF8FF"
    HEX_SUCCESS = "C6F6D5"

    # Institutional Header
    p_inst = doc.add_paragraph()
    p_inst.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_inst = p_inst.add_run("ABES ENGINEERING COLLEGE, GHAZIABAD\nDEPARTMENT OF COMPUTER APPLICATIONS (MCA)")
    r_inst.font.name = "Calibri"
    r_inst.font.size = Pt(11)
    r_inst.font.bold = True
    r_inst.font.color.rgb = COLOR_SECONDARY

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_title = p_title.add_run("CAPSTONE VIVA PRESENTATION SLIDE DECK & SPEAKER SCRIPT")
    r_title.font.name = "Calibri"
    r_title.font.size = Pt(20)
    r_title.font.bold = True
    r_title.font.color.rgb = COLOR_PRIMARY

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sub = p_sub.add_run("Store Wise Sales Comparison Dashboard | Official 11-Slide Master Viva Deck\nDeliverable: STEP 18 — Viva Presentation Slide Deck & Rehearsal Guide (v1.0)")
    r_sub.font.name = "Calibri"
    r_sub.font.size = Pt(11)
    r_sub.font.italic = True
    r_sub.font.color.rgb = COLOR_MUTED

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # Document Control Block
    ctrl_table = doc.add_table(rows=5, cols=4)
    ctrl_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    ctrl_data = [
        ("Project Name", "Store Wise Sales Comparison Dashboard", "Deliverable ID", "DOC-18 / STEP 18"),
        ("Student Author", "Anu Tyagi (Roll 2500320140025)", "Academic Degree", "MCA 2nd Year (Data Analytics)"),
        ("Faculty Mentor", "Mr. Dinesh", "Institution", "ABES Engineering College, Ghaziabad"),
        ("Total Slides", "11 Slides (16:9 Widescreen)", "Presentation Time", "5 Minutes (30 Seconds / Slide)"),
        ("PowerPoint File", "docs/2500320140025_Presentation.pptx", "Live App URL", "store-wise-comparison-dashboard-tfgpqzxbwqggso9wax3wfm.streamlit.app")
    ]
    for row_idx, data in enumerate(ctrl_data):
        row = ctrl_table.rows[row_idx]
        for col_idx in range(4):
            cell = row.cells[col_idx]
            cell.text = data[col_idx]
            p = cell.paragraphs[0]
            p.runs[0].font.name = "Calibri"
            p.runs[0].font.size = Pt(9.5)
            if col_idx in [0, 2]:
                p.runs[0].font.bold = True
                p.runs[0].font.color.rgb = COLOR_PRIMARY
                set_cell_background(cell, HEX_BANNER)
            else:
                p.runs[0].font.color.rgb = COLOR_DARK
                set_cell_background(cell, "FFFFFF")
            set_cell_margins(cell, top=70, bottom=70, left=100, right=100)

    doc.add_paragraph().paragraph_format.space_after = Pt(14)

    # Section 1: 11-Slide Breakdown with Speaker Notes
    doc.add_heading(level=1).add_run("1. Master 11-Slide Presentation Content & Speaker Scripts").font.color.rgb = COLOR_PRIMARY

    slides_info = [
        (1, "Title Slide", "Store Wise Sales Comparison Dashboard",
         [
             "Candidate Name: Anu Tyagi | University Roll No: 2500320140025",
             "Academic Program: MCA 2nd Year (Data Analytics Specialization), ABES Engineering College",
             "Faculty Project Mentor: Mr. Dinesh",
             "Project Title: Store Wise Sales Comparison Dashboard (Enterprise Retail BI Platform)",
             "Public GitHub Repository: https://github.com/annutyagi/Store-wise-comparison-dashboard",
             "Live Streamlit Web Application: https://store-wise-comparison-dashboard-tfgpqzxbwqggso9wax3wfm.streamlit.app/"
         ],
         "Good morning respected mentor and examiners. I am Anu Tyagi, Roll 2500320140025, MCA 2nd Year student at ABES Engineering College. Today I am presenting my Capstone project: Store Wise Sales Comparison Dashboard, an enterprise retail business intelligence system developed under the guidance of Mr. Dinesh. The complete source code and live application are accessible on GitHub and Streamlit Cloud."),

        (2, "Problem & HMWs", "Business Problem & Operational Challenge",
         [
             "Commercial Blindspot: Multi-unit retail chains suffer from delayed, siloed spreadsheet reporting masking store-level margin leaks.",
             "Real-Estate Productivity Gap: Total gross sales fails to reflect square-foot capital efficiency across varying store formats.",
             "Uncontrolled Margin Erosion: Heavy promotional markdowns stimulate unit volume but severely erode operating profitability.",
             "Top HMW-01: How might we benchmark store revenue vs profit margins to identify margin-diluting stores? (OQ-01)",
             "Top HMW-02: How might we calculate Sales Velocity per SqFt and ATV across store sizes to evaluate real-estate ROI? (OQ-02)",
             "Operational Goal: Centralize 10 stores' data into a real-time BI dashboard with sub-second (< 1.0s) query response time."
         ],
         "Retail enterprises operating distributed physical stores face major operational blindspots. High-grossing stores often conceal negative operating margins, and square-foot productivity is rarely benchmarked against real-estate costs. We formulated 10 Operational Questions and 12 How Might We statements. Our two primary focus areas are: first, benchmarking revenue against profit margins, and second, calculating sales velocity per square foot to optimize store footprint capital allocation."),

        (3, "Stakeholders", "Key Stakeholders & Governance Alignment",
         [
             "SH-01 Regional Retail Director (High Influence / High Interest): Accountable for territory profitability, store expansion, and monthly quota achievement.",
             "SH-02 Retail Store Manager (Medium Influence / High Interest): Manages floor staff shifts, daily checkout wait times, and weekend inventory staging.",
             "SH-03 Category Merchandiser (High Influence / Medium Interest): Oversees category product assortment, promotional discount caps, and margin protection.",
             "SH-04 Lead Business Analyst: Conducts customer spending distribution profiling, outlier detection, and statistical hypothesis testing.",
             "SH-05 & SH-06 Faculty Mentor & Evaluator: Guide architectural standards, mathematical rigor, and final academic capstone clearance.",
             "Governance: All project deliverables are governed by an 8-phase RACI matrix ensuring clear operational ownership."
         ],
         "Our stakeholder ecosystem comprises 6 key roles. The Regional Director requires high-level territory profitability and quota variance tracking. Store Managers need hourly customer volume trends to optimize staff shifts and weekend logistics. Category Merchandisers require discount elasticity analytics to prevent margin erosion. The Lead Analyst relies on automated statistical profiling. This alignment is governed by a formal RACI matrix across all SDLC stages."),

        (4, "Traceability", "Requirements Traceability Matrix (RTM)",
         [
             "BR-01 (Multi-Store Benchmarking): HMW-01 -> US-03, 05 -> UI-01, 02 -> TC-F-01, 03, TC-T-07 | Must (Verified)",
             "BR-02 (Velocity per SqFt & ATV): HMW-02 -> US-06 -> UI-02 -> TC-F-04, TC-T-09 | Must (Verified)",
             "BR-03 (Category & Demographics): HMW-03 -> US-07 -> UI-03 -> TC-F-05 | Must (Verified)",
             "BR-04 (Temporal MoM & Festive Surge): HMW-04 -> US-04, 09 -> UI-01, 04 -> TC-F-02, 07 | Must (Verified)",
             "MoSCoW Balance: 14 Business Requirements (64.3% Must, 21.4% Should, 7.1% Could, 7.1% Won't/Deferred).",
             "Traceability Audit: 100% bidirectional coverage from business needs to automated test cases with zero gaps."
         ],
         "Slide 4 shows our Requirements Traceability Matrix. We have achieved 100% bidirectional traceability across all 14 Business Requirements. Every operational question traces directly through an HMW statement, a User Story, a UI Screen, and automated unit and functional test cases. For instance, BR-01 benchmarks store gross sales and margins, traced through User Story 05 and verified by tests TC-F-03 and TC-T-07. There are zero orphaned requirements in our architecture."),

        (5, "Data Architecture", "Relational Data Architecture & Hygiene Engine",
         [
             "Database Scope: Relational SQLite database (data/app.db) comprising 5 tables and 1 analytical view across FY 2024.",
             "Data Volumes: 10 stores, 50 products, 120 targets, 200 orders, 620 order items, and 36-column view_store_sales_analytics.",
             "Synthetic Realism: Fictional Indian retail dataset (seed 42) incorporating 5 empirical retail behavioral patterns.",
             "Cleaning Fix 1 (Quantity Invariant): Clamped all order item quantities to quantity >= 1 to prevent non-positive rows.",
             "Cleaning Fix 2 (Margin Invariant): Enforced selling_price >= cost_price across all 620 line items to prevent negative margin corruption.",
             "Cleaning Fix 3 (Imputation & Tie-Out): Median imputation for missing amounts; zero-variance double-entry order reconciliation."
         ],
         "Our data layer operates on a normalized SQLite database containing 5 relational tables: dim_stores, dim_products, dim_sales_targets, fact_orders, and fact_order_items. We synthesized realistic Indian retail data using random seed 42. In our data cleaning pipeline, we enforced three critical invariants: quantities must be at least one, selling prices must exceed unit cost prices, and double-entry order gross totals must tie out to line item sums with zero variance. Foreign key integrity was validated with zero orphan records."),

        (6, "System Architecture", "System Architecture & Processing Pipeline",
         [
             "Tier 1 (Storage): SQLite database (data/app.db) with PRAGMA foreign_keys = ON and pre-joined 36-column analytical view.",
             "Tier 2 (Ingestion & Cache): src/data_loader.py utilizing @st.cache_data for instant in-memory retrieval (0.038s latency).",
             "Tier 3 (Cleaning & Hygiene): src/cleaning.py enforcing domain invariants and logging missing value imputations.",
             "Tier 4 (Analytics Core): src/analytics.py implementing BURL-01..06 formulas, Welch's t-test, IQR, and skewness.",
             "Tier 5 (Presentation): Streamlit multipage app (app.py + 7 pages) with isolated Tailwind CSS components.",
             "Performance SLA: Average SQL analytical view query benchmark executes in 0.038 seconds, well below 1.0s limit."
         ],
         "Slide 6 outlines our 5-tier decoupled system architecture. Storage is managed in SQLite with foreign keys strictly enabled. Our data loading layer utilizes Streamlit caching so that the 36-column analytical view is loaded once into memory, delivering sub-second response times of 0.038 seconds. The analytics core in src/analytics.py calculates standardized retail formulas, Welch's t-test, and non-parametric statistics without UI dependencies, ensuring separation of concerns."),

        (7, "Multipage Dashboard", "Multipage Dashboard & Presentation Architecture",
         [
             "Corporate Theme: Styled with corporate navy (#1A365D) and isolated iframe Tailwind CSS components.",
             "Global Filter Sync: Shared sidebar (src/sidebar.py) manages store, date, tier, and category filters via st.session_state.",
             "UI-01 Executive Overview (app.py): 4 high-level KPI cards, Regional Sales bar chart, MoM velocity line, strategic insight box.",
             "UI-02 Store Benchmarking: Store revenue rankings vs profit margin, Sales Velocity per SqFt scatter plot, store scorecard.",
             "UI-03 Category Demographics: Stacked category tier sales (Urban/Suburban/Rural), discount vs margin regression scatter.",
             "UI-04 to UI-08: Hourly trading rush, quota variance alerts, payment tender dynamics, data export, and quality audit."
         ],
         "Our dashboard provides an intuitive, executive-grade user experience built on Streamlit multipage architecture. We incorporated a corporate navy palette and modern Tailwind CSS KPI cards. A global sidebar synchronizes store, date range, tier, and category filters across all screens via session state. The 8 screens cover everything from high-level regional benchmarking and space productivity to hourly trading heatmaps, payment tender adoption, and one-click Excel data exporting."),

        (8, "Key Analytical Findings", "Key Analytical Findings & Statistical Proof",
         [
             "Finding 1 (+37.7% Q4 Festive Surge): October-November Diwali festival season generates 37.7% of annual retail net sales.",
             "Finding 2 (+26.1% Weekend Basket Premium): Statistically validated weekend ticket size surge via Welch's t-test (t = -11.42, p < 0.001).",
             "Finding 3 (Discount Elasticity Tipping Point, r = -0.584): Discounts > 25% cause sharp margin erosion of up to 14 percentage points.",
             "Real-Estate Space Productivity: Delhi Flagship commands top sales velocity (Rs 120/sqft), but Pune leads margin retention (42.4%).",
             "Tender Dynamics: UPI represents 52% of transaction volume, while Credit Card drives highest Average Transaction Value (+44% ATV).",
             "Statistical Rigor: All findings backed by unpooled standard error calculations and non-parametric skewness profiling."
         ],
         "Our analytics uncovered three major empirical insights. First, a dramatic 37.7% sales spike during the Q4 Diwali festive season. Second, a 26.1% premium in weekend average transaction values, which we proved statistically significant using Welch's t-test with t = -11.42 and p < 0.001. Third, an empirical negative correlation of r = -0.584 between discounts and margins, proving that discounts beyond 25% severely erode profitability. These findings provide clear operational guidelines for management."),

        (9, "Testing & QA", "Automated Quality Assurance & Verification Suite",
         [
             "Execution Summary: 15/15 Technical Unit Tests Passed via run_tests.py in 0.065 seconds (100% Pass Rate).",
             "Functional Verification: 16/16 Functional Test Cases (TC-F-01..16) verified covering filters, charts, and file downloads.",
             "TC-T-04 (Referential Integrity): Executed PRAGMA foreign_key_check; verified zero orphan child records in SQLite database.",
             "TC-T-10 (Quota Alert Engine): Verified +/- 5% threshold logic correctly triggers dynamic warning banners on underperformance.",
             "TC-T-12 (Statistical Test): Verified Welch's t-test unpooled variance calculations and Welch-Satterthwaite df (13.44).",
             "TC-F-16 (Safety Guard): Verified graceful warning banner rendering when filter queries return zero records, preventing UI crashes."
         ],
         "Quality assurance was built into every layer. Our automated test suite executes 15 technical unit and integration tests covering database schema validity, referential integrity, business formula precision, statistical algorithms, and query latency. All 15 tests pass in 0.065 seconds with zero failures. Additionally, all 16 functional user acceptance tests pass, including test TC-F-16 which ensures the application never crashes when conflicting filters produce an empty dataset."),

        (10, "Recommendations & Roadmap", "Business Recommendations, Limitations & Roadmap",
         [
             "Recommendation 1 (Discount Ceiling): Enforce a strict 25% promotional discount cap to protect gross margins against erosion.",
             "Recommendation 2 (Shift Scheduling): Align floor staff shift rotas to bimodal peak hours (13:00-15:00 and 18:00-21:00).",
             "Recommendation 3 (Expansion Capital): Prioritize suburban boutique formats over downtown mega-stores to maximize margin ROI.",
             "Limitation 1: Single-currency INR scope; cross-border multi-currency FX engine (BR-14) deferred to Phase 2.",
             "Limitation 2: SQLite database runs on local read-only file; concurrent writes require client-server RDBMS.",
             "Future Roadmap: PostgreSQL migration, live POS IoT barcode scanner integration, and automated ERP replenishment triggers."
         ],
         "Based on our findings, we offer three strategic recommendations: enforce a 25% promotional discount ceiling, align store shift staffing to peak lunch and evening trading hours, and focus expansion capital on high-margin suburban formats. We also note our project limitations: it operates in single-currency INR and uses embedded SQLite. Our Phase 2 roadmap includes migrating to cloud PostgreSQL, live POS barcode integration, and automated inventory replenishment."),

        (11, "Live Demo & Artifacts", "Live Demonstration & Project Artifacts",
         [
             "Live Streamlit Web Application: https://store-wise-comparison-dashboard-tfgpqzxbwqggso9wax3wfm.streamlit.app/",
             "GitHub Source Code Repository: https://github.com/annutyagi/Store-wise-comparison-dashboard",
             "Master Documentation Artifact: docs/2500320140025_Project_Document.docx (Complete 10-Section Capstone Report)",
             "Full Requirements Traceability: docs/04_Traceability_Matrix.xlsx (100% Bidirectional Coverage)",
             "Automated Verification Suite: python run_tests.py (15/15 Technical Tests Passing, 100% Pass Rate)",
             "Thank You! Open for Faculty Mentor & University Viva Examination."
         ],
         "In conclusion, the Store Wise Sales Comparison Dashboard is fully deployed, tested, and live at our Streamlit URL. All 18 SDLC deliverables—from problem understanding and database design to automated testing, deployment, and our master project report—are committed to GitHub. Thank you to Mr. Dinesh and the examination committee. I am now ready for the live demonstration and your questions.")
    ]

    for s_num, s_tag, s_title, s_bullets, s_notes in slides_info:
        h = doc.add_heading(level=2)
        r = h.add_run(f"Slide {s_num}: {s_title} ({s_tag})")
        r.font.bold = True
        r.font.color.rgb = COLOR_PRIMARY
        
        # Slide content box
        tbl = doc.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell = tbl.rows[0].cells[0]
        set_cell_background(cell, HEX_ZEBRA)
        set_cell_margins(cell, top=100, bottom=100, left=140, right=140)
        
        p_c = cell.paragraphs[0]
        r_c = p_c.add_run("SLIDE CONTENT (Max 6 Bullet Points):\n")
        r_c.font.bold = True
        r_c.font.size = Pt(9.5)
        r_c.font.color.rgb = COLOR_SECONDARY
        
        for b in s_bullets:
            p_b = cell.add_paragraph()
            p_b.paragraph_format.space_before = Pt(2)
            p_b.paragraph_format.space_after = Pt(2)
            r_b = p_b.add_run(f"• {b}")
            r_b.font.size = Pt(8.5)
            r_b.font.color.rgb = COLOR_DARK
            
        doc.add_paragraph().paragraph_format.space_after = Pt(4)
        
        # Speaker script box
        tbl_n = doc.add_table(rows=1, cols=1)
        tbl_n.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell_n = tbl_n.rows[0].cells[0]
        set_cell_background(cell_n, HEX_BANNER)
        set_cell_margins(cell_n, top=80, bottom=80, left=120, right=120)
        
        p_n = cell_n.paragraphs[0]
        r_nt = p_n.add_run("SPEAKER SCRIPT (What to Say — 30 Seconds Rehearsal):\n")
        r_nt.font.bold = True
        r_nt.font.size = Pt(9.0)
        r_nt.font.color.rgb = COLOR_PRIMARY
        
        p_ns = cell_n.add_paragraph()
        r_nss = p_ns.add_run(f'"{s_notes}"')
        r_nss.font.size = Pt(9.0)
        r_nss.font.italic = True
        r_nss.font.color.rgb = COLOR_DARK
        
        doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # Rehearsal & Presentation Tips
    doc.add_heading(level=1).add_run("2. 5-Minute Viva Rehearsal Strategy & Examiners Guide").font.color.rgb = COLOR_PRIMARY
    tips = [
        ("Timing Discipline", "Spend strictly 25-30 seconds per slide. Slides 1-3 establish context (1.5 min), Slides 4-7 present technical architecture (1.5 min), and Slides 8-11 demonstrate empirical proof and live demo (2 min)."),
        ("Live Demonstration Transition", "At Slide 11, transition seamlessly to the open browser tab running store-wise-comparison-dashboard-tfgpqzxbwqggso9wax3wfm.streamlit.app. Demonstrate the multi-store filter, date range slider, and CSV download button."),
        ("Verification Proof", "Keep terminal open running 'python run_tests.py' showing 15/15 tests passing in 0.065 seconds as tangible proof of engineering quality."),
        ("Defense of Metrics", "When asked about Sales Velocity per SqFt or Welch's t-test, reference the exact mathematical formulations (BURL-04, Welch-Satterthwaite equation) from the Master Project Document.")
    ]
    for title, desc in tips:
        p_t = doc.add_paragraph()
        r1 = p_t.add_run(f"• {title}: ")
        r1.font.bold = True
        r1.font.color.rgb = COLOR_PRIMARY
        r2 = p_t.add_run(desc)
        r2.font.color.rgb = COLOR_DARK

    doc.save(filename)
    print(f"Presentation Companion Word Document successfully created at: {filename}")

if __name__ == "__main__":
    import sys
    target = sys.argv[1] if len(sys.argv) > 1 else "docs/18_Presentation.docx"
    create_presentation_docx(target)
