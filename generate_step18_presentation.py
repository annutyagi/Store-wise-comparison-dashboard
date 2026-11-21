import os
import zipfile
import xml.etree.ElementTree as ET

def build_presentation(output_path):
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    
    # 11 Slides definitions: (title, subtitle, category, bullet_points, table_data, notes)
    slides_data = [
        # Slide 1: Title
        {
            "type": "title",
            "title": "STORE WISE SALES COMPARISON DASHBOARD",
            "subtitle": "Enterprise Multi-Store Business Intelligence & Retail Velocity Analytics",
            "category": "ABES Engineering College | Department of Computer Applications | MCA Capstone Project",
            "bullets": [
                "Candidate Name: Anu Tyagi | University Roll No: 2500320140025",
                "Project ID: P052 | Domain: Data Analytics | Batch: ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01",
                "Faculty Project Mentor: Mr. Dinesh",
                "Academic Institution: ABES Engineering College, Ghaziabad",
                "GitHub Repository: github.com/annutyagi/Store-wise-comparison-dashboard",
                "Live Streamlit Application: store-wise-comparison-dashboard-tfgpqzxbwqggso9wax3wfm.streamlit.app"
            ],
            "notes": "Good morning respected mentor and examiners. I am Anu Tyagi, Roll 2500320140025, MCA 2nd Year student at ABES Engineering College. Today I am presenting my Capstone project: Store Wise Sales Comparison Dashboard, an enterprise retail business intelligence system developed under the guidance of Mr. Dinesh. The complete source code and live application are accessible on GitHub and Streamlit Cloud."
        },
        # Slide 2: Problem & HMW
        {
            "type": "content",
            "title": "Business Problem & Operational Challenge",
            "category": "Step 1 & Step 2: Problem Framing & Opportunity Formulation",
            "bullets": [
                "Commercial Blindspot: Multi-unit retail chains suffer from delayed, siloed spreadsheet reporting masking store-level margin leaks.",
                "Real-Estate Productivity Gap: Total gross sales fails to reflect square-foot capital efficiency across varying store formats.",
                "Uncontrolled Margin Erosion: Heavy promotional markdowns stimulate unit volume but severely erode operating profitability.",
                "Top HMW-01: How might we benchmark store revenue vs profit margins to identify margin-diluting stores? (OQ-01)",
                "Top HMW-02: How might we calculate Sales Velocity per SqFt and ATV across store sizes to evaluate real-estate ROI? (OQ-02)",
                "Operational Goal: Centralize 10 stores' data into a real-time BI dashboard with sub-second (< 1.0s) query response time."
            ],
            "notes": "Retail enterprises operating distributed physical stores face major operational blindspots. High-grossing stores often conceal negative operating margins, and square-foot productivity is rarely benchmarked against real-estate costs. We formulated 10 Operational Questions and 12 How Might We statements. Our two primary focus areas are: first, benchmarking revenue against profit margins, and second, calculating sales velocity per square foot to optimize store footprint capital allocation."
        },
        # Slide 3: Stakeholders
        {
            "type": "content",
            "title": "Key Stakeholders & Governance Alignment",
            "category": "Step 3 & Step 5: Stakeholder Ecosystem & RACI Matrix",
            "bullets": [
                "SH-01 Regional Retail Director (High Influence / High Interest): Accountable for territory profitability, store expansion, and monthly quota achievement.",
                "SH-02 Retail Store Manager (Medium Influence / High Interest): Manages floor staff shifts, daily checkout wait times, and weekend inventory staging.",
                "SH-03 Category Merchandiser (High Influence / Medium Interest): Oversees category product assortment, promotional discount caps, and margin protection.",
                "SH-04 Lead Business Analyst: Conducts customer spending distribution profiling, outlier detection, and statistical hypothesis testing.",
                "SH-05 & SH-06 Faculty Mentor & Evaluator: Guide architectural standards, mathematical rigor, and final academic capstone clearance.",
                "Governance: All project deliverables are governed by an 8-phase RACI matrix ensuring clear operational ownership."
            ],
            "notes": "Our stakeholder ecosystem comprises 6 key roles. The Regional Director requires high-level territory profitability and quota variance tracking. Store Managers need hourly customer volume trends to optimize staff shifts and weekend logistics. Category Merchandisers require discount elasticity analytics to prevent margin erosion. The Lead Analyst relies on automated statistical profiling. This alignment is governed by a formal RACI matrix across all SDLC stages."
        },
        # Slide 4: Traceability
        {
            "type": "content",
            "title": "Requirements Traceability Matrix (RTM)",
            "category": "Step 4 & Step 13: Bidirectional Requirements Traceability",
            "bullets": [
                "BR-01 (Multi-Store Benchmarking): HMW-01 -> US-03, 05 -> UI-01, 02 -> TC-F-01, 03, TC-T-07 | Must (Verified)",
                "BR-02 (Velocity per SqFt & ATV): HMW-02 -> US-06 -> UI-02 -> TC-F-04, TC-T-09 | Must (Verified)",
                "BR-03 (Category & Demographics): HMW-03 -> US-07 -> UI-03 -> TC-F-05 | Must (Verified)",
                "BR-04 (Temporal MoM & Festive Surge): HMW-04 -> US-04, 09 -> UI-01, 04 -> TC-F-02, 07 | Must (Verified)",
                "MoSCoW Balance: 14 Business Requirements (64.3% Must, 21.4% Should, 7.1% Could, 7.1% Won't/Deferred).",
                "Traceability Audit: 100% bidirectional coverage from business needs to automated test cases with zero gaps."
            ],
            "notes": "Slide 4 shows our Requirements Traceability Matrix. We have achieved 100% bidirectional traceability across all 14 Business Requirements. Every operational question traces directly through an HMW statement, a User Story, a UI Screen, and automated unit and functional test cases. For instance, BR-01 benchmarks store gross sales and margins, traced through User Story 05 and verified by tests TC-F-03 and TC-T-07. There are zero orphaned requirements in our architecture."
        },
        # Slide 5: Data Architecture
        {
            "type": "content",
            "title": "Relational Data Architecture & Hygiene Engine",
            "category": "Step 8 & Step 10: Database Schema & Dataset Validation",
            "bullets": [
                "Database Scope: Relational SQLite database (data/app.db) comprising 5 tables and 1 analytical view across FY 2024.",
                "Data Volumes: 10 stores, 50 products, 120 targets, 200 orders, 620 order items, and 36-column view_store_sales_analytics.",
                "Synthetic Realism: Fictional Indian retail dataset (seed 42) incorporating 5 empirical retail behavioral patterns.",
                "Cleaning Fix 1 (Quantity Invariant): Clamped all order item quantities to quantity >= 1 to prevent non-positive rows.",
                "Cleaning Fix 2 (Margin Invariant): Enforced selling_price >= cost_price across all 620 line items to prevent negative margin corruption.",
                "Cleaning Fix 3 (Imputation & Tie-Out): Median imputation for missing amounts; zero-variance double-entry order reconciliation."
            ],
            "notes": "Our data layer operates on a normalized SQLite database containing 5 relational tables: dim_stores, dim_products, dim_sales_targets, fact_orders, and fact_order_items. We synthesized realistic Indian retail data using random seed 42. In our data cleaning pipeline, we enforced three critical invariants: quantities must be at least one, selling prices must exceed unit cost prices, and double-entry order gross totals must tie out to line item sums with zero variance. Foreign key integrity was validated with zero orphan records."
        },
        # Slide 6: System Architecture
        {
            "type": "content",
            "title": "System Architecture & Processing Pipeline",
            "category": "Step 7 & Step 15: 5-Tier Decoupled Application Architecture",
            "bullets": [
                "Tier 1 (Storage): SQLite database (data/app.db) with PRAGMA foreign_keys = ON and pre-joined 36-column analytical view.",
                "Tier 2 (Ingestion & Cache): src/data_loader.py utilizing @st.cache_data for instant in-memory retrieval (0.038s latency).",
                "Tier 3 (Cleaning & Hygiene): src/cleaning.py enforcing domain invariants and logging missing value imputations.",
                "Tier 4 (Analytics Core): src/analytics.py implementing BURL-01..06 formulas, Welch's t-test, IQR, and skewness.",
                "Tier 5 (Presentation): Streamlit multipage app (app.py + 7 pages) with isolated Tailwind CSS components.",
                "Performance SLA: Average SQL analytical view query benchmark executes in 0.038 seconds, well below 1.0s limit."
            ],
            "notes": "Slide 6 outlines our 5-tier decoupled system architecture. Storage is managed in SQLite with foreign keys strictly enabled. Our data loading layer utilizes Streamlit caching so that the 36-column analytical view is loaded once into memory, delivering sub-second response times of 0.038 seconds. The analytics core in src/analytics.py calculates standardized retail formulas, Welch's t-test, and non-parametric statistics without UI dependencies, ensuring separation of concerns."
        },
        # Slide 7: Multipage Dashboard
        {
            "type": "content",
            "title": "Multipage Dashboard & Presentation Architecture",
            "category": "Step 12 & Step 15: UI Framework & Multipage Navigation",
            "bullets": [
                "Corporate Theme: Styled with corporate navy (#1A365D) and isolated iframe Tailwind CSS components.",
                "Global Filter Sync: Shared sidebar (src/sidebar.py) manages store, date, tier, and category filters via st.session_state.",
                "UI-01 Executive Overview (app.py): 4 high-level KPI cards, Regional Sales bar chart, MoM velocity line, strategic insight box.",
                "UI-02 Store Benchmarking: Store revenue rankings vs profit margin, Sales Velocity per SqFt scatter plot, store scorecard.",
                "UI-03 Category Demographics: Stacked category tier sales (Urban/Suburban/Rural), discount vs margin regression scatter.",
                "UI-04 to UI-08: Hourly trading rush, quota variance alerts, payment tender dynamics, data export, and quality audit."
            ],
            "notes": "Our dashboard provides an intuitive, executive-grade user experience built on Streamlit multipage architecture. We incorporated a corporate navy palette and modern Tailwind CSS KPI cards. A global sidebar synchronizes store, date range, tier, and category filters across all screens via session state. The 8 screens cover everything from high-level regional benchmarking and space productivity to hourly trading heatmaps, payment tender adoption, and one-click Excel data exporting."
        },
        # Slide 8: Key Analytical Findings
        {
            "type": "content",
            "title": "Key Analytical Findings & Statistical Proof",
            "category": "Empirical Discoveries from Store Transaction Analytics",
            "bullets": [
                "Finding 1 (+37.7% Q4 Festive Surge): October-November Diwali festival season generates 37.7% of annual retail net sales.",
                "Finding 2 (+26.1% Weekend Basket Premium): Statistically validated weekend ticket size surge via Welch's t-test (t = -11.42, p < 0.001).",
                "Finding 3 (Discount Elasticity Tipping Point, r = -0.584): Discounts > 25% cause sharp margin erosion of up to 14 percentage points.",
                "Real-Estate Space Productivity: Delhi Flagship commands top sales velocity (Rs 120/sqft), but Pune leads margin retention (42.4%).",
                "Tender Dynamics: UPI represents 52% of transaction volume, while Credit Card drives highest Average Transaction Value (+44% ATV).",
                "Statistical Rigor: All findings backed by unpooled standard error calculations and non-parametric skewness profiling."
            ],
            "notes": "Our analytics uncovered three major empirical insights. First, a dramatic 37.7% sales spike during the Q4 Diwali festive season. Second, a 26.1% premium in weekend average transaction values, which we proved statistically significant using Welch's t-test with t = -11.42 and p < 0.001. Third, an empirical negative correlation of r = -0.584 between discounts and margins, proving that discounts beyond 25% severely erode profitability. These findings provide clear operational guidelines for management."
        },
        # Slide 9: Testing & Quality Assurance
        {
            "type": "content",
            "title": "Automated Quality Assurance & Verification Suite",
            "category": "Step 13 & Step 15: Unit, Integration & Functional Verification",
            "bullets": [
                "Execution Summary: 15/15 Technical Unit Tests Passed via run_tests.py in 0.065 seconds (100% Pass Rate).",
                "Functional Verification: 16/16 Functional Test Cases (TC-F-01..16) verified covering filters, charts, and file downloads.",
                "TC-T-04 (Referential Integrity): Executed PRAGMA foreign_key_check; verified zero orphan child records in SQLite database.",
                "TC-T-10 (Quota Alert Engine): Verified +/- 5% threshold logic correctly triggers dynamic warning banners on underperformance.",
                "TC-T-12 (Statistical Test): Verified Welch's t-test unpooled variance calculations and Welch-Satterthwaite df (13.44).",
                "TC-F-16 (Safety Guard): Verified graceful warning banner rendering when filter queries return zero records, preventing UI crashes."
            ],
            "notes": "Quality assurance was built into every layer. Our automated test suite executes 15 technical unit and integration tests covering database schema validity, referential integrity, business formula precision, statistical algorithms, and query latency. All 15 tests pass in 0.065 seconds with zero failures. Additionally, all 16 functional user acceptance tests pass, including test TC-F-16 which ensures the application never crashes when conflicting filters produce an empty dataset."
        },
        # Slide 10: Recommendations & Roadmap
        {
            "type": "content",
            "title": "Business Recommendations, Limitations & Roadmap",
            "category": "Strategic Takeaways & Future Engineering Enhancements",
            "bullets": [
                "Recommendation 1 (Discount Ceiling): Enforce a strict 25% promotional discount cap to protect gross margins against erosion.",
                "Recommendation 2 (Shift Scheduling): Align floor staff shift rotas to bimodal peak hours (13:00-15:00 and 18:00-21:00).",
                "Recommendation 3 (Expansion Capital): Prioritize suburban boutique formats over downtown mega-stores to maximize margin ROI.",
                "Limitation 1: Single-currency INR scope; cross-border multi-currency FX engine (BR-14) deferred to Phase 2.",
                "Limitation 2: SQLite database runs on local read-only file; concurrent writes require client-server RDBMS.",
                "Future Roadmap: PostgreSQL migration, live POS IoT barcode scanner integration, and automated ERP replenishment triggers."
            ],
            "notes": "Based on our findings, we offer three strategic recommendations: enforce a 25% promotional discount ceiling, align store shift staffing to peak lunch and evening trading hours, and focus expansion capital on high-margin suburban formats. We also note our project limitations: it operates in single-currency INR and uses embedded SQLite. Our Phase 2 roadmap includes migrating to cloud PostgreSQL, live POS barcode integration, and automated inventory replenishment."
        },
        # Slide 11: Live Demo & Artifacts
        {
            "type": "title",
            "title": "Live Demonstration & Project Artifacts",
            "category": "ABES MCA Capstone Project Evaluation | Academic Year 2026-2027",
            "bullets": [
                "Live Streamlit Web Application: https://store-wise-comparison-dashboard-tfgpqzxbwqggso9wax3wfm.streamlit.app/",
                "GitHub Source Code Repository: https://github.com/annutyagi/Store-wise-comparison-dashboard",
                "Master Documentation Artifact: docs/2500320140025_Project_Document.docx (Complete 10-Section Capstone Report)",
                "Full Requirements Traceability: docs/04_Traceability_Matrix.xlsx (100% Bidirectional Coverage)",
                "Automated Verification Suite: python run_tests.py (15/15 Technical Tests Passing, 100% Pass Rate)",
                "Thank You! Open for Faculty Mentor & University Viva Examination."
            ],
            "notes": "In conclusion, the Store Wise Sales Comparison Dashboard is fully deployed, tested, and live at our Streamlit URL. All 18 SDLC deliverables—from problem understanding and database design to automated testing, deployment, and our master project report—are committed to GitHub. Thank you to Mr. Dinesh and the examination committee. I am now ready for the live demonstration and your questions."
        }
    ]

    # OOXML Templates for presentation
    # 1. [Content_Types].xml
    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
    <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
    <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
    <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
"""
    for i in range(1, 12):
        content_types += f'    <Override PartName="/ppt/slides/slide{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>\n'
        content_types += f'    <Override PartName="/ppt/notesSlides/notesSlide{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>\n'
    content_types += "</Types>"

    # 2. _rels/.rels
    root_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>"""

    # 3. ppt/presentation.xml
    pres_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
    <p:sldMasterIdLst>
        <p:sldMasterId id="2147483648" r:id="rId1"/>
    </p:sldMasterIdLst>
    <p:sldIdLst>
"""
    for i in range(1, 12):
        pres_xml += f'        <p:sldId id="{255 + i}" r:id="rId{i + 1}"/>\n'
    pres_xml += """    </p:sldIdLst>
    <p:sldSz cx="12192000" cy="6858000" type="screen16x9"/>
    <p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>"""

    # 4. ppt/_rels/presentation.xml.rels
    pres_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
"""
    for i in range(1, 12):
        pres_rels += f'    <Relationship Id="rId{i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{i}.xml"/>\n'
    pres_rels += "</Relationships>"

    # 5. ppt/theme/theme1.xml
    theme_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="ABES Theme">
    <a:themeElements>
        <a:clrScheme name="ABES Navy">
            <a:dk1><a:srgbClr val="1A365D"/></a:dk1>
            <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
            <a:dk2><a:srgbClr val="2D3748"/></a:dk2>
            <a:lt2><a:srgbClr val="F7FAFC"/></a:lt2>
            <a:accent1><a:srgbClr val="1A365D"/></a:accent1>
            <a:accent2><a:srgbClr val="2B6CB0"/></a:accent2>
            <a:accent3><a:srgbClr val="319795"/></a:accent3>
            <a:accent4><a:srgbClr val="DD6B20"/></a:accent4>
            <a:accent5><a:srgbClr val="3182CE"/></a:accent5>
            <a:accent6><a:srgbClr val="718096"/></a:accent6>
            <a:hlink><a:srgbClr val="2B6CB0"/></a:hlink>
            <a:folHlink><a:srgbClr val="4A5568"/></a:folHlink>
        </a:clrScheme>
        <a:fontScheme name="Calibri Scheme">
            <a:majorFont><a:latin typeface="Calibri"/></a:majorFont>
            <a:minorFont><a:latin typeface="Calibri"/></a:minorFont>
        </a:fontScheme>
        <a:fmtScheme name="Office">
            <a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>
            <a:lnStyleLst><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>
            <a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>
            <a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>
        </a:fmtScheme>
    </a:themeElements>
</a:theme>"""

    # 6. ppt/slideMasters/slideMaster1.xml
    sm_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
    <p:cSld>
        <p:bg><p:bgRef idx="1001"><a:schemeClr val="lt2"/></p:bgRef></p:bg>
        <p:spTree>
            <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
            <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
        </p:spTree>
    </p:cSld>
    <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
    <p:sldLayoutIdLst>
        <p:sldLayoutId id="2147483649" r:id="rId1"/>
    </p:sldLayoutIdLst>
</p:sldMaster>"""

    sm_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>"""

    # 7. ppt/slideLayouts/slideLayout1.xml
    sl_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank" preserve="1">
    <p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld>
    <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>"""

    sl_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>"""

    # Create ZIP archive in memory
    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", root_rels)
        z.writestr("ppt/presentation.xml", pres_xml)
        z.writestr("ppt/_rels/presentation.xml.rels", pres_rels)
        z.writestr("ppt/theme/theme1.xml", theme_xml)
        z.writestr("ppt/slideMasters/slideMaster1.xml", sm_xml)
        z.writestr("ppt/slideMasters/_rels/slideMaster1.xml.rels", sm_rels)
        z.writestr("ppt/slideLayouts/slideLayout1.xml", sl_xml)
        z.writestr("ppt/slideLayouts/_rels/slideLayout1.xml.rels", sl_rels)

        for i, s_data in enumerate(slides_data, start=1):
            is_title = s_data["type"] == "title"
            title_text = s_data["title"]
            cat_text = s_data["category"]
            bullets = s_data["bullets"]
            notes_text = s_data["notes"]

            # Escape XML entities
            def clean_xml(txt):
                return txt.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")

            # Slide XML
            # Header Box (Top Banner)
            # Body Card Box
            # Footer Tag
            slide_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
    <p:cSld>
        <p:spTree>
            <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
            <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
            
            <!-- Background Shape -->
            <p:sp>
                <p:nvSpPr><p:cNvPr id="2" name="Slide Background"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr/></p:nvSpPr>
                <p:spPr>
                    <a:xfrm><a:off x="0" y="0"/><a:ext cx="12192000" cy="6858000"/></a:xfrm>
                    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                    <a:solidFill><a:srgbClr val="F7FAFC"/></a:solidFill>
                </p:spPr>
            </p:sp>

            <!-- Header Banner -->
            <p:sp>
                <p:nvSpPr><p:cNvPr id="3" name="Header Banner"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
                <p:spPr>
                    <a:xfrm><a:off x="609600" y="457200"/><a:ext cx="10972800" cy="1143000"/></a:xfrm>
                    <a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 12000"/></a:avLst></a:prstGeom>
                    <a:solidFill><a:srgbClr val="1A365D"/></a:solidFill>
                </p:spPr>
                <p:txBody>
                    <a:bodyPr vert="horz" lIns="254000" tIns="127000" rIns="254000" bIns="127000" anchor="ctr"/>
                    <a:lstStyle/>
                    <a:p>
                        <a:pPr algn="l"/>
                        <a:r>
                            <a:rPr lang="en-US" sz="1100" b="1">
                                <a:solidFill><a:srgbClr val="63B3ED"/></a:solidFill>
                            </a:rPr>
                            <a:t>{clean_xml(cat_text)}</a:t>
                        </a:r>
                    </a:p>
                    <a:p>
                        <a:pPr algn="l"/>
                        <a:r>
                            <a:rPr lang="en-US" sz="2200" b="1">
                                <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
                            </a:rPr>
                            <a:t>{clean_xml(title_text)}</a:t>
                        </a:r>
                    </a:p>
                </p:txBody>
            </p:sp>

            <!-- Body Card Shape -->
            <p:sp>
                <p:nvSpPr><p:cNvPr id="4" name="Content Card"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
                <p:spPr>
                    <a:xfrm><a:off x="609600" y="1828800"/><a:ext cx="10972800" cy="4419600"/></a:xfrm>
                    <a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 8000"/></a:avLst></a:prstGeom>
                    <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
                    <a:ln w="12700"><a:solidFill><a:srgbClr val="CBD5E0"/></a:solidFill></a:ln>
                </p:spPr>
                <p:txBody>
                    <a:bodyPr vert="horz" lIns="381000" tIns="304800" rIns="381000" bIns="304800" anchor="t"/>
                    <a:lstStyle/>
"""
            for bullet in bullets:
                # Highlight bold prefixes before colon
                if ":" in bullet:
                    prefix, rest = bullet.split(":", 1)
                    slide_xml += f"""
                    <a:p>
                        <a:pPr marL="285750" indent="-285750" spaceBefore="120000" spaceAfter="60000">
                            <a:buFont typeface="Arial"/>
                            <a:buChar char="▪"/>
                        </a:pPr>
                        <a:r>
                            <a:rPr lang="en-US" sz="1350" b="1">
                                <a:solidFill><a:srgbClr val="1A365D"/></a:solidFill>
                            </a:rPr>
                            <a:t>{clean_xml(prefix)}:</a:t>
                        </a:r>
                        <a:r>
                            <a:rPr lang="en-US" sz="1350">
                                <a:solidFill><a:srgbClr val="2D3748"/></a:solidFill>
                            </a:rPr>
                            <a:t>{clean_xml(rest)}</a:t>
                        </a:r>
                    </a:p>
"""
                else:
                    slide_xml += f"""
                    <a:p>
                        <a:pPr marL="285750" indent="-285750" spaceBefore="120000" spaceAfter="60000">
                            <a:buFont typeface="Arial"/>
                            <a:buChar char="▪"/>
                        </a:pPr>
                        <a:r>
                            <a:rPr lang="en-US" sz="1350">
                                <a:solidFill><a:srgbClr val="2D3748"/></a:solidFill>
                            </a:rPr>
                            <a:t>{clean_xml(bullet)}</a:t>
                        </a:r>
                    </a:p>
"""

            # Footer text inside slide
            slide_xml += f"""
                </p:txBody>
            </p:sp>

            <!-- Slide Footer -->
            <p:sp>
                <p:nvSpPr><p:cNvPr id="5" name="Slide Footer"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
                <p:spPr>
                    <a:xfrm><a:off x="609600" y="6324600"/><a:ext cx="10972800" cy="381000"/></a:xfrm>
                    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                </p:spPr>
                <p:txBody>
                    <a:bodyPr vert="horz" lIns="0" tIns="0" rIns="0" bIns="0" anchor="ctr"/>
                    <a:lstStyle/>
                    <a:p>
                        <a:pPr algn="l"/>
                        <a:r>
                            <a:rPr lang="en-US" sz="900" b="1">
                                <a:solidFill><a:srgbClr val="718096"/></a:solidFill>
                            </a:rPr>
                            <a:t>Anu Tyagi | Roll: 2500320140025 | MCA 2nd Year, ABES Engineering College | Mentor: Mr. Dinesh</a:t>
                        </a:r>
                        <a:r>
                            <a:rPr lang="en-US" sz="900" b="1">
                                <a:solidFill><a:srgbClr val="2B6CB0"/></a:solidFill>
                            </a:rPr>
                            <a:t>                                                          Slide {i} of 11</a:t>
                        </a:r>
                    </a:p>
                </p:txBody>
            </p:sp>
        </p:spTree>
    </p:cSld>
    <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>"""

            # Slide Rels
            slide_rels = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide{i}.xml"/>
</Relationships>"""

            # Notes Slide XML
            notes_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
    <p:cSld>
        <p:spTree>
            <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
            <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
            <p:sp>
                <p:nvSpPr><p:cNvPr id="2" name="Notes Text Box"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>
                <p:spPr><a:xfrm><a:off x="431800" y="431800"/><a:ext cx="5994400" cy="8280400"/></a:xfrm></p:spPr>
                <p:txBody>
                    <a:bodyPr/><a:lstStyle/>
                    <a:p>
                        <a:pPr spaceAfter="100000"/>
                        <a:r>
                            <a:rPr lang="en-US" sz="1200" b="1">
                                <a:solidFill><a:srgbClr val="1A365D"/></a:solidFill>
                            </a:rPr>
                            <a:t>SPEAKER NOTES (Spoken Script ~30 Seconds):</a:t>
                        </a:r>
                    </a:p>
                    <a:p>
                        <a:r>
                            <a:rPr lang="en-US" sz="1100">
                                <a:solidFill><a:srgbClr val="2D3748"/></a:solidFill>
                            </a:rPr>
                            <a:t>{clean_xml(notes_text)}</a:t>
                        </a:r>
                    </a:p>
                </p:txBody>
            </p:sp>
        </p:spTree>
    </p:cSld>
    <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:notes>"""

            # Notes Slide Rels
            notes_rels = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="../slides/slide{i}.xml"/>
</Relationships>"""

            z.writestr(f"ppt/slides/slide{i}.xml", slide_xml)
            z.writestr(f"ppt/slides/_rels/slide{i}.xml.rels", slide_rels)
            z.writestr(f"ppt/notesSlides/notesSlide{i}.xml", notes_xml)
            z.writestr(f"ppt/notesSlides/_rels/notesSlide{i}.xml.rels", notes_rels)

    print(f"Presentation successfully created at: {output_path}")

if __name__ == "__main__":
    import sys
    out = sys.argv[1] if len(sys.argv) > 1 else "docs/2500320140025_Presentation.pptx"
    build_presentation(out)
