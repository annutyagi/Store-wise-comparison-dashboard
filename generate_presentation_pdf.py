import os
import sys
from reportlab.lib.pagesizes import landscape, A4
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, PageBreak, Table, TableStyle
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfgen import canvas

# Import slide definitions from generate_step18_presentation.py
import generate_step18_presentation as pres_src

class SlideNumberedCanvas(canvas.Canvas):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.draw_slide_chrome(num_pages)
            super().showPage()
        super().save()

    def draw_slide_chrome(self, page_count):
        self.saveState()
        # Footer
        self.setFont("Helvetica", 8)
        self.setFillColor(colors.HexColor("#718096"))
        self.drawString(36, 18, "Anu Tyagi | Roll: 2500320140025 | MCA 2nd Year, ABES Engineering College, Ghaziabad | Mentor: Mr. Dinesh")
        self.drawRightString(806, 18, f"Slide {self._pageNumber} of {page_count}")
        self.setStrokeColor(colors.HexColor("#CBD5E0"))
        self.setLineWidth(0.5)
        self.line(36, 28, 806, 28)
        self.restoreState()

def build_presentation_pdf(output_paths):
    # Dummy runner to extract slides data
    # We inspect generate_step18_presentation.py data
    slides = [
        # Slide 1
        {
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
            "notes": "Good morning respected mentor Mr. Dinesh and examiners. I am Anu Tyagi, Roll 2500320140025, MCA 2nd Year student at ABES Engineering College, Batch ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01. Today I am presenting my Capstone project: Store Wise Sales Comparison Dashboard (Project ID: P052) in the Data Analytics domain, developed under the guidance of Mr. Dinesh. The complete source code and live application are accessible on GitHub and Streamlit Cloud."
        },
        # Slide 2
        {
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
        # Slide 3
        {
            "title": "Key Stakeholders & Governance Alignment",
            "category": "Step 3 & Step 5: Stakeholder Ecosystem & RACI Matrix",
            "bullets": [
                "SH-01 Regional Retail Director (High Influence / High Interest): Accountable for territory profitability, store expansion, and monthly quota achievement.",
                "SH-02 Retail Store Manager (Medium Influence / High Interest): Manages floor staff shifts, daily checkout wait times, and weekend inventory staging.",
                "SH-03 Category Merchandiser (High Influence / Medium Interest): Oversees category product assortment, promotional discount caps, and margin protection.",
                "SH-04 Lead Business Analyst: Conducts customer spending distribution profiling, outlier detection, and statistical hypothesis testing.",
                "SH-05 Faculty Project Mentor (Mr. Dinesh): Evaluates architectural rigor, data integrity, and compliance with academic capstone standards.",
                "Governance Model: Interactive filter cascading and role-specific drilldowns deliver targeted operational value across all decision tiers."
            ],
            "notes": "Our stakeholder governance model aligns executive, operational, and merchandising needs. The Regional Retail Director relies on high-level margin scorecards; Store Managers monitor peak footfall rush hours for staff scheduling; and Merchandisers protect profit margins by analyzing discount elasticity. Academic and engineering rigor is overseen by our project mentor Mr. Dinesh."
        },
        # Slide 4
        {
            "title": "End-to-End Requirements Traceability (RTM)",
            "category": "Step 4: 100% Bidirectional SDLC Alignment",
            "table": [
                ["HMW ID", "BR ID", "User Story", "Test Case", "UI Screen", "Priority / Status"],
                ["HMW-01", "BR-01", "US-03", "TC-F-01 / TC-T-07", "UI-01 / UI-02", "Must | VERIFIED"],
                ["HMW-02", "BR-02", "US-04", "TC-F-02 / TC-T-08", "UI-02", "Must | VERIFIED"],
                ["HMW-05", "BR-05", "US-07", "TC-F-05 / TC-T-09", "UI-03", "Must | VERIFIED"],
                ["HMW-08", "BR-08", "US-11", "TC-F-08 / TC-T-10", "UI-04", "Must | VERIFIED"]
            ],
            "bullets": [
                "100% Traceability: Every business problem maps directly from HMW to Requirements, User Stories, UI Screens, and Automated Tests.",
                "Zero Orphan Requirements: All 14 Business Requirements (BR-01..14) possess verified test cases and database lineage.",
                "Audit-Ready Governance: Changes at any tier can be evaluated instantly for regression impact across the 36-column schema."
            ],
            "notes": "Slide 4 showcases our Requirements Traceability Matrix. We have achieved 100% bidirectional coverage across all 18 SDLC artifacts. There are zero orphan requirements. Every feature implemented in Streamlit traces directly back through user stories, functional and non-functional requirements, to a core business operational question."
        },
        # Slide 5
        {
            "title": "Data Foundation, Quality & Cleansing",
            "category": "Step 8, 9 & 10: Relational Schema & Quality Engine",
            "bullets": [
                "Dataset Scale: 10,000 order transactions, 21,374 order item lines across 12 physical stores and 40 product SKUs (FY 2024).",
                "Schema Architecture: 3NF relational SQLite database comprising 5 core tables and 1 master analytical view (36 columns).",
                "Fix 1 (Negative Quantities): Replaced 107 zero/negative item counts with category median quantities (Invariant: Quantity >= 1).",
                "Fix 2 (Uncapped Discounts): Capped 62 rogue discount percentages exceeding 50% to prevent margin inversion (Invariant: Discount in [0, 50]).",
                "Fix 3 (Referential Integrity): Cleaned 31 orphaned customer IDs to ensure 100% foreign key enforcement (PRAGMA foreign_key_check = 0).",
                "Zero Leakage Tie-Out: Double-entry order line profit sums reconcile perfectly against parent header totals (Variance: ₹0.00)."
            ],
            "notes": "Our analytical foundation is an enterprise-scale relational database stored in SQLite. We synthetically generated and cleansed 10,000 transactions and over 21,000 item lines. Three primary automated cleaning interventions were executed: clamping negative order quantities, enforcing a 50% promotional discount ceiling, and eliminating orphan foreign keys. Our data quality engine confirms zero orphaned child records."
        },
        # Slide 6
        {
            "title": "System Architecture & Technical Flow",
            "category": "Step 7: 5-Tier Decoupled Architecture",
            "bullets": [
                "Storage Tier: SQLite relational database (data/app.db) with 3NF referential integrity and optimized multi-column indexes.",
                "Data Access Tier: src/data_loader.py leveraging @st.cache_data for sub-second query caching and instant dashboard responsiveness.",
                "Business Logic Tier: src/analytics.py and src/cleaning.py implementing business formulas BURL-01..06 and pure NumPy statistics.",
                "Presentation Tier: Multipage Streamlit application (app.py + 7 sub-pages in pages/) with corporate Tailwind CSS design tokens.",
                "Filter Control Tier: src/sidebar.py synchronizing global multi-select filters across all 8 screens via st.session_state.",
                "Production Performance: Average end-to-end analytical query latency is 0.042 seconds, well below the 3.0s SLA."
            ],
            "notes": "Our software follows a 5-tier modular architecture. We decoupled the presentation layer in Streamlit from our data access and business computation tiers. By utilizing Streamlit's caching decorators and SQLite indexed views, analytical query round-trips execute in just 42 milliseconds, well exceeding our 3.0-second non-functional performance requirement."
        },
        # Slide 7
        {
            "title": "Interactive Dashboard Capabilities",
            "category": "Step 12 & Step 15: Multipage Streamlit UI Library",
            "bullets": [
                "UI-01 Executive Overview (app.py): 4 Tailwind KPI metric cards, regional revenue bar chart, and MoM sales velocity trendline.",
                "UI-02 Multi-Store Benchmarking (pages/1_Store_Comparison.py): Store net revenue rankings and sales velocity (₹/sqft) vs footprint scatter plot.",
                "UI-03 Category Demographics (pages/2_Category_Demographics.py): 100% stacked bar chart by demographic tier and discount elasticity scatter.",
                "UI-04 Temporal Trends (pages/3_Temporal_Trends.py): 24-hour trading rush heatmap and weekday vs weekend Welch's t-test card.",
                "UI-05 Quota Variance & Leakage (pages/4_Target_Variance.py): Target variance gauge with +/- 5% alert triggers and return/cancellation pie chart.",
                "UI-06 Tender & Basket Dynamics (pages/5_Basket_Dynamics.py): UPI vs Card adoption donut chart and order amount IQR box plots.",
                "UI-07 Tabular Explorer & Quality Audit (pages/6 & 7): Full 36-column filtered CSV/Excel export and automated relational integrity scorecard."
            ],
            "notes": "The dashboard features 8 dedicated analytical screens built with Tailwind-styled iframes, custom CSS, and interactive Plotly figures. Key capabilities include store-level revenue and margin benchmarking, promotional discount elasticity modeling, hourly transaction rush analysis, and automated quota variance alerting."
        },
        # Slide 8
        {
            "title": "Key Empirical Findings & Business Insights",
            "category": "Step 15: Data-Driven Retail Discoveries",
            "bullets": [
                "Finding 1 (Real-Estate Efficiency): Delhi Flagship generates top total revenue (₹1.84 Cr), but Suburban Pune achieves +12% higher profit margin retention (45.2% vs 33.1%) due to lower occupancy overheads (₹38/sqft).",
                "Finding 2 (Promotional Erosion): Discounts beyond 25% trigger severe profit margin erosion (r = -0.584), confirming promotional markdowns stimulate unit velocity at the direct expense of commercial sustainability.",
                "Finding 3 (Weekend Surge Lift): Weekend trading windows demonstrate a statistically significant +18.4% ATV surge (Welch's t = 4.12, p < 0.001), warranting strategic weekend floor staffing reallocation.",
                "Executive Action: Freeze metro megastore expansions; redirect 40% of FY25 capital budget to suburban formats and enforce a hard 25% promotional discount ceiling."
            ],
            "notes": "Our analysis revealed three major commercial insights: First, prime metro stores produce high gross revenue but suffer from margin dilution, whereas suburban outlets deliver superior margin efficiency. Second, promotional markdowns beyond 25% actively destroy net profitability. Third, weekend trading exhibits a statistically significant 18.4% basket size expansion, verified by Welch's t-test with p-value under 0.001."
        },
        # Slide 9
        {
            "title": "Automated Testing & Verification Suite",
            "category": "Step 13: 100% Test Pass Rate Across 15 Suites",
            "table": [
                ["Test ID", "Target Layer", "Validation Assertion", "Result", "Execution Time"],
                ["TC-T-01..04", "Database & FKs", "5 Tables Exist, Row Counts Met, PRAGMA foreign_key_check = 0", "PASS (100%)", "0.045 s"],
                ["TC-T-05..06", "Views & Cleaning", "36 View Columns Match Schema, Invariants Held (Qty>=1, Disc<=50)", "PASS (100%)", "0.082 s"],
                ["TC-T-07..10", "Business Formulas", "Net Sales, Gross Margin %, ATV, Velocity/SqFt, +/-5% Alert Check", "PASS (100%)", "0.012 s"],
                ["TC-T-11..12", "Statistics Engine", "Non-Parametric IQR/Skewness, Welch's t-test Satterthwaite df", "PASS (100%)", "0.024 s"],
                ["TC-T-13..15", "App & Latency", "Streamlit Config Present, UI Library Clean, Analytical Query < 1.0s", "PASS (100%)", "0.729 s"]
            ],
            "bullets": [
                "Zero External Test Dependencies: Executed via Python's native unittest module (run_tests.py) and 100% pytest-compatible.",
                "Deterministic Test Vectors: All mathematical formulas validated against pre-computed synthetic truth matrices."
            ],
            "notes": "Software quality is verified by our automated test suite. 15 technical tests cover data layer integrity, database views, business math formulas, statistical routines, and performance benchmarks. The entire test suite executes in under 1.2 seconds with a 100% pass rate."
        },
        # Slide 10
        {
            "title": "Strategic Recommendations & Roadmap",
            "category": "Step 15: Commercial Value & Future Evolution",
            "bullets": [
                "Strategic Assortment Optimization: Shift Rural inventory allocations toward FMCG/Grocery staples and concentrate premium Electronics in Urban centers.",
                "Algorithmic Markdown Governance: Implement real-time discount guardrails preventing store managers from discounting items beyond 25%.",
                "Dynamic Staff Rostering: Reallocate 30% of weekday store staffing hours to peak weekend windows (16:00 - 20:00) to reduce checkout friction.",
                "Known Limitations: Historical scope is 1 fiscal year (FY24); demand forecasting does not incorporate weather or macroeconomic inflation indices.",
                "Phase 2 Roadmap: Integrate machine learning Prophet time-series models for store-level sales forecasting and automated replenishment triggers.",
                "Multi-Tenant RBAC: Implement role-based security filters for Store Managers, Regional Directors, and Executive Leadership."
            ],
            "notes": "Based on our findings, we recommend three actionable corporate strategies: demographic assortment re-balancing, hard algorithmic discount caps at 25%, and dynamic weekend staffing schedules. For future work, we plan to incorporate predictive time-series machine learning and role-based access control."
        },
        # Slide 11
        {
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

    for out_path in output_paths:
        os.makedirs(os.path.dirname(out_path), exist_ok=True)
        # Landscape A4: 841.89 x 595.28 points
        doc = SimpleDocTemplate(
            out_path,
            pagesize=landscape(A4),
            leftMargin=36,
            rightMargin=36,
            topMargin=32,
            bottomMargin=36
        )

        styles = getSampleStyleSheet()

        title_style = ParagraphStyle(
            "SlideTitle",
            fontName="Helvetica-Bold",
            fontSize=18,
            leading=22,
            textColor=colors.HexColor("#1A365D"),
            spaceAfter=3
        )
        cat_style = ParagraphStyle(
            "SlideCategory",
            fontName="Helvetica-Bold",
            fontSize=9,
            leading=12,
            textColor=colors.HexColor("#2B6CB0"),
            spaceAfter=10
        )
        bullet_style = ParagraphStyle(
            "SlideBullet",
            fontName="Helvetica",
            fontSize=9.5,
            leading=14,
            textColor=colors.HexColor("#2D3748"),
            spaceAfter=6
        )
        notes_header_style = ParagraphStyle(
            "NotesHeader",
            fontName="Helvetica-Bold",
            fontSize=8,
            leading=10,
            textColor=colors.HexColor("#2C5282")
        )
        notes_body_style = ParagraphStyle(
            "NotesBody",
            fontName="Helvetica-Oblique",
            fontSize=8,
            leading=11,
            textColor=colors.HexColor("#4A5568")
        )

        story = []

        for s_idx, slide in enumerate(slides, start=1):
            # Top Banner Header
            header_table = Table([[
                Paragraph(slide["title"], title_style),
                Paragraph(f"<b>SLIDE {s_idx} / {len(slides)}</b>", ParagraphStyle("Right", fontName="Helvetica-Bold", fontSize=10, alignment=2, textColor=colors.HexColor("#718096")))
            ]], colWidths=[620, 150])
            header_table.setStyle(TableStyle([
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('BOTTOMPADDING', (0,0), (-1,-1), 0),
            ]))
            story.append(header_table)
            story.append(Paragraph(slide["category"], cat_style))
            story.append(Spacer(1, 4))

            # Table or Bullets
            if "table" in slide:
                raw_table = slide["table"]
                formatted_table = []
                for r_idx, row in enumerate(raw_table):
                    row_cells = []
                    for c_idx, cell in enumerate(row):
                        f_style = ParagraphStyle(
                            f"Cell_{r_idx}_{c_idx}",
                            fontName="Helvetica-Bold" if r_idx == 0 else "Helvetica",
                            fontSize=8 if r_idx == 0 else 7.5,
                            leading=10,
                            textColor=colors.white if r_idx == 0 else colors.HexColor("#2D3748")
                        )
                        row_cells.append(Paragraph(cell, f_style))
                    formatted_table.append(row_cells)

                t_widths = [80, 70, 90, 240, 150, 140] if len(raw_table[0]) == 6 else [770 / len(raw_table[0])] * len(raw_table[0])
                t_obj = Table(formatted_table, colWidths=t_widths)
                t_obj.setStyle(TableStyle([
                    ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#1A365D")),
                    ('BACKGROUND', (0,1), (-1,-1), colors.HexColor("#F7FAFC")),
                    ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E0")),
                    ('PADDING', (0,0), (-1,-1), 4),
                    ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ]))
                story.append(t_obj)
                story.append(Spacer(1, 8))

            if "bullets" in slide:
                for b in slide["bullets"]:
                    story.append(Paragraph(f"&bull; {b}", bullet_style))
                story.append(Spacer(1, 6))

            # Speaker Notes Callout Box at bottom
            if "notes" in slide:
                notes_table = Table([[
                    Paragraph("<b>SPEAKER SCRIPT / VIVA NOTES (30s REHEARSAL):</b>", notes_header_style),
                ], [
                    Paragraph(f'"{slide["notes"]}"', notes_body_style)
                ]], colWidths=[770])
                notes_table.setStyle(TableStyle([
                    ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#EBF8FF")),
                    ('BOX', (0,0), (-1,-1), 0.75, colors.HexColor("#BEE3F8")),
                    ('PADDING', (0,0), (-1,-1), 4),
                ]))
                story.append(notes_table)

            if s_idx < len(slides):
                story.append(PageBreak())

        doc.build(story, canvasmaker=SlideNumberedCanvas)
        print(f"Presentation PDF successfully created at: {out_path}")

if __name__ == "__main__":
    build_presentation_pdf([
        "docs/2500320140025_Presentation.pdf",
        "docs/Presentation.pdf"
    ])
