import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

def create_ui_framework_doc(filename):
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
    s_run = sub_p.add_run("Step 12: User Interface (UI) Framework & Wireframe Design Document")
    s_run.font.size = Pt(12)
    s_run.font.color.rgb = COLOR_MUTED

    # Metadata Block
    meta_table = doc.add_table(rows=4, cols=2)
    meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta_data = [
        [("Learner Name:", "Anu Tyagi"), ("Roll Number:", "2500320140025")],
        [("Batch ID:", "ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01"), ("Institution:", "ABES Engineering College")],
        [("Project Mentor:", "Mr. Dinesh"), ("UI Architecture:", "Streamlit Multipage + Tailwind CDN iframes")],
        [("Screen Inventory:", "UI-01 to UI-08 (8 Dedicated Screens)"), ("Design Theme:", "Corporate Deep Navy (#1A365D)")]
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

    # Section 1: UI Screens Inventory Table
    add_custom_heading("1. UI Screens Inventory & Functional Scope", level=1)
    doc.add_paragraph(
        "The web application is structured into eight focused screens implementing Streamlit's native multipage architecture (`pages/`). "
        "Each screen addresses specific operational inquiries while sharing global session state filters:"
    )

    screens_table = doc.add_table(rows=9, cols=6)
    screens_headers = ["UI ID", "Screen Name & Route", "Core Operational Purpose", "Key Visual Components", "Active Filters", "Linked User Stories"]
    for i, h in enumerate(screens_headers):
        screens_table.rows[0].cells[i].paragraphs[0].text = h

    screens_data = [
        ["UI-01", "Executive Overview\n(app.py)", "High-level retail performance summary across all operating stores.", "4 Tailwind KPI cards, regional sales breakdown donut, top-5 store leaderboards.", "Date Range, Region, Store Tier", "US-03, US-04"],
        ["UI-02", "Store Comparison\n(01_Store_Comparison.py)", "Multi-store benchmarking, sales rankings & floor-space productivity.", "Dual-axis ranked bar/line chart, 4-quadrant space productivity scatter plot.", "Date Range, Region, Tier, Store Multi-select", "US-05, US-06"],
        ["UI-03", "Category Drilldown\n(02_Category_Drilldown.py)", "Category revenue contribution by demographic tier and discount elasticity.", "100% stacked bar chart by tier, discount elasticity scatter with OLS trendline.", "Date Range, Region, Category Dropdown", "US-07, US-08"],
        ["UI-04", "Temporal Trends\n(03_Temporal_Trends.py)", "Time-series trajectory, weekday-weekend split & peak trading hours.", "Multi-line MoM regional trend, grouped day-type bar, 24-hr trading heatmap.", "Date Range, Region, Metric Toggle", "US-09, US-10, US-11"],
        ["UI-05", "Target & Leakage\n(04_Target_Variance.py)", "Monthly budget quota attainment, +/- 5% alerts & revenue leakage audit.", "Target vs actual bullet charts, amber/red alert badges, return Pareto chart.", "Month Picker, Region, Store Selector", "US-12, US-13"],
        ["UI-06", "Statistical EDA\n(05_Statistical_EDA.py)", "Non-parametric transaction distributions, IQR, skewness & tender analysis.", "Tukey boxplots by tender, log-normal distribution histogram, ANOVA test table.", "Date Range, Store Selector, Tender Filter", "US-14, US-15"],
        ["UI-07", "Data Explorer\n(06_Data_Explorer.py)", "Raw transactional tabular query, search, filtering and data export.", "Paginated 36-column data table, column selector, single-click CSV export.", "Date Range, Full Global Filters, Search Box", "US-16"],
        ["UI-08", "System & About\n(07_System_About.py)", "Database schema audit, referential assertions, metadata & project details.", "Table row count badges, foreign key verification status, author/mentor info.", "None (System Diagnostics View)", "US-01, US-02"]
    ]
    for r_idx, row_vals in enumerate(screens_data, start=1):
        for c_idx, val in enumerate(row_vals):
            screens_table.rows[r_idx].cells[c_idx].paragraphs[0].text = val
    style_table(screens_table, [0.6, 1.4, 1.5, 1.6, 0.9, 0.8], [WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER])

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # Section 2: ASCII Wireframes
    add_custom_heading("2. High-Fidelity ASCII Wireframes per Screen", level=1)
    
    wireframes = [
        ("UI-01: Executive Overview Landing Page (app.py)",
"""+---------------------------------------------------------------------------------------------------+
|  [LOGO] STORE WISE SALES COMPARISON DASHBOARD — EXECUTIVE OVERVIEW            Analyst: Anu Tyagi  |
|  Unified Multi-Store Analytics Platform | Active Fiscal 2024                 Mentor: Mr. Dinesh   |
+------------------------------------+--------------------------------------------------------------+
| SIDEBAR FILTERS                    | [TAILWIND KPI CARDS BANNER]                                  |
| Date Range: [2024-01-01 to 12-31]  | +-----------------+ +---------------+ +------------+ +-----+ |
| Region:     [All Regions     v]    | | GROSS REVENUE   | | COMPLETED ORD | | AVG BASKET | | MAR%| |
| Store Tier: [All Tiers       v]    | | ₹3,245,600.00   | | 324 Orders    | | ₹10,017.28 | |46.8%| |
| Store Multi:[Select Stores   v]    | | +14.2% MoM  [↑] | | +8.5% YoY [↑] | | +2.1%  [↑] | |[OK] | |
|                                    | +-----------------+ +---------------+ +------------+ +-----+ |
| ---------------------------------- |--------------------------------------------------------------+
| NAVIGATION                         | [VISUAL ROW 1: REGIONAL MIX & LEADERBOARD]                   |
| > Executive Overview (Active)      | +-------------------------------+ +------------------------+ |
| - Store Comparison                 | | REGIONAL SALES SHARE (Donut)  | | TOP 5 STORES (Bar Chart)   | |
| - Category Drilldown               | | - North: 45.2% (₹1.46M)       | | 1. Metro Noida:   ₹480k    | |
| - Temporal Trends                  | | - South: 28.6% (₹928k)        | | 2. Connaught DLI: ₹440k    | |
| - Target & Leakage                 | | - West:  26.2% (₹850k)        | | 3. Bandra Mumbai: ₹410k    | |
| - Statistical EDA                  | +-------------------------------+ +------------------------+ |
| - Data Explorer & Download         |--------------------------------------------------------------+
| - System Quality & About           | [EXECUTIVE INSIGHT CALLOUT BOX]                              |
|                                    | [INSIGHT] Q4 Festival surge drove 37.7% uplift over baseline.|
+------------------------------------+--------------------------------------------------------------+"""),

        ("UI-02: Store Comparison & Space Productivity (01_Store_Comparison.py)",
"""+---------------------------------------------------------------------------------------------------+
|  STORE COMPARISON & SPACE PRODUCTIVITY BENCHMARKING (UI-02)                                       |
+------------------------------------+--------------------------------------------------------------+
| SIDEBAR FILTERS                    | [CHART 1: DUAL-AXIS RANKED STORE REVENUE & GROSS MARGIN %]  |
| Cohort Stores: [12 Stores Selected]| Bars: Total Net Sales (₹) | Line: Gross Profit Margin (%)    |
| Metric: [Revenue & Margin    v]    | 500k | [Bar: STR-001]                     --* (Margin: 48%) |
| Benchmark Line: [Show Regional Avg]| 400k | [Bar: STR-003]                 --* (Margin: 46%)     |
|                                    | 300k | [Bar: STR-006]             --* (Margin: 47%)         |
|                                    | 200k | [Bar: STR-007]         --* (Margin: 44%)             |
|                                    | 100k | [Bar: STR-012]     --* (Margin: 42%)                 |
|                                    |    0 +------------------------------------------------------ |
|                                    |--------------------------------------------------------------+
|                                    | [CHART 2: SPACE PRODUCTIVITY SCATTER MATRIX (OQ-02)]         |
|                                    | Y: Sales Velocity (₹ / Sq.Ft.)  vs  X: Store Size (Sq.Ft.)   |
|                                    | Vel  |                                                       |
|                                    | High |     [STR-001] (High Area, High Vel)                  |
|                                    |      |          [STR-005]                                   |
|                                    | Med  |  [STR-010]                                            |
|                                    | Low  |                 [STR-008] (Space Inefficient)         |
|                                    |    0 +------------------------------------------------ SqFt  |
+------------------------------------+--------------------------------------------------------------+"""),

        ("UI-03: Category Drilldown & Discount Elasticity (02_Category_Drilldown.py)",
"""+---------------------------------------------------------------------------------------------------+
|  CATEGORY MERCHANDISING & DISCOUNT ELASTICITY DRILLDOWN (UI-03)                                   |
+------------------------------------+--------------------------------------------------------------+
| FILTERS                            | [CHART 1: 100% STACKED CATEGORY CONTRIBUTION BY STORE TIER]  |
| Tier: [Urban, Suburban, Rural]     | Urban   | [Electronics: 35%] [Apparel: 30%] [Footwear] [Groc]|
| Category: [All Categories    v]    | Suburban| [Electronics: 22%] [Apparel: 25%] [Home Goods]     |
| Discount Range: [0% to 50%]        | Rural   | [Groceries: 45%]     [Home Goods: 30%] [Apparel]   |
|                                    |         0%---------------------------------------------100%  |
|                                    |--------------------------------------------------------------+
|                                    | [CHART 2: DISCOUNT ELASTICITY VS GROSS MARGIN % (OQ-07)]     |
|                                    | Margin%|                                                     |
|                                    |    60% | * (0% Disc)                                         |
|                                    |    50% |     * * (5-10% Disc)                                |
|                                    |    40% |         \ OLS Linear Regression Trendline           |
|                                    |    30% |             * * * (25% Disc - Margin Erosion)     |
|                                    |    20% |                     * (35% Deep Markdowns)          |
|                                    |      0 +-------------------------------------------- Disc %  |
|                                    | Pearson r: -0.584 | [WARNING] Markdowns > 20% destroy margin |
+------------------------------------+--------------------------------------------------------------+"""),

        ("UI-04: Temporal Trends & Rush Heatmaps (03_Temporal_Trends.py)",
"""+---------------------------------------------------------------------------------------------------+
|  TEMPORAL VELOCITY, WEEKEND SPLIT & 24-HOUR TRADING RUSH HEATMAP (UI-04)                          |
+------------------------------------+--------------------------------------------------------------+
| FILTERS                            | [CHART 1: MONTH-OVER-MONTH (MoM) REGIONAL SALES TRAJECTORY]  |
| Granularity: [Monthly        v]    | Sales|                                         /\ (Diwali)   |
| Trend Lines: [By Region      v]    | 500k |                             /\         /  \ (Q4 Peak) |
| Year:        [2024           v]    | 300k | ---------------------------/--\-------/----\--------- |
|                                    | 100k | Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec       |
|                                    |--------------------------------------------------------------+
|                                    | [CHART 2: WEEKDAY VS WEEKEND] | [CHART 3: 24-HR TRADING RUSH]|
|                                    | Weekend ATV: ₹10,540 (+26.1%) | Store x Hour Sales Heatmap   |
|                                    | Weekday ATV: ₹8,358           | STR-001 | [..Low..][HIGH][..]|
|                                    | Welch's t-test: p = 0.004 **  | STR-002 | [..Low..][HIGH][..]|
|                                    |                               | Hours:   06  09  12 15 18 21 |
+------------------------------------+--------------------------------------------------------------+"""),

        ("UI-05: Target Quota Variance & Leakage Audit (04_Target_Variance.py)",
"""+---------------------------------------------------------------------------------------------------+
|  TARGET QUOTA VARIANCE & REVENUE LEAKAGE AUDIT (UI-05)                                            |
+------------------------------------+--------------------------------------------------------------+
| FILTERS                            | [SECTION 1: TARGET ATTAINMENT & DYNAMIC THRESHOLD ALERTS]    |
| Month: [2024-11              v]    | Store       | Target (₹) | Actual (₹) | Variance% | Status    |
| Quota Threshold: [+/- 5.0%]        | STR-001 (N) | ₹1,120,000 | ₹1,210,000 | +8.0%     | [GREEN OK]|
| Leakage Filter: [Returns & Cancels]| STR-004 (N) | ₹1,050,000 | ₹1,085,000 | +3.3%     | [GREEN OK]|
|                                    | STR-008 (W) |   ₹780,000 |   ₹725,000 | -7.1%     | [ALERT ⚠️]|
|                                    | STR-011 (W) |   ₹390,000 |   ₹345,000 | -11.5%    | [CRITICAL]|
|                                    |--------------------------------------------------------------+
|                                    | [SECTION 2: REVENUE LEAKAGE AUDIT & RETURN PARETO CHART]     |
|                                    | Gross Sales: ₹3,410,000 - Returns: ₹164,400 = Net Recognized |
|                                    | Pareto Bar: Electronics (42% of returns) > Apparel (31%)     |
+------------------------------------+--------------------------------------------------------------+"""),

        ("UI-06: Statistical EDA & Tender Boxplots (05_Statistical_EDA.py)",
"""+---------------------------------------------------------------------------------------------------+
|  STATISTICAL DISTRIBUTION & PAYMENT TENDER ANALYSIS (UI-06)                                       |
+------------------------------------+--------------------------------------------------------------+
| FILTERS                            | [CHART 1: TENDER SHARE]      | [CHART 2: BASKET VALUE BOX]  |
| Tender Type: [All Methods    v]    | - UPI:  48.2% (₹1.56M)       | Value|    [Tukey Whiskers]   |
| Outlier Trim: [None (Show All) v]  | - Card: 28.5% (₹925k)        | 25k  |      o o (Outliers)   |
| Distribution: [Log-Normal Fit v]   | - Cash: 14.8% (₹480k)        | 15k  |    +---+  +---+       |
|                                    | - Debit: 8.5% (₹275k)        | 10k  |    |   |  |   | (IQR) |
|                                    |                              |      |    +---+  +---+       |
|                                    |                              |    0 +----------------------- |
|                                    |                              |        UPI   Card Cash  Debit |
|                                    |--------------------------------------------------------------+
|                                    | [PARAMETRIC & NON-PARAMETRIC DISTRIBUTION SUMMARY TABLE]     |
|                                    | Mean: ₹9,220 | Median: ₹7,850 | IQR: ₹5,400 | Skewness: +1.82|
+------------------------------------+--------------------------------------------------------------+"""),

        ("UI-07: Data Explorer & Download Utility (06_Data_Explorer.py)",
"""+---------------------------------------------------------------------------------------------------+
|  DATA EXPLORER & CSV EXPORT UTILITY (UI-07)                                                       |
+------------------------------------+--------------------------------------------------------------+
| FILTERS & SEARCH                   | [ACTIVE DATASET: view_store_sales_analytics (853 Rows)]       |
| Search: [Headphones________]       | [BUTTON: Download Filtered CSV]  [BUTTON: Download Excel]    |
| Select Columns: [36 Columns  v]    |--------------------------------------------------------------+
| Page Size:      [25 Rows     v]    | Item ID | Order ID | Date | Store | Product | Qty | Net (₹)|
| Sort By:        [Date Desc   v]    | 001     | ORD-0001 |11-23 | STR04 | Noise-Can| 1   | 5,499  |
|                                    | 002     | ORD-0001 |11-23 | STR04 | Earbuds  | 2   | 4,748  |
|                                    | 003     | ORD-0002 |10-29 | STR02 | Sandals  | 1   |   799  |
|                                    | Showing 1 - 25 of 853 matching records | Page: [< 1 2 3 >]   |
+------------------------------------+--------------------------------------------------------------+"""),

        ("UI-08: System Diagnostics, Data Quality & About (07_System_About.py)",
"""+---------------------------------------------------------------------------------------------------+
|  SYSTEM DIAGNOSTICS, DATA QUALITY AUDIT & ABOUT (UI-08)                                           |
+------------------------------------+--------------------------------------------------------------+
| DIAGNOSTIC TABS                    | [SECTION 1: DATABASE INTEGRITY VERIFICATION (PRAGMA)]        |
| > Schema Integrity (Active)        | dim_stores: 12 rows | dim_products: 60 rows | Orders: 352    |
| - Referential Assertions           | Foreign Key Violations: 0 (PASS) | Null Primary Keys: 0 (PASS)|
| - Academic Project Metadata        | B-Tree Index Status: 4 Active Indexes (Optimal)              |
|                                    |--------------------------------------------------------------+
|                                    | [SECTION 2: ACADEMIC PROJECT METADATA & SIGN-OFF]            |
|                                    | Project: Store Wise Sales Comparison Dashboard               |
|                                    | Student: Anu Tyagi (Roll: 2500320140025, MCA 2nd Year, ABES) |
|                                    | Mentor:  Mr. Dinesh | Tech: Python, Streamlit, SQL |
+------------------------------------+--------------------------------------------------------------+""")
    ]

    for title, wireframe_text in wireframes:
        add_custom_heading(title, level=2, space_before=10, space_after=2)
        box = doc.add_table(rows=1, cols=1)
        box.alignment = WD_TABLE_ALIGNMENT.CENTER
        c = box.rows[0].cells[0]
        set_cell_background(c, 'EDF2F7')
        set_cell_margins(c, top=60, bottom=60, left=100, right=100)
        c.width = Inches(6.5)
        cp = c.paragraphs[0]
        cp.paragraph_format.space_after = Pt(0)
        cr = cp.add_run(wireframe_text)
        cr.font.name = 'Consolas'
        cr.font.size = Pt(7.0)
        cr.font.color.rgb = COLOR_TEXT

    # Section 3: Global Layout & Visual Design System
    add_custom_heading("3. Global Layout & Visual Design System", level=1)
    doc.add_paragraph(
        "• Color Palette: Primary Corporate Navy (`#1A365D`), Slate Blue (`#2B6CB0`), Charcoal Text (`#2D3748`), "
        "Emerald Success (`#38A169`), Amber Warning (`#DD6B20`), and Crimson Alert (`#E53E3E`).\n"
        "• Typography Hierarchy: Clean sans-serif system fonts (`Inter`, `Segoe UI`, `system-ui`) for UI chrome, paired with `Consolas` "
        "for numeric currency indicators, SKU codes, and data tables.\n"
        "• State Persistence Architecture: Global filters are persisted across multipage routes using `st.session_state` keys "
        "(`st.session_state.filter_date_range`, `st.session_state.filter_region`, `st.session_state.filter_tier`). When a user updates "
        "a filter on any screen, all other pages automatically inherit the modified scope.\n"
        "• Responsive & Phone View Layout: On wide viewports (>= 1280px), metrics render in 4 columns and charts in dual columns. "
        "On tablet/mobile viewports (< 768px), components automatically collapse into a single-column stacked layout with full horizontal touch scrolling."
    )

    # Section 4: Reusable UI Component Specifications
    add_custom_heading("4. Reusable UI Component Specifications (Tailwind CSS)", level=1)
    doc.add_paragraph(
        "UI components are implemented in `src/ui.py` using `st.components.v1.html` and Tailwind CDN to isolate styling from Streamlit DOM:\n"
        "1. Executive Header Component: Multi-tone gradient card (`bg-gradient-to-r from-slate-900 via-blue-900 to-indigo-900`) with rounded borders, title, metadata badges, and mentor/author labels.\n"
        "2. KPI Metric Card Component: Elevated white container (`bg-white rounded-xl p-4 shadow-sm border border-slate-200`) displaying uppercase metric label, bold rupee amount, and percentage delta indicator pill (`text-emerald-700 bg-emerald-50 border-emerald-200`).\n"
        "3. Diagnostic Insight Box: Contextual colored alert container (`rounded-xl p-3.5 border shadow-sm`) with uppercase status badges (`INFO`, `WARNING`, `ALERT`, `INSIGHT`)."
    )

    # Section 5: Best Chart for Each OQ
    add_custom_heading("5. Optimal Visualization Selection Matrix (OQ-01 to OQ-10)", level=1)
    doc.add_paragraph(
        "The visualization type for each operational question was selected based on strict graphical perception and data density principles:"
    )

    chart_table = doc.add_table(rows=11, cols=4)
    chart_headers = ["OQ ID", "Operational Question", "Best Chart Selection", "Analytical Justification & Why"]
    for i, h in enumerate(chart_headers):
        chart_table.rows[0].cells[i].paragraphs[0].text = h

    chart_data = [
        ["OQ-01", "Store Gross Sales Revenue & Profit Margin", "Dual-Axis Ranked Bar & Line Chart", "Ranked horizontal bars order sales volume cleanly, while the overlay margin line immediately reveals high-volume stores suffering from margin dilution."],
        ["OQ-02", "Space Productivity (Velocity/SqFt) & ATV", "4-Quadrant Scatter Plot", "Plots Sales Velocity against Store Size (Sq.Ft.) with crosshair benchmarks, categorizing stores into 4 quadrants (Space-Efficient Leaders vs. Space Inefficient Outliers)."],
        ["OQ-03", "Category Revenue Contribution by Tier", "100% Stacked Horizontal Bar Chart", "Normalizes category sales to 100% per demographic tier, allowing direct comparison of category share across Urban, Suburban, and Rural formats irrespective of total revenue scale."],
        ["OQ-04", "Month-over-Month (MoM) Regional Trends", "Multi-Line Time-Series Trajectory", "Line slopes effectively convey directional trajectory and seasonal growth rates across North, South, and West clusters without bar clutter."],
        ["OQ-05", "Weekday vs. Weekend Sales Split", "Grouped Comparative Bar + Boxplot", "Contrasts aggregated totals side-by-side while boxplot whiskers show ticket distribution variance, supported by Welch's t-test hypothesis verification."],
        ["OQ-06", "Return Rates & Net Revenue Leakage", "Waterfall Chart & Pareto Defect Bar", "Waterfall chart visualizes financial leakage from Gross to Net (Gross - Returns - Cancels = Net), while Pareto bar isolates the vital few product categories driving 80% of returns."],
        ["OQ-07", "Discount Elasticity & Margin Preservation", "Bivariate Scatter Plot with OLS Line", "Scatter plot displays individual order-item discounts against realized profit margins, with the linear regression trendline proving the rate of margin destruction."],
        ["OQ-08", "Target Quota Attainment & Variance", "Bullet Graph & Diverging Variance Bar", "Bullet graphs display actual performance against budgeted targets and qualitative ranges; diverging bars highlight positive/negative percentage deviations."],
        ["OQ-09", "Payment Tender Share & Basket Distribution", "Donut Breakdown + Tukey Boxplots", "Donut chart displays tender adoption shares (UPI, Card, Cash), while adjacent Tukey boxplots display median basket size and dispersion by payment type."],
        ["OQ-10", "24-Hour Trading Rush Patterns", "Store-by-Hour Intensity Heatmap", "A two-dimensional matrix (Stores on Y-axis, Hours 00-23 on X-axis) uses color saturation to instantly highlight trading peak rush windows (17:00 to 21:00)."]
    ]
    for r_idx, row_vals in enumerate(chart_data, start=1):
        for c_idx, val in enumerate(row_vals):
            chart_table.rows[r_idx].cells[c_idx].paragraphs[0].text = val
    style_table(chart_table, [0.7, 1.8, 1.8, 2.2], [WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT])

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # Section 6: Configuration & Component Files
    add_custom_heading("6. Configuration & Starter Component Files", level=1)
    doc.add_paragraph(
        "• Production Theme File: `.streamlit/config.toml` configuring primary `#1A365D`, background `#F8FAFC`, secondary `#EDF2F7`, text `#2D3748`, and disabling telemetry.\n"
        "• Reusable UI Library: `src/ui.py` providing `render_tailwind_header()`, `render_kpi_cards()`, and `render_insight_box()` using isolated HTML components."
    )

    doc.save(filename)
    print(f"UI Framework Document successfully created at: {filename}")

if __name__ == "__main__":
    import sys
    target = sys.argv[1] if len(sys.argv) > 1 else "docs/12_UI_Framework.docx"
    create_ui_framework_doc(target)
