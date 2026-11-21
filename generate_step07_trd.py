import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

def create_trd_doc(filename):
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
    style_normal.font.size = Pt(10.5)
    style_normal.font.color.rgb = COLOR_TEXT
    style_normal.paragraph_format.line_spacing = 1.15
    style_normal.paragraph_format.space_after = Pt(5)

    def set_cell_background(cell, hex_color):
        shading_xml = f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>'
        cell._tc.get_or_add_tcPr().append(parse_xml(shading_xml))

    def set_cell_margins(cell, top=100, bottom=100, left=140, right=140):
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
        # Header formatting
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
        
        # Body rows formatting
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
    s_run = sub_p.add_run("Step 7: Technical Requirements Document (TRD) & Solution Blueprint")
    s_run.font.size = Pt(12)
    s_run.font.color.rgb = COLOR_MUTED

    # Document Metadata Block
    meta_table = doc.add_table(rows=4, cols=2)
    meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta_data = [
        [("Learner Name:", "Anu Tyagi"), ("Roll Number:", "2500320140025")],
        [("Batch ID:", "ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01"), ("Institution:", "ABES Engineering College")],
        [("Project Mentor:", "Mr. Dinesh"), ("Document Type:", "Technical Requirements & System Architecture")],
        [("Fixed Tech Stack:", "Python 3.12, Streamlit, Plotly, SQLite"), ("Deployment Target:", "Streamlit Community Cloud")]
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

    # Section 1: End-to-End System Architecture (Mermaid Description)
    add_custom_heading("1. End-to-End System Architecture", level=1)
    doc.add_paragraph(
        "The application follows a modular 5-tier analytical data pipeline: Data Ingestion -> Caching & Loading -> "
        "Data Validation & Cleaning -> Business Analytics & Statistics -> Interactive Multipage Presentation. "
        "The architecture is designed specifically to operate within the 1 GB memory boundary of Streamlit Community Cloud."
    )
    
    # Architecture Pipeline Box
    arch_box = doc.add_table(rows=1, cols=1)
    arch_box.alignment = WD_TABLE_ALIGNMENT.CENTER
    c = arch_box.rows[0].cells[0]
    set_cell_background(c, 'F0FFF4')
    set_cell_margins(c, top=100, bottom=100, left=150, right=150)
    c.width = Inches(6.5)
    cp = c.paragraphs[0]
    cp.paragraph_format.space_after = Pt(0)
    r_arch_title = cp.add_run("SYSTEM ARCHITECTURE PIPELINE FLOW (Mermaid Structure):\n")
    r_arch_title.bold = True
    r_arch_title.font.size = Pt(9.5)
    r_arch_title.font.color.rgb = RGBColor(34, 84, 61)
    r_arch = cp.add_run(
        "flowchart TD\n"
        "  subgraph Data_Storage [Persistent Storage Tier]\n"
        "    DB[(SQLite: data/app.db)]\n"
        "  end\n"
        "  subgraph Ingestion_Caching [Data Access & Caching Layer]\n"
        "    Loader[src/data_loader.py: load_data]\n"
        "    Cache{{@st.cache_data LRU Cache}}\n"
        "    DB --> Loader --> Cache\n"
        "  end\n"
        "  subgraph Data_Cleansing [Validation & Cleansing Engine]\n"
        "    Validator[src/validator.py: validate_integrity]\n"
        "    Cache --> Validator\n"
        "  end\n"
        "  subgraph Analytics_Engine [Business Intelligence & Statistical Layer]\n"
        "    Analytics[src/analytics.py: KPI & Aggregation Engine]\n"
        "    Stats[src/stats.py: Statistical & Hypothesis Engine]\n"
        "    Validator --> Analytics\n"
        "    Validator --> Stats\n"
        "  end\n"
        "  subgraph Presentation_Tier [Streamlit Multipage UI Layer]\n"
        "    App[app.py: Executive Overview]\n"
        "    P1[pages/01_Store_Comparison.py]\n"
        "    P2[pages/02_Category_Drilldown.py]\n"
        "    P3[pages/03_Temporal_Trends.py]\n"
        "    P4[pages/04_Target_Variance.py]\n"
        "    P5[pages/05_Statistical_EDA.py]\n"
        "    Tailwind[src/components.py: Tailwind CDN iframe Cards]\n"
        "    Analytics --> App & P1 & P2 & P3 & P4\n"
        "    Stats --> P5\n"
        "    Tailwind --> App & P1 & P2 & P3 & P4 & P5\n"
        "  end"
    )
    r_arch.font.name = 'Consolas'
    r_arch.font.size = Pt(8.0)
    r_arch.font.color.rgb = RGBColor(45, 55, 72)

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # Section 2: Repository Folder Structure
    add_custom_heading("2. Repository Directory Structure", level=1)
    doc.add_paragraph(
        "The project repository is structured strictly around clean code principles, separating source logic, multipage presentation, "
        "relational data storage, documentation, and automated testing suites:"
    )

    tree_p = doc.add_paragraph()
    tree_p.paragraph_format.left_indent = Inches(0.2)
    tree_run = tree_p.add_run(
        "hcl/\n"
        "├── .streamlit/\n"
        "│   └── config.toml                # Custom dark navy theme & server layout settings\n"
        "├── data/\n"
        "│   ├── app.db                     # Normalized SQLite relational database file\n"
        "│   └── schema.sql                 # DDL schema, table constraints & B-tree indexes\n"
        "├── docs/                          # Academic & professional deliverables (Steps 01-18)\n"
        "│   ├── 01_Project_Understanding.docx\n"
        "│   ├── 02_HMW_Statements.docx\n"
        "│   ├── 03_Needs_Analysis.docx\n"
        "│   ├── 04_Traceability_Matrix.xlsx\n"
        "│   ├── 05_BRD.docx\n"
        "│   └── 07_TRD.docx\n"
        "├── pages/\n"
        "│   ├── 01_Store_Comparison.py      # Multi-store ranking, ATV & space productivity\n"
        "│   ├── 02_Category_Drilldown.py    # Demographic store-tier category contribution\n"
        "│   ├── 03_Temporal_Trends.py       # MoM, weekday-weekend split, hourly rush heatmaps\n"
        "│   ├── 04_Target_Variance.py       # Monthly budget vs actual quota variance alerts\n"
        "│   └── 05_Statistical_EDA.py       # Boxplots, skewness, ANOVA, elasticity regression\n"
        "├── src/\n"
        "│   ├── __init__.py\n"
        "│   ├── analytics.py               # Aggregations, ranking, BURL formula implementations\n"
        "│   ├── components.py              # Tailwind CSS HTML/iframe KPI cards & headers\n"
        "│   ├── data_loader.py             # SQLite query execution & @st.cache_data wrappers\n"
        "│   ├── stats.py                   # SciPy/Statsmodels statistical test implementations\n"
        "│   └── validator.py               # Schema integrity, referential checks & assertions\n"
        "├── tests/\n"
        "│   ├── __init__.py\n"
        "│   ├── test_analytics.py          # Unit tests for KPI calculations and rankings\n"
        "│   ├── test_app.py                # Streamlit AppTest headless UI interaction tests\n"
        "│   ├── test_database.py           # SQLite constraints, foreign keys & non-null checks\n"
        "│   └── test_stats.py              # Unit tests for statistical formulas and tests\n"
        "├── app.py                         # Main landing page: Executive Retail KPI Overview\n"
        "├── requirements.txt               # Pinned production dependency manifest\n"
        "└── README.md                      # Project setup, execution and cloud deployment guide"
    )
    tree_run.font.name = 'Consolas'
    tree_run.font.size = Pt(8.5)
    tree_run.font.color.rgb = COLOR_TEXT

    # Section 3: Functional Requirements (FR-01 to FR-14)
    add_custom_heading("3. Functional Requirements & Module Mapping", level=1)
    doc.add_paragraph(
        "Functional requirements (FR-01 through FR-14) bridge upstream business requirements (BR) into concrete software capabilities. "
        "Each requirement is mapped to its dedicated Python source module and presentation view:"
    )

    fr_table = doc.add_table(rows=15, cols=5)
    fr_headers = ["FR ID", "Functional Requirement Description", "Linked BR", "Priority", "Target Python Module & Page"]
    for i, h in enumerate(fr_headers):
        fr_table.rows[0].cells[i].paragraphs[0].text = h

    fr_data = [
        ["FR-01", "Compute and render descending store rankings by gross revenue and gross profit margin with interactive region filters.", "BR-01", "Must", "src/analytics.py | pages/01_Store_Comparison.py"],
        ["FR-02", "Calculate Sales Velocity per Sq.Ft. and Average Transaction Value (ATV) across store cohorts using interactive scatter plots.", "BR-02", "Must", "src/analytics.py | pages/01_Store_Comparison.py"],
        ["FR-03", "Generate 100% stacked category contribution charts segmented across Urban, Suburban, and Rural store demographic tiers.", "BR-03", "Must", "src/analytics.py | pages/02_Category_Drilldown.py"],
        ["FR-04", "Calculate and plot month-over-month (MoM) sales velocity trends and percentage deltas across regional store clusters.", "BR-04", "Must", "src/analytics.py | pages/03_Temporal_Trends.py"],
        ["FR-05", "Aggregate transaction volumes into 24 distinct hourly bins (00-23) and render an interactive Store vs. Hour sales density heatmap.", "BR-05", "Must", "src/analytics.py | pages/03_Temporal_Trends.py"],
        ["FR-06", "Partition order volume and basket amounts into Weekday (Mon-Fri) and Weekend (Sat-Sun) cohorts with comparative metric cards.", "BR-06", "Should", "src/analytics.py | pages/03_Temporal_Trends.py"],
        ["FR-07", "Audit order cancellations and returns to reconcile Net Sales: Gross Sales - (Returns + Cancellations) with Pareto leakage visualization.", "BR-07", "Must", "src/analytics.py | pages/04_Target_Variance.py"],
        ["FR-08", "Compute Pearson and Spearman correlation between discount percentages, unit sales volume, and gross profit margins with OLS regression.", "BR-08", "Must", "src/stats.py | pages/02_Category_Drilldown.py"],
        ["FR-09", "Calculate monthly target sales variance (Actual - Target) and dynamically trigger amber/red visual alerts for variances exceeding +/- 5%.", "BR-09", "Must", "src/analytics.py | pages/04_Target_Variance.py"],
        ["FR-10", "Analyze transaction frequency and basket size distributions across tender types (UPI, Credit Card, Debit Card, Cash) using boxplots.", "BR-10", "Should", "src/analytics.py | pages/05_Statistical_EDA.py"],
        ["FR-11", "Compute parametric and non-parametric statistical metrics (Mean, Median, StdDev, IQR, Skewness, Outlier counts) for store transaction tickets.", "BR-11", "Should", "src/stats.py | pages/05_Statistical_EDA.py"],
        ["FR-12", "Execute automated schema validation checking primary key uniqueness, foreign key integrity, non-negative monetary amounts, and date validity.", "BR-12", "Must", "src/validator.py | tests/test_database.py"],
        ["FR-13", "Provide dynamic CSV/Excel export utility downloading currently filtered tabular aggregated store summaries.", "BR-13", "Could", "src/components.py | pages/01_Store_Comparison.py"],
        ["FR-14", "Real-time foreign currency conversion and exchange rate API integration.", "BR-14", "Won't", "Formally Deferred to Phase 2 (Domestic INR Scope)"]
    ]
    for r_idx, row_vals in enumerate(fr_data, start=1):
        for c_idx, val in enumerate(row_vals):
            fr_table.rows[r_idx].cells[c_idx].paragraphs[0].text = val
    style_table(fr_table, [0.7, 2.7, 0.7, 0.7, 1.7], [WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT])

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # Section 4: Data Layer Specification
    add_custom_heading("4. Data Access Layer & Caching Architecture", level=1)
    doc.add_paragraph(
        "To satisfy the < 2.0s latency requirement on free cloud tiers, the data access layer utilizes SQLite with B-tree indexes "
        "and aggressive in-memory caching via `@st.cache_data`:"
    )
    doc.add_paragraph(
        "• SQLite Read-Only Pragmas: Database connections are opened using URI mode `file:data/app.db?mode=ro` with `PRAGMA query_only = ON;` "
        "and `PRAGMA cache_size = -16000;` (16 MB page cache), preventing concurrency locks and accidental database mutations.\n"
        "• Two-Tier Caching Strategy: Raw table queries are cached via `@st.cache_data(ttl=3600, show_spinner=False)`. User-filtered subset "
        "aggregations are computed in-memory using vectorized Pandas operations, guaranteeing sub-50ms UI re-renders upon filter changes.\n"
        "• Data Cleansing & Sanity Assertions: In `src/validator.py`, automated assertions ensure: (1) `quantity > 0`, (2) `unit_price >= cost_price`, "
        "(3) `0 <= discount_percentage <= 50.0`, (4) valid ISO-8601 timestamps, and (5) zero orphaned transaction foreign keys."
    )

    # Section 5: Statistical Methods for Each OQ
    add_custom_heading("5. Statistical Methodology & Formulation (OQ-01 to OQ-10)", level=1)
    doc.add_paragraph(
        "Every Operational Question is resolved through deterministic statistical formulas and hypothesis-testing techniques:"
    )

    stats_table = doc.add_table(rows=11, cols=4)
    stats_headers = ["OQ ID", "Core Analytical Objective", "Statistical Method & Formula", "Python Implementation Library"]
    for i, h in enumerate(stats_headers):
        stats_table.rows[0].cells[i].paragraphs[0].text = h

    stats_data = [
        ["OQ-01", "Store Gross Sales & Profit Margin Ranking", "Ranked aggregation & Pareto CDF: Margin% = ((Selling - Cost) / Selling) * 100", "pandas.DataFrame.rank(), numpy"],
        ["OQ-02", "Space Productivity & ATV Outlier Detection", "ATV = Sum(Sales)/N_orders; Velocity = Sum(Sales)/SqFt; Modified Z-Score: Mi = 0.6745*(x - median)/MAD", "scipy.stats.median_abs_deviation"],
        ["OQ-03", "Category Share by Demographic Store Tier", "Contingency Table & Chi-Square Independence Test: Chi2 = Sum((O - E)^2 / E); p < 0.05", "scipy.stats.chi2_contingency"],
        ["OQ-04", "Month-over-Month (MoM) Regional Trajectory", "First-Order Differencing: g_t = ((Y_t - Y_{t-1}) / Y_{t-1}) * 100 with 3-Month Moving Average (SMA-3)", "pandas.Series.pct_change(), rolling()"],
        ["OQ-05", "Weekday vs. Weekend Spending Variance", "Two-Sample Independent Welch's t-test (unequal variances): t = (X1 - X2) / sqrt(s1^2/n1 + s2^2/n2)", "scipy.stats.ttest_ind(equal_var=False)"],
        ["OQ-06", "Return Rate & Revenue Leakage Audit", "Return Ratio: Sum(Returns)/Sum(Gross); Net Sales = Gross - Returns - Cancellations; Pareto 80/20", "pandas, numpy"],
        ["OQ-07", "Discount Elasticity & Margin Preservation", "Pearson (r) & Spearman rank correlation; OLS Linear Regression: Margin = Beta_0 + Beta_1 * Discount", "statsmodels.api.OLS, scipy.stats.pearsonr"],
        ["OQ-08", "Monthly Target Quota Attainment & Variance", "Attainment% = (Actual/Target) * 100; Variance Z-score across store network to isolate chronic outliers", "scipy.stats.zscore, numpy"],
        ["OQ-09", "Payment Tender vs. Basket Size Distribution", "One-Way ANOVA (F-test) or Kruskal-Wallis H-test; IQR = Q3 - Q1; Whiskers: Q1-1.5*IQR, Q3+1.5*IQR", "scipy.stats.f_oneway, scipy.stats.kruskal"],
        ["OQ-10", "Temporal 24-Hour Trading Rush Patterns", "Temporal Kernel Density Estimation (KDE) and Hourly Binned Matrix: H_{store, hour}; Peak-to-Trough Ratio", "scipy.stats.gaussian_kde, plotly.heatmap"]
    ]
    for r_idx, row_vals in enumerate(stats_data, start=1):
        for c_idx, val in enumerate(row_vals):
            stats_table.rows[r_idx].cells[c_idx].paragraphs[0].text = val
    style_table(stats_table, [0.7, 2.0, 2.5, 1.3], [WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT])

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # Section 6: UI Framework & Tailwind CSS Integration
    add_custom_heading("6. User Interface Architecture & Styling Specifications", level=1)
    doc.add_paragraph(
        "To deliver an executive, modern presentation superior to standard Streamlit layouts, UI presentation is enhanced through "
        "Tailwind CSS integration in isolated `components.v1.html` iframes:\n"
        "• Color Palette: Primary Deep Navy (`#1A365D`), Slate Blue (`#2B6CB0`), Charcoal Text (`#2D3748`), Emerald Accent (`#38A169`), Amber Alert (`#DD6B20`), Crimson Danger (`#E53E3E`).\n"
        "• Custom Tailwind KPI Component (`src/components.py`): Header banners and metric cards are rendered via HTML iframes importing Tailwind CDN "
        "(`https://cdn.tailwindcss.com`), delivering clean shadow cards (`shadow-md rounded-xl p-5 border border-slate-200 bg-white`), typography hierarchy, and delta badges.\n"
        "• Responsive Multipage Sidebar Navigation: Seamless page navigation via Streamlit native `pages/` directory structure with uniform top-bar global date and region filters.\n"
        "• Plotly Visual Styling: Unified chart template (`plotly_white`) utilizing project corporate colors, custom hover templates (`₹%{y:,.2f}`), and disabled zoom clutter for frictionless readability."
    )

    # Section 7: Performance, Error Handling & Security
    add_custom_heading("7. Performance, Error Handling, and Security Architecture", level=1)
    doc.add_paragraph(
        "• Performance Benchmarks: System initial cold-load < 2.0 seconds; cached filter recalculations < 0.3 seconds; total RAM footprint < 350 MB.\n"
        "• Graceful Degradation & Error Handling: All analytical functions are wrapped in deterministic try-except blocks. Database access errors or empty filter states surface contextual Streamlit notifications (`st.warning('No transactions match the selected filters.')`) without crashing the application.\n"
        "• Security Architecture: Parameterized SQL queries only (zero string concatenation / zero SQL injection risk); SQLite database file configured as read-only; zero persistent user input fields preventing Cross-Site Scripting (XSS)."
    )

    # Section 8: Pinned requirements.txt
    add_custom_heading("8. Pinned Dependency Specification (requirements.txt)", level=1)
    doc.add_paragraph(
        "To guarantee 100% build reproducibility across local developer machines and Streamlit Community Cloud, exact package versions are pinned:"
    )

    req_p = doc.add_paragraph()
    req_p.paragraph_format.left_indent = Inches(0.2)
    req_run = req_p.add_run(
        "streamlit==1.32.2\n"
        "pandas==2.2.1\n"
        "numpy==1.26.4\n"
        "plotly==5.20.0\n"
        "scipy==1.12.0\n"
        "statsmodels==0.14.1\n"
        "openpyxl==3.1.5\n"
        "python-docx==1.2.0\n"
        "pytest==8.1.1\n"
        "pytest-cov==4.1.0"
    )
    req_run.font.name = 'Consolas'
    req_run.font.size = Pt(9.0)
    req_run.font.color.rgb = COLOR_PRIMARY

    # Section 9: Deployment Requirements
    add_custom_heading("9. Cloud Deployment & CI/CD Specification", level=1)
    doc.add_paragraph(
        "• Host Platform: Streamlit Community Cloud connected directly to the project GitHub repository.\n"
        "• Entry Point: `app.py` at the root directory.\n"
        "• Configuration File: `.streamlit/config.toml` configuring server headless mode, custom theme colors, and disabling telemetry.\n"
        "• Automated CI/CD Workflow: GitHub Actions pipeline configured to run `pytest --cov=src` on every push to the `main` branch, blocking deployment if test coverage drops below 85%."
    )

    # Section 10: Technical Risks & Mitigations
    add_custom_heading("10. Technical Risk Assessment & Mitigation", level=1)
    
    trsk_table = doc.add_table(rows=6, cols=5)
    trsk_headers = ["Risk ID", "Technical Risk Description", "Likelihood", "Impact", "Architectural Mitigation Strategy"]
    for i, h in enumerate(trsk_headers):
        trsk_table.rows[0].cells[i].paragraphs[0].text = h

    trsk_data = [
        ["TRSK-01", "Streamlit session state resets causing loss of active filter context.", "Medium", "Medium", "Store active filter parameters in st.query_params to preserve state across page transitions."],
        ["TRSK-02", "Iframe component resizing glitches on mobile viewports.", "Medium", "Low", "Set explicit minimum desktop viewport requirements (>=1280px) and fixed iframe component heights (120px to 160px)."],
        ["TRSK-03", "Memory leak from repeated Plotly figure generation during continuous user interactions.", "Low", "High", "Explicitly release figure handles and utilize Streamlit native garbage collection between run turns."],
        ["TRSK-04", "Database file locking during parallel read queries on Streamlit Cloud.", "Low", "Medium", "Enforce WAL mode (PRAGMA journal_mode=WAL;) and read-only connection strings in SQLite."],
        ["TRSK-05", "Scientific package compilation incompatibilities on Linux cloud containers.", "Low", "High", "Pin standard wheel-supported package releases in requirements.txt (NumPy 1.26.x, Pandas 2.2.x)."]
    ]
    for r_idx, row_vals in enumerate(trsk_data, start=1):
        for c_idx, val in enumerate(row_vals):
            trsk_table.rows[r_idx].cells[c_idx].paragraphs[0].text = val
    style_table(trsk_table, [0.8, 2.0, 0.8, 0.8, 2.1], [WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT])

    doc.save(filename)
    print(f"TRD Document successfully created at: {filename}")

if __name__ == "__main__":
    import sys
    target = sys.argv[1] if len(sys.argv) > 1 else "docs/07_TRD.docx"
    create_trd_doc(target)
