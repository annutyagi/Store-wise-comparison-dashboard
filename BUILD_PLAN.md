# BUILD PLAN: Store Wise Sales Comparison Dashboard
**Version:** 1.0  
**Date:** November 20, 2026  
**Author:** Anu Tyagi (Roll No: 2500320140025, MCA 2nd Year, ABES Engineering College)  
**Academic Mentor:** Mr. Dinesh  
**Project ID:** P052  
**Batch ID:** ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01  
**Domain:** Data Analytics  
**Target Platform:** Streamlit Multipage + Plotly + Tailwind CDN + SQLite (`data/app.db`)

---

## 1. Project Summary & Master Deliverable Links

The Store Wise Sales Comparison Dashboard is an enterprise-grade business intelligence and retail data analytics platform designed to solve operational visibility gaps across a multi-store retail chain. Operating on a relational SQLite database (`data/app.db`) with 5 core tables and a 36-column analytical view (`view_store_sales_analytics`), the system tracks gross revenue, net profit margins, sales velocity per square foot, and average transaction value (ATV) across 10 retail stores in 3 demographic tiers (Urban, Suburban, Rural). The platform addresses 10 Operational Questions (`OQ-01` to `OQ-10`), implements 14 Business Requirements (`BR-01` to `BR-14`, with 64.3% Must priority), delivers 16 User Stories (`US-01` to `US-16`, 44 story points), and visualizes insights across 8 responsive Tailwind-styled Streamlit screens (`UI-01` to `UI-08`). The analytical engine incorporates Welch's t-test for weekday vs. weekend sales lift, non-parametric distribution profiling, quota variance alerting ($\pm 5\%$), and discount elasticity regression. The system is verified through 16 Functional Test Cases (`TC-F-01` to `TC-F-16`) and 15 Automated Technical Test Cases (`TC-T-01` to `TC-T-15`) achieving 100% pass rate in 0.136 seconds.

### Master Document Links
1. **Project Understanding:** [docs/01_Project_Understanding.docx](file:///c:/Users/annut/OneDrive/Desktop/hcl/docs/01_Project_Understanding.docx)
2. **HMW Statements:** [docs/02_HMW_Statements.docx](file:///c:/Users/annut/OneDrive/Desktop/hcl/docs/02_HMW_Statements.docx)
3. **Needs Analysis:** [docs/03_Needs_Analysis.docx](file:///c:/Users/annut/OneDrive/Desktop/hcl/docs/03_Needs_Analysis.docx)
4. **Traceability Matrix:** [docs/04_Traceability_Matrix.xlsx](file:///c:/Users/annut/OneDrive/Desktop/hcl/docs/04_Traceability_Matrix.xlsx)
5. **Business Requirements Document (BRD):** [docs/05_BRD.docx](file:///c:/Users/annut/OneDrive/Desktop/hcl/docs/05_BRD.docx)
6. **Technical Requirements Document (TRD):** [docs/07_TRD.docx](file:///c:/Users/annut/OneDrive/Desktop/hcl/docs/07_TRD.docx)
7. **Database Design & Data Dictionary:** [docs/08_Database_Design.xlsx](file:///c:/Users/annut/OneDrive/Desktop/hcl/docs/08_Database_Design.xlsx)
8. **Entity-Relationship Diagram:** [docs/09_ER_Diagram.docx](file:///c:/Users/annut/OneDrive/Desktop/hcl/docs/09_ER_Diagram.docx) | Graphic: [docs/09_ER_Diagram.png](file:///c:/Users/annut/OneDrive/Desktop/hcl/docs/09_ER_Diagram.png)
9. **Dataset Generation Report:** [docs/10_Dataset_Generation.docx](file:///c:/Users/annut/OneDrive/Desktop/hcl/docs/10_Dataset_Generation.docx)
10. **User Stories & Sprint Backlog:** [docs/11_User_Stories.docx](file:///c:/Users/annut/OneDrive/Desktop/hcl/docs/11_User_Stories.docx)
11. **UI Framework & Tailwind Design:** [docs/12_UI_Framework.docx](file:///c:/Users/annut/OneDrive/Desktop/hcl/docs/12_UI_Framework.docx)
12. **Test Cases Specification:** [docs/13_Test_Cases.docx](file:///c:/Users/annut/OneDrive/Desktop/hcl/docs/13_Test_Cases.docx)
13. **Build Plan Specification:** [docs/14_BUILD_PLAN.docx](file:///c:/Users/annut/OneDrive/Desktop/hcl/docs/14_BUILD_PLAN.docx)

---

## 2. Final Project Folder Structure

```text
hcl/
├── .streamlit/
│   └── config.toml                  # Streamlit corporate navy theme configuration (#1A365D)
├── data/
│   ├── app.db                       # Relational SQLite database with DDL, foreign keys, and analytical view
│   ├── store_wise_sales_data.xlsx   # 5-sheet consolidated Excel dataset (seed 42)
│   ├── dim_stores.csv               # 10 retail store master records
│   ├── dim_products.csv             # 50 product catalog master records
│   ├── fact_orders.csv              # 200 customer order transactional records
│   ├── fact_order_items.csv         # 620 order line item records
│   └── dim_sales_targets.csv        # 120 store monthly quota records
├── docs/                            # Capstone documentation deliverables (DOCX, XLSX, PNG)
│   ├── 01_Project_Understanding.docx
│   ├── 02_HMW_Statements.docx
│   ├── 03_Needs_Analysis.docx
│   ├── 04_Traceability_Matrix.xlsx
│   ├── 05_BRD.docx
│   ├── 07_TRD.docx
│   ├── 08_Database_Design.xlsx
│   ├── 09_ER_Diagram.docx
│   ├── 09_ER_Diagram.png
│   ├── 10_Dataset_Generation.docx
│   ├── 11_User_Stories.docx
│   ├── 12_UI_Framework.docx
│   ├── 13_Test_Cases.docx
│   └── 14_BUILD_PLAN.docx
├── pages/                           # Streamlit multipage application modules
│   ├── 1_Store_Comparison.py        # UI-02: Multi-Store Benchmarking & Real-Estate Productivity
│   ├── 2_Category_Demographics.py   # UI-03: Demographic Tier & Category Drill-Down
│   ├── 3_Temporal_Trends.py         # UI-04: MoM Velocity, Hourly Trading Windows, Weekend Surge
│   ├── 4_Target_Variance.py         # UI-05: Budgeted Quota Variance & Net Revenue Leakage
│   ├── 5_Basket_Dynamics.py         # UI-06: Payment Tender & Non-Parametric Distribution Stats
│   ├── 6_Data_Export.py             # UI-07: Filtered CSV / Excel Tabular Data Export
│   └── 7_Data_Quality.py            # UI-08: Relational Hygiene & Anomaly Imputation Audit
├── src/                             # Core Python analytical and UI modules
│   ├── __init__.py
│   ├── analytics.py                 # Business formulas BURL-01..06, Welch's t-test, skewness, IQR
│   ├── cleaning.py                  # Domain invariant checks and anomaly imputation
│   ├── data_loader.py               # Cached SQLite loading via st.cache_data
│   ├── generate_data.py             # Synthetic retail data generator (seed 42)
│   ├── sidebar.py                   # Reusable sidebar filters synchronized with st.session_state
│   └── ui.py                        # Tailwind CDN header, responsive KPI cards, insight box
├── tests/                           # Automated verification suite (unittest & pytest compatible)
│   ├── __init__.py
│   ├── test_analysis.py             # TC-T-07 to TC-T-12: Formulas, Welch's t-test, statistics
│   ├── test_app.py                  # TC-T-13 to TC-T-15: Theming, UI components, query latency
│   └── test_data.py                 # TC-T-01 to TC-T-06: Database schema, row counts, PK/FK checks
├── app.py                           # UI-01: Main entrypoint and Executive Overview Dashboard
├── BUILD_PLAN.md                    # Step-by-step master build plan for execution
├── CLAUDE.md                        # Project development rules and governance guidelines
├── requirements.txt                 # Pinned dependencies for deployment
├── run_tests.py                     # Single-command automated test execution runner
└── README.md                        # Production project overview and deployment guide
```

---

## 3. Step-by-Step Build Tasks (`BUILD-01` to `BUILD-12`)

The implementation in Step 15 strictly executes the following 12 tasks in linear dependency order.

### `BUILD-01`: Environment Setup & Dependency Pinning
- **Goal:** Establish production Python configuration, pin runtime dependencies in `requirements.txt`, and configure corporate theme in `.streamlit/config.toml`.
- **Files Affected:** `requirements.txt`, `.streamlit/config.toml`
- **Linked Epics / Stories / Features:** Epic 1 (`US-01`), `F-06`, `FR-12`
- **Exact AI Prompt:**
  ```text
  Create requirements.txt pinned with:
  streamlit>=1.35.0, pandas>=2.2.0, numpy>=1.26.0, plotly>=5.22.0, openpyxl>=3.1.2, python-docx>=1.1.0
  Ensure .streamlit/config.toml defines primaryColor="#1A365D", backgroundColor="#FFFFFF", secondaryBackgroundColor="#F7FAFC", textColor="#2D3748".
  ```
- **How to Check:** Verify file presence and inspect color hex values.
- **Tests That Must Pass:** `TC-T-13`

---

### `BUILD-02`: Relational Data Loader, Caching & Invariant Cleaning Engine
- **Goal:** Build robust SQLite database loading with `st.cache_data`, automatic foreign key verification, missing value imputation, and domain invariant enforcement.
- **Files Affected:** `src/data_loader.py`, `src/cleaning.py`
- **Linked Epics / Stories / Features:** Epic 1 (`US-01`, `US-02`), `F-06`, `FR-12`, `BR-12`
- **Exact AI Prompt:**
  ```text
  Write src/data_loader.py to connect to data/app.db with st.cache_data, loading view_store_sales_analytics, dim_stores, dim_products, fact_orders, fact_order_items, and dim_sales_targets.
  Write src/cleaning.py to enforce domain invariants: quantity >= 1, discount between 0 and 50, selling price >= cost price, and impute missing values (mean for sales, mode for category).
  ```
- **How to Check:** Call `load_analytical_data()` in Python, assert shape has 36 columns and > 0 rows.
- **Tests That Must Pass:** `TC-T-01`, `TC-T-02`, `TC-T-03`, `TC-T-04`, `TC-T-05`, `TC-T-06`

---

### `BUILD-03`: Business Metrics & Statistical Analysis Core Engine
- **Goal:** Implement the mathematical core functions for Net Sales (`BURL-01`), Gross Margin % (`BURL-02`), ATV (`BURL-03`), Sales Velocity per SqFt (`BURL-04`), Target Quota Variance % (`BURL-05`), Welch's t-test for weekend sales lift, and non-parametric distribution metrics (IQR, Skewness, Median).
- **Files Affected:** `src/analytics.py`
- **Linked Epics / Stories / Features:** Epic 3 (`US-05`, `US-06`, `US-08`), Epic 4 (`US-11`, `US-13`, `US-15`), `F-01`, `F-02`, `F-03`, `F-04`, `F-05`, `BR-01..11`
- **Exact AI Prompt:**
  ```text
  Write src/analytics.py containing:
  1. calculate_net_sales(gross, returns, cancels) -> round(gross - returns - cancels, 2)
  2. calculate_gross_margin_pct(selling, cost) -> round(((selling - cost)/selling)*100, 2)
  3. calculate_atv(net_sales, orders_count) -> round(net_sales / orders_count, 2)
  4. calculate_sales_velocity_sqft(net_sales, sqft) -> round(net_sales / sqft, 4)
  5. check_target_variance_alert(actual, target, threshold=0.05) -> (var_pct, bool_alert)
  6. calculate_distribution_stats(series) -> {mean, median, iqr, skewness}
  7. welchs_ttest_stat(s1, s2) -> (t_stat, df)
  Ensure functions handle zero-division gracefully without throwing unhandled exceptions.
  ```
- **How to Check:** Run `& "C:\Users\annut\AppData\Local\Python\bin\python.exe" -m unittest tests/test_analysis.py`.
- **Tests That Must Pass:** `TC-T-07`, `TC-T-08`, `TC-T-09`, `TC-T-10`, `TC-T-11`, `TC-T-12`

---

### `BUILD-04`: Reusable Tailwind UI Component Library & Shared Sidebar
- **Goal:** Develop reusable executive UI components (gradient header banner, responsive multi-column KPI cards, styled insight boxes) and global sidebar filter controller synchronized via `st.session_state`.
- **Files Affected:** `src/ui.py`, `src/sidebar.py`
- **Linked Epics / Stories / Features:** Epic 2 (`US-03`, `US-04`), `F-01`, `F-03`, `BR-01`, `BR-04`
- **Exact AI Prompt:**
  ```text
  Enhance src/ui.py with render_tailwind_header(), render_kpi_cards(), and render_insight_box() utilizing Tailwind CSS CDN via st.components.v1.html with safe fallback when Streamlit is not installed.
  Create src/sidebar.py to render global sidebar filters: Store multiselect, Date Range slider, Demographic Tier checkboxes, and Product Category select, binding selections into st.session_state.
  ```
- **How to Check:** Verify `src/ui.py` imports cleanly and `src/sidebar.py` initializes filter defaults.
- **Tests That Must Pass:** `TC-T-14`, `TC-F-01`, `TC-F-02`

---

### `BUILD-05`: Main Entrypoint & Executive Overview Dashboard (`app.py`, `UI-01`)
- **Goal:** Build the main application landing page displaying the corporate header, global filters, 4 high-level KPI cards (Total Net Sales, Average Gross Margin %, Overall ATV, Total Completed Orders), Regional Sales bar chart, and MoM sales velocity trendline.
- **Files Affected:** `app.py`
- **Linked Epics / Stories / Features:** Epic 2 (`US-03`, `US-04`), Epic 3 (`US-05`), `F-01` (`FR-01`), `F-03` (`FR-04`), `BR-01`, `BR-04`, `OQ-01`, `OQ-04`
- **Exact AI Prompt:**
  ```text
  Construct app.py as the primary dashboard:
  1. Configure page layout='wide' and title='Store Wise Sales Comparison Dashboard'.
  2. Integrate src/sidebar.py for shared session state filtering.
  3. Render Tailwind header with author 'Anu Tyagi' and mentor 'Mr. Dinesh'.
  4. Display 4 executive KPI cards (Net Sales, Margin %, ATV, Total Orders).
  5. Render Plotly bar chart: Regional Sales Comparison (North, South, West, East).
  6. Render Plotly line chart: Month-over-Month Velocity Trendline (highlighting Q4 spike).
  7. Handle empty filtered datasets gracefully with a styled warning banner.
  ```
- **How to Check:** Verify layout elements, KPI calculations, and chart responsiveness.
- **Tests That Must Pass:** `TC-F-01`, `TC-F-02`, `TC-F-03`, `TC-F-07`, `TC-F-16`

---

### `BUILD-06`: Multi-Store Benchmarking & Real-Estate Comparison Page (`pages/1_Store_Comparison.py`, `UI-02`)
- **Goal:** Develop screen `UI-02` benchmarking all 10 stores: gross revenue vs net profit margin rankings, real-estate productivity (Sales Velocity per SqFt), and store expansion ROI insights.
- **Files Affected:** `pages/1_Store_Comparison.py`
- **Linked Epics / Stories / Features:** Epic 3 (`US-05`, `US-06`), `F-01` (`FR-01`, `FR-02`), `BR-01`, `BR-02`, `OQ-01`, `OQ-02`
- **Exact AI Prompt:**
  ```text
  Build pages/1_Store_Comparison.py:
  1. Header: 'Multi-Store Benchmarking & Real-Estate Productivity'.
  2. Plotly Horizontal Bar Chart: Store Rankings by Gross Sales with Profit Margin color-coding. Highlight margin-diluting stores (< 30%).
  3. Plotly Scatter Plot: Store Size (SqFt) vs Sales Velocity per SqFt with bubble size representing Total Net Sales.
  4. Benchmarking summary table displaying Store Name, City, Tier, Size SqFt, Net Sales, Margin %, and Velocity/SqFt.
  5. Insight callout identifying Delhi Flagship as velocity leader and highlighting suburban real-estate efficiency.
  ```
- **How to Check:** Check store ranking order, assert margin dilution visual flags, and inspect velocity calculations.
- **Tests That Must Pass:** `TC-F-03`, `TC-F-04`, `TC-T-08`, `TC-T-09`

---

### `BUILD-07`: Demographic Tier & Category Drill-Down Page (`pages/2_Category_Demographics.py`, `UI-03`)
- **Goal:** Develop screen `UI-03` exploring product category contribution across Urban, Suburban, and Rural tiers, plus discount elasticity vs gross profit margin erosion.
- **Files Affected:** `pages/2_Category_Demographics.py`
- **Linked Epics / Stories / Features:** Epic 3 (`US-07`, `US-08`), `F-02` (`FR-03`, `FR-08`), `BR-03`, `BR-08`, `OQ-03`, `OQ-07`
- **Exact AI Prompt:**
  ```text
  Build pages/2_Category_Demographics.py:
  1. Header: 'Product Category Demographics & Discount Elasticity'.
  2. Plotly Stacked Bar Chart: Category Sales Revenue contribution across Urban, Suburban, and Rural tiers.
  3. Plotly Scatter Plot: Discount % (0-50%) vs Line Margin % with linear regression trendline (demonstrating margin erosion, r = -0.584).
  4. KPI cards: Electronics Share in Urban %, Staples Share in Rural %, Margin Dilution Threshold (25% discount).
  5. Merchandising insight box recommending inventory allocation rules by demographic tier.
  ```
- **How to Check:** Verify category breakdown matches empirical demographic patterns and regression line is negative.
- **Tests That Must Pass:** `TC-F-05`, `TC-F-06`, `TC-T-08`

---

### `BUILD-08`: Temporal Trends & Peak Trading Windows Page (`pages/3_Temporal_Trends.py`, `UI-04`)
- **Goal:** Develop screen `UI-04` tracking Month-over-Month sales velocity, Q4 festive spike (+37.7%), hourly transaction volume trading windows, and weekday vs weekend sales lift (+26.1% ATV).
- **Files Affected:** `pages/3_Temporal_Trends.py`
- **Linked Epics / Stories / Features:** Epic 3 (`US-09`), Epic 4 (`US-10`, `US-11`), `F-03` (`FR-04`, `FR-05`, `FR-06`), `BR-04`, `BR-05`, `BR-06`, `OQ-04`, `OQ-05`, `OQ-10`
- **Exact AI Prompt:**
  ```text
  Build pages/3_Temporal_Trends.py:
  1. Header: 'Temporal Velocity, Peak Trading Windows & Seasonality'.
  2. Plotly Line Chart: 12-Month Sales Velocity trendline annotating the October-November Diwali festival surge (+37.7%).
  3. Plotly Bar / Heatmap: Hourly Customer Transaction Volumes (10:00 to 22:00) identifying lunch (13:00-15:00) and evening peak trading rush (18:00-21:00).
  4. Plotly Grouped Bar Chart: Weekday vs Weekend Sales & ATV Comparison.
  5. Statistical Callout: Welch's t-test hypothesis result (t = -11.42, p < 0.001) confirming weekend basket lift is statistically significant.
  ```
- **How to Check:** Verify Q4 festival annotation, peak hourly volume peaks, and Welch's t-test statistical readout.
- **Tests That Must Pass:** `TC-F-07`, `TC-F-08`, `TC-F-09`, `TC-T-12`

---

### `BUILD-09`: Target Variance & Net Revenue Leakage Page (`pages/4_Target_Variance.py`, `UI-05`)
- **Goal:** Develop screen `UI-05` calculating actual vs budgeted monthly sales quotas, $\pm 5\%$ threshold alerting, and return/cancellation leakage rate monitoring.
- **Files Affected:** `pages/4_Target_Variance.py`
- **Linked Epics / Stories / Features:** Epic 4 (`US-12`, `US-13`), `F-04` (`FR-07`, `FR-09`), `BR-07`, `BR-09`, `OQ-06`, `OQ-08`
- **Exact AI Prompt:**
  ```text
  Build pages/4_Target_Variance.py:
  1. Header: 'Sales Quota Variance & Revenue Leakage Monitor'.
  2. Plotly Gauge / Waterfall Chart: Monthly Actual vs Target Quota achievement.
  3. Status Banner: Dynamic alert flagging underperforming stores (< -5% variance) in red and overperforming stores (> +5%) in green.
  4. Plotly Bar Chart: Return Amount INR and Order Cancellation Rate by Store and Product Category.
  5. Operational insight box detailing inventory leakage prevention and regional intervention protocols.
  ```
- **How to Check:** Verify alert triggering on test thresholds and correct calculation of net leakage percentages.
- **Tests That Must Pass:** `TC-F-10`, `TC-F-11`, `TC-T-07`, `TC-T-10`

---

### `BUILD-10`: Payment Tender & Basket Dynamics Page (`pages/5_Basket_Dynamics.py`, `UI-06`)
- **Goal:** Develop screen `UI-06` analyzing payment method adoption (UPI, Credit Card, Cash), transaction value distribution (mean, median, IQR, skewness), and outlier box plots.
- **Files Affected:** `pages/5_Basket_Dynamics.py`
- **Linked Epics / Stories / Features:** Epic 4 (`US-14`, `US-15`), `F-05` (`FR-10`, `FR-11`), `BR-10`, `BR-11`, `OQ-09`
- **Exact AI Prompt:**
  ```text
  Build pages/5_Basket_Dynamics.py:
  1. Header: 'Payment Tender Dynamics & Basket Size Distribution'.
  2. Plotly Donut Chart: Payment Method Transaction Share (UPI, Credit Card, Debit Card, Cash).
  3. Plotly Bar Chart: Average Transaction Value (ATV) by Payment Method (demonstrating Credit Card ATV premium).
  4. Plotly Box Plot: Order Amount distribution with Parametric (Mean, Std Dev) and Non-Parametric (Median, IQR, Skewness) readout cards.
  5. Commercial insight box outlining payment gateway fee optimization and EMI promotional strategy.
  ```
- **How to Check:** Verify UPI volume dominance, Credit Card ticket premium, and positive skewness metrics.
- **Tests That Must Pass:** `TC-F-12`, `TC-F-13`, `TC-T-11`

---

### `BUILD-11`: Tabular Data Export & Data Hygiene Audit Pages (`pages/6_Data_Export.py`, `pages/7_Data_Quality.py`, `UI-07`, `UI-08`)
- **Goal:** Develop screen `UI-07` for filtered CSV/Excel downloads and screen `UI-08` displaying database relational integrity, imputation logs, and error-handling verification.
- **Files Affected:** `pages/6_Data_Export.py`, `pages/7_Data_Quality.py`
- **Linked Epics / Stories / Features:** Epic 1 (`US-01`, `US-02`), Epic 5 (`US-16`), `F-01` (`FR-13`), `F-06` (`FR-12`), `BR-12`, `BR-13`, `OQ-01..10`
- **Exact AI Prompt:**
  ```text
  Build pages/6_Data_Export.py:
  1. Header: 'Tabular Data Explorer & Export Engine'.
  2. Filtered interactive st.dataframe showing analytical records.
  3. st.download_button for CSV and Excel format with date-stamped filenames.
  Build pages/7_Data_Quality.py:
  1. Header: 'Relational Data Governance & Quality Audit'.
  2. Integrity summary cards: 0 Orphan Foreign Keys, 0 Duplicate PKs, 100% Invariant Compliance.
  3. Missing value imputation audit table and anomaly treatment log.
  4. System reconciliation summary verifying fact_orders to fact_order_items tie-out.
  ```
- **How to Check:** Verify file download payloads, inspect data quality metrics, and test zero-result filter handling.
- **Tests That Must Pass:** `TC-F-14`, `TC-F-15`, `TC-F-16`, `TC-T-01..06`

---

### `BUILD-12`: Comprehensive Verification, Production README & Deployment Package
- **Goal:** Execute the full automated test suite (`run_tests.py`), generate comprehensive `README.md` with system architecture diagrams, user guide, and viva notes, and prepare GitHub/Streamlit Cloud deployment package.
- **Files Affected:** `README.md`, `run_tests.py`
- **Linked Epics / Stories / Features:** All Epics 1-5, All US-01..16, All FR-01..14
- **Exact AI Prompt:**
  ```text
  Execute run_tests.py to verify 100% test pass rate across all 15 technical tests.
  Write a comprehensive, publication-grade README.md containing:
  1. Project Title, Student & Mentor details, ABES Engineering College MCA Project.
  2. Architecture overview with Mermaid diagram.
  3. Key Findings & Empirical Retail Patterns (Festival Spike, Weekend Lift, Discount Elasticity).
  4. Quickstart commands (install, run app, run tests).
  5. Multipage navigation guide with screenshots / descriptions.
  6. Requirements Traceability Summary (100% coverage).
  7. 10 Comprehensive Viva Questions & Answers for final academic evaluation.
  ```
- **How to Check:** Run `& "C:\Users\annut\AppData\Local\Python\bin\python.exe" run_tests.py` (assert exit code 0) and inspect `README.md`.
- **Tests That Must Pass:** `TC-T-01` to `TC-T-15`, `TC-F-01` to `TC-F-16`

---

## 4. Definition of Done (DoD)

The project reaches completion only when all the following criteria are objectively verified:
1. **Requirements Coverage:** All 9 "Must Have" and 3 "Should Have" Business Requirements (`BR-01` to `BR-13`) are fully implemented and verifiable in the UI.
2. **Automated Testing:** All 15 Technical Unit & Integration Tests (`TC-T-01` to `TC-T-15`) execute and pass with 100% success rate under `run_tests.py`.
3. **Functional UI Completeness:** All 8 screens (`UI-01` to `UI-08`) render without uncaught Python exceptions, supporting dynamic sidebar filtering and graceful handling of empty queries (`TC-F-16`).
4. **Data Integrity:** SQLite database (`data/app.db`) passes `PRAGMA foreign_key_check` with zero orphan records and zero duplicate primary keys.
5. **Performance SLA:** Query execution latency against `view_store_sales_analytics` remains strictly under 1.0 second (well within the 3.0s SLA).
6. **Documentation Completeness:** All documentation deliverables (`01` through `14`) are saved in `docs/` in Word (`.docx`) and Excel (`.xlsx`) formats.
7. **Production Documentation:** Comprehensive `README.md` is present in the root directory with complete installation, execution, architecture, and viva notes.

---

## 5. Terminal Execution Commands

### A. Environment Installation
```powershell
pip install -r requirements.txt
```

### B. Launch Application Locally
```powershell
streamlit run app.py
```

### C. Execute Automated Test Suite
```powershell
# Run via unified standalone runner:
python run_tests.py

# Or run via Python unittest module:
python -m unittest discover -s tests -p "test_*.py"

# Or run via pytest (if installed in target environment):
pytest tests/ -v
```

---

## 6. Comprehensive Academic Viva Notes (10 Likely Questions & Answers)

#### Q1: What is the business problem your Store Wise Sales Comparison Dashboard solves?
**Answer:** Retail chains operating multi-unit stores often suffer from fragmented, delayed spreadsheet reporting. Leadership lacks visibility into store-level profitability variations, retail real-estate productivity, margin erosion caused by excessive discounting, demographic inventory misalignments, and net revenue leakage from returns. This platform provides an automated, centralized BI dashboard offering real-time multi-store benchmarking, quota variance alerting, and prescriptive merchandising insights.

#### Q2: Explain the database design and why a star/snowflake schema was chosen over a single flat file.
**Answer:** The database (`data/app.db`) uses a relational star/snowflake schema comprising 3 dimension tables (`dim_stores`, `dim_products`, `dim_sales_targets`) and 2 fact tables (`fact_orders` and `fact_order_items`). Normalizing into dimensions and facts prevents data anomalies, enforces referential integrity through foreign keys, eliminates redundant string storage, and allows granular order-item level margin calculations while maintaining order-level payment dynamics. A 36-column analytical view (`view_store_sales_analytics`) joins these entities for performant, zero-friction dashboard querying.

#### Q3: Why is Sales Velocity per Square Foot a critical retail KPI, and how is it calculated?
**Answer:** Gross sales revenue alone can be deceptive; a 12,000 sqft flagship store will naturally gross more than a 2,000 sqft boutique. Sales Velocity per SqFt measures retail real-estate capital efficiency:
$$\text{Sales Velocity per SqFt} = \frac{\text{Total Net Sales (INR)}}{\text{Store Area (SqFt)}}$$
It normalizes commercial performance across store footprints, revealing whether suburban or boutique locations actually yield higher return on physical space investment than massive metropolitan flagships.

#### Q4: Why did you use Welch's t-test instead of standard Student's t-test for weekday vs. weekend sales?
**Answer:** Retail footfall and spending patterns on weekends exhibit substantially higher variance than weekday shopping. Student's two-sample t-test assumes homoscedasticity (equal variance between groups). When sample variances are unequal, Student's t-test suffers inflated Type I error rates. Welch's t-test calculates separate sample variances and uses the Welch-Satterthwaite equation for adjusted degrees of freedom, guaranteeing valid statistical conclusions ($t < -10.0, p < 0.001$).

#### Q5: How does your application protect gross profit margins against discount elasticity?
**Answer:** In screen `UI-03` and `src/analytics.py`, the system plots discount percentage against line profit margin percentage. Empirical analysis reveals a negative linear correlation ($r = -0.584$). While discounts below 20% stimulate unit volume without harming profitability, discounts exceeding 25% cause severe gross margin erosion. The dashboard provides visual warning thresholds to prevent merchandisers from implementing destructive margin-diluting promotions.

#### Q6: How does the dashboard detect and handle seasonal demand shifts?
**Answer:** In screen `UI-04`, the temporal velocity engine tracks Month-over-Month (MoM) sales velocity across the 12-month calendar. The analysis captures the Indian retail festive surge (Q4 Diwali festival spike in October-November), which generates a +37.7% sales lift over baseline. This enables supply chain and store managers to initiate pre-festival inventory staging and temporary staffing weeks in advance.

#### Q7: What mechanism enforces data quality and referential integrity in your architecture?
**Answer:** Data governance is enforced at three levels:
1. **Relational Constraints:** SQLite `CREATE TABLE` DDL defines `PRIMARY KEY`, `FOREIGN KEY ... REFERENCES`, `NOT NULL`, `UNIQUE`, and `CHECK` constraints with `PRAGMA foreign_keys = ON;`.
2. **Automated Tests:** `tests/test_data.py` executes `PRAGMA foreign_key_check;` and uniqueness GROUP BY queries (`TC-T-03`, `TC-T-04`).
3. **Application Invariant Engine:** `src/cleaning.py` verifies business invariants (quantities $\ge 1$, discounts $\le 50\%$, selling price $\ge$ cost price) and logs imputation events.

#### Q8: How is Streamlit caching utilized to optimize dashboard performance?
**Answer:** Streamlit reruns the Python script on every user interaction. Without caching, querying SQLite and executing table joins on every slider movement would cause unacceptable latency. We apply `@st.cache_data` to data-loading functions in `src/data_loader.py`. The joined analytical view is cached in memory, reducing query benchmark latency to just 0.038 seconds, well below our 1.0-second performance SLA (`TC-T-15`).

#### Q9: What is the purpose of the Requirements Traceability Matrix (RTM), and how was it maintained?
**Answer:** The RTM (`docs/04_Traceability_Matrix.xlsx`) provides end-to-end forward and backward traceability across the software development lifecycle. It maps each Operational Question (`OQ`) to a How Might We statement (`HMW`), Business Requirement (`BR`), Feature (`F`), Database column, User Story (`US`), UI screen (`UI`), and Test Case (`TC-F` / `TC-T`). This guarantees that no requirement is orphaned, no gold-plating occurs, and 100% of functional requirements have passing verification tests.

#### Q10: How does the application handle edge cases such as empty filter selections or missing data?
**Answer:** Negative and edge-case handling is enforced by `TC-F-16`. If a user selects mutually exclusive filter criteria (e.g., Rural Tier combined with a Luxury category having zero sales), the dashboard intercepts the empty dataframe and renders a graceful, user-friendly notification banner (*"No records match selected filters"*). It suppresses Plotly rendering crashes, preventing unhandled exceptions or broken visual layouts.
