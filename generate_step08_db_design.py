import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def create_database_design_excel(filename):
    os.makedirs(os.path.dirname(filename), exist_ok=True)
    wb = openpyxl.Workbook()
    
    FONT_FAMILY = "Calibri"
    font_title = Font(name=FONT_FAMILY, size=15, bold=True, color="1A365D")
    font_sub = Font(name=FONT_FAMILY, size=10, italic=True, color="718096")
    font_header = Font(name=FONT_FAMILY, size=10, bold=True, color="FFFFFF")
    font_data = Font(name=FONT_FAMILY, size=9.5, color="2D3748")
    font_data_bold = Font(name=FONT_FAMILY, size=9.5, bold=True, color="1A365D")
    font_code = Font(name="Consolas", size=9.0, color="1A365D")
    font_code_ddl = Font(name="Consolas", size=9.0, color="2D3748")

    fill_header = PatternFill(start_color="1A365D", end_color="1A365D", fill_type="solid")
    fill_sub_banner = PatternFill(start_color="EBF8FF", end_color="EBF8FF", fill_type="solid")
    fill_zebra_even = PatternFill(start_color="F7FAFC", end_color="F7FAFC", fill_type="solid")
    fill_zebra_odd = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    fill_pk = PatternFill(start_color="FED7D7", end_color="FED7D7", fill_type="solid")
    fill_fk = PatternFill(start_color="FEEBC8", end_color="FEEBC8", fill_type="solid")
    fill_pass = PatternFill(start_color="C6F6D5", end_color="C6F6D5", fill_type="solid")

    thin_border_side = Side(border_style="thin", color="CBD5E0")
    border_cell = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    border_header = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=Side(border_style="medium", color="1A365D"))

    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)

    def apply_title(ws, title_text, sub_text, col_span="F"):
        ws.views.sheetView[0].showGridLines = True
        ws.merge_cells(f"A1:{col_span}1")
        ws["A1"] = title_text
        ws["A1"].font = font_title
        ws["A1"].alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[1].height = 24

        ws.merge_cells(f"A2:{col_span}2")
        ws["A2"] = sub_text
        ws["A2"].font = font_sub
        ws["A2"].alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[2].height = 16

    # ====================================================
    # SHEET 1: Overview
    # ====================================================
    ws1 = wb.active
    ws1.title = "Overview"
    apply_title(ws1, "DATABASE DESIGN ARCHITECTURE & RELATIONAL OVERVIEW", "Store Wise Sales Comparison Dashboard | Normalized 3NF SQLite Relational Store", "E")

    ws1.merge_cells("A4:E4")
    ws1["A4"] = "1. SCHEMA SPECIFICATION & ARCHITECTURAL SUMMARY"
    ws1["A4"].font = Font(name=FONT_FAMILY, size=11, bold=True, color="1A365D")
    ws1["A4"].fill = fill_sub_banner
    ws1.row_dimensions[4].height = 22

    overview_headers = ["Schema Component", "Specification Details", "Governing Standard / Design Pattern"]
    ws1.row_dimensions[5].height = 24
    for c_idx, h in enumerate(overview_headers, start=1):
        cell = ws1.cell(row=5, column=c_idx, value=h)
        cell.font = font_header; cell.fill = fill_header; cell.alignment = align_center; cell.border = border_header

    overview_rows = [
        ("Database Engine", "SQLite 3.35+ (file: data/app.db)", "Self-contained serverless relational database engine with zero external connection overhead"),
        ("Normalization Level", "Third Normal Form (3NF)", "Zero transitive dependencies; decoupled store dimensions, product catalog, and transaction fact tables"),
        ("Table Count", "5 Relational Tables (3 Dimensions + 2 Fact Tables)", "dim_stores, dim_products, fact_orders, fact_order_items, dim_sales_targets"),
        ("Joined Analysis View", "view_store_sales_analytics (36 Columns)", "Denormalized analytics view joining header facts, item line facts, product categories, and store metadata"),
        ("Key Integrity", "Enforced Foreign Key Constraints (PRAGMA foreign_keys = ON;)", "RESTRICT on delete/update to preserve historical sales audit trails and referential integrity"),
        ("Index Strategy", "Explicit B-Tree Indexes on all Foreign Keys & Filter Columns", "B-tree indexes on order_date, store_id, product_id, category_name, payment_method for sub-millisecond queries"),
        ("Storage Footprint", "Estimated Database Size: < 5 MB for seed data; < 25 MB for 50,000 rows", "Optimized column datatypes (TEXT, INTEGER, REAL) operating comfortably under Streamlit Cloud 1GB limit")
    ]
    for r_idx, (comp, spec, gov) in enumerate(overview_rows, start=6):
        ws1.row_dimensions[r_idx].height = 24
        c1 = ws1.cell(row=r_idx, column=1, value=comp); c1.font = font_data_bold; c1.alignment = align_left; c1.border = border_cell
        c2 = ws1.cell(row=r_idx, column=2, value=spec); c2.font = font_data; c2.alignment = align_left; c2.border = border_cell
        c3 = ws1.cell(row=r_idx, column=3, value=gov); c3.font = font_data; c3.alignment = align_left; c3.border = border_cell

    ws1_widths = [22, 42, 58]
    for c_idx, width in enumerate(ws1_widths, start=1):
        ws1.column_dimensions[get_column_letter(c_idx)].width = width

    # ====================================================
    # SHEET 2: Tables_Summary
    # ====================================================
    ws2 = wb.create_sheet(title="Tables_Summary")
    apply_title(ws2, "RELATIONAL TABLES CATALOG & ENTITY GRAIN SUMMARY", "Table Grain, Primary Keys, Foreign Keys, Column Counts & Planned Volumes", "G")

    tbl_headers = ["Table Name", "Table Type", "Entity Grain", "Primary Key (PK)", "Foreign Keys (FKs)", "Column Count", "Planned Seed Rows"]
    ws2.row_dimensions[4].height = 24
    for c_idx, h in enumerate(tbl_headers, start=1):
        cell = ws2.cell(row=4, column=c_idx, value=h)
        cell.font = font_header; cell.fill = fill_header; cell.alignment = align_center; cell.border = border_header

    tbl_summary_data = [
        ("dim_stores", "Dimension (Master)", "One row per retail store location", "store_id", "None", 10, "12 stores (Urban, Suburban, Rural)"),
        ("dim_products", "Dimension (Master)", "One row per sellable retail SKU", "product_id", "None", 10, "60 SKUs across 5 core categories"),
        ("fact_orders", "Fact (Transaction Header)", "One row per customer transaction checkout", "order_id", "store_id -> dim_stores(store_id)", 13, "250 - 300 seed orders (10k+ in Step 10)"),
        ("fact_order_items", "Fact (Line Item Detail)", "One row per individual product item within an order", "order_item_id", "order_id -> fact_orders(order_id), product_id -> dim_products(product_id)", 10, "600 - 800 line items (25k+ in Step 10)"),
        ("dim_sales_targets", "Dimension / Fact (Budget)", "One row per store per calendar month", "target_id", "store_id -> dim_stores(store_id)", 6, "144 target records (12 stores x 12 months)"),
        ("view_store_sales_analytics", "SQL Analysis View", "Denormalized line-item level with store & product dimensions", "N/A (Derived View)", "N/A (Joined on FKs)", 36, "Dynamically queried view over all line items")
    ]
    for r_idx, row_vals in enumerate(tbl_summary_data, start=5):
        ws2.row_dimensions[r_idx].height = 24
        fill_cur = fill_zebra_even if r_idx % 2 == 0 else fill_zebra_odd
        for c_idx, val in enumerate(row_vals, start=1):
            cell = ws2.cell(row=r_idx, column=c_idx, value=val)
            cell.font = font_data_bold if c_idx in (1, 4) else font_data
            cell.alignment = align_left if c_idx in (1, 3, 5, 7) else align_center
            cell.fill = fill_cur; cell.border = border_cell

    ws2_widths = [24, 22, 34, 18, 32, 14, 28]
    for c_idx, width in enumerate(ws2_widths, start=1):
        ws2.column_dimensions[get_column_letter(c_idx)].width = width
    ws2.auto_filter.ref = f"A4:G{len(tbl_summary_data) + 4}"

    # Helper function for Data Dictionary sheets
    def build_dd_sheet(sheet_title, table_name, desc_text, columns_data):
        ws = wb.create_sheet(title=sheet_title)
        apply_title(ws, f"DATA DICTIONARY: {table_name}", desc_text, "G")
        dd_headers = ["Column Name", "Data Type", "Key Constraint", "Nullable", "Allowed Values / Format", "Business Description", "OQ / BR Supported"]
        ws.row_dimensions[4].height = 24
        for c_idx, h in enumerate(dd_headers, start=1):
            cell = ws.cell(row=4, column=c_idx, value=h)
            cell.font = font_header; cell.fill = fill_header; cell.alignment = align_center; cell.border = border_header

        for r_idx, row in enumerate(columns_data, start=5):
            ws.row_dimensions[r_idx].height = 24
            fill_cur = fill_zebra_even if r_idx % 2 == 0 else fill_zebra_odd
            col_name, dtype, key, null_val, allowed, desc, oq_br = row
            c1 = ws.cell(row=r_idx, column=1, value=col_name); c1.font = font_code; c1.alignment = align_left; c1.border = border_cell
            c2 = ws.cell(row=r_idx, column=2, value=dtype); c2.font = font_data; c2.alignment = align_center; c2.border = border_cell
            c3 = ws.cell(row=r_idx, column=3, value=key); c3.alignment = align_center; c3.border = border_cell
            if key == "PK":
                c3.font = Font(name=FONT_FAMILY, size=9.5, bold=True, color="9B2C2C"); c3.fill = fill_pk
            elif "FK" in key:
                c3.font = Font(name=FONT_FAMILY, size=9.5, bold=True, color="C05621"); c3.fill = fill_fk
            else:
                c3.font = font_data; c3.fill = fill_cur

            c4 = ws.cell(row=r_idx, column=4, value=null_val); c4.font = font_data; c4.alignment = align_center; c4.border = border_cell; c4.fill = fill_cur
            c5 = ws.cell(row=r_idx, column=5, value=allowed); c5.font = font_data; c5.alignment = align_left; c5.border = border_cell; c5.fill = fill_cur
            c6 = ws.cell(row=r_idx, column=6, value=desc); c6.font = font_data; c6.alignment = align_left; c6.border = border_cell; c6.fill = fill_cur
            c7 = ws.cell(row=r_idx, column=7, value=oq_br); c7.font = font_data_bold; c7.alignment = align_center; c7.border = border_cell; c7.fill = fill_cur

        ws_widths = [22, 14, 14, 12, 28, 48, 20]
        for c_idx, width in enumerate(ws_widths, start=1):
            ws.column_dimensions[get_column_letter(c_idx)].width = width
        ws.auto_filter.ref = f"A4:G{len(columns_data) + 4}"

    # Sheet 3: DD_dim_stores
    dim_stores_cols = [
        ("store_id", "TEXT", "PK", "No", "'STR-001' to 'STR-012'", "Unique primary key identifying the retail store outlet", "OQ-01, OQ-02, BR-01"),
        ("store_name", "TEXT", "None", "No", "Alphanumeric string (e.g., 'Metro Central Hub')", "Official public commercial name of the store", "OQ-01, BR-01"),
        ("region_id", "TEXT", "None", "No", "'REG-N', 'REG-S', 'REG-W'", "Code identifying the geographic sales territory", "OQ-04, BR-04"),
        ("region_name", "TEXT", "None", "No", "'North', 'South', 'West'", "Human-readable geographic region cluster name", "OQ-04, BR-04"),
        ("city", "TEXT", "None", "No", "'Noida', 'Delhi', 'Bengaluru', 'Mumbai', etc.", "City location of the retail store outlet", "OQ-01, BR-01"),
        ("state", "TEXT", "None", "No", "Standard Indian State Names (e.g., 'Uttar Pradesh')", "State jurisdiction of the retail store outlet", "OQ-01, BR-01"),
        ("store_tier", "TEXT", "None", "No", "'Urban', 'Suburban', 'Rural'", "Demographic format classification of the store footprint", "OQ-03, BR-03"),
        ("store_size_sqft", "INTEGER", "None", "No", "Integer between 1,200 and 8,500 sq.ft.", "Total physical retail floor space area in square feet", "OQ-02, BR-02"),
        ("opening_date", "TEXT", "None", "No", "ISO Date: 'YYYY-MM-DD'", "Historical calendar date when the store launched operations", "BR-12"),
        ("is_active", "INTEGER", "None", "No", "1 (Active), 0 (Closed)", "Operational status flag indicating active trading", "BR-12")
    ]
    build_dd_sheet("DD_dim_stores", "dim_stores", "Master dimension containing retail store outlet locations, physical dimensions & demographic tiers", dim_stores_cols)

    # Sheet 4: DD_dim_products
    dim_products_cols = [
        ("product_id", "TEXT", "PK", "No", "'PRD-001' to 'PRD-060'", "Unique primary key identifying each sellable retail SKU", "OQ-03, OQ-07, BR-03"),
        ("product_name", "TEXT", "None", "No", "Descriptive SKU Name (e.g., 'Wireless Noise-Canceling Headphones')", "Commercial brand name of the retail merchandise item", "OQ-03, BR-03"),
        ("category_id", "TEXT", "None", "No", "'CAT-01' to 'CAT-05'", "Code identifying product merchandise classification", "OQ-03, BR-03"),
        ("category_name", "TEXT", "None", "No", "'Electronics', 'Apparel', 'Groceries', 'Home Goods', 'Footwear'", "Primary product category cluster name", "OQ-03, BR-03"),
        ("sub_category", "TEXT", "None", "No", "'Audio', 'Menswear', 'Dairy', 'Kitchenware', etc.", "Secondary granular sub-category classification", "OQ-03, BR-03"),
        ("cost_price_inr", "REAL", "None", "No", "Monetary value > 0.0 in INR", "Cost of goods sold (COGS) base wholesale acquisition cost", "OQ-01, OQ-07, BR-01"),
        ("base_selling_price_inr", "REAL", "None", "No", "Monetary value >= cost_price_inr in INR", "Standard manufacturer suggested retail selling price (MSRP)", "OQ-01, OQ-07, BR-01"),
        ("target_margin_pct", "REAL", "None", "No", "Percentage: 10.0% to 65.0%", "Target gross profitability margin threshold for the SKU", "OQ-07, BR-08"),
        ("sku_code", "TEXT", "UNIQUE", "No", "'SKU-EL-001', etc.", "Standard barcoded stock-keeping unit inventory identifier", "BR-12"),
        ("is_active", "INTEGER", "None", "No", "1 (Active), 0 (Discontinued)", "Inventory lifecycle flag indicating active catalog item", "BR-12")
    ]
    build_dd_sheet("DD_dim_products", "dim_products", "Master product catalog containing merchandise taxonomy, wholesale costs & baseline retail prices", dim_products_cols)

    # Sheet 5: DD_fact_orders
    fact_orders_cols = [
        ("order_id", "TEXT", "PK", "No", "'ORD-2024-0001' upwards", "Unique transaction header identifier per customer checkout", "OQ-02, BR-02, BR-12"),
        ("store_id", "TEXT", "FK", "No", "Must exist in dim_stores(store_id)", "Foreign key linking transaction to the executing store", "OQ-01, OQ-02, BR-01"),
        ("order_datetime", "TEXT", "None", "No", "ISO-8601: 'YYYY-MM-DD HH:MM:SS'", "Precise transactional checkout timestamp", "OQ-10, BR-05"),
        ("order_date", "TEXT", "None", "No", "ISO Date: 'YYYY-MM-DD'", "Calendar date of checkout (Indexed for date filtering)", "OQ-04, BR-04"),
        ("order_time", "TEXT", "None", "No", "Time: 'HH:MM:SS'", "Local time component of checkout", "OQ-10, BR-05"),
        ("order_hour", "INTEGER", "None", "No", "0 to 23", "Extracted 24-hour integer bucket for rush analysis", "OQ-10, BR-05"),
        ("day_of_week", "TEXT", "None", "No", "'Monday', 'Tuesday', ..., 'Sunday'", "Day of the week name for temporal aggregation", "OQ-05, BR-06"),
        ("is_weekend_flag", "INTEGER", "None", "No", "1 (Saturday/Sunday), 0 (Monday-Friday)", "Binary flag partitioning weekday vs. weekend transactions", "OQ-05, BR-06"),
        ("payment_method", "TEXT", "None", "No", "'UPI', 'Credit Card', 'Debit Card', 'Cash'", "Customer tender payment method used for settlement", "OQ-09, BR-10"),
        ("order_status", "TEXT", "None", "No", "'Completed', 'Returned', 'Cancelled'", "Lifecycle audit status of the transaction", "OQ-06, BR-07"),
        ("gross_amount_inr", "REAL", "None", "No", "Sum of items before discounts/returns (>= 0.0)", "Total order value before promotional markdowns", "OQ-01, BR-01"),
        ("discount_amount_inr", "REAL", "None", "No", "Total promotional markdown value (>= 0.0)", "Total price discount conceded on the order", "OQ-07, BR-08"),
        ("return_amount_inr", "REAL", "None", "No", "Value refunded upon return/cancellation (>= 0.0)", "Revenue deducted due to product returns/cancellations", "OQ-06, BR-07"),
        ("net_amount_inr", "REAL", "None", "No", "Gross - Discount - Return (Formula BURL-01)", "Net recognized sales revenue for the transaction", "OQ-01, OQ-02, BR-01")
    ]
    build_dd_sheet("DD_fact_orders", "fact_orders", "Transaction header fact table recording checkout timestamps, store links, tenders & net totals", fact_orders_cols)

    # Sheet 6: DD_fact_order_items
    fact_order_items_cols = [
        ("order_item_id", "INTEGER", "PK", "No", "Auto-incrementing integer (1, 2, 3...)", "Unique surrogate key identifying each order line item", "BR-12"),
        ("order_id", "TEXT", "FK", "No", "Must exist in fact_orders(order_id)", "Foreign key linking line item to parent transaction header", "OQ-02, BR-02"),
        ("product_id", "TEXT", "FK", "No", "Must exist in dim_products(product_id)", "Foreign key linking line item to the purchased SKU", "OQ-03, OQ-07, BR-03"),
        ("quantity", "INTEGER", "None", "No", "Integer >= 1 (Typically 1 to 10)", "Number of SKU units purchased in this order line", "OQ-07, BR-08"),
        ("unit_cost_price_inr", "REAL", "None", "No", "Historical unit wholesale cost at sale time", "Cost price recorded at the moment of transaction", "OQ-01, BR-01"),
        ("unit_selling_price_inr", "REAL", "None", "No", "Historical unit base retail price at sale time", "Gross retail price per unit before line discount", "OQ-01, BR-01"),
        ("discount_pct", "REAL", "None", "No", "0.0% to 50.0% (Enforced by CHECK constraint)", "Promotional markdown percentage applied to the item", "OQ-07, BR-08"),
        ("line_gross_amount_inr", "REAL", "None", "No", "quantity * unit_selling_price_inr", "Gross total for the line item prior to discount", "OQ-01, BR-01"),
        ("line_discount_amount_inr", "REAL", "None", "No", "line_gross_amount * (discount_pct / 100.0)", "Total monetary discount deducted from the line", "OQ-07, BR-08"),
        ("line_net_amount_inr", "REAL", "None", "No", "line_gross_amount - line_discount_amount", "Net recognized revenue realized from the line item", "OQ-01, OQ-03, BR-01"),
        ("line_cogs_inr", "REAL", "None", "No", "quantity * unit_cost_price_inr", "Total cost of goods sold attributed to the line item", "OQ-01, BR-01"),
        ("line_profit_inr", "REAL", "None", "No", "line_net_amount_inr - line_cogs_inr", "Gross profit contribution realized from this line item", "OQ-01, BR-01"),
        ("line_margin_pct", "REAL", "None", "No", "(line_profit / line_net_amount) * 100.0", "Gross profit margin percentage yielded by the line", "OQ-01, OQ-07, BR-01")
    ]
    build_dd_sheet("DD_fact_order_items", "fact_order_items", "Transaction line item fact table recording item quantities, unit prices, discounts & net margins", fact_order_items_cols)

    # Sheet 7: DD_dim_sales_targets
    dim_sales_targets_cols = [
        ("target_id", "TEXT", "PK", "No", "'TGT-STR001-202401', etc.", "Unique primary key identifying monthly store budget quota", "OQ-08, BR-09"),
        ("store_id", "TEXT", "FK", "No", "Must exist in dim_stores(store_id)", "Foreign key linking budget target to specific retail store", "OQ-08, BR-09"),
        ("target_month", "TEXT", "None", "No", "Format 'YYYY-MM' (e.g., '2024-01')", "Fiscal calendar month assigned for the revenue quota", "OQ-08, BR-09"),
        ("target_revenue_inr", "REAL", "None", "No", "Budgeted sales target in INR (e.g., 850,000.0)", "Assigned commercial revenue benchmark quota", "OQ-08, BR-09"),
        ("target_orders_count", "INTEGER", "None", "No", "Expected order checkout volume (e.g., 450)", "Assigned transaction volume quota", "OQ-08, BR-09"),
        ("target_atv_inr", "REAL", "None", "No", "target_revenue_inr / target_orders_count", "Planned average transaction basket value benchmark", "OQ-08, BR-09")
    ]
    build_dd_sheet("DD_dim_sales_targets", "dim_sales_targets", "Monthly store revenue budget and sales quota target dimension for variance tracking", dim_sales_targets_cols)

    # Sheet 8: Relationships
    ws8 = wb.create_sheet(title="Relationships")
    apply_title(ws8, "ENTITY RELATIONSHIPS & REFERENTIAL INTEGRITY RULES", "Parent-Child Constraints, Foreign Keys, Cardinalities & Deletion Policies", "F")

    rel_headers = ["Parent Entity (1)", "Child Entity (M)", "Foreign Key Reference", "Cardinality", "On Delete / On Update Action", "Business Integrity Rule"]
    ws8.row_dimensions[4].height = 24
    for c_idx, h in enumerate(rel_headers, start=1):
        cell = ws8.cell(row=4, column=c_idx, value=h)
        cell.font = font_header; cell.fill = fill_header; cell.alignment = align_center; cell.border = border_header

    rel_data = [
        ("dim_stores", "fact_orders", "fact_orders.store_id = dim_stores.store_id", "1 to Many (1:N)", "ON DELETE RESTRICT\nON UPDATE CASCADE", "A store cannot be deleted while historical transaction records exist."),
        ("fact_orders", "fact_order_items", "fact_order_items.order_id = fact_orders.order_id", "1 to Many (1:N)", "ON DELETE CASCADE\nON UPDATE CASCADE", "Deleting an order header purges all associated line items to prevent orphaned facts."),
        ("dim_products", "fact_order_items", "fact_order_items.product_id = dim_products.product_id", "1 to Many (1:N)", "ON DELETE RESTRICT\nON UPDATE CASCADE", "A product cannot be deleted if historical sales line items reference its product_id."),
        ("dim_stores", "dim_sales_targets", "dim_sales_targets.store_id = dim_stores.store_id", "1 to Many (1:N)", "ON DELETE RESTRICT\nON UPDATE CASCADE", "A store cannot be deleted if active monthly budget quota targets are assigned.")
    ]
    for r_idx, (p, c, fk, card, act, rule) in enumerate(rel_data, start=5):
        ws8.row_dimensions[r_idx].height = 32
        fill_cur = fill_zebra_even if r_idx % 2 == 0 else fill_zebra_odd
        c1 = ws8.cell(row=r_idx, column=1, value=p); c1.font = font_data_bold; c1.alignment = align_left; c1.fill = fill_cur; c1.border = border_cell
        c2 = ws8.cell(row=r_idx, column=2, value=c); c2.font = font_data_bold; c2.alignment = align_left; c2.fill = fill_cur; c2.border = border_cell
        c3 = ws8.cell(row=r_idx, column=3, value=fk); c3.font = font_code; c3.alignment = align_left; c3.fill = fill_cur; c3.border = border_cell
        c4 = ws8.cell(row=r_idx, column=4, value=card); c4.font = font_data_bold; c4.alignment = align_center; c4.fill = fill_cur; c4.border = border_cell
        c5 = ws8.cell(row=r_idx, column=5, value=act); c5.font = font_code; c5.alignment = align_center; c5.fill = fill_cur; c5.border = border_cell
        c6 = ws8.cell(row=r_idx, column=6, value=rule); c6.font = font_data; c6.alignment = align_left; c6.fill = fill_cur; c6.border = border_cell

    ws8_widths = [20, 20, 36, 18, 26, 44]
    for c_idx, width in enumerate(ws8_widths, start=1):
        ws8.column_dimensions[get_column_letter(c_idx)].width = width
    ws8.auto_filter.ref = f"A4:F{len(rel_data) + 4}"

    # Sheet 9: KPIs_Measures
    ws9 = wb.create_sheet(title="KPIs_Measures")
    apply_title(ws9, "BUSINESS METRICS & ANALYTICAL FORMULATION REPOSITORY", "Key Performance Indicators (KPIs), Mathematical Formulas, Target Dimensions & OQs", "E")

    kpi_headers = ["Measure / KPI Name", "Canonical Mathematical Formula", "Underlying Database Source Fields", "Target Visual Representation", "Supported OQ"]
    ws9.row_dimensions[4].height = 24
    for c_idx, h in enumerate(kpi_headers, start=1):
        cell = ws9.cell(row=4, column=c_idx, value=h)
        cell.font = font_header; cell.fill = fill_header; cell.alignment = align_center; cell.border = border_header

    kpi_measures_data = [
        ("Gross Sales Revenue", "SUM(fact_orders.gross_amount_inr)", "fact_orders.gross_amount_inr", "KPI Metric Card & Ranked Bar Chart", "OQ-01"),
        ("Net Sales Revenue", "SUM(gross_amount_inr) - SUM(discount_amount_inr) - SUM(return_amount_inr)", "fact_orders.gross_amount_inr, return_amount_inr", "KPI Metric Card & Waterfall Leakage Chart", "OQ-01, OQ-06"),
        ("Gross Profit Margin %", "(SUM(line_profit_inr) / SUM(line_net_amount_inr)) * 100.0", "fact_order_items.line_profit_inr, line_net_amount_inr", "Comparative Dual-Axis Bar & Line Chart", "OQ-01"),
        ("Average Transaction Value (ATV)", "SUM(net_amount_inr) / COUNT(DISTINCT order_id)", "fact_orders.net_amount_inr, fact_orders.order_id", "Metric Card & Outlier Scatter Plot", "OQ-02"),
        ("Sales Velocity per Sq.Ft.", "SUM(net_amount_inr) / dim_stores.store_size_sqft", "fact_orders.net_amount_inr, dim_stores.store_size_sqft", "Space Productivity Scatter Plot (SqFt vs Sales)", "OQ-02"),
        ("Category Revenue Contribution %", "(SUM(category_sales) / SUM(store_tier_total_sales)) * 100.0", "dim_products.category_name, dim_stores.store_tier", "100% Stacked Horizontal Bar Chart", "OQ-03"),
        ("Month-over-Month (MoM) Growth %", "((Current_Month_Sales - Prior_Month_Sales) / Prior_Month_Sales) * 100.0", "fact_orders.order_date, fact_orders.net_amount_inr", "Multi-Line Regional Trajectory Trend Chart", "OQ-04"),
        ("Weekday vs. Weekend Sales Split", "SUM(CASE WHEN is_weekend_flag=1 THEN net_amount ELSE 0 END) vs. Weekday", "fact_orders.is_weekend_flag, fact_orders.net_amount_inr", "Grouped Comparative Bar Chart & Welch's t-test", "OQ-05"),
        ("Product Return & Cancellation Rate %", "(SUM(return_amount_inr) / SUM(gross_amount_inr)) * 100.0", "fact_orders.order_status, return_amount_inr", "Return Leakage Metric Badge & Pareto Chart", "OQ-06"),
        ("Discount Elasticity Coefficient (PED)", "Pearson Correlation (discount_pct, quantity) & OLS Slope Beta_1", "fact_order_items.discount_pct, quantity, margin", "Regression Trendline with 95% Confidence Band", "OQ-07"),
        ("Target Quota Variance %", "((SUM(actual_net_amount) - target_revenue_inr) / target_revenue_inr) * 100.0", "fact_orders.net_amount_inr, dim_sales_targets.target", "Bullet Graph & Dynamic Amber/Red Alert Badges", "OQ-08"),
        ("Payment Tender Distribution %", "(COUNT(order_id by payment_method) / Total_Orders) * 100.0", "fact_orders.payment_method, fact_orders.order_id", "Donut Breakdown Chart & Basket Value Boxplot", "OQ-09"),
        ("Hourly Trading Density (Rush Windows)", "COUNT(order_id) and SUM(net_amount_inr) grouped by order_hour (0-23)", "fact_orders.order_hour, dim_stores.store_id", "24-Hour Intensity Heatmap Matrix (Store x Hour)", "OQ-10")
    ]
    for r_idx, row_vals in enumerate(kpi_measures_data, start=5):
        ws9.row_dimensions[r_idx].height = 24
        fill_cur = fill_zebra_even if r_idx % 2 == 0 else fill_zebra_odd
        for c_idx, val in enumerate(row_vals, start=1):
            cell = ws9.cell(row=r_idx, column=c_idx, value=val)
            cell.font = font_data_bold if c_idx in (1, 5) else font_data
            cell.alignment = align_left if c_idx in (1, 2, 3, 4) else align_center
            cell.fill = fill_cur; cell.border = border_cell

    ws9_widths = [26, 46, 38, 36, 16]
    for c_idx, width in enumerate(ws9_widths, start=1):
        ws9.column_dimensions[get_column_letter(c_idx)].width = width
    ws9.auto_filter.ref = f"A4:E{len(kpi_measures_data) + 4}"

    # Sheet 10: SQLite_DDL
    ws10 = wb.create_sheet(title="SQLite_DDL")
    apply_title(ws10, "SQLITE PRODUCTION DDL (CREATE TABLE, INDEXES & VIEWS)", "Executable DDL Schema with Foreign Keys, Check Constraints & 36-Column Analytical View", "C")

    ddl_headers = ["DDL Object Name", "Object Type", "SQL Execution Statement"]
    ws10.row_dimensions[4].height = 24
    for c_idx, h in enumerate(ddl_headers, start=1):
        cell = ws10.cell(row=4, column=c_idx, value=h)
        cell.font = font_header; cell.fill = fill_header; cell.alignment = align_center; cell.border = border_header

    ddl_statements = [
        ("PRAGMA foreign_keys", "Configuration", "PRAGMA foreign_keys = ON;"),
        ("dim_stores", "Table DDL", 
"""CREATE TABLE IF NOT EXISTS dim_stores (
    store_id TEXT PRIMARY KEY,
    store_name TEXT NOT NULL,
    region_id TEXT NOT NULL,
    region_name TEXT NOT NULL CHECK(region_name IN ('North', 'South', 'West')),
    city TEXT NOT NULL,
    state TEXT NOT NULL,
    store_tier TEXT NOT NULL CHECK(store_tier IN ('Urban', 'Suburban', 'Rural')),
    store_size_sqft INTEGER NOT NULL CHECK(store_size_sqft > 0),
    opening_date TEXT NOT NULL,
    is_active INTEGER NOT NULL DEFAULT 1 CHECK(is_active IN (0, 1))
);"""),
        ("dim_products", "Table DDL",
"""CREATE TABLE IF NOT EXISTS dim_products (
    product_id TEXT PRIMARY KEY,
    product_name TEXT NOT NULL,
    category_id TEXT NOT NULL,
    category_name TEXT NOT NULL CHECK(category_name IN ('Electronics', 'Apparel', 'Groceries', 'Home Goods', 'Footwear')),
    sub_category TEXT NOT NULL,
    cost_price_inr REAL NOT NULL CHECK(cost_price_inr > 0.0),
    base_selling_price_inr REAL NOT NULL CHECK(base_selling_price_inr >= cost_price_inr),
    target_margin_pct REAL NOT NULL,
    sku_code TEXT NOT NULL UNIQUE,
    is_active INTEGER NOT NULL DEFAULT 1 CHECK(is_active IN (0, 1))
);"""),
        ("fact_orders", "Table DDL",
"""CREATE TABLE IF NOT EXISTS fact_orders (
    order_id TEXT PRIMARY KEY,
    store_id TEXT NOT NULL,
    order_datetime TEXT NOT NULL,
    order_date TEXT NOT NULL,
    order_time TEXT NOT NULL,
    order_hour INTEGER NOT NULL CHECK(order_hour BETWEEN 0 AND 23),
    day_of_week TEXT NOT NULL,
    is_weekend_flag INTEGER NOT NULL CHECK(is_weekend_flag IN (0, 1)),
    payment_method TEXT NOT NULL CHECK(payment_method IN ('UPI', 'Credit Card', 'Debit Card', 'Cash')),
    order_status TEXT NOT NULL CHECK(order_status IN ('Completed', 'Returned', 'Cancelled')),
    gross_amount_inr REAL NOT NULL CHECK(gross_amount_inr >= 0.0),
    discount_amount_inr REAL NOT NULL DEFAULT 0.0 CHECK(discount_amount_inr >= 0.0),
    return_amount_inr REAL NOT NULL DEFAULT 0.0 CHECK(return_amount_inr >= 0.0),
    net_amount_inr REAL NOT NULL,
    FOREIGN KEY (store_id) REFERENCES dim_stores(store_id) ON DELETE RESTRICT ON UPDATE CASCADE
);"""),
        ("fact_order_items", "Table DDL",
"""CREATE TABLE IF NOT EXISTS fact_order_items (
    order_item_id INTEGER PRIMARY KEY AUTOINCREMENT,
    order_id TEXT NOT NULL,
    product_id TEXT NOT NULL,
    quantity INTEGER NOT NULL CHECK(quantity >= 1),
    unit_cost_price_inr REAL NOT NULL CHECK(unit_cost_price_inr > 0.0),
    unit_selling_price_inr REAL NOT NULL CHECK(unit_selling_price_inr > 0.0),
    discount_pct REAL NOT NULL DEFAULT 0.0 CHECK(discount_pct BETWEEN 0.0 AND 50.0),
    line_gross_amount_inr REAL NOT NULL,
    line_discount_amount_inr REAL NOT NULL DEFAULT 0.0,
    line_net_amount_inr REAL NOT NULL,
    line_cogs_inr REAL NOT NULL,
    line_profit_inr REAL NOT NULL,
    line_margin_pct REAL NOT NULL,
    FOREIGN KEY (order_id) REFERENCES fact_orders(order_id) ON DELETE CASCADE ON UPDATE CASCADE,
    FOREIGN KEY (product_id) REFERENCES dim_products(product_id) ON DELETE RESTRICT ON UPDATE CASCADE
);"""),
        ("dim_sales_targets", "Table DDL",
"""CREATE TABLE IF NOT EXISTS dim_sales_targets (
    target_id TEXT PRIMARY KEY,
    store_id TEXT NOT NULL,
    target_month TEXT NOT NULL,
    target_revenue_inr REAL NOT NULL CHECK(target_revenue_inr > 0.0),
    target_orders_count INTEGER NOT NULL CHECK(target_orders_count > 0),
    target_atv_inr REAL NOT NULL,
    FOREIGN KEY (store_id) REFERENCES dim_stores(store_id) ON DELETE RESTRICT ON UPDATE CASCADE
);"""),
        ("idx_orders_date_store", "B-Tree Index", "CREATE INDEX IF NOT EXISTS idx_orders_date_store ON fact_orders(order_date, store_id);"),
        ("idx_orders_hour", "B-Tree Index", "CREATE INDEX IF NOT EXISTS idx_orders_hour ON fact_orders(order_hour);"),
        ("idx_order_items_product", "B-Tree Index", "CREATE INDEX IF NOT EXISTS idx_order_items_product ON fact_order_items(product_id);"),
        ("idx_products_category", "B-Tree Index", "CREATE INDEX IF NOT EXISTS idx_products_category ON dim_products(category_name);"),
        ("view_store_sales_analytics", "SQL View (36 Columns)",
"""CREATE VIEW IF NOT EXISTS view_store_sales_analytics AS
SELECT 
    foi.order_item_id,
    fo.order_id,
    fo.order_datetime,
    fo.order_date,
    fo.order_time,
    fo.order_hour,
    strftime('%Y', fo.order_date) AS order_year,
    strftime('%m', fo.order_date) AS order_month,
    strftime('%Y-%m', fo.order_date) AS order_year_month,
    CASE 
        WHEN strftime('%m', fo.order_date) IN ('01', '02', '03') THEN 'Q1'
        WHEN strftime('%m', fo.order_date) IN ('04', '05', '06') THEN 'Q2'
        WHEN strftime('%m', fo.order_date) IN ('07', '08', '09') THEN 'Q3'
        ELSE 'Q4'
    END AS order_quarter,
    fo.day_of_week,
    fo.is_weekend_flag,
    ds.store_id,
    ds.store_name,
    ds.region_name,
    ds.city,
    ds.state,
    ds.store_tier,
    ds.store_size_sqft,
    dp.product_id,
    dp.product_name,
    dp.category_name,
    dp.sub_category,
    foi.quantity,
    foi.unit_cost_price_inr,
    foi.unit_selling_price_inr,
    foi.discount_pct,
    foi.line_gross_amount_inr,
    foi.line_discount_amount_inr,
    foi.line_net_amount_inr,
    foi.line_cogs_inr,
    foi.line_profit_inr,
    foi.line_margin_pct,
    fo.payment_method,
    fo.order_status,
    ROUND(foi.line_net_amount_inr / ds.store_size_sqft, 4) AS sales_velocity_sqft_contrib
FROM fact_order_items foi
JOIN fact_orders fo ON foi.order_id = fo.order_id
JOIN dim_stores ds ON fo.store_id = ds.store_id
JOIN dim_products dp ON foi.product_id = dp.product_id;""")
    ]

    for r_idx, (obj_name, obj_type, sql_code) in enumerate(ddl_statements, start=5):
        ws10.row_dimensions[r_idx].height = 42 if "\n" not in sql_code else min(140, 18 * sql_code.count("\n") + 20)
        fill_cur = fill_zebra_even if r_idx % 2 == 0 else fill_zebra_odd
        c1 = ws10.cell(row=r_idx, column=1, value=obj_name); c1.font = font_data_bold; c1.alignment = align_left; c1.fill = fill_cur; c1.border = border_cell
        c2 = ws10.cell(row=r_idx, column=2, value=obj_type); c2.font = font_data; c2.alignment = align_center; c2.fill = fill_cur; c2.border = border_cell
        c3 = ws10.cell(row=r_idx, column=3, value=sql_code); c3.font = font_code_ddl; c3.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True); c3.fill = fill_cur; c3.border = border_cell

    ws10_widths = [24, 18, 80]
    for c_idx, width in enumerate(ws10_widths, start=1):
        ws10.column_dimensions[get_column_letter(c_idx)].width = width

    wb.save(filename)
    print(f"Database Design Excel successfully created at: {filename}")

if __name__ == "__main__":
    import sys
    target = sys.argv[1] if len(sys.argv) > 1 else "docs/08_Database_Design.xlsx"
    create_database_design_excel(target)
