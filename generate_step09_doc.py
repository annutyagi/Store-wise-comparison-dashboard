import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

def create_er_diagram_doc(filename, image_path):
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
    s_run = sub_p.add_run("Step 9: Entity Relationship (ER) Diagram & Structural Integrity Specification")
    s_run.font.size = Pt(12)
    s_run.font.color.rgb = COLOR_MUTED

    # Document Metadata Block
    meta_table = doc.add_table(rows=4, cols=2)
    meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta_data = [
        [("Learner Name:", "Anu Tyagi"), ("Roll Number:", "2500320140025")],
        [("Batch ID:", "ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01"), ("Institution:", "ABES Engineering College")],
        [("Project Mentor:", "Mr. Dinesh"), ("Document Type:", "Entity Relationship (ER) Architecture")],
        [("Upstream Artifact:", "docs/08_Database_Design.xlsx"), ("Schema Standard:", "Third Normal Form (3NF) Relational")]
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

    # Section 1: Visual Entity Relationship Diagram Embedding
    add_custom_heading("1. Visual Entity Relationship (ER) Diagram", level=1)
    doc.add_paragraph(
        "The following diagram illustrates the complete 3NF relational data model engineered for the Store Wise Sales Comparison Dashboard. "
        "It depicts the 5 core relational entities, primary/foreign keys, cardinality constraints, and the denormalized 36-column analytical view:"
    )

    if os.path.exists(image_path):
        img_p = doc.add_paragraph()
        img_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        img_p.paragraph_format.space_before = Pt(6)
        img_p.paragraph_format.space_after = Pt(6)
        img_p.add_run().add_picture(image_path, width=Inches(6.5))
        
        cap_p = doc.add_paragraph()
        cap_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cap_p.paragraph_format.space_after = Pt(10)
        c_run = cap_p.add_run("Figure 9.1: Comprehensive Relational Entity Relationship Diagram (3NF Schema)")
        c_run.font.size = Pt(9.0)
        c_run.font.italic = True
        c_run.font.color.rgb = COLOR_MUTED

    # Section 2: Canonical Mermaid ER Diagram Source Code
    add_custom_heading("2. Canonical Mermaid erDiagram Specification", level=1)
    doc.add_paragraph(
        "Below is the complete, production-grade Mermaid code representing the database model. "
        "This specification can be copied directly into mermaid.live or any compatible Markdown viewer:"
    )

    mermaid_code = (
        "erDiagram\n"
        "    dim_stores ||--o{ fact_orders : \"operates / hosts\"\n"
        "    fact_orders ||--|{ fact_order_items : \"contains\"\n"
        "    dim_products ||--o{ fact_order_items : \"referenced_in\"\n"
        "    dim_stores ||--o{ dim_sales_targets : \"assigned\"\n\n"
        "    dim_stores {\n"
        "        TEXT store_id PK \"Primary Key (e.g. STR-001)\"\n"
        "        TEXT store_name \"Commercial Store Name\"\n"
        "        TEXT region_id \"Region Code\"\n"
        "        TEXT region_name \"North | South | West\"\n"
        "        TEXT city \"Store City Location\"\n"
        "        TEXT state \"State Jurisdiction\"\n"
        "        TEXT store_tier \"Urban | Suburban | Rural\"\n"
        "        INTEGER store_size_sqft \"Selling Space Area\"\n"
        "        TEXT opening_date \"Launch Date YYYY-MM-DD\"\n"
        "        INTEGER is_active \"1=Active, 0=Closed\"\n"
        "    }\n\n"
        "    dim_products {\n"
        "        TEXT product_id PK \"Primary Key (e.g. PRD-001)\"\n"
        "        TEXT product_name \"Commercial SKU Name\"\n"
        "        TEXT category_id \"Category Identifier\"\n"
        "        TEXT category_name \"Electronics | Apparel | Groceries...\"\n"
        "        TEXT sub_category \"Granular Classification\"\n"
        "        REAL cost_price_inr \"Wholesale Cost Price (COGS)\"\n"
        "        REAL base_selling_price_inr \"Base Retail Selling Price\"\n"
        "        REAL target_margin_pct \"Target Gross Profit Margin %\"\n"
        "        TEXT sku_code UK \"Unique Inventory Barcode\"\n"
        "        INTEGER is_active \"1=Active, 0=Discontinued\"\n"
        "    }\n\n"
        "    fact_orders {\n"
        "        TEXT order_id PK \"Unique Transaction Identifier\"\n"
        "        TEXT store_id FK \"Foreign Key to dim_stores\"\n"
        "        TEXT order_datetime \"ISO-8601 Timestamp\"\n"
        "        TEXT order_date \"Date YYYY-MM-DD\"\n"
        "        TEXT order_time \"Time HH:MM:SS\"\n"
        "        INTEGER order_hour \"Hour (0 to 23)\"\n"
        "        TEXT day_of_week \"Monday to Sunday\"\n"
        "        INTEGER is_weekend_flag \"1=Weekend, 0=Weekday\"\n"
        "        TEXT payment_method \"UPI | Card | Cash\"\n"
        "        TEXT order_status \"Completed | Returned | Cancelled\"\n"
        "        REAL gross_amount_inr \"Gross Order Total\"\n"
        "        REAL discount_amount_inr \"Total Conceded Discount\"\n"
        "        REAL return_amount_inr \"Refund Value on Return\"\n"
        "        REAL net_amount_inr \"Recognized Net Sales\"\n"
        "    }\n\n"
        "    fact_order_items {\n"
        "        INTEGER order_item_id PK \"Surrogate Key Auto-Increment\"\n"
        "        TEXT order_id FK \"Foreign Key to fact_orders\"\n"
        "        TEXT product_id FK \"Foreign Key to dim_products\"\n"
        "        INTEGER quantity \"Purchased Quantity (>= 1)\"\n"
        "        REAL unit_cost_price_inr \"Snapshot Unit Cost\"\n"
        "        REAL unit_selling_price_inr \"Snapshot Retail Price\"\n"
        "        REAL discount_pct \"Promotional Discount (0-50%)\"\n"
        "        REAL line_gross_amount_inr \"Quantity * Unit Price\"\n"
        "        REAL line_discount_amount_inr \"Discount Amount\"\n"
        "        REAL line_net_amount_inr \"Recognized Line Revenue\"\n"
        "        REAL line_cogs_inr \"Quantity * Unit Cost\"\n"
        "        REAL line_profit_inr \"Line Net - Line COGS\"\n"
        "        REAL line_margin_pct \"Gross Profit Margin %\"\n"
        "    }\n\n"
        "    dim_sales_targets {\n"
        "        TEXT target_id PK \"Target Quota Record ID\"\n"
        "        TEXT store_id FK \"Foreign Key to dim_stores\"\n"
        "        TEXT target_month \"Target Month YYYY-MM\"\n"
        "        REAL target_revenue_inr \"Budgeted Sales Target (INR)\"\n"
        "        INTEGER target_orders_count \"Target Transaction Volume\"\n"
        "        REAL target_atv_inr \"Benchmark Basket Size\"\n"
        "    }"
    )

    box_table = doc.add_table(rows=1, cols=1)
    box_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    c = box_table.rows[0].cells[0]
    set_cell_background(c, 'EDF2F7')
    set_cell_margins(c, top=80, bottom=80, left=120, right=120)
    c.width = Inches(6.5)
    cp = c.paragraphs[0]
    cp.paragraph_format.space_after = Pt(0)
    cr = cp.add_run(mermaid_code)
    cr.font.name = 'Consolas'
    cr.font.size = Pt(8.0)
    cr.font.color.rgb = COLOR_TEXT

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # Section 3: Plain-English Relationship Statements
    add_custom_heading("3. Plain-English Relationship Formulations", level=1)
    doc.add_paragraph(
        "In compliance with structural data modeling standards, every relationship is articulated in unambiguous natural language:"
    )

    rel_statements = [
        ("Relationship 1 (Stores to Orders): ", "One retail store operates and hosts zero or many customer checkout orders (1 to 0..N). Conversely, each checkout order is transacted at exactly one retail store."),
        ("Relationship 2 (Orders to Order Items): ", "One customer checkout order contains one or many item line records (1 to 1..N). Conversely, each item line record belongs to exactly one customer checkout order."),
        ("Relationship 3 (Products to Order Items): ", "One product catalog SKU is referenced in zero or many item line records (1 to 0..N). Conversely, each item line record refers to exactly one product catalog SKU."),
        ("Relationship 4 (Stores to Sales Targets): ", "One retail store is assigned zero or many monthly revenue target quotas (1 to 0..N). Conversely, each monthly target quota is assigned to exactly one retail store.")
    ]
    for tag, stmt in rel_statements:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.2)
        p.paragraph_format.space_after = Pt(3)
        r_b = p.add_run(tag)
        r_b.bold = True
        r_b.font.color.rgb = COLOR_PRIMARY
        r_t = p.add_run(stmt)
        r_t.font.color.rgb = COLOR_TEXT

    # Section 4: Business Rules Enforced by the Design
    add_custom_heading("4. Business Rules Enforced by the Relational Schema", level=1)
    doc.add_paragraph(
        "The relational schema acts as the foundational governance layer, enforcing core enterprise rules at the database engine level:"
    )

    schema_rules = [
        ("Rule 1: Referential Integrity & Deletion Protection: ", "Foreign key constraints are enforced unconditionally via PRAGMA foreign_keys = ON;. Master records in dim_stores and dim_products have ON DELETE RESTRICT, prohibiting deletion while referenced by historical transactions."),
        ("Rule 2: Cascade Deletion for Dependent Child Facts: ", "The relationship between fact_orders and fact_order_items enforces ON DELETE CASCADE, guaranteeing that if an order header is purged during administrative testing, all dependent line items are automatically deleted, preventing orphaned fact records."),
        ("Rule 3: Financial Non-Negative & Positive Boundaries: ", "Quantity must be >= 1; unit_cost_price_inr must be > 0.0; and unit_selling_price_inr must be >= unit_cost_price_inr, ensuring zero negative margin anomalies at catalog entry."),
        ("Rule 4: Validated Markdown Ceiling: ", "Promotional discount percentages are strictly constrained: 0.0 <= discount_pct <= 50.0 via database CHECK constraints, guarding against erroneous markdown entries."),
        ("Rule 5: Temporal and Categorical Domain Validity: ", "CHECK constraints restrict regions to ('North', 'South', 'West'), store tiers to ('Urban', 'Suburban', 'Rural'), payment methods to ('UPI', 'Credit Card', 'Debit Card', 'Cash'), and order hours between 0 and 23.")
    ]
    for tag, stmt in schema_rules:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.2)
        p.paragraph_format.space_after = Pt(3)
        r_b = p.add_run(tag)
        r_b.bold = True
        r_b.font.color.rgb = COLOR_SECONDARY
        r_t = p.add_run(stmt)
        r_t.font.color.rgb = COLOR_TEXT

    # Section 5: Design Problems Identified & Solutions Implemented
    add_custom_heading("5. Architectural Design Flaws Identified & Engineered Solutions", level=1)
    doc.add_paragraph(
        "During iterative architectural reviews, four significant design traps were identified and resolved:"
    )

    flaw_table = doc.add_table(rows=5, cols=4)
    flaw_headers = ["Identified Design Trap / Flaw", "Potential Operational Risk", "Engineered Architectural Fix", "Resulting Structural Benefit"]
    for i, h in enumerate(flaw_headers):
        flaw_table.rows[0].cells[i].paragraphs[0].text = h

    flaws_data = [
        [
            "Direct Many-to-Many (M:N) link between Orders and Products",
            "Inability to record line item quantities, line discounts, or distinct product prices within an order.",
            "Decomposed M:N into associative entity fact_order_items with two 1:N relationships.",
            "Delivers full line-item granularity, enabling SKU margin tracking and discount elasticity calculations (OQ-07)."
        ],
        [
            "Dynamic price lookups pointing directly to master catalog",
            "If product prices change in dim_products, historical sales totals and profit margins alter retroactively (Type 2 SCD flaw).",
            "Captured point-in-time snapshot prices (unit_cost_price_inr and unit_selling_price_inr) inside fact_order_items.",
            "Historical transactional immutability; past financial performance remains untouched by future pricing updates."
        ],
        [
            "Strict pure normalization requiring on-the-fly math across 50k rows",
            "Computing line_net, line_profit, and margin dynamically on every user interaction degrades UI response to > 3.0s.",
            "Selectively persisted calculated fields in fact_order_items and engineered view_store_sales_analytics.",
            "Sub-50ms analytical response times on Streamlit Cloud while respecting single-source calculation formulas."
        ],
        [
            "Date string parsing overhead during time-series aggregations",
            "Repeatedly parsing ISO timestamp strings into year, month, quarter, and day-type on every chart refresh causes CPU bottlenecks.",
            "Pre-extracted integer order_hour, day_of_week, and is_weekend_flag at ingestion; indexed via B-trees.",
            "Sub-second temporal heatmaps (OQ-10) and instant weekday-weekend partition queries (OQ-05)."
        ]
    ]
    for r_idx, row_vals in enumerate(flaws_data, start=1):
        for c_idx, val in enumerate(row_vals):
            flaw_table.rows[r_idx].cells[c_idx].paragraphs[0].text = val
    style_table(flaw_table, [1.5, 1.6, 1.8, 1.6], [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT])

    doc.save(filename)
    print(f"ER Diagram Document successfully created at: {filename}")

if __name__ == "__main__":
    import sys
    target_doc = sys.argv[1] if len(sys.argv) > 1 else "docs/09_ER_Diagram.docx"
    img = sys.argv[2] if len(sys.argv) > 2 else "docs/09_ER_Diagram.png"
    create_er_diagram_doc(target_doc, img)
