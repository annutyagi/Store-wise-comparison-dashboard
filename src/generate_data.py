import os
import random
import sqlite3
import datetime
import numpy as np
import pandas as pd
import openpyxl

def generate_synthetic_dataset():
    # Set deterministic random seed
    random.seed(42)
    np.random.seed(42)
    
    os.makedirs("data", exist_ok=True)
    os.makedirs("src", exist_ok=True)
    
    # ====================================================
    # 1. GENERATE dim_stores (12 Fictional Indian Stores)
    # ====================================================
    stores_data = [
        {"store_id": "STR-001", "store_name": "Metro Central Hub - Noida", "region_id": "REG-N", "region_name": "North", "city": "Noida", "state": "Uttar Pradesh", "store_tier": "Urban", "store_size_sqft": 7500, "opening_date": "2020-03-15", "is_active": 1},
        {"store_id": "STR-002", "store_name": "Hazratganj Heritage Store - Lucknow", "region_id": "REG-N", "region_name": "North", "city": "Lucknow", "state": "Uttar Pradesh", "store_tier": "Urban", "store_size_sqft": 6200, "opening_date": "2020-08-20", "is_active": 1},
        {"store_id": "STR-003", "store_name": "Connaught Galleria - New Delhi", "region_id": "REG-N", "region_name": "North", "city": "New Delhi", "state": "Delhi", "store_tier": "Urban", "store_size_sqft": 8200, "opening_date": "2019-11-10", "is_active": 1},
        {"store_id": "STR-004", "store_name": "Cyber Park Mart - Gurugram", "region_id": "REG-N", "region_name": "North", "city": "Gurugram", "state": "Haryana", "store_tier": "Urban", "store_size_sqft": 6800, "opening_date": "2021-02-14", "is_active": 1},
        {"store_id": "STR-005", "store_name": "Indiranagar High Street - Bengaluru", "region_id": "REG-S", "region_name": "South", "city": "Bengaluru", "state": "Karnataka", "store_tier": "Urban", "store_size_sqft": 7100, "opening_date": "2020-01-25", "is_active": 1},
        {"store_id": "STR-006", "store_name": "Bandra Trend Outlet - Mumbai", "region_id": "REG-W", "region_name": "West", "city": "Mumbai", "state": "Maharashtra", "store_tier": "Urban", "store_size_sqft": 7800, "opening_date": "2019-09-05", "is_active": 1},
        {"store_id": "STR-007", "store_name": "Koregaon Park Market - Pune", "region_id": "REG-W", "region_name": "West", "city": "Pune", "state": "Maharashtra", "store_tier": "Suburban", "store_size_sqft": 4500, "opening_date": "2021-06-18", "is_active": 1},
        {"store_id": "STR-008", "store_name": "SG Highway Express - Ahmedabad", "region_id": "REG-W", "region_name": "West", "city": "Ahmedabad", "state": "Gujarat", "store_tier": "Suburban", "store_size_sqft": 4800, "opening_date": "2021-09-12", "is_active": 1},
        {"store_id": "STR-009", "store_name": "Pink City Bazaar - Jaipur", "region_id": "REG-N", "region_name": "North", "city": "Jaipur", "state": "Rajasthan", "store_tier": "Suburban", "store_size_sqft": 4200, "opening_date": "2022-04-10", "is_active": 1},
        {"store_id": "STR-010", "store_name": "Alwar bypass Mart - Alwar", "region_id": "REG-N", "region_name": "North", "city": "Alwar", "state": "Rajasthan", "store_tier": "Rural", "store_size_sqft": 2400, "opening_date": "2022-11-05", "is_active": 1},
        {"store_id": "STR-011", "store_name": "Baramati Community Hub - Baramati", "region_id": "REG-W", "region_name": "West", "city": "Baramati", "state": "Maharashtra", "store_tier": "Rural", "store_size_sqft": 2200, "opening_date": "2023-02-18", "is_active": 1},
        {"store_id": "STR-012", "store_name": "Mandya Farmers Outlet - Mandya", "region_id": "REG-S", "region_name": "South", "city": "Mandya", "state": "Karnataka", "store_tier": "Rural", "store_size_sqft": 1900, "opening_date": "2023-05-22", "is_active": 1}
    ]
    df_stores = pd.DataFrame(stores_data)

    # ====================================================
    # 2. GENERATE dim_products (60 Retail SKUs)
    # ====================================================
    categories_spec = {
        "Electronics": {
            "id": "CAT-01",
            "items": [
                ("Noise-Cancelling Wireless Headphones", "Audio", 3200.0, 5499.0, 41.8),
                ("True Wireless Earbuds Pro", "Audio", 1400.0, 2499.0, 44.0),
                ("Smart Fitness Watch Ultra", "Wearables", 2100.0, 3999.0, 47.5),
                ("Mechanical Gaming Keyboard", "Peripherals", 1800.0, 3299.0, 45.4),
                ("Ergonomic Wireless Mouse", "Peripherals", 650.0, 1199.0, 45.8),
                ("Fast Charging Power Bank 20000mAh", "Accessories", 950.0, 1799.0, 47.2),
                ("4K Ultra-HD Action Camera", "Cameras", 5200.0, 8999.0, 42.2),
                ("High-Speed USB-C Hub 7-in-1", "Accessories", 1100.0, 1999.0, 45.0),
                ("Portable Bluetooth Speaker 20W", "Audio", 1600.0, 2899.0, 44.8),
                ("Smart WiFi LED Desk Lamp", "Smart Home", 850.0, 1599.0, 46.8),
                ("Dual Band WiFi 6 Router", "Networking", 2400.0, 4299.0, 44.2),
                ("1080p Full HD Webcam with Mic", "Peripherals", 1250.0, 2299.0, 45.6)
            ]
        },
        "Apparel": {
            "id": "CAT-02",
            "items": [
                ("Slim Fit Cotton Formal Shirt", "Menswear", 650.0, 1499.0, 56.6),
                ("Classic Indigo Denim Jeans", "Menswear", 850.0, 1999.0, 57.5),
                ("Embroidered Anarkali Kurti", "Womenswear", 750.0, 1799.0, 58.3),
                ("Handloom Cotton Saree", "Ethnic", 1200.0, 2899.0, 58.6),
                ("Breathable Running Track Pants", "Activewear", 550.0, 1299.0, 57.7),
                ("Dry-Fit Athletic Gym T-Shirt", "Activewear", 380.0, 899.0, 57.7),
                ("Chinos Casual Stretch Trousers", "Menswear", 780.0, 1849.0, 57.8),
                ("Floral Printed Summer Maxi Dress", "Womenswear", 920.0, 2199.0, 58.2),
                ("Pure Wool Winter Pullover", "Winterwear", 1100.0, 2599.0, 57.7),
                ("Kids Organic Cotton Casual Set", "Kidswear", 420.0, 999.0, 58.0),
                ("Structured Blazer Slim Fit", "Formalwear", 2100.0, 4999.0, 58.0),
                ("Festive Silk Blend Kurta Pajama", "Ethnic", 1300.0, 2999.0, 56.7)
            ]
        },
        "Groceries": {
            "id": "CAT-03",
            "items": [
                ("Organic Royal Basmati Rice 5kg", "Staples", 480.0, 649.0, 26.0),
                ("Cold Pressed Mustard Oil 1L", "Cooking Oil", 140.0, 195.0, 28.2),
                ("Desi Cow Ghee A2 500ml", "Dairy", 420.0, 580.0, 27.6),
                ("Unpolished Toor Dal 1kg", "Staples", 115.0, 155.0, 25.8),
                ("Premium California Almonds 500g", "Dry Fruits", 380.0, 520.0, 26.9),
                ("Roasted Cashews Salted 250g", "Dry Fruits", 220.0, 310.0, 29.0),
                ("Organic Wild Raw Honey 500g", "Sweeteners", 240.0, 340.0, 29.4),
                ("Assam Orthodox Tea Leaf 500g", "Beverages", 190.0, 275.0, 30.9),
                ("Filter Coffee Ground Roasted 250g", "Beverages", 145.0, 210.0, 31.0),
                ("Multigrain Atta Flour 5kg", "Staples", 220.0, 295.0, 25.4),
                ("Organic Green Cardamom 100g", "Spices", 280.0, 395.0, 29.1),
                ("Dark Chocolate Cocoa Bar 100g", "Snacks", 95.0, 145.0, 34.5)
            ]
        },
        "Home Goods": {
            "id": "CAT-04",
            "items": [
                ("Double Bed Comforter All-Season", "Bedding", 1250.0, 2499.0, 50.0),
                ("Egyptian Cotton 300TC Bed Sheet", "Bedding", 850.0, 1799.0, 52.8),
                ("Non-Stick Tri-Ply Kadai 2.5L", "Cookware", 980.0, 1999.0, 51.0),
                ("Stainless Steel Knife Set 5-Piece", "Kitchenware", 540.0, 1149.0, 53.0),
                ("Automatic Electric Kettle 1.8L", "Small Appliances", 620.0, 1299.0, 52.3),
                ("Memory Foam Ergonomic Pillow", "Bedding", 680.0, 1449.0, 53.1),
                ("Ceramic Dinnerware Set 18-Piece", "Tableware", 1650.0, 3499.0, 52.8),
                ("Microfiber Bath Towel Set 4-Piece", "Bath", 420.0, 899.0, 53.3),
                ("Aroma Essential Oil Diffuser 500ml", "Decor", 750.0, 1599.0, 53.1),
                ("Insulated Stainless Water Bottle 1L", "Kitchenware", 380.0, 799.0, 52.4),
                ("Cast Iron Skillet Pre-Seasoned 10-inch", "Cookware", 720.0, 1499.0, 52.0),
                ("Storage Basket Fabric Foldable Set", "Storage", 310.0, 699.0, 55.7)
            ]
        },
        "Footwear": {
            "id": "CAT-05",
            "items": [
                ("Mesh Breathable Running Shoes", "Sports", 950.0, 2199.0, 56.8),
                ("Genuine Leather Formal Derby Shoes", "Formal", 1450.0, 3299.0, 56.0),
                ("Cushioned Casual Walking Loafers", "Casual", 780.0, 1799.0, 56.6),
                ("Lightweight EVA Sliders Comfort", "Casual", 280.0, 649.0, 56.9),
                ("Orthopedic Support Daily Sandals", "Casual", 520.0, 1199.0, 56.6),
                ("Ankle High Trekking Boots", "Outdoor", 1650.0, 3799.0, 56.6),
                ("Classic White Canvas Sneakers", "Casual", 620.0, 1399.0, 55.7),
                ("Ethnic Embroidered Mojaris Juttis", "Ethnic", 490.0, 1099.0, 55.4),
                ("Slip-On Memory Foam Work Shoes", "Formal", 1100.0, 2499.0, 56.0),
                ("Waterproof Monsoon Rain Sandals", "Monsoon", 340.0, 799.0, 57.4),
                ("High Grip Court Badminton Shoes", "Sports", 1250.0, 2799.0, 55.3),
                ("Soft Fur Lined Winter Slippers", "Indoor", 310.0, 699.0, 55.7)
            ]
        }
    }

    products_list = []
    prod_counter = 1
    for cat_name, cat_info in categories_spec.items():
        cat_id = cat_info["id"]
        for p_name, sub_cat, cost, price, target_margin in cat_info["items"]:
            prod_id = f"PRD-{prod_counter:03d}"
            sku = f"SKU-{cat_name[:2].upper()}-{prod_counter:03d}"
            products_list.append({
                "product_id": prod_id,
                "product_name": p_name,
                "category_id": cat_id,
                "category_name": cat_name,
                "sub_category": sub_cat,
                "cost_price_inr": cost,
                "base_selling_price_inr": price,
                "target_margin_pct": target_margin,
                "sku_code": sku,
                "is_active": 1
            })
            prod_counter += 1
    df_products = pd.DataFrame(products_list)

    # ====================================================
    # 3. GENERATE fact_orders & fact_order_items (350 Orders, 900+ Items)
    # With 5 Realistic Business Patterns
    # ====================================================
    orders_list = []
    order_items_list = []
    item_id_counter = 1
    
    start_date = datetime.date(2024, 1, 1)
    end_date = datetime.date(2024, 12, 31)
    date_range_days = (end_date - start_date).days

    # Order weights by month (Pattern 1: Q4 Festival Spike)
    # Months 1-9: normal (weights 1.0); Months 10-12: festive peak (weights 1.5 - 1.8)
    month_weights = {
        1: 0.9, 2: 0.85, 3: 0.95, 4: 1.0, 5: 1.05, 6: 1.0,
        7: 1.1, 8: 1.15, 9: 1.1, 10: 1.65, 11: 1.75, 12: 1.55
    }

    # Store popularity weights (Urban stores have higher volume than Rural)
    store_volume_weights = {
        "STR-001": 1.4, "STR-002": 1.2, "STR-003": 1.5, "STR-004": 1.3,
        "STR-005": 1.35, "STR-006": 1.45, "STR-007": 1.0, "STR-008": 0.95,
        "STR-009": 0.9, "STR-010": 0.55, "STR-011": 0.5, "STR-012": 0.45
    }
    store_ids = list(store_volume_weights.keys())
    store_probs = np.array(list(store_volume_weights.values()))
    store_probs = store_probs / store_probs.sum()

    total_target_orders = 350
    for order_idx in range(1, total_target_orders + 1):
        order_id = f"ORD-2024-{order_idx:04d}"
        
        # Pick store
        assigned_store_id = np.random.choice(store_ids, p=store_probs)
        store_meta = df_stores[df_stores["store_id"] == assigned_store_id].iloc[0]
        store_tier = store_meta["store_tier"]
        
        # Pick date with Q4 seasonal weighting (Pattern 1)
        while True:
            rand_day_offset = random.randint(0, date_range_days)
            dt = start_date + datetime.timedelta(days=rand_day_offset)
            m_weight = month_weights[dt.month]
            if random.random() < (m_weight / 2.0):
                order_date = dt
                break

        # Day of week & Weekend flag (Pattern 2: Weekend Spike)
        dow_names = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
        day_of_week = dow_names[order_date.weekday()]
        is_weekend = 1 if order_date.weekday() >= 5 else 0

        # Pattern 5: Hourly Distribution (Evening Rush Window 17:00 - 21:00)
        # Probability distribution over 24 hours
        hour_weights = [
            0.01, 0.01, 0.01, 0.01, 0.01, 0.01, # 00 - 05
            0.02, 0.03, 0.04, 0.06, 0.08, 0.09, # 06 - 11
            0.08, 0.07, 0.06, 0.08, 0.10, 0.15, # 12 - 17
            0.18, 0.16, 0.12, 0.08, 0.04, 0.02  # 18 - 23
        ]
        hour_probs = np.array(hour_weights) / sum(hour_weights)
        order_hour = int(np.random.choice(range(24), p=hour_probs))
        order_minute = random.randint(0, 59)
        order_second = random.randint(0, 59)
        order_time_str = f"{order_hour:02d}:{order_minute:02d}:{order_second:02d}"
        order_datetime_str = f"{order_date.isoformat()} {order_time_str}"

        # Payment Tender Distribution (Pattern: UPI & Credit Card dominant in Urban, Cash in Rural)
        if store_tier == "Urban":
            tender_p = [0.50, 0.32, 0.10, 0.08] # UPI, Card, Debit, Cash
        elif store_tier == "Suburban":
            tender_p = [0.45, 0.22, 0.18, 0.15]
        else: # Rural
            tender_p = [0.30, 0.10, 0.15, 0.45]
        payment_method = np.random.choice(["UPI", "Credit Card", "Debit Card", "Cash"], p=tender_p)

        # Order Status (92% Completed, 5% Returned, 3% Cancelled)
        order_status = np.random.choice(["Completed", "Returned", "Cancelled"], p=[0.92, 0.05, 0.03])

        # Generate Line Items for this Order
        # Pattern 2 & Q4: Weekends and Q4 have higher items count (2 to 5 items) vs Weekday (1 to 3 items)
        if is_weekend or order_date.month in [10, 11, 12]:
            num_items = random.choices([1, 2, 3, 4, 5], weights=[0.10, 0.35, 0.30, 0.15, 0.10])[0]
        else:
            num_items = random.choices([1, 2, 3, 4], weights=[0.40, 0.40, 0.15, 0.05])[0]

        # Pattern 3: Demographic Tier Category Affinity
        # Urban prefers Electronics (35%) & Apparel (30%); Rural prefers Groceries (45%) & Home Goods (30%)
        if store_tier == "Urban":
            cat_weights = {"Electronics": 0.35, "Apparel": 0.30, "Footwear": 0.15, "Home Goods": 0.12, "Groceries": 0.08}
        elif store_tier == "Suburban":
            cat_weights = {"Electronics": 0.22, "Apparel": 0.25, "Footwear": 0.18, "Home Goods": 0.20, "Groceries": 0.15}
        else: # Rural
            cat_weights = {"Groceries": 0.45, "Home Goods": 0.30, "Apparel": 0.12, "Footwear": 0.08, "Electronics": 0.05}

        chosen_cats = list(cat_weights.keys())
        chosen_cat_probs = np.array(list(cat_weights.values())) / sum(cat_weights.values())

        order_gross = 0.0
        order_discount = 0.0
        
        # Track items in this order
        selected_prods = []
        for _ in range(num_items):
            cat_pick = np.random.choice(chosen_cats, p=chosen_cat_probs)
            prod_subset = df_products[df_products["category_name"] == cat_pick]
            prod_row = prod_subset.sample(n=1).iloc[0]
            selected_prods.append(prod_row)

        for prod in selected_prods:
            qty = random.choices([1, 2, 3, 4], weights=[0.70, 0.20, 0.07, 0.03])[0]
            cost_price = prod["cost_price_inr"]
            selling_price = prod["base_selling_price_inr"]

            # Pattern 4: Discount Elasticity (Promotions apply discount 0%, 5%, 10%, 15%, 20%, 30%)
            # High discounts (>= 20%) increase quantity but squeeze line profit margin
            discount_choice = random.choices([0.0, 5.0, 10.0, 15.0, 25.0, 35.0], weights=[0.45, 0.20, 0.15, 0.10, 0.07, 0.03])[0]
            if discount_choice >= 25.0:
                qty = random.choices([2, 3, 4, 5], weights=[0.35, 0.35, 0.20, 0.10])[0]

            line_gross = round(qty * selling_price, 2)
            line_disc = round(line_gross * (discount_choice / 100.0), 2)
            line_net = round(line_gross - line_disc, 2)
            line_cogs = round(qty * cost_price, 2)
            line_profit = round(line_net - line_cogs, 2)
            line_margin = round((line_profit / line_net) * 100.0, 2) if line_net > 0 else 0.0

            order_items_list.append({
                "order_item_id": item_id_counter,
                "order_id": order_id,
                "product_id": prod["product_id"],
                "quantity": qty,
                "unit_cost_price_inr": cost_price,
                "unit_selling_price_inr": selling_price,
                "discount_pct": discount_choice,
                "line_gross_amount_inr": line_gross,
                "line_discount_amount_inr": line_disc,
                "line_net_amount_inr": line_net,
                "line_cogs_inr": line_cogs,
                "line_profit_inr": line_profit,
                "line_margin_pct": line_margin
            })
            item_id_counter += 1
            order_gross += line_gross
            order_discount += line_disc

        # Calculate final order financials
        order_gross = round(order_gross, 2)
        order_discount = round(order_discount, 2)
        if order_status == "Returned":
            order_return = round(order_gross - order_discount, 2)
            order_net = 0.0
        elif order_status == "Cancelled":
            order_return = round(order_gross - order_discount, 2)
            order_net = 0.0
        else:
            order_return = 0.0
            order_net = round(order_gross - order_discount, 2)

        orders_list.append({
            "order_id": order_id,
            "store_id": assigned_store_id,
            "order_datetime": order_datetime_str,
            "order_date": order_date.isoformat(),
            "order_time": order_time_str,
            "order_hour": order_hour,
            "day_of_week": day_of_week,
            "is_weekend_flag": is_weekend,
            "payment_method": payment_method,
            "order_status": order_status,
            "gross_amount_inr": order_gross,
            "discount_amount_inr": order_discount,
            "return_amount_inr": order_return,
            "net_amount_inr": order_net
        })

    df_orders = pd.DataFrame(orders_list)
    df_order_items = pd.DataFrame(order_items_list)

    # Controlled Data Quality Quirks (for Step 15 cleaning & test validations)
    # Add 2 intentional outlier corporate orders with high basket value
    outlier_order_1 = df_orders.iloc[10].copy()
    outlier_order_1["order_id"] = "ORD-2024-9001"
    outlier_order_1["gross_amount_inr"] = 145000.0
    outlier_order_1["net_amount_inr"] = 135000.0
    outlier_order_1["discount_amount_inr"] = 10000.0
    
    outlier_order_2 = df_orders.iloc[20].copy()
    outlier_order_2["order_id"] = "ORD-2024-9002"
    outlier_order_2["gross_amount_inr"] = 182000.0
    outlier_order_2["net_amount_inr"] = 170000.0
    outlier_order_2["discount_amount_inr"] = 12000.0

    df_orders = pd.concat([df_orders, pd.DataFrame([outlier_order_1, outlier_order_2])], ignore_index=True)

    # ====================================================
    # 4. GENERATE dim_sales_targets (144 Records: 12 Stores x 12 Months)
    # ====================================================
    targets_list = []
    tgt_counter = 1
    # Monthly base revenue targets calibrated by store tier & size
    base_tier_targets = {"Urban": 1100000.0, "Suburban": 750000.0, "Rural": 380000.0}
    base_tier_orders = {"Urban": 550, "Suburban": 380, "Rural": 220}

    for store in stores_data:
        s_id = store["store_id"]
        tier = store["store_tier"]
        base_rev = base_tier_targets[tier]
        base_ord = base_tier_orders[tier]
        
        for m in range(1, 13):
            month_str = f"2024-{m:02d}"
            # Seasonality multiplier on target
            s_mult = month_weights[m]
            target_rev = round(base_rev * s_mult * (1.0 + random.uniform(-0.04, 0.05)), 2)
            target_ord = int(base_ord * s_mult * (1.0 + random.uniform(-0.03, 0.04)))
            target_atv = round(target_rev / target_ord, 2)
            
            tgt_id = f"TGT-{s_id.replace('-', '')}-{m:02d}"
            targets_list.append({
                "target_id": tgt_id,
                "store_id": s_id,
                "target_month": month_str,
                "target_revenue_inr": target_rev,
                "target_orders_count": target_ord,
                "target_atv_inr": target_atv
            })
            tgt_counter += 1

    df_targets = pd.DataFrame(targets_list)

    # ====================================================
    # 5. CREATE SQLITE DATABASE (data/app.db) & POPULATE
    # ====================================================
    db_path = "data/app.db"
    if os.path.exists(db_path):
        os.remove(db_path)

    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    cur.execute("PRAGMA foreign_keys = ON;")

    # Execute DDL
    ddl_script = """
    CREATE TABLE dim_stores (
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
    );

    CREATE TABLE dim_products (
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
    );

    CREATE TABLE fact_orders (
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
    );

    CREATE TABLE fact_order_items (
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
    );

    CREATE TABLE dim_sales_targets (
        target_id TEXT PRIMARY KEY,
        store_id TEXT NOT NULL,
        target_month TEXT NOT NULL,
        target_revenue_inr REAL NOT NULL CHECK(target_revenue_inr > 0.0),
        target_orders_count INTEGER NOT NULL CHECK(target_orders_count > 0),
        target_atv_inr REAL NOT NULL,
        FOREIGN KEY (store_id) REFERENCES dim_stores(store_id) ON DELETE RESTRICT ON UPDATE CASCADE
    );

    CREATE INDEX idx_orders_date_store ON fact_orders(order_date, store_id);
    CREATE INDEX idx_orders_hour ON fact_orders(order_hour);
    CREATE INDEX idx_order_items_product ON fact_order_items(product_id);
    CREATE INDEX idx_products_category ON dim_products(category_name);

    CREATE VIEW view_store_sales_analytics AS
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
    JOIN dim_products dp ON foi.product_id = dp.product_id;
    """
    cur.executescript(ddl_script)

    # Insert Data into Tables
    df_stores.to_sql("dim_stores", conn, if_exists="append", index=False)
    df_products.to_sql("dim_products", conn, if_exists="append", index=False)
    df_orders.to_sql("fact_orders", conn, if_exists="append", index=False)
    df_order_items.to_sql("fact_order_items", conn, if_exists="append", index=False)
    df_targets.to_sql("dim_sales_targets", conn, if_exists="append", index=False)
    conn.commit()

    # Query View
    df_view = pd.read_sql_query("SELECT * FROM view_store_sales_analytics", conn)
    conn.close()

    # ====================================================
    # 6. EXPORT CSVs & EXCEL WORKBOOK
    # ====================================================
    df_stores.to_csv("data/dim_stores.csv", index=False)
    df_products.to_csv("data/dim_products.csv", index=False)
    df_orders.to_csv("data/fact_orders.csv", index=False)
    df_order_items.to_csv("data/fact_order_items.csv", index=False)
    df_targets.to_csv("data/dim_sales_targets.csv", index=False)
    df_view.to_csv("data/view_store_sales_analytics.csv", index=False)

    excel_path = "data/store_wise_sales_data.xlsx"
    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        df_stores.to_excel(writer, sheet_name="dim_stores", index=False)
        df_products.to_excel(writer, sheet_name="dim_products", index=False)
        df_orders.to_excel(writer, sheet_name="fact_orders", index=False)
        df_order_items.to_excel(writer, sheet_name="fact_order_items", index=False)
        df_targets.to_excel(writer, sheet_name="dim_sales_targets", index=False)
        df_view.to_excel(writer, sheet_name="view_analytics", index=False)

    # Also save schema.sql
    with open("data/schema.sql", "w", encoding="utf-8") as f:
        f.write(ddl_script)

    print("Synthetic dataset successfully generated:")
    print(f"  • Stores rows: {len(df_stores)}")
    print(f"  • Products rows: {len(df_products)}")
    print(f"  • Orders rows: {len(df_orders)}")
    print(f"  • Order Items rows: {len(df_order_items)}")
    print(f"  • Sales Targets rows: {len(df_targets)}")
    print(f"  • Joined View rows: {len(df_view)} (Columns: {len(df_view.columns)})")
    print(f"  • SQLite DB: {db_path}")
    print(f"  • Excel Workbook: {excel_path}")

if __name__ == "__main__":
    generate_synthetic_dataset()
