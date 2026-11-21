
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
    