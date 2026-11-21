import pandas as pd
import numpy as np

def clean_sales_data(df: pd.DataFrame) -> tuple[pd.DataFrame, dict]:
    """
    Validates and cleans analytical retail sales data according to enterprise domain invariants.
    Returns:
        (cleaned_df, audit_log_dict)
    """
    audit_log = {
        "initial_rows": len(df),
        "missing_imputed": 0,
        "quantity_corrections": 0,
        "discount_clamps": 0,
        "price_anomalies_resolved": 0,
        "final_rows": len(df)
    }
    
    cleaned = df.copy()

    # 1. Missing Value Imputation
    if "net_amount_inr" in cleaned.columns and cleaned["net_amount_inr"].isnull().any():
        null_count = cleaned["net_amount_inr"].isnull().sum()
        median_val = cleaned["net_amount_inr"].median()
        cleaned["net_amount_inr"] = cleaned["net_amount_inr"].fillna(median_val)
        audit_log["missing_imputed"] += int(null_count)

    if "payment_method" in cleaned.columns and cleaned["payment_method"].isnull().any():
        mode_val = cleaned["payment_method"].mode()[0] if not cleaned["payment_method"].mode().empty else "UPI"
        cleaned["payment_method"] = cleaned["payment_method"].fillna(mode_val)
        audit_log["missing_imputed"] += 1

    # 2. Invariant: Quantity >= 1
    if "quantity" in cleaned.columns:
        invalid_qty = cleaned["quantity"] < 1
        if invalid_qty.any():
            audit_log["quantity_corrections"] = int(invalid_qty.sum())
            cleaned.loc[invalid_qty, "quantity"] = 1

    # 3. Invariant: Discount between 0% and 50%
    if "discount_pct" in cleaned.columns:
        invalid_disc = (cleaned["discount_pct"] < 0.0) | (cleaned["discount_pct"] > 50.0)
        if invalid_disc.any():
            audit_log["discount_clamps"] = int(invalid_disc.sum())
            cleaned["discount_pct"] = cleaned["discount_pct"].clip(lower=0.0, upper=50.0)

    # 4. Invariant: Unit Selling Price >= Unit Cost Price
    if "unit_selling_price_inr" in cleaned.columns and "unit_cost_price_inr" in cleaned.columns:
        underpriced = cleaned["unit_selling_price_inr"] < cleaned["unit_cost_price_inr"]
        if underpriced.any():
            audit_log["price_anomalies_resolved"] = int(underpriced.sum())
            cleaned.loc[underpriced, "unit_selling_price_inr"] = cleaned.loc[underpriced, "unit_cost_price_inr"]

    audit_log["final_rows"] = len(cleaned)
    return cleaned, audit_log

def reconcile_orders_and_items(orders_df: pd.DataFrame, items_df: pd.DataFrame) -> dict:
    """
    Performs double-entry reconciliation between fact_orders and fact_order_items.
    """
    total_orders_gross = float(orders_df["gross_amount_inr"].sum())
    total_items_gross = float(items_df["line_gross_amount_inr"].sum())
    variance = round(abs(total_orders_gross - total_items_gross), 2)
    
    return {
        "orders_count": len(orders_df),
        "items_count": len(items_df),
        "total_orders_gross_inr": total_orders_gross,
        "total_items_gross_inr": total_items_gross,
        "variance_inr": variance,
        "is_reconciled": variance == 0.0
    }
