import numpy as np
import pandas as pd

def calculate_net_sales(gross: float, returns: float, cancels: float) -> float:
    """Formula BURL-01: Net Sales = Gross Sales - Returns - Cancellations."""
    return round(float(gross) - (float(returns) + float(cancels)), 2)

def calculate_gross_margin_pct(selling_price: float, cost_price: float) -> float:
    """Formula BURL-02: Margin % = ((Selling - Cost) / Selling) * 100."""
    if selling_price <= 0:
        return 0.0
    return round(((float(selling_price) - float(cost_price)) / float(selling_price)) * 100.0, 2)

def calculate_atv(net_sales: float, completed_orders: int) -> float:
    """Formula BURL-03: Average Transaction Value = Total Net Sales / Count of Completed Orders."""
    if completed_orders <= 0:
        return 0.0
    return round(float(net_sales) / int(completed_orders), 2)

def calculate_sales_velocity_sqft(net_sales: float, sqft: int) -> float:
    """Formula BURL-04: Sales Velocity per SqFt = Total Net Sales / Area."""
    if sqft <= 0:
        return 0.0
    return round(float(net_sales) / int(sqft), 4)

def check_target_variance_alert(actual: float, target: float, threshold: float = 0.05) -> tuple[float, bool]:
    """Formula BURL-05: Target Quota Variance % = ((Actual - Target) / Target) * 100."""
    if target <= 0:
        return (0.0, False)
    var_pct = round(((float(actual) - float(target)) / float(target)) * 100.0, 2)
    is_alert = abs(float(actual) - float(target)) / float(target) >= threshold
    return (var_pct, is_alert)

def calculate_distribution_stats(series: list[float] | pd.Series) -> dict:
    """Non-parametric and parametric distribution metrics using pure NumPy."""
    arr = np.array(series, dtype=float)
    if len(arr) == 0:
        return {"mean": 0.0, "median": 0.0, "iqr": 0.0, "skewness": 0.0}
    
    q75, q25 = np.percentile(arr, [75, 25])
    iqr = round(float(q75 - q25), 2)
    mean_val = float(np.mean(arr))
    std_val = float(np.std(arr, ddof=1)) if len(arr) > 1 else 1.0
    if std_val > 0:
        skew_val = float(np.mean(((arr - mean_val) / std_val) ** 3))
    else:
        skew_val = 0.0

    return {
        "mean": round(mean_val, 2),
        "median": round(float(np.median(arr)), 2),
        "iqr": iqr,
        "skewness": round(skew_val, 2)
    }

def welchs_ttest_stat(sample1: list[float] | pd.Series, sample2: list[float] | pd.Series) -> tuple[float, float]:
    """Welch's t-test statistic and Welch-Satterthwaite degrees of freedom."""
    s1 = np.array(sample1, dtype=float)
    s2 = np.array(sample2, dtype=float)
    n1, n2 = len(s1), len(s2)
    if n1 <= 1 or n2 <= 1:
        return (0.0, 0.0)
    m1, m2 = np.mean(s1), np.mean(s2)
    v1 = np.var(s1, ddof=1)
    v2 = np.var(s2, ddof=1)
    se = np.sqrt(v1 / n1 + v2 / n2)
    if se == 0:
        return (0.0, 0.0)
    t_stat = (m1 - m2) / se
    df = ((v1 / n1 + v2 / n2) ** 2) / (
        ((v1 / n1) ** 2) / (n1 - 1) + ((v2 / n2) ** 2) / (n2 - 1)
    )
    return round(float(t_stat), 4), round(float(df), 2)

def compute_store_benchmarks(df: pd.DataFrame) -> pd.DataFrame:
    """Aggregates store-level metrics for Multi-Store Benchmarking (UI-02)."""
    if df.empty:
        return pd.DataFrame()
    city_col = "city" if "city" in df.columns else ("city_name" if "city_name" in df.columns else None)
    group_cols = ["store_id", "store_name"]
    if city_col:
        group_cols.append(city_col)
    for c in ["region_name", "store_tier", "store_size_sqft"]:
        if c in df.columns:
            group_cols.append(c)

    grouped = df.groupby(group_cols).agg(
        gross_sales=("line_gross_amount_inr", "sum"),
        net_sales=("line_net_amount_inr", "sum"),
        total_profit=("line_profit_inr", "sum"),
        total_orders=("order_id", "nunique"),
        avg_discount=("discount_pct", "mean")
    ).reset_index()

    grouped["margin_pct"] = (grouped["total_profit"] / grouped["net_sales"] * 100.0).round(2)
    grouped["sales_velocity_sqft"] = (grouped["net_sales"] / grouped["store_size_sqft"]).round(2)
    grouped["atv"] = (grouped["net_sales"] / grouped["total_orders"]).round(2)
    return grouped.sort_values(by="net_sales", ascending=False)

def compute_mom_velocity(df: pd.DataFrame) -> pd.DataFrame:
    """Calculates Month-over-Month sales velocity and growth rates (UI-04)."""
    if df.empty:
        return pd.DataFrame()
    df_copy = df.copy()
    if not pd.api.types.is_datetime64_any_dtype(df_copy["order_date"]):
        df_copy["order_date"] = pd.to_datetime(df_copy["order_date"])
    
    df_copy["year_month"] = df_copy["order_date"].dt.to_period("M").astype(str)
    mom = df_copy.groupby("year_month").agg(
        net_sales=("line_net_amount_inr", "sum"),
        total_orders=("order_id", "nunique")
    ).reset_index()
    mom["mom_growth_pct"] = mom["net_sales"].pct_change() * 100.0
    return mom
