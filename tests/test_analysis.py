import unittest
import numpy as np
import pandas as pd

def calculate_net_sales(gross: float, returns: float, cancels: float) -> float:
    """Formula BURL-01: Net Sales = Gross Sales - Returns - Cancellations."""
    return round(gross - (returns + cancels), 2)

def calculate_gross_margin_pct(selling_price: float, cost_price: float) -> float:
    """Formula BURL-02: Margin % = ((Selling - Cost) / Selling) * 100."""
    if selling_price <= 0:
        return 0.0
    return round(((selling_price - cost_price) / selling_price) * 100.0, 2)

def calculate_atv(net_sales: float, completed_orders: int) -> float:
    """Formula BURL-03: Average Transaction Value = Total Net Sales / Count of Completed Orders."""
    if completed_orders <= 0:
        return 0.0
    return round(net_sales / completed_orders, 2)

def calculate_sales_velocity_sqft(net_sales: float, sqft: int) -> float:
    """Formula BURL-04: Sales Velocity per SqFt = Total Net Sales / Area."""
    if sqft <= 0:
        return 0.0
    return round(net_sales / sqft, 4)

def check_target_variance_alert(actual: float, target: float, threshold: float = 0.05) -> tuple[float, bool]:
    """Formula BURL-05: Target Quota Variance % = ((Actual - Target) / Target) * 100."""
    if target <= 0:
        return (0.0, False)
    var_pct = round(((actual - target) / target) * 100.0, 2)
    is_alert = abs(actual - target) / target >= threshold
    return (var_pct, is_alert)

def calculate_distribution_stats(series: list[float]) -> dict:
    """Non-parametric and parametric distribution metrics using pure NumPy."""
    arr = np.array(series, dtype=float)
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

def welchs_ttest_stat(sample1: list[float], sample2: list[float]) -> tuple[float, float]:
    """Welch's t-test statistic and Welch-Satterthwaite degrees of freedom."""
    s1 = np.array(sample1, dtype=float)
    s2 = np.array(sample2, dtype=float)
    n1, n2 = len(s1), len(s2)
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


class TestKPIAndAnalysisFormulas(unittest.TestCase):
    """TC-T-07 to TC-T-12: Business Metric and Statistical Formula Tests."""

    def test_tc_t_07_net_sales_calculation(self):
        """TC-T-07: Verify BURL-01 Net Sales formula on known test vectors."""
        self.assertEqual(calculate_net_sales(10000.0, 500.0, 300.0), 9200.0)
        self.assertEqual(calculate_net_sales(5000.0, 0.0, 0.0), 5000.0)
        self.assertEqual(calculate_net_sales(2000.0, 2000.0, 0.0), 0.0)

    def test_tc_t_08_gross_margin_calculation(self):
        """TC-T-08: Verify BURL-02 Gross Margin % formula on known test vectors."""
        self.assertEqual(calculate_gross_margin_pct(100.0, 60.0), 40.0)
        self.assertEqual(calculate_gross_margin_pct(2500.0, 1400.0), 44.0)
        self.assertEqual(calculate_gross_margin_pct(0.0, 50.0), 0.0)

    def test_tc_t_09_atv_and_sales_velocity(self):
        """TC-T-09: Verify BURL-03 ATV and BURL-04 Velocity/SqFt formulas."""
        self.assertEqual(calculate_atv(100000.0, 10), 10000.0)
        self.assertEqual(calculate_atv(5000.0, 0), 0.0)
        self.assertEqual(calculate_sales_velocity_sqft(500000.0, 5000), 100.0)
        self.assertEqual(calculate_sales_velocity_sqft(75000.0, 1500), 50.0)

    def test_tc_t_10_target_quota_variance_alert(self):
        """TC-T-10: Verify BURL-05 +/- 5% variance alert triggering."""
        # Underperformance > 5% (Alert)
        var_pct, alert = check_target_variance_alert(92000.0, 100000.0)
        self.assertEqual(var_pct, -8.0)
        self.assertTrue(alert)

        # Slight variance within 5% (No Alert)
        var_pct, alert = check_target_variance_alert(102000.0, 100000.0)
        self.assertEqual(var_pct, 2.0)
        self.assertFalse(alert)

        # Overperformance > 5% (Alert/Flag)
        var_pct, alert = check_target_variance_alert(108000.0, 100000.0)
        self.assertEqual(var_pct, 8.0)
        self.assertTrue(alert)

    def test_tc_t_11_statistical_distribution_metrics(self):
        """TC-T-11: Verify non-parametric metrics (IQR, Median, Skewness)."""
        data = [1000, 1200, 1500, 1800, 2000, 2200, 2500, 3000, 50000] # High outlier
        metrics = calculate_distribution_stats(data)
        self.assertEqual(metrics["median"], 2000.0)
        self.assertEqual(metrics["iqr"], 1000.0)
        self.assertGreater(metrics["skewness"], 1.5) # Right-skewed distribution

    def test_tc_t_12_welchs_ttest_weekend_split(self):
        """TC-T-12: Verify Welch's t-test calculation for unequal variances."""
        weekday_sales = [1200, 1400, 1100, 1350, 1500, 1250, 1300, 1450, 1150, 1400]
        weekend_sales = [2500, 2800, 2400, 3100, 2900, 2700, 3200, 2600, 3000, 2850]
        t_stat, df = welchs_ttest_stat(weekday_sales, weekend_sales)
        self.assertLess(t_stat, -10.0) # Statistically massive difference between weekday & weekend
        self.assertGreater(df, 10.0) # Welch-Satterthwaite df is ~13.44
        self.assertAlmostEqual(df, 13.44, places=1)

if __name__ == "__main__":
    unittest.main()
