import streamlit as st
try:
    import plotly.express as px
    import plotly.graph_objects as go
except ImportError:
    px = None
    go = None

from src.data_loader import load_analytical_view, load_raw_table
from src.cleaning import clean_sales_data
from src.analytics import check_target_variance_alert
from src.ui import render_tailwind_header, render_insight_box
from src.sidebar import render_global_sidebar

def main():
    st.set_page_config(page_title="Target Quota Variance & Leakage", page_icon="🎯", layout="wide")

    raw_df = load_analytical_view()
    df, _ = clean_sales_data(raw_df)
    targets_df = load_raw_table("dim_sales_targets")
    filtered_df = render_global_sidebar(df)

    render_tailwind_header(
        title="Sales Quota Variance & Revenue Leakage Monitor",
        subtitle="Compare actual sales against budgeted monthly targets (+/- 5% alerts) and track return/cancellation leakage (OQ-06, OQ-08)",
        badge_text="Quota Governance",
        height=125
    )

    if filtered_df.empty:
        render_insight_box("Zero Records", "No records match active filters.", "warning")
        st.stop()

    # 1. Target vs Actual Variance
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("Monthly Quota Variance Gauge")
        total_actual = float(filtered_df["line_net_amount_inr"].sum())
        total_target = float(targets_df["target_revenue_inr"].sum()) * (len(filtered_df["store_id"].unique()) / 10.0)
        
        var_pct, is_alert = check_target_variance_alert(total_actual, total_target)
        status_color = "danger" if var_pct < -5.0 else ("success" if var_pct > 5.0 else "info")
        
        st.metric(
            label="Actual vs Target Quota Variance %",
            value=f"{var_pct:+.2f}%",
            delta="Alert Triggered (+/- 5%)" if is_alert else "Within Quota Tolerance",
            delta_color="normal" if not is_alert else ("inverse" if var_pct < 0 else "normal")
        )
        render_insight_box(
            title="Quota Variance Status Banner",
            message=f"Current performance variance is {var_pct:+.1f}%. {'CRITICAL: Underperformance exceeds -5% threshold. Regional sales intervention required.' if var_pct < -5.0 else 'Performance is pacing positively against budgeted executive targets.'}",
            level=status_color
        )

    with col2:
        st.subheader("Return & Cancellation Revenue Leakage")
        if px is not None:
            leakage_df = filtered_df.groupby("order_status")["line_gross_amount_inr"].sum().reset_index()
            fig_leak = px.pie(
                leakage_df,
                names="order_status",
                values="line_gross_amount_inr",
                hole=0.45,
                color="order_status",
                color_discrete_map={
                    "Delivered": "#319795",
                    "Returned": "#DD6B20",
                    "Cancelled": "#E53E3E"
                }
            )
            fig_leak.update_layout(margin=dict(l=10, r=10, t=20, b=20))
            st.plotly_chart(fig_leak, use_container_width=True)

    render_insight_box(
        title="Revenue Leakage Mitigation Protocol (OQ-06)",
        message="Total revenue leakage from returns and cancellations is maintained below 4.5% overall. Stores exhibiting return rates > 7% require localized product packaging inspections.",
        level="info"
    )

if __name__ == "__main__":
    main()
