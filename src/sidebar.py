import pandas as pd
import datetime

try:
    import streamlit as st
except ImportError:
    st = None

def render_global_sidebar(df: pd.DataFrame) -> pd.DataFrame:
    """
    Renders global interactive sidebar filters synchronized with st.session_state.
    Returns the filtered analytical DataFrame.
    """
    if st is None or df.empty:
        return df

    st.sidebar.markdown(
        """
        <div style="padding: 10px 0px 15px 0px;">
            <h3 style="color: #1A365D; margin: 0; font-size: 1.15rem; font-weight: 700;">
                📊 Global Filters
            </h3>
            <p style="color: #718096; margin: 2px 0 0 0; font-size: 0.75rem;">
                Cross-dimensional retail filters
            </p>
        </div>
        """,
        unsafe_allow_html=True
    )

    # 1. Store Multi-select
    all_stores = sorted(df["store_name"].dropna().unique().tolist())
    selected_stores = st.sidebar.multiselect(
        "Select Stores:",
        options=all_stores,
        default=all_stores,
        help="Filter metrics across one or multiple retail stores."
    )

    # 2. Date Range Filter
    if "order_date" in df.columns:
        min_date = df["order_date"].min()
        max_date = df["order_date"].max()
        if isinstance(min_date, pd.Timestamp):
            min_date = min_date.date()
        if isinstance(max_date, pd.Timestamp):
            max_date = max_date.date()
            
        selected_date_range = st.sidebar.date_input(
            "Order Date Range:",
            value=(min_date, max_date),
            min_value=min_date,
            max_value=max_date,
            help="Filter transactions by date window."
        )
    else:
        selected_date_range = None

    # 3. Store Demographic Tier
    all_tiers = sorted(df["store_tier"].dropna().unique().tolist())
    selected_tiers = st.sidebar.multiselect(
        "Demographic Tier:",
        options=all_tiers,
        default=all_tiers,
        help="Urban, Suburban, or Rural retail environments."
    )

    # 4. Product Category
    all_categories = sorted(df["category_name"].dropna().unique().tolist())
    selected_categories = st.sidebar.multiselect(
        "Product Category:",
        options=all_categories,
        default=all_categories,
        help="Filter by merchandising category."
    )

    # 5. Payment Tender Method
    all_payments = sorted(df["payment_method"].dropna().unique().tolist())
    selected_payments = st.sidebar.multiselect(
        "Payment Method:",
        options=all_payments,
        default=all_payments,
        help="UPI, Credit Card, Debit Card, Cash."
    )

    # Reset Filter Button
    if st.sidebar.button("🔄 Reset All Filters", use_container_width=True):
        st.session_state.clear()
        st.rerun()

    # Apply Filtering Logic
    filtered_df = df.copy()

    if selected_stores:
        filtered_df = filtered_df[filtered_df["store_name"].isin(selected_stores)]
    else:
        filtered_df = filtered_df.iloc[0:0]

    if selected_date_range and len(selected_date_range) == 2:
        start_d, end_d = selected_date_range
        filtered_df = filtered_df[
            (filtered_df["order_date"] >= start_d) & (filtered_df["order_date"] <= end_d)
        ]

    if selected_tiers:
        filtered_df = filtered_df[filtered_df["store_tier"].isin(selected_tiers)]
    else:
        filtered_df = filtered_df.iloc[0:0]

    if selected_categories:
        filtered_df = filtered_df[filtered_df["category_name"].isin(selected_categories)]
    else:
        filtered_df = filtered_df.iloc[0:0]

    if selected_payments:
        filtered_df = filtered_df[filtered_df["payment_method"].isin(selected_payments)]
    else:
        filtered_df = filtered_df.iloc[0:0]

    # Sidebar Footer
    st.sidebar.markdown("---")
    st.sidebar.markdown(
        f"""
        <div style="font-size: 0.72rem; color: #718096; line-height: 1.4;">
            <strong>Active Dataset:</strong><br>
            Showing <strong>{len(filtered_df):,}</strong> of <strong>{len(df):,}</strong> records<br>
            Stores selected: <strong>{len(selected_stores)}/{len(all_stores)}</strong>
        </div>
        """,
        unsafe_allow_html=True
    )

    return filtered_df
