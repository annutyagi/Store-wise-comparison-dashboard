import streamlit as st
import io
import datetime

from src.data_loader import load_analytical_view
from src.cleaning import clean_sales_data
from src.ui import render_tailwind_header, render_insight_box
from src.sidebar import render_global_sidebar

def main():
    st.set_page_config(page_title="Tabular Data & Export", page_icon="📥", layout="wide")

    raw_df = load_analytical_view()
    df, _ = clean_sales_data(raw_df)
    filtered_df = render_global_sidebar(df)

    render_tailwind_header(
        title="Tabular Data Explorer & Export Engine",
        subtitle="Download filtered, audit-ready analytical datasets in CSV and Excel formats (BR-13, US-16)",
        badge_text="Data Export",
        height=125
    )

    if filtered_df.empty:
        render_insight_box("Zero Records", "No records match active filters.", "warning")
        st.stop()

    st.subheader(f"Filtered Records: {len(filtered_df):,} Rows Selected")
    st.dataframe(filtered_df, use_container_width=True)

    # Export Buttons
    col1, col2, _ = st.columns([1, 1, 2])

    today_str = datetime.date.today().strftime("%Y%m%d")

    with col1:
        csv_data = filtered_df.to_csv(index=False).encode("utf-8")
        st.download_button(
            label="📄 Download CSV",
            data=csv_data,
            file_name=f"retail_sales_export_{today_str}.csv",
            mime="text/csv",
            use_container_width=True
        )

    with col2:
        buffer = io.BytesIO()
        try:
            filtered_df.to_excel(buffer, index=False, engine="openpyxl")
            excel_data = buffer.getvalue()
            st.download_button(
                label="📊 Download Excel",
                data=excel_data,
                file_name=f"retail_sales_export_{today_str}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        except Exception as e:
            st.warning(f"Excel export unavailable: {e}. Please use CSV download.")

    render_insight_box(
        title="Data Export Governance (BR-13)",
        message="Exported datasets include the complete 36-column analytical view with full store demographics, order line profits, and sales velocity metrics ready for executive board reporting.",
        level="info"
    )

if __name__ == "__main__":
    import os
    main()
