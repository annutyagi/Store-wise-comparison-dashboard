import os
import sqlite3
import pandas as pd
from pathlib import Path

try:
    import streamlit as st
    cache_data_decorator = st.cache_data
except ImportError:
    # Fallback decorator when running outside Streamlit / in test suite
    def cache_data_decorator(func):
        return func

BASE_DIR = Path(__file__).resolve().parent.parent
DB_PATH = BASE_DIR / "data" / "app.db"

def get_db_connection():
    """Returns a read-only SQLite connection with foreign keys enabled."""
    if not DB_PATH.exists():
        raise FileNotFoundError(f"Database not found at {DB_PATH}")
    conn = sqlite3.connect(f"file:{DB_PATH.as_posix()}?mode=ro", uri=True)
    conn.execute("PRAGMA foreign_keys = ON;")
    return conn

@cache_data_decorator
def load_analytical_view() -> pd.DataFrame:
    """Loads the 36-column denormalized analytical view for fast dashboard analytics."""
    conn = get_db_connection()
    try:
        query = "SELECT * FROM view_store_sales_analytics"
        df = pd.read_sql_query(query, conn)
        # Parse datetime columns
        if "order_datetime" in df.columns:
            df["order_datetime"] = pd.to_datetime(df["order_datetime"])
        if "order_date" in df.columns:
            df["order_date"] = pd.to_datetime(df["order_date"]).dt.date
        return df
    finally:
        conn.close()

@cache_data_decorator
def load_raw_table(table_name: str) -> pd.DataFrame:
    """Loads an individual relational master or fact table from data/app.db."""
    allowed_tables = {
        "dim_stores", "dim_products", "fact_orders", 
        "fact_order_items", "dim_sales_targets"
    }
    if table_name not in allowed_tables:
        raise ValueError(f"Unauthorized table query: {table_name}")
    
    conn = get_db_connection()
    try:
        df = pd.read_sql_query(f"SELECT * FROM {table_name}", conn)
        return df
    finally:
        conn.close()

@cache_data_decorator
def check_foreign_key_violations() -> list:
    """Executes PRAGMA foreign_key_check asserting zero orphan records."""
    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        violations = cursor.execute("PRAGMA foreign_key_check;").fetchall()
        return violations
    finally:
        conn.close()
