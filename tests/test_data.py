import os
import sqlite3
import unittest
import pandas as pd
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
DB_PATH = BASE_DIR / "data" / "app.db"

class TestDataLayerAndIntegrity(unittest.TestCase):
    """TC-T-01 to TC-T-06: Database Schema, Referential Integrity, and Data Hygiene Tests."""

    @classmethod
    def setUpClass(cls):
        if not DB_PATH.exists():
            raise FileNotFoundError(f"Database file not found at {DB_PATH}")
        cls.conn = sqlite3.connect(f"file:{DB_PATH.as_posix()}?mode=ro", uri=True)
        cls.conn.execute("PRAGMA foreign_keys = ON;")

    @classmethod
    def tearDownClass(cls):
        if hasattr(cls, "conn") and cls.conn:
            cls.conn.close()

    def test_tc_t_01_db_file_exists(self):
        """TC-T-01: Verify SQLite database file exists, is accessible, and has non-zero size."""
        self.assertTrue(DB_PATH.exists(), f"Database file not found at {DB_PATH}")
        size_bytes = os.path.getsize(DB_PATH)
        self.assertGreater(size_bytes, 1024, f"Database file size {size_bytes} is unexpectedly small.")

    def test_tc_t_02_table_row_counts(self):
        """TC-T-02: Verify minimum planned row counts across all 5 relational tables."""
        expected_minimums = {
            "dim_stores": 10,
            "dim_products": 50,
            "fact_orders": 150,
            "fact_order_items": 500,
            "dim_sales_targets": 100
        }
        for table, min_count in expected_minimums.items():
            df = pd.read_sql_query(f"SELECT COUNT(*) as cnt FROM {table}", self.conn)
            actual_count = int(df["cnt"].iloc[0])
            self.assertGreaterEqual(
                actual_count, min_count,
                f"Table {table} has {actual_count} rows, expected at least {min_count}"
            )

    def test_tc_t_03_primary_key_uniqueness(self):
        """TC-T-03: Verify primary key uniqueness across all dimension and fact tables."""
        pk_map = {
            "dim_stores": "store_id",
            "dim_products": "product_id",
            "fact_orders": "order_id",
            "fact_order_items": "order_item_id",
            "dim_sales_targets": "target_id"
        }
        for table, pk in pk_map.items():
            query = f"""
            SELECT {pk}, COUNT(*) as cnt 
            FROM {table} 
            GROUP BY {pk} 
            HAVING COUNT(*) > 1
            """
            duplicates = pd.read_sql_query(query, self.conn)
            self.assertEqual(len(duplicates), 0, f"Found duplicate primary keys in {table}: {duplicates}")

    def test_tc_t_04_foreign_key_integrity(self):
        """TC-T-04: Execute PRAGMA foreign_key_check asserting zero orphan child records."""
        cursor = self.conn.cursor()
        fk_violations = cursor.execute("PRAGMA foreign_key_check;").fetchall()
        self.assertEqual(len(fk_violations), 0, f"Foreign key integrity violations detected: {fk_violations}")

    def test_tc_t_05_view_store_sales_analytics(self):
        """TC-T-05: Verify view_store_sales_analytics contains exactly 36 expected columns."""
        df_sample = pd.read_sql_query("SELECT * FROM view_store_sales_analytics LIMIT 5", self.conn)
        self.assertEqual(len(df_sample.columns), 36, f"Expected 36 columns in view, got {len(df_sample.columns)}")
        self.assertIn("order_id", df_sample.columns)
        self.assertIn("store_name", df_sample.columns)
        self.assertIn("sales_velocity_sqft_contrib", df_sample.columns)
        self.assertIn("line_margin_pct", df_sample.columns)

    def test_tc_t_06_data_cleaning_invariants(self):
        """TC-T-06: Assert cleaning domain invariants (quantity >= 1, discount between 0 and 50)."""
        df_items = pd.read_sql_query(
            "SELECT quantity, discount_pct, unit_selling_price_inr, unit_cost_price_inr FROM fact_order_items",
            self.conn
        )
        self.assertTrue((df_items["quantity"] >= 1).all(), "Found invalid order quantity < 1")
        self.assertTrue(
            (df_items["discount_pct"] >= 0.0).all() and (df_items["discount_pct"] <= 50.0).all(),
            "Discount percentage out of allowed [0, 50] range"
        )
        self.assertTrue(
            (df_items["unit_selling_price_inr"] >= df_items["unit_cost_price_inr"]).all(),
            "Found selling price less than cost price in fact_order_items"
        )

if __name__ == "__main__":
    unittest.main()
