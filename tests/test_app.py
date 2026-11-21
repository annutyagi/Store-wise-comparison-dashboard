import os
import time
import sqlite3
import unittest
import pandas as pd
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
DB_PATH = BASE_DIR / "data" / "app.db"

class TestAppArchitectureAndUI(unittest.TestCase):
    """TC-T-13 to TC-T-15: Streamlit App Configuration, UI Components, and Latency Benchmark."""

    def test_tc_t_13_streamlit_config_toml(self):
        """TC-T-13: Verify .streamlit/config.toml exists and contains corporate theme."""
        config_path = BASE_DIR / ".streamlit" / "config.toml"
        self.assertTrue(config_path.exists(), ".streamlit/config.toml not found")
        with open(config_path, "r", encoding="utf-8") as f:
            content = f.read()
        self.assertIn("[theme]", content)
        self.assertIn("#1A365D", content) # Corporate Navy Primary

    def test_tc_t_14_ui_component_library(self):
        """TC-T-14: Verify src/ui.py imports cleanly and provides reusable Tailwind components."""
        import src.ui as ui
        self.assertTrue(hasattr(ui, "render_tailwind_header"))
        self.assertTrue(hasattr(ui, "render_kpi_cards"))
        self.assertTrue(hasattr(ui, "render_insight_box"))

    def test_tc_t_15_query_latency_benchmark(self):
        """TC-T-15: Verify analytical view query latency is under 1.0 second (well below 3.0s limit)."""
        start_time = time.perf_counter()
        conn = sqlite3.connect(f"file:{DB_PATH.as_posix()}?mode=ro", uri=True)
        df = pd.read_sql_query("SELECT * FROM view_store_sales_analytics", conn)
        conn.close()
        duration = time.perf_counter() - start_time
        self.assertLess(duration, 1.0, f"Query took {duration:.3f}s, expected < 1.0s")
        self.assertGreater(len(df), 0, "Analytical view returned empty dataframe")

if __name__ == "__main__":
    unittest.main()
