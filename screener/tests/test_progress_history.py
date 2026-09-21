"""シャドウI-v2: 過去5年の進捗率と R（進捗率 ÷ 過去中央値進捗率）。

固定するもの: 進捗率の作り方（累計 ÷ その年度の着地）、中央値の年数、
R の定義（= 推定着地 ÷ 会社予想）、PIT（as_of 以降の開示を見ない）、除外条件。
"""
import sqlite3
import unittest

from screener.signals import progress_history as P

DDL = """
CREATE TABLE fins_summary (
  code TEXT, disc_date TEXT, doc_type TEXT, disc_time TEXT, per_type TEXT,
  per_start TEXT, per_end TEXT, fy_end TEXT, sales REAL, op REAL, odp REAL, np REAL,
  f_sales REAL, f_op REAL, fnc_sales REAL, fnc_op REAL, nxf_sales REAL, nxf_op REAL,
  fetched_at TEXT);
"""
TAN = "2QFinancialStatements_Consolidated_JP"
FYD = "FYFinancialStatements_Consolidated_JP"


def _row(con, code, d, per_type, fy_end, op, f_op=None, doc=None, sales=None, f_sales=None):
    con.execute("INSERT INTO fins_summary (code, disc_date, doc_type, per_type, fy_end,"
                " op, f_op, sales, f_sales) VALUES (?,?,?,?,?,?,?,?,?)",
                (code, d, doc or (FYD if per_type == "FY" else TAN), per_type, fy_end,
                 op, f_op, sales, f_sales))


class TestProgress(unittest.TestCase):
    def setUp(self):
        self.con = sqlite3.connect(":memory:")
        self.con.executescript(DDL)
        # 過去4年: 2Q 累計 30 → 着地 100（進捗 30%）
        for i, fy in enumerate(("2022-03-31", "2023-03-31", "2024-03-31", "2025-03-31")):
            _row(self.con, "1111", "%d-11-01" % (2021 + i), "2Q", fy, 30.0, f_op=100.0)
            _row(self.con, "1111", "%d-05-01" % (2022 + i), "FY", fy, 100.0)
        # 当期 2026-03: 2Q 累計 60・会予 120（進捗 50%）
        _row(self.con, "1111", "2025-11-01", "2Q", "2026-03-31", 60.0, f_op=120.0)

    def test_series_and_median(self):
        rows = P.load_rows(self.con, "1111")
        s = P.progress_series(rows)
        self.assertEqual(len(s), 4)
        self.assertAlmostEqual(s[("2023-03-31", "2Q")], 0.30)
        med, years = P.median_progress(s, "2Q", exclude_fy="2026-03-31")
        self.assertAlmostEqual(med, 0.30)
        self.assertEqual(len(years), 4)

    def test_r_equals_landing_over_forecast(self):
        r = P.evaluate(self.con, "1111")
        self.assertTrue(r["available"])
        self.assertAlmostEqual(r["progress"], 0.50)
        self.assertAlmostEqual(r["r"], 0.50 / 0.30)
        self.assertAlmostEqual(r["landing_op"], 60.0 / 0.30)          # 200
        self.assertAlmostEqual(r["r"], r["landing_op"] / r["forecast_op"])
        self.assertTrue(r["fired"])                                    # R 1.67 >= 1.30
        self.assertEqual(r["score"], 0.0)

    def test_needs_three_years(self):
        con = sqlite3.connect(":memory:")
        con.executescript(DDL)
        for i, fy in enumerate(("2024-03-31", "2025-03-31")):
            _row(con, "2222", "%d-11-01" % (2023 + i), "2Q", fy, 30.0, f_op=100.0)
            _row(con, "2222", "%d-05-01" % (2024 + i), "FY", fy, 100.0)
        _row(con, "2222", "2025-11-01", "2Q", "2026-03-31", 60.0, f_op=120.0)
        r = P.evaluate(con, "2222")
        self.assertFalse(r["available"])
        self.assertIn("3年に満たない", r["evidence"])

    def test_pit_ignores_later_disclosures(self):
        # as_of を当期2Qの前にすると、当期は前年の2Qになる
        r = P.evaluate(self.con, "1111", as_of="2025-06-30")
        self.assertTrue(r["available"])
        self.assertEqual(r["fy_end"], "2025-03-31")

    def test_landing_missing_year_is_skipped(self):
        con = sqlite3.connect(":memory:")
        con.executescript(DDL)
        _row(con, "3333", "2023-11-01", "2Q", "2024-03-31", 30.0, f_op=100.0)  # 着地なし
        for i, fy in enumerate(("2025-03-31", "2026-03-31")):
            _row(con, "3333", "%d-11-01" % (2024 + i), "2Q", fy, 30.0, f_op=100.0)
            _row(con, "3333", "%d-05-01" % (2025 + i), "FY", fy, 100.0)
        s = P.progress_series(P.load_rows(con, "3333"))
        self.assertNotIn(("2024-03-31", "2Q"), s)

    def test_loss_landing_is_excluded(self):
        con = sqlite3.connect(":memory:")
        con.executescript(DDL)
        _row(con, "4444", "2023-11-01", "2Q", "2024-03-31", 30.0, f_op=100.0)
        _row(con, "4444", "2024-05-01", "FY", "2024-03-31", -50.0)      # 着地が赤字
        s = P.progress_series(P.load_rows(con, "4444"))
        self.assertEqual(s, {})

    def test_extreme_progress_is_excluded(self):
        con = sqlite3.connect(":memory:")
        con.executescript(DDL)
        _row(con, "5555", "2023-11-01", "2Q", "2024-03-31", 300.0, f_op=100.0)
        _row(con, "5555", "2024-05-01", "FY", "2024-03-31", 100.0)      # 進捗 300%
        self.assertEqual(P.progress_series(P.load_rows(con, "5555")), {})

    def test_parameters_match_the_registration(self):
        self.assertEqual(P.FIRE_R, 1.30)
        self.assertEqual(P.NOTE_R_SALES, 1.10)
        self.assertEqual((P.MIN_YEARS, P.MAX_YEARS), (3, 5))
        self.assertEqual((P.PROGRESS_MIN, P.PROGRESS_MAX), (0.05, 2.00))


if __name__ == "__main__":
    unittest.main()
