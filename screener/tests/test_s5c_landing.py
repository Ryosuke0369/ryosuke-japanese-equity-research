"""S5c（着地推定と会社予想の乖離）。閾値と対象外の条件を固定する。

フィクスチャの数値は 3441 山王の形（Q3 進捗は高いが Q4 が赤字で着地は会予割れ）と、
その逆（前年並みの季節性なら着地が会予を大きく超える形）。
"""
import sqlite3
import unittest

from screener.signals import s5c_landing as S5C

DDL = """
CREATE TABLE quarterly_standalone_all (
  ticker TEXT, fiscal_year INTEGER, quarter_type TEXT, period_end TEXT,
  span_q INTEGER, period_start TEXT, sales REAL, operating_profit REAL,
  is_valid INTEGER DEFAULT 1, generation INTEGER DEFAULT 1);
CREATE TABLE company_forecasts (
  ticker TEXT, fiscal_year INTEGER, forecast_sales REAL, forecast_op REAL,
  source_date TEXT);
CREATE TABLE filings (ticker TEXT, filing_date TEXT, period_end TEXT);
CREATE TABLE pl_adjustments (ticker TEXT, period_end TEXT, amount REAL);
"""


def _q(con, ticker, fy, q, sales, op):
    con.execute("INSERT INTO quarterly_standalone_all (ticker, fiscal_year, quarter_type,"
                " period_end, span_q, period_start, sales, operating_profit)"
                " VALUES (?,?,?,?,1,?,?,?)",
                (ticker, fy, "Q%d" % q, "FY%d-Q%d" % (fy, q), "FY%d-Q%d" % (fy, q),
                 sales, op))
    con.execute("INSERT INTO filings VALUES (?,?,?)",
                (ticker, "%d-01-01" % fy, "FY%d-Q%d" % (fy, q)))


class TestS5c(unittest.TestCase):
    def setUp(self):
        self.con = sqlite3.connect(":memory:")
        self.con.row_factory = sqlite3.Row
        self.con.executescript(DDL)
        # 前年（FY2025）: Q1〜Q4 = 10/10/10/70 → Q3までの進捗 30/100 = 30%
        for q, op in ((1, 10), (2, 10), (3, 10), (4, 70)):
            _q(self.con, "1111", 2025, q, op * 10, op)
        # 当期（FY2026）: Q1〜Q3 = 20/20/20（累計60）
        for q in (1, 2, 3):
            _q(self.con, "1111", 2026, q, 200, 20)

    def _fc(self, op, sales=2000):
        self.con.execute("INSERT INTO company_forecasts VALUES ('1111',2026,?,?,'2026-01-01')",
                         (sales, op))

    def test_landing_uses_prior_year_progress(self):
        self._fc(op=120)                 # 着地推定 60 / 0.30 = 200 → 乖離 +66.7%
        r = S5C.evaluate(self.con, "1111")
        self.assertTrue(r["available"])
        self.assertAlmostEqual(r["progress_prior"], 0.30)
        self.assertAlmostEqual(r["landing_op"], 200.0)
        self.assertAlmostEqual(r["gap_op"], 200 / 120 - 1)
        self.assertTrue(r["fired"])

    def test_below_the_disclosure_line_does_not_fire(self):
        self._fc(op=180)                 # 乖離 +11.1% < 30%
        r = S5C.evaluate(self.con, "1111")
        self.assertTrue(r["available"])
        self.assertFalse(r["fired"])

    def test_score_is_always_zero(self):
        self._fc(op=120)
        self.assertEqual(S5C.evaluate(self.con, "1111")["score"], 0.0)

    def test_negative_forecast_is_excluded(self):
        self._fc(op=-50)
        r = S5C.evaluate(self.con, "1111")
        self.assertFalse(r["available"])
        self.assertIn("0以下", r["evidence"])

    def test_no_prior_year_is_not_guessed(self):
        con = sqlite3.connect(":memory:")
        con.row_factory = sqlite3.Row
        con.executescript(DDL)
        for q in (1, 2, 3):
            _q(con, "2222", 2026, q, 200, 20)
        con.execute("INSERT INTO company_forecasts VALUES ('2222',2026,2000,120,'2026-01-01')")
        r = S5C.evaluate(con, "2222")
        self.assertFalse(r["available"])
        self.assertIn("前年同期の進捗率", r["evidence"])

    def test_prior_year_loss_is_excluded(self):
        con = sqlite3.connect(":memory:")
        con.row_factory = sqlite3.Row
        con.executescript(DDL)
        for q, op in ((1, -10), (2, -10), (3, -10), (4, -10)):   # 前年通期が赤字
            _q(con, "3333", 2025, q, 100, op)
        for q in (1, 2, 3):
            _q(con, "3333", 2026, q, 200, 20)
        con.execute("INSERT INTO company_forecasts VALUES ('3333',2026,2000,120,'2026-01-01')")
        self.assertFalse(S5C.evaluate(con, "3333")["available"])

    def test_extreme_prior_progress_is_excluded(self):
        con = sqlite3.connect(":memory:")
        con.row_factory = sqlite3.Row
        con.executescript(DDL)
        # 前年 Q1〜Q3 の累計が通期の 300%（Q4 に大赤字）→ 進捗率の上限超で対象外
        for q, op in ((1, 100), (2, 100), (3, 100), (4, -200)):
            _q(con, "4444", 2025, q, 100, op)
        for q in (1, 2, 3):
            _q(con, "4444", 2026, q, 200, 20)
        con.execute("INSERT INTO company_forecasts VALUES ('4444',2026,2000,120,'2026-01-01')")
        r = S5C.evaluate(con, "4444")
        self.assertFalse(r["available"])

    def test_q4_is_out_of_scope(self):
        con = sqlite3.connect(":memory:")
        con.row_factory = sqlite3.Row
        con.executescript(DDL)
        for q, op in ((1, 10), (2, 10), (3, 10), (4, 70)):
            _q(con, "5555", 2025, q, 100, op)
        for q in (1, 2, 3, 4):
            _q(con, "5555", 2026, q, 200, 20)
        con.execute("INSERT INTO company_forecasts VALUES ('5555',2026,2000,120,'2026-01-01')")
        r = S5C.evaluate(con, "5555")
        self.assertFalse(r["available"])

    def test_parameters_match_the_registration(self):
        self.assertEqual(S5C.FIRE_OP_PCT, 0.30)
        self.assertEqual(S5C.NOTE_SALES_PCT, 0.10)
        self.assertEqual(S5C.MAX_PRIOR_YEARS, 2)
        self.assertEqual((S5C.PROGRESS_MIN, S5C.PROGRESS_MAX), (0.05, 2.00))
        self.assertEqual(S5C.MAX_Q, 3)


if __name__ == "__main__":
    unittest.main()
