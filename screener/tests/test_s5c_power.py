"""シャドウI-v2 追補: 修正イベントの検出と、判定日・窓の扱いを固定する。"""
import sqlite3
import unittest

from screener.report import s5c_power as W

DDL = """
CREATE TABLE fins_summary (
  code TEXT, disc_date TEXT, doc_type TEXT, disc_time TEXT, per_type TEXT,
  per_start TEXT, per_end TEXT, fy_end TEXT, sales REAL, op REAL, odp REAL, np REAL,
  f_sales REAL, f_op REAL, fnc_sales REAL, fnc_op REAL, nxf_sales REAL, nxf_op REAL,
  fetched_at TEXT);
"""


def _row(con, code, d, fy, f_op, per="2Q"):
    con.execute("INSERT INTO fins_summary (code, disc_date, doc_type, per_type, fy_end,"
                " f_op) VALUES (?,?,?,?,?,?)",
                (code, d, "2QFinancialStatements_Consolidated_JP", per, fy, f_op))


class TestRevisionEvents(unittest.TestCase):
    def setUp(self):
        self.con = sqlite3.connect(":memory:")
        self.con.executescript(DDL)

    def test_change_is_an_event_and_first_disclosure_is_not(self):
        _row(self.con, "1111", "2025-05-01", "2026-03-31", 100.0)   # initial
        _row(self.con, "1111", "2025-08-01", "2026-03-31", 100.0)   # 据置
        _row(self.con, "1111", "2025-11-01", "2026-03-31", 130.0)   # up
        _row(self.con, "1111", "2026-02-01", "2026-03-31", 90.0)    # down
        evs, skipped = W.revision_events(self.con)
        self.assertEqual([(d, k) for d, k, _m in evs["1111"]],
                         [("2025-11-01", "up"), ("2026-02-01", "down")])
        self.assertEqual(skipped["initial（前の予想が無い）"], 1)
        self.assertEqual(skipped["据置"], 1)
        self.assertAlmostEqual(evs["1111"][0][2], 0.30)

    def test_new_fiscal_year_restarts(self):
        _row(self.con, "2222", "2025-05-01", "2026-03-31", 100.0)
        _row(self.con, "2222", "2026-05-01", "2027-03-31", 300.0)   # 別年度 = initial
        evs, _ = W.revision_events(self.con)
        self.assertEqual(evs.get("2222", []), [])

    def test_zero_previous_is_skipped(self):
        _row(self.con, "3333", "2025-05-01", "2026-03-31", 0.0)
        _row(self.con, "3333", "2025-11-01", "2026-03-31", 50.0)
        evs, skipped = W.revision_events(self.con)
        self.assertEqual(evs.get("3333", []), [])
        self.assertEqual(skipped["前回予想が0"], 1)


class TestWindow(unittest.TestCase):
    def test_outcome_takes_the_first_event_in_window(self):
        evs = {"1111": [("2025-06-01", "down", -0.1), ("2025-07-01", "up", 0.4)]}
        self.assertEqual(W.outcome(evs, "1111", "2025-05-01", "2025-08-01"), "down")
        self.assertEqual(W.outcome(evs, "1111", "2025-06-15", "2025-08-01"), "up")
        self.assertEqual(W.outcome(evs, "1111", "2025-08-02", "2025-11-01"), "none")

    def test_event_on_the_judgment_day_is_excluded(self):
        evs = {"1111": [("2025-06-01", "up", 0.4)]}
        self.assertEqual(W.outcome(evs, "1111", "2025-06-01", "2025-08-30"), "none")


class TestJudgmentDates(unittest.TestCase):
    def test_first_business_day_of_each_month(self):
        cal = ["2021-09-30", "2021-10-01", "2021-10-04", "2021-11-01",
               "2026-06-01", "2026-06-02", "2026-07-01"]
        got = W.judgment_dates(cal)
        self.assertEqual(got, ["2021-10-01", "2021-11-01", "2026-06-01"])


class TestParameters(unittest.TestCase):
    def test_registered_values(self):
        self.assertEqual(W.WINDOW_DAYS, 90)
        self.assertEqual(W.MIN_N, 100)
        self.assertEqual([b[2] for b in W.R_BINS],
                         ["R<1.0", "1.0-1.1", "1.1-1.3", "1.3-1.5", "1.5-2.0", "2.0+"])


if __name__ == "__main__":
    unittest.main()
