"""テーゼ破綻の観測ログ（記録専用）。売買の記録に触れないことを固定する。"""
import sqlite3
import unittest
from unittest import mock

from screener.report import thesis_observe as T


class _P:
    """thesis_break の代わり。投影DBは使わない。"""

    def __init__(self, table):
        self.table = table


def _fake_break(pcon, code, event_date, cache=None):
    return pcon.table[(code, event_date)]


class TestObserve(unittest.TestCase):
    def setUp(self):
        self.con = sqlite3.connect(":memory:")
        self.con.executescript(
            "CREATE TABLE paper_trades (trade_id INTEGER PRIMARY KEY, code TEXT,"
            " event_date TEXT, ret_net REAL, exit_rule TEXT);")
        self.con.executemany(
            "INSERT INTO paper_trades VALUES (?,?,?,?,?)",
            [(1, "1111", "2026-08-10", 0.02, "T+2"),
             (2, "2222", "2026-08-12", -0.05, "T+2"),
             (3, "3333", "2026-12-01", None, None)])      # まだ決算前
        self.pcon = _P({
            ("1111", "2026-08-10"): {"evaluable": True, "break_or": False,
                                     "break_and": False, "sales_yoy": 10.0, "op_yoy": 5.0},
            ("2222", "2026-08-12"): {"evaluable": True, "break_or": True,
                                     "break_and": False, "sales_yoy": 4.0, "op_yoy": -8.0},
            ("3333", "2026-12-01"): {"evaluable": False, "break_or": False,
                                     "break_and": False, "sales_yoy": None, "op_yoy": None},
        })

    def _run(self, as_of=None):
        with mock.patch.object(T.X, "thesis_break", _fake_break):
            return T.observe(self.con, self.pcon, as_of)

    def test_records_only_disclosed_events(self):
        st = self._run(as_of="2026-09-20")
        self.assertEqual(st["new"], 2)
        got = self.con.execute(
            "SELECT trade_id, break_or, op_yoy_diff FROM thesis_observations "
            "ORDER BY trade_id").fetchall()
        self.assertEqual(got, [(1, 0, 5.0), (2, 1, -8.0)])

    def test_idempotent(self):
        self._run(as_of="2026-09-20")
        st = self._run(as_of="2026-09-20")
        self.assertEqual((st["new"], st["updated"]), (0, 2))
        self.assertEqual(self.con.execute(
            "SELECT COUNT(*) FROM thesis_observations").fetchone()[0], 2)

    def test_does_not_touch_trade_records(self):
        before = self.con.execute("SELECT * FROM paper_trades ORDER BY trade_id").fetchall()
        self._run(as_of="2026-09-20")
        after = self.con.execute("SELECT * FROM paper_trades ORDER BY trade_id").fetchall()
        self.assertEqual(before, after)

    def test_report_groups(self):
        self._run(as_of="2026-09-20")
        txt = T.report(self.con)
        self.assertIn("テーゼ健全", txt)
        self.assertIn("テーゼ破綻(OR)", txt)


if __name__ == "__main__":
    unittest.main()
