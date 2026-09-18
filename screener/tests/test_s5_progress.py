"""tests/test_s5_progress.py — S5 進捗率・死んだガイダンスの単体テスト（§33）。

実運用で人手検証した2例の数値をそのまま固定する。**同じ進捗113%が、
一方は上方修正後、一方は通期未達で終わっている**ので、両方を残す。

  6838 多摩川HD（2026年10月期）: Q3累計OP 928.5 / 通期予想OP 820 → 暗黙のQ4 −108.5
  3441 山王    （2026年7月期）: Q3累計OP 1,582 / 通期予想OP 1,400 → 暗黙のQ4 −182
"""
import os
import sqlite3
import sys
import unittest
from datetime import date

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, ROOT)

from screener.projection import materialize as MZ          # noqa: E402
from screener.signals import span_scorers as SP             # noqa: E402


def _db(fy, cum_sales, cum_op, fc_sales, fc_op, span=3, fym=7):
    con = sqlite3.connect(":memory:")
    con.row_factory = sqlite3.Row
    con.executescript(MZ.DDL)
    con.execute("INSERT INTO universe (ticker, fiscal_year_end) VALUES ('T', ?)", (fym,))
    con.execute(
        "INSERT INTO quarterly_standalone_all (ticker, fiscal_year, quarter_type,"
        " period_end, span_q, period_start, sales, operating_profit, is_valid)"
        " VALUES ('T', ?, ?, ?, ?, ?, ?, ?, 1)",
        (fy, "%dQ" % span, "FY%d-Q%d" % (fy, span), span, "FY%d-Q1" % fy,
         cum_sales, cum_op))
    con.execute(
        "INSERT INTO company_forecasts (ticker, fiscal_year, forecast_sales,"
        " forecast_op, source_date) VALUES ('T', ?, ?, ?, '2026-06-01')",
        (fy, fc_sales, fc_op))
    # **PIT の可視集合は filings から作られる。**filings が空の DB だと
    # すべての期が「公知でない」になり、スコアラーは何も見えない。
    con.execute(
        "INSERT INTO filings (ticker, period_end, filing_date, fiscal_year,"
        " quarter_type, source, doc_id)"
        " VALUES ('T', ?, '2026-06-01', ?, ?, 'tdnet', 'D1')",
        ("FY%d-Q%d" % (fy, span), fy, "%dQ" % span))
    con.commit()
    return con


class TestDeadGuidance(unittest.TestCase):
    def test_6838型_Q3累計OPが通期予想を超える(self):
        con = _db(2026, cum_sales=5302.1, cum_op=928.5,
                  fc_sales=6800.0, fc_op=820.0, fym=10)
        r = SP.s5_progress(con, "T", as_of=date(2026, 9, 18))
        self.assertTrue(r["available"], r["evidence"])
        d = r["details"]
        self.assertEqual(d["elapsed_q"], 3)
        self.assertAlmostEqual(d["pace"], 0.75)
        self.assertAlmostEqual(d["progress_op"], 928.5 / 820.0, places=3)
        self.assertAlmostEqual(d["implied_rest_op"], -108.5, places=1)
        self.assertAlmostEqual(d["implied_rest_op_per_q"], -108.5, places=1)
        self.assertTrue(d["guidance_dead"])
        self.assertTrue(r["guidance_dead"])
        self.assertIn("死んだガイダンス", r["evidence"])
        # 進捗113% / 経過四半期比75% → 超過 +38pt
        self.assertAlmostEqual(d["pace_excess_pt"], 38.2, places=0)

    def test_3441型_同じ113パーセントで暗黙のQ4は赤字(self):
        con = _db(2026, cum_sales=15000.0, cum_op=1582.0,
                  fc_sales=20000.0, fc_op=1400.0, fym=7)
        r = SP.s5_progress(con, "T", as_of=date(2026, 6, 1))
        d = r["details"]
        self.assertAlmostEqual(d["progress_op"], 1582.0 / 1400.0, places=3)
        self.assertAlmostEqual(d["implied_rest_op"], -182.0, places=1)
        self.assertTrue(d["guidance_dead"])
        # **6838 と同じ形**であることを固定する（結末は逆だった）。
        self.assertAlmostEqual(d["progress_op"], 1.13, places=2)

    def test_進捗が経過四半期どおりならガイダンスは死んでいない(self):
        con = _db(2026, cum_sales=15000.0, cum_op=1050.0,
                  fc_sales=20000.0, fc_op=1400.0, fym=7)
        r = SP.s5_progress(con, "T", as_of=date(2026, 6, 1))
        d = r["details"]
        self.assertFalse(d["guidance_dead"])
        self.assertAlmostEqual(d["implied_rest_op"], 350.0, places=1)
        self.assertAlmostEqual(d["pace_excess_pt"], 0.0, places=1)

    def test_予想OPが無い銘柄では進捗列が空になる(self):
        con = _db(2026, cum_sales=15000.0, cum_op=1050.0,
                  fc_sales=20000.0, fc_op=0.0, fym=7)
        r = SP.s5_progress(con, "T", as_of=date(2026, 6, 1))
        d = r["details"]
        self.assertIsNone(d["progress_op"])
        self.assertIsNone(d["implied_rest_op"])
        self.assertFalse(d["guidance_dead"])


class TestS5bUnscored(unittest.TestCase):
    """§36: 陳腐化は加点しない。数値は S5b として 0点で残す。"""

    def _score(self, con, as_of):
        import sys as _s
        from screener.report import backtest_eval as V1
        root = str(V1._external_root())
        if root not in _s.path:
            _s.path.insert(0, root)
        from module_b.run_scorers import SCORERS_ALL
        from screener.signals import span_runner as SR
        return SR.score_ticker(con, "T", SCORERS_ALL, as_of=as_of,
                               policy="evidence_strict")

    def test_陳腐化したS5は加点されずS5bに残る(self):
        con = _db(2026, cum_sales=5302.1, cum_op=928.5,
                  fc_sales=6800.0, fc_op=820.0, fym=10)
        r = SP.s5_progress(con, "T", as_of=date(2026, 9, 18))
        self.assertTrue(r["guidance_dead"])
        self.assertIsNotNone(r["dead_signal"])
        self.assertEqual(r["dead_signal"]["score"], 0.0)
        self.assertIn("陳腐化", r["dead_signal"]["evidence"])
        # policy 側で不採用になる
        s5 = self._score(con, date(2026, 9, 18))["S5"]
        self.assertFalse(s5["available"])
        self.assertEqual(s5["score"], 0.0)
        self.assertIn("guidance_dead_unscored", s5["strict_flags"])
        # **減点にはしない。**方向が測れていないものに符号を与えない
        self.assertGreaterEqual(s5["score"], 0.0)

    def test_陳腐化していないS5は従来どおり点になる(self):
        # 6466 型: 上方修正後で進捗が正常域に戻っている
        con = _db(2026, cum_sales=8543.5, cum_op=1076.6,
                  fc_sales=11200.0, fc_op=1150.0, fym=9)
        r = SP.s5_progress(con, "T", as_of=date(2026, 9, 18))
        self.assertFalse(r["guidance_dead"])
        self.assertIsNone(r["dead_signal"])
        self.assertAlmostEqual(r["details"]["implied_rest_op"], 73.4, places=1)
        s5 = self._score(con, date(2026, 9, 18))["S5"]
        self.assertNotIn("guidance_dead_unscored", s5.get("strict_flags") or [])


if __name__ == "__main__":
    unittest.main(verbosity=2)
