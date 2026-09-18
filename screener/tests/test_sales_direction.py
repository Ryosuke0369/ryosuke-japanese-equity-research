"""tests/test_sales_direction.py — S1 売上方向ガードの単体テスト（§32）。

実DBに依存しない合成データで、規則そのものを固定する。
実銘柄での回帰は tests/test_signal_design_fixtures.py（フィクスチャ）側。

守る不変条件:
  - 6ヶ月の値と3ヶ月の値をそのまま比べない（run-rate に正規化する）
  - 重なる区間を「直前」として採らない（H1 の直前は前年 H2 であって Q2 ではない）
  - qoq / yoy のどちらも取れなければ down にしない（unverified）
  - 累計 span でしか判定できていないことは quarter_level=False で分かる
"""
import os
import sqlite3
import sys
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, ROOT)

from screener.projection import materialize as MZ          # noqa: E402
from screener.signals import sales_direction as SD          # noqa: E402


def _db(rows):
    """rows: (period_end, period_start, span_q, sales) の並び。"""
    con = sqlite3.connect(":memory:")
    con.row_factory = sqlite3.Row
    con.executescript(MZ.DDL)
    for pe, ps, sp, sales in rows:
        con.execute(
            "INSERT INTO quarterly_standalone_all (ticker, fiscal_year, "
            " quarter_type, period_end, span_q, period_start, sales, is_valid) "
            "VALUES ('T', ?, ?, ?, ?, ?, ?, 1)",
            (int(pe[2:6]), "%dQ" % int(pe[-1]), pe, sp, ps, sales))
    con.commit()
    return con


class TestRunRate(unittest.TestCase):
    def test_半期と四半期を素のまま比べない(self):
        # H1 3,743（6ヶ月）→ Q3 1,559（3ヶ月）。素の金額では -58% だが、
        # 1四半期あたりに直すと 1,871 → 1,559 で -16.7%。
        con = _db([("FY2026-Q2", "FY2026-Q1", 2, 3742.889),
                   ("FY2026-Q3", "FY2026-Q3", 1, 1559.2)])
        ts = SD.tiles(con, "T")
        self.assertEqual([t["span_q"] for t in ts], [2, 1])
        d = SD._direction_at(ts, 1)
        self.assertAlmostEqual(d["qoq_pct"], -16.7, places=1)
        self.assertEqual(d["qoq_peer"], "FY2026-Q2")

    def test_重なる区間を直前として採らない(self):
        # 同じ FY2026-Q2 に span=1 と span=2 が並ぶ形。finest 規則で span=1 が
        # 採られ、その直前は FY2026-Q1 になる（H1 ではない）。
        con = _db([("FY2026-Q1", "FY2026-Q1", 1, 100.0),
                   ("FY2026-Q2", "FY2026-Q2", 1, 90.0),
                   ("FY2026-Q2", "FY2026-Q1", 2, 190.0)])
        ts = SD.tiles(con, "T")
        self.assertEqual([(t["period_end"], t["span_q"]) for t in ts],
                         [("FY2026-Q1", 1), ("FY2026-Q2", 1)])
        d = SD._direction_at(ts, 1)
        self.assertEqual(d["qoq_peer"], "FY2026-Q1")
        self.assertAlmostEqual(d["qoq_pct"], -10.0, places=1)


class TestStatus(unittest.TestCase):
    def test_3期連続減は_down(self):
        con = _db([("FY2026-Q1", "FY2026-Q1", 1, 2051.0),
                   ("FY2026-Q2", "FY2026-Q2", 1, 1691.9),
                   ("FY2026-Q3", "FY2026-Q3", 1, 1559.2)])
        r = SD.evaluate(con, "T")
        self.assertEqual(r["status"], "down")
        self.assertTrue(r["quarter_level"])
        self.assertLess(r["latest"]["qoq_pct"], 0)

    def test_qoqが正なら_up(self):
        con = _db([("FY2026-Q1", "FY2026-Q1", 1, 100.0),
                   ("FY2026-Q2", "FY2026-Q2", 1, 120.0)])
        self.assertEqual(SD.evaluate(con, "T")["status"], "up")

    def test_yoyだけ正でも_up(self):
        # qoq は -10% だが yoy は +20% → 「いずれかが正」なので up
        con = _db([("FY2025-Q2", "FY2025-Q2", 1, 100.0),
                   ("FY2026-Q1", "FY2026-Q1", 1, 133.3),
                   ("FY2026-Q2", "FY2026-Q2", 1, 120.0)])
        r = SD.evaluate(con, "T")
        self.assertEqual(r["status"], "up")
        self.assertLess(r["latest"]["qoq_pct"], 0)
        self.assertGreater(r["latest"]["yoy_pct"], 0)

    def test_比較相手が無ければ_unverified(self):
        con = _db([("FY2026-Q2", "FY2026-Q1", 2, 3742.889)])
        r = SD.evaluate(con, "T")
        self.assertEqual(r["status"], "unverified")
        self.assertFalse(r["quarter_level"])

    def test_累計spanでしか判定できないことが分かる(self):
        # 半期しか無い銘柄。方向は出せるが quarter_level は False。
        con = _db([("FY2025-Q2", "FY2025-Q1", 2, 2576.767),
                   ("FY2025-Q4", "FY2025-Q3", 2, 3019.351),
                   ("FY2026-Q2", "FY2026-Q1", 2, 3742.889)])
        r = SD.evaluate(con, "T")
        self.assertEqual(r["status"], "up")
        self.assertFalse(r["quarter_level"])
        self.assertIn("Q単独では未確認", r["note"])


class TestTrend(unittest.TestCase):
    def test_推移テキストにspanが出る(self):
        con = _db([("FY2026-Q2", "FY2026-Q1", 2, 3742.889),
                   ("FY2026-Q3", "FY2026-Q3", 1, 1559.2)])
        t = SD.trend_text(SD.tiles(con, "T"))
        self.assertIn("span=2", t)
        self.assertIn("1559", t.replace(",", ""))


if __name__ == "__main__":
    unittest.main(verbosity=2)
