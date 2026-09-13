"""tests/test_false_positive_fixtures.py — 2026-09-13 に確定した偽陽性の回帰テスト。

フィクスチャ: tests/fixtures/false_positive_20260913.json（実DBから抽出、as_of 2026-09-13）
  3475 グッドコムアセット: 根拠期なし（別枠の位置ベース比較）で S1=1.0、DSO 0日→0日
  2776 新都HD           : 売上5倍 × DSO 113.9→30.8日、根拠書類 2025-09-11
  5136 tripla            : 対照。S2 が span_matched・根拠期 FY2026-Q2（直前四半期）で残るべき

守る不変条件（calibration_backlog §31）:
  - 根拠期が特定できない結果は点にならない
  - 直前四半期から2期以上古い根拠は点にならない
  - 売上の前年比が 2.0 倍以上/0.5 倍以下の S1/S2 は点にならない
  - 健全な根拠（直前四半期・照合可能）は残る
"""
import json
import os
import sqlite3
import sys
import unittest
from datetime import date

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, ROOT)

from screener.projection import materialize as MZ          # noqa: E402
from screener.projection import span_matched as SM          # noqa: E402
from screener.signals import span_runner as SR               # noqa: E402

FIXTURE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fixtures",
                       "false_positive_20260913.json")
AS_OF = date(2026, 9, 13)


def _external_scorers():
    from screener.report import backtest_eval as V1
    try:
        root = str(V1._external_root())
    except SystemExit:
        return None
    if root not in sys.path:
        sys.path.insert(0, root)
    from module_b.run_scorers import SCORERS_ALL
    return SCORERS_ALL


@unittest.skipUnless(os.path.exists(FIXTURE), "fixture が無い環境ではスキップ")
class TestFalsePositiveFixtures(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.scorers = _external_scorers()
        if cls.scorers is None:
            raise unittest.SkipTest("別枠 earnings_screener が無い環境")
        with open(FIXTURE, encoding="utf-8") as fh:
            fx = json.load(fh)
        con = sqlite3.connect(":memory:")
        con.row_factory = sqlite3.Row
        con.executescript(MZ.DDL)
        for table, rows in fx["projection"].items():
            for r in rows:
                cols = list(r)
                con.execute("INSERT INTO %s (%s) VALUES (%s)" % (
                    table, ",".join(cols), ",".join("?" * len(cols))), [r[c] for c in cols])
        con.commit()
        cls.con = con

    def setUp(self):
        SM.clear_cache()

    def score(self, code, policy):
        return SR.score_ticker(self.con, code, self.scorers, as_of=AS_OF, policy=policy)

    # ---- 修正前の挙動を記録（フィクスチャが症例を再現していることの確認） ----
    def test_修正前_3475はS1が根拠期なしで1点(self):
        r = self.score("3475", "prefer_span")
        self.assertTrue(r["S1"]["available"])
        self.assertEqual(r["_source"]["S1"], "external")
        self.assertFalse(r["S1"].get("period"))
        self.assertGreaterEqual(r["S1"]["score"], 0.99)

    def test_修正前_2776はS1が売上5倍で1点(self):
        r = self.score("2776", "prefer_span")
        self.assertTrue(r["S1"]["available"])
        self.assertEqual(r["S1"]["period"], "FY2026-Q2")
        self.assertGreaterEqual(r["S1"]["score"], 0.99)

    # ---- 修正後 ----
    def test_3475は根拠期なしのシグナルが点にならない(self):
        r = self.score("3475", "evidence_strict")
        for sid in ("S1", "S2"):
            self.assertFalse(r[sid]["available"], sid)
            self.assertIn("no_period", r[sid]["strict_flags"], sid)
        fired = [k for k, v in r.items() if isinstance(v, dict)
                 and v.get("available") and (v.get("score") or 0) > 0]
        for k in fired:
            self.assertTrue(r[k].get("period"), "根拠期なしで発火しているシグナル %s" % k)
        self.assertTrue(r["evidence_score"] is None or r["evidence_score"] < 0.10)

    def test_2776は売上5倍と古い根拠で点にならない(self):
        r = self.score("2776", "evidence_strict")
        s1 = r["S1"]
        self.assertFalse(s1["available"])
        flags = ",".join(s1["strict_flags"])
        self.assertIn("scope_change_suspect", flags)
        self.assertIn("stale_lag", flags)
        self.assertTrue(r["evidence_score"] is None or r["evidence_score"] < 0.10)

    def test_5136は健全な根拠で残る(self):
        r = self.score("5136", "evidence_strict")
        s2 = r["S2"]
        self.assertTrue(s2["available"], s2.get("evidence"))
        self.assertEqual(s2["period"], "FY2026-Q2")
        self.assertEqual(s2["evidence_lag_q"], 0)
        self.assertGreater(s2["score"], 0.10)
        self.assertGreaterEqual(r["evidence_score"], 0.10)


class TestRecency(unittest.TestCase):
    def test_直前四半期の決め方(self):
        # 10月期: 四半期末は 1/4/7/10 月。9/13 時点では 7/31+45日=9/14 が未到来 → 4月
        self.assertEqual(SR.expected_latest_ym(date(2026, 9, 13), 10), (2026, 4))
        self.assertEqual(SR.expected_latest_ym(date(2026, 9, 14), 10), (2026, 7))
        # 3月期: 6/30+45日=8/14 → 9/13 時点は 6月
        self.assertEqual(SR.expected_latest_ym(date(2026, 9, 13), 3), (2026, 6))

    def test_lag(self):
        self.assertEqual(SR.lag_quarters("FY2026-Q2", date(2026, 9, 13), 10), 0)   # 2026-04
        self.assertEqual(SR.lag_quarters("FY2026-Q2", date(2026, 9, 13), 1), 3)    # 2025-07 vs 2026-04
        self.assertEqual(SR.lag_quarters("FY2027-Q1", date(2026, 9, 13), 3), 0)
        self.assertEqual(SR.lag_quarters("FY2026-Q4", date(2026, 9, 13), 3), 1)
        self.assertIsNone(SR.lag_quarters("FY2026-Q4", date(2026, 9, 13), None))


if __name__ == "__main__":
    unittest.main(verbosity=2)
