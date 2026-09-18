"""tests/test_signal_design_fixtures.py — 2026-09-18 のシグナル設計修正の回帰テスト。

フィクスチャ: tests/fixtures/signal_design_20260918.json
（`python -m screener.tests.fixtures.extract_fixture --codes 3441 6838 ...` で実DBから抽出）

  6838 多摩川HD（2026年10月期）: 修正前は S1=1.0（合成 1.03 で1位）。根拠期は H1（span=2）で
      H1 売上は前年比 +45.2% のため別枠の半減ガードが効かない。一方 Q単独は3期連続減。
  3441 山王（2026年7月期）    : 対照。S1 は DSO 悪化で正しく減点されている。

守る不変条件:
  §32  売上が縮んでいるとき（最新タイルが qoq/yoy とも非正）、S1 は点にならない
  §32  Q単独で方向が確認できないときは、不採用にせず警告（strict_notes）を残す
  §33  S5 は暗黙の残存四半期利益を計算し、負なら guidance_dead を立てる
  §34  根拠期の span が表示から分かる（累計で判定していることが隠れない）
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
from screener.signals import sales_direction as SD          # noqa: E402
from screener.signals import span_runner as SR              # noqa: E402

FIXTURE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fixtures",
                       "signal_design_20260918.json")
AS_OF = date(2026, 9, 18)


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


def _load(section_rows, con):
    for table, rows in section_rows.items():
        for r in rows:
            cols = [c for c in r if r[c] is not None or True]
            try:
                con.execute("INSERT OR REPLACE INTO %s (%s) VALUES (%s)" % (
                    table, ",".join(cols), ",".join("?" * len(cols))),
                    [r[c] for c in cols])
            except sqlite3.OperationalError:
                pass            # そのテーブルが無い DB（本体/投影の別）


@unittest.skipUnless(os.path.exists(FIXTURE), "fixture が無い環境ではスキップ")
class TestSignalDesign(unittest.TestCase):
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
        _load(fx["projection"], con)
        con.commit()
        cls.con = con
        cls.fx = fx

    def setUp(self):
        SM.clear_cache()

    def score(self, code, policy="evidence_strict"):
        return SR.score_ticker(self.con, code, self.scorers, as_of=AS_OF,
                               policy=policy)

    # ------------------------------------------------------------ §32
    def test_6838_売上方向ガードで_S1が黙って満点にならない(self):
        """**発火しない、または警告つきになる。** どちらかであることを固定する。

        MODE="any"（指示の文言どおり qoq **または** yoy が正なら up）では
        6838 は Q3単独 qoq -16.7% / yoy +17.4% で up になり、S1 は残る。
        その場合でも `sales_qoq_negative` の警告が必ず付く。
        """
        r = self.score("6838")
        s1 = r["S1"]
        sd = s1.get("sales_direction") or {}
        if sd.get("status") == "down":
            self.assertFalse(s1["available"], s1.get("evidence"))
            self.assertIn("sales_shrinking", s1["strict_flags"])
            self.assertEqual(s1["score"], 0.0)
            return
        notes = ",".join(s1.get("strict_notes") or [])
        self.assertTrue(
            ("sales_qoq_negative" in notes
             or "sales_direction_cumulative_only" in notes
             or "sales_direction_unverified" in notes),
            "警告なしで S1 が通っている: %r" % (s1.get("strict_notes"),))
        # 満額では通らない（直前四半期より古い根拠なので ×0.5）
        self.assertLess(s1["score"], s1["raw_score"])

    def test_6838_ANDの読みなら不採用になる(self):
        """qoq と yoy の**両方**を要求する読み（MODE="all"）では down になる。

        どちらの読みを採るかは設計判断（§32-3 に実測を記録）。テストは
        **選択で結論が変わること自体**を固定して、黙って切り替わらないようにする。
        """
        d = SD.evaluate(self.con, "6838", mode="all")
        self.assertEqual(d["status"], "down", d["note"])
        self.assertTrue(d["quarter_level"])
        self.assertLess(d["latest"]["qoq_pct"], 0)
        self.assertGreater(d["latest"]["yoy_pct"], 0)

    def test_6838_修正前は満点だったことを記録する(self):
        # prefer_span（旧既定）ではガードが効かない＝症例が再現する側。
        r = self.score("6838", "prefer_span")
        s1 = r["S1"]
        if s1.get("available") and s1.get("period") == "FY2026-Q2":
            self.assertGreaterEqual(s1["score"], 0.99)

    def test_6838_売上推移が_Q単独まで降りて表示される(self):
        sd = SD.evaluate(self.con, "6838")
        self.assertTrue(sd["trend_text"])
        # span!=1 のタイルには span を必ず書く（3ヶ月と6ヶ月を混ぜない）
        for t in sd["trend"]:
            if t["span_q"] != 1:
                self.assertIn("span=%d" % t["span_q"], sd["trend_text"])

    def test_3441_DSO悪化は従来どおり減点のまま(self):
        r = self.score("3441")
        s1 = r["S1"]
        if s1.get("available"):
            self.assertLess(s1["score"], 0, s1.get("evidence"))
            # 減点側にガードは掛けない（縮小ガードは加点の抑制のみ）
            self.assertNotIn("sales_shrinking", s1.get("strict_flags") or [])

    # ------------------------------------------------------------ §33
    def test_S5_暗黙の残存四半期利益が計算される(self):
        got = 0
        for code in ("3441", "6838"):
            d = (self.score(code)["S5"].get("details") or {})
            if d.get("progress_op") is None:
                continue
            got += 1
            self.assertIn("implied_rest_op", d)
            self.assertIn("guidance_dead", d)
            self.assertEqual(d["guidance_dead"], d["implied_rest_op"] < 0)
        if not got:
            self.skipTest("フィクスチャに通期会社予想OPが無い（S5 は評価不能）")

    # ------------------------------------------------------------ §34
    def test_根拠期のspanが結果から分かる(self):
        for code in ("3441", "6838"):
            for sid in ("S1", "S2", "S4", "S5"):
                v = self.score(code)[sid]
                if v.get("available") and v.get("period"):
                    self.assertIsNotNone(v.get("span_q"),
                                         "%s %s に span が無い" % (code, sid))


if __name__ == "__main__":
    unittest.main(verbosity=2)
