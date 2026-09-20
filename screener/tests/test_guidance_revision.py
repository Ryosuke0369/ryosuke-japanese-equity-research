"""業績予想の修正開示を guidance に取り込む（calibration_backlog §41）。

修正開示（tse-rvfc-*）は、同じ ForecastMember の文脈に「今回の予想」と
「前回の予想」を CurrentMember / PreviousMember で分けて載せる。これを内訳
（member）として扱っていたため、予想が financials_dim に落ちて guidance に
1行も入らず、`revision_direction` は全行 `initial` のままだった。

フィクスチャは実際の開示（1382 ホーブ 2026-08-06）の値。
**営業利益の予想が +24百万円 → −22百万円（赤字転落）** で、比率で判定すると
符号を取り違える形になっている。
"""
import json
import os
import sqlite3
import unittest

from screener.extract import xbrl_parser as X

FIX = os.path.join(os.path.dirname(__file__), "fixtures", "guidance_revision_20260920.json")


def _fixture():
    with open(FIX, encoding="utf-8") as fh:
        return json.load(fh)


class TestRevisionDirection(unittest.TestCase):
    def test_direction_rules(self):
        self.assertEqual(X.revision_direction(-22e6, 24e6), "down")   # 1382 実例
        self.assertEqual(X.revision_direction(30.0, 24.0), "up")
        self.assertEqual(X.revision_direction(24.0, 24.0), "flat")
        self.assertEqual(X.revision_direction(24.0, None), "initial")
        self.assertEqual(X.revision_direction(None, 24.0), "initial")

    def test_loss_forecast_is_not_an_upgrade(self):
        # 赤字幅の縮小は up、拡大は down。絶対値で比べると逆になる
        self.assertEqual(X.revision_direction(-10.0, -50.0), "up")
        self.assertEqual(X.revision_direction(-50.0, -10.0), "down")


class TestContextAxis(unittest.TestCase):
    def test_current_previous_are_axes_not_breakdowns(self):
        m = X.load_mapping()
        for ctx, want in _fixture()["contexts"].items():
            with self.subTest(ctx=ctx):
                d = m.parse_context(ctx, "tdnet")
                self.assertEqual(d["role"], want["role"])
                self.assertEqual(d["revision"], want["revision"])
                # rest に残ると financials_dim 行になってしまう
                self.assertEqual(X.dims_of(ctx, m, "tdnet"), [])


class TestWriteGuidance(unittest.TestCase):
    def setUp(self):
        self.con = sqlite3.connect(":memory:")
        self.con.execute(
            "CREATE TABLE guidance (code TEXT, date TEXT, fy TEXT, item TEXT, "
            " value REAL, revision_direction TEXT, filing_id INTEGER, prev_value REAL,"
            " PRIMARY KEY (code, date, fy, item))")
        self.row = {"id": 1, "code": "1382", "date": "2026-08-06"}

    def _write(self, forecasts):
        X.write_guidance(self.con, self.row, forecasts)
        return {(r[0], r[1]): r[2:] for r in self.con.execute(
            "SELECT fy, item, value, prev_value, revision_direction FROM guidance")}

    def test_revision_carries_previous_and_direction(self):
        fx = _fixture()["forecasts"]
        got = self._write({tuple(k.split("|")): v for k, v in fx.items()})
        self.assertEqual(got[("FY2026", "operating_income")],
                         (-22e6, 24e6, "down"))
        self.assertEqual(got[("FY2026", "revenue")], (2372e6, 2482e6, "down"))

    def test_plain_forecast_stays_initial(self):
        got = self._write({("FY2027", "operating_income"): {"revised": 500.0}})
        self.assertEqual(got[("FY2027", "operating_income")], (500.0, None, "initial"))

    def test_range_only_forecast_is_not_written(self):
        # Upper/Lower しか無い項目を1点に潰さない
        got = self._write({("FY2027", "revenue"): {"previous": 100.0}})
        self.assertEqual(got, {})


class TestRolePrecedence(unittest.TestCase):
    """レンジ予想（Upper/Lower）が通常の予想を上書きしないこと。"""

    def test_plain_forecast_wins_over_range(self):
        m = X.load_mapping()
        plain = "CurrentYearDuration_ConsolidatedMember_ForecastMember"
        upper = "CurrentYearDuration_ConsolidatedMember_UpperMember"
        self.assertEqual(m.parse_context(plain, "tdnet")["role"], "forecast")
        self.assertEqual(m.parse_context(upper, "tdnet")["role"], "forecast_upper")
        self.assertLess(X._ROLE_RANK["forecast"], X._ROLE_RANK["forecast_upper"])


if __name__ == "__main__":
    unittest.main()
