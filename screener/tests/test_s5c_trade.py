"""シャドウK: 建て方・出口・コスト・集計の定義を固定する。"""
import unittest
from unittest import mock

from screener.report import s5c_trade as K


class TestSummarise(unittest.TestCase):
    def setUp(self):
        # 2判定日 × 2銘柄。判定日ごとの平均は +5% と -1%
        self.rows = [
            {"as_of": "2025-01-06", "code": "1111", "ret_net": 0.10, "excess": 0.08,
             "r": 1.4, "fired": 1, "rev_in_hold": "up"},
            {"as_of": "2025-01-06", "code": "2222", "ret_net": 0.00, "excess": -0.01,
             "r": 1.6, "fired": 1, "rev_in_hold": "none"},
            {"as_of": "2025-02-03", "code": "1111", "ret_net": -0.02, "excess": -0.03,
             "r": 2.5, "fired": 1, "rev_in_hold": "none"},
            {"as_of": "2025-02-03", "code": "2222", "ret_net": 0.00, "excess": 0.00,
             "r": 1.35, "fired": 1, "rev_in_hold": "none"},
        ]

    def test_clustered_average_is_by_judgment_day(self):
        m, t, nd = K.clustered(self.rows)
        self.assertAlmostEqual(m, (0.05 + -0.01) / 2)
        self.assertEqual(nd, 2)

    def test_mdd_uses_monthly_equal_weight_curve(self):
        dd, nav = K.mdd(self.rows)
        self.assertAlmostEqual(nav, 1.05 * 0.99)
        self.assertAlmostEqual(dd, -0.01)

    def test_summary_fields(self):
        s = K.summarise("x", self.rows)
        self.assertEqual((s["n"], s["codes"]), (4, 2))
        self.assertAlmostEqual(s["mean"], 0.02)
        self.assertAlmostEqual(s["win"], 0.25)


class TestBuildTrades(unittest.TestCase):
    """出口は次の判定日の終値。コストは往復0.4%。価格が無ければ落とす。"""

    def setUp(self):
        self.dates = ["2025-01-06", "2025-02-03", "2025-03-03"]
        self.px = {"1111": {"2025-01-06": 100.0, "2025-02-03": 110.0, "2025-03-03": 99.0},
                   "2222": {"2025-01-06": 50.0}}          # 出口の価格が無い
        self.idx = {"2025-01-06": 1000.0, "2025-02-03": 1020.0, "2025-03-03": 1010.0}
        self.evs = {"1111": [("2025-01-20", "up", 0.4)]}

    def _run(self):
        def fake_eval(con, code, as_of):
            return {"available": True, "r": 1.5, "fired": True, "period": "2Q",
                    "median_progress": 0.5}
        with mock.patch.object(K.P, "evaluate", fake_eval):
            return K.build_trades(None, self.dates, ["1111", "2222"], self.evs,
                                  self.px, self.idx)

    def test_exit_is_next_judgment_day_and_cost_is_applied(self):
        rows, miss, _no_r = self._run()
        first = [r for r in rows if r["as_of"] == "2025-01-06"][0]
        self.assertEqual(first["exit_date"], "2025-02-03")
        self.assertAlmostEqual(first["ret_gross"], 0.10)
        self.assertAlmostEqual(first["ret_net"], 0.10 - K.COST)
        self.assertAlmostEqual(first["bench"], 0.02)
        self.assertAlmostEqual(first["excess"], 0.10 - K.COST - 0.02)
        # 建玉を作るのは dates[:-1] の2判定日。2222 はどちらも出口の価格が無い
        self.assertEqual(miss, 2)
        self.assertEqual({r["code"] for r in rows}, {"1111"})

    def test_last_judgment_day_has_no_trade(self):
        rows, _m, _n = self._run()
        self.assertEqual({r["as_of"] for r in rows}, {"2025-01-06", "2025-02-03"})

    def test_revision_inside_holding_is_recorded(self):
        rows, _m, _n = self._run()
        first = [r for r in rows if r["as_of"] == "2025-01-06"][0]
        self.assertEqual(first["rev_in_hold"], "up")
        second = [r for r in rows if r["as_of"] == "2025-02-03"][0]
        self.assertEqual(second["rev_in_hold"], "none")

    def test_cost_matches_v2(self):
        self.assertEqual(K.COST, 0.004)


if __name__ == "__main__":
    unittest.main()
