"""シャドウL: 「修正が出たら翌営業日の始値」「出なければ N 営業日後の終値」を固定する。"""
import unittest

from screener.report import s5c_event_exit as L

CAL = ["2025-01-06", "2025-01-07", "2025-01-08", "2025-01-09", "2025-01-10",
       "2025-01-14", "2025-01-15", "2025-01-16", "2025-01-17", "2025-01-20",
       "2025-01-21"]
# (始値, 終値)
PX = {"1111": {d: (100.0 + i, 101.0 + i) for i, d in enumerate(CAL)}}
IDX = {d: 1000.0 for d in CAL}


class TestExitRules(unittest.TestCase):
    def setUp(self):
        self.entries = [{"as_of": "2025-01-06", "code": "1111", "r": 1.5}]

    def test_no_revision_exits_at_n_days_close(self):
        rows, why = L.simulate(self.entries, {}, PX, IDX, CAL, 5)
        r = rows[0]
        self.assertEqual(r["exit_date"], CAL[5])
        self.assertEqual(r["exit_by"], "N日")
        self.assertAlmostEqual(r["ret_gross"], PX["1111"][CAL[5]][1] / PX["1111"][CAL[0]][1] - 1)
        self.assertEqual(why["N日で降りた"], 1)

    def test_revision_exits_next_day_open(self):
        evs = {"1111": [("2025-01-08", "up", 0.4)]}
        rows, why = L.simulate(self.entries, evs, PX, IDX, CAL, 10)
        r = rows[0]
        self.assertEqual(r["exit_date"], "2025-01-09")          # 開示の翌営業日
        self.assertEqual(r["exit_by"], "修正")
        self.assertAlmostEqual(r["ret_gross"],
                               PX["1111"]["2025-01-09"][0] / PX["1111"][CAL[0]][1] - 1)
        self.assertEqual(why["修正で降りた"], 1)

    def test_revision_after_the_window_does_not_trigger(self):
        evs = {"1111": [("2025-01-20", "up", 0.4)]}
        rows, _why = L.simulate(self.entries, evs, PX, IDX, CAL, 5)
        self.assertEqual(rows[0]["exit_by"], "N日")

    def test_down_revision_triggers_in_main_but_not_in_up_only(self):
        evs = {"1111": [("2025-01-08", "down", -0.4)]}
        main, _ = L.simulate(self.entries, evs, PX, IDX, CAL, 10)
        self.assertEqual(main[0]["exit_by"], "修正")
        up_only, _ = L.simulate(self.entries, evs, PX, IDX, CAL, 10, up_only=True)
        self.assertEqual(up_only[0]["exit_by"], "N日")

    def test_cost_and_excess(self):
        rows, _ = L.simulate(self.entries, {}, PX, IDX, CAL, 5)
        r = rows[0]
        self.assertAlmostEqual(r["ret_net"], r["ret_gross"] - L.COST)
        self.assertAlmostEqual(r["excess"], r["ret_net"] - 0.0)   # 指数は横ばい

    def test_registered_values(self):
        self.assertEqual(L.N_DAYS, (5, 10, 20))
        self.assertEqual(L.COST, 0.004)


if __name__ == "__main__":
    unittest.main()
