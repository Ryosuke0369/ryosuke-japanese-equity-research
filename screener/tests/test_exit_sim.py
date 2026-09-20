"""シャドウE フェーズB（backtest_acceptance_criteria.md B-1/B-2）の出口4分岐と受容帯。

固定するもの: 受容帯の算術（一様配分・POC からの隣接拡張・VAL 割れの連続日数）、
分岐の優先順、テーゼ破綻の OR/AND、受容帯がエントリー日以降の価格を見ないこと。
"""
import sqlite3
import unittest

from screener.report import exit_sim as X
from screener.technical import acceptance_zone as AZ


def _bar(h, l, c, t):
    return {"high": h, "low": l, "close": c, "turnover": t}


class TestAcceptanceZone(unittest.TestCase):
    def test_point_bars_value_area(self):
        bars = [_bar(100, 100, 100, 50), _bar(105, 105, 105, 30), _bar(110, 110, 110, 20)]
        z = AZ.value_area(bars, n_bins=10)
        self.assertAlmostEqual(z["val"], 100.0)
        self.assertAlmostEqual(z["vah"], 106.0)
        self.assertAlmostEqual(z["share"], 0.80)
        self.assertAlmostEqual(z["poc"], 100.5)

    def test_turnover_spread_uniformly_over_range(self):
        z = AZ.value_area([_bar(110, 100, 105, 100)], n_bins=10)
        self.assertAlmostEqual(z["val"], 100.0)
        self.assertAlmostEqual(z["vah"], 107.0)
        self.assertAlmostEqual(z["share"], 0.70)

    def test_flat_range_has_no_zone(self):
        self.assertIsNone(AZ.value_area([_bar(100, 100, 100, 10)] * 3))

    def test_first_break_needs_consecutive_closes(self):
        closes = [("d1", 99), ("d2", 101), ("d3", 99), ("d4", 99), ("d5", 99)]
        self.assertIsNone(AZ.first_break(100, closes[:2], 3))
        self.assertEqual(AZ.first_break(100, closes, 3)[0], "d5")
        self.assertEqual(AZ.first_break(100, closes, 1)[0], "d1")


class TestZoneBarsArePriorOnly(unittest.TestCase):
    def setUp(self):
        self.con = sqlite3.connect(":memory:")
        self.con.execute("CREATE TABLE prices (code TEXT, date TEXT, close REAL, "
                         "adj_close REAL, high REAL, low REAL, turnover_value REAL)")
        rows = [("1111", "2026-01-%02d" % d, 200.0, 100.0, 210.0, 190.0, 1000.0)
                for d in range(1, 11)]
        rows.append(("1111", "2026-01-11", 400.0, 400.0, 420.0, 380.0, 9e9))   # エントリー日
        self.con.executemany("INSERT INTO prices VALUES (?,?,?,?,?,?,?)", rows)

    def test_excludes_entry_day_and_adjusts(self):
        bars = X.zone_bars(self.con, "1111", "2026-01-11", n=60)
        self.assertEqual(len(bars), 10)
        self.assertEqual(bars[-1]["date"], "2026-01-10")
        self.assertAlmostEqual(bars[0]["high"], 105.0)      # 210 × (100/200)
        self.assertAlmostEqual(bars[0]["close"], 100.0)


class TestThesisRule(unittest.TestCase):
    """OR は「取れたもののどれかが前年割れ」。AND は比較用（B-2b）。"""

    def _res(self, sales, op):
        have = [v for v in (sales, op) if v is not None]
        return {"break_or": bool(have) and any(v <= 0 for v in have),
                "break_and": bool(have) and all(v <= 0 for v in have),
                "evaluable": bool(have)}

    def test_or_and_and(self):
        # 3441 山王 Q4 の形: 売上は過去最大でも営業利益は前年割れ
        r = self._res(+500.0, -300.0)
        self.assertTrue(r["break_or"])
        self.assertFalse(r["break_and"])
        self.assertTrue(self._res(-1.0, -1.0)["break_and"])
        self.assertFalse(self._res(+1.0, +1.0)["break_or"])
        self.assertFalse(self._res(None, None)["evaluable"])
        self.assertTrue(self._res(None, -5.0)["break_or"])


class TestBusinessDayExitDates(unittest.TestCase):
    """出口日は別枠と同じ営業日カレンダー（祝日込み）で event_date から数える。"""

    def test_skips_weekend_and_holiday(self):
        self.assertEqual(X.bd_after("2026-09-18", 1), "2026-09-24")   # 9/21-23 は休場
        self.assertEqual(X.bd_after("2026-09-18", 2), "2026-09-25")

    def test_idx_at_snaps_forward(self):
        path = [("2026-09-18", 1.0), ("2026-09-25", 2.0)]
        self.assertEqual(X.idx_at(path, "2026-09-24"), 1)
        self.assertIsNone(X.idx_at(path, "2026-10-01"))


class TestDecideExits(unittest.TestCase):
    def setUp(self):
        # 出口 index は呼び出し側が営業日カレンダーで解決する（T+0=15 / T+1=16 / T+2=17）
        self.ix = dict(base_i=17, t1_i=16, drift_i=35)
        # 終値は既定で受容帯の上（105）。必要なテストだけ下に落とす
        self.path = [("d%02d" % i, 105.0) for i in range(40)]
        self.zone = {"val": 100.0, "vah": 110.0, "poc": 105.0}
        self.no_break = {"break_or": False, "break_and": False, "evaluable": True}
        self.broke = {"break_or": True, "break_and": False, "evaluable": True}

    def test_base_exit_is_v2(self):
        ex = X.decide_exits(self.path, self.zone, self.no_break, False, **self.ix)
        self.assertEqual(ex["E0"], (17, "基準T+2"))
        self.assertEqual(ex["E4"], (17, "基準T+2"))

    def test_thesis_break_exits_t1(self):
        ex = X.decide_exits(self.path, self.zone, self.broke, False, **self.ix)
        self.assertEqual(ex["E2"], (16, "テーゼ破綻"))
        self.assertEqual(ex["E1"], (17, "基準T+2"))       # 分岐1だけなら効かない
        self.assertEqual(ex["E4"], (16, "テーゼ破綻"))

    def test_revision_exits_t1(self):
        ex = X.decide_exits(self.path, self.zone, self.no_break, True, **self.ix)
        self.assertEqual(ex["E3"], (16, "初動利確"))
        self.assertEqual(ex["E4"], (16, "初動利確"))

    def test_structural_break_wins_and_can_fire_before_t0(self):
        path = list(self.path)
        for i in (5, 6, 7):
            path[i] = ("d%02d" % i, 90.0)                 # 3日連続で VAL 割れ
        ex = X.decide_exits(path, self.zone, self.broke, True, **self.ix)
        self.assertEqual(ex["struct_i"], 7)
        self.assertEqual(ex["E1"], (7, "構造破綻"))
        self.assertEqual(ex["E4"], (7, "構造破綻"))
        self.assertEqual(ex["E2"], (16, "テーゼ破綻"))    # 分岐2だけなら T+1 のまま

    def test_entry_below_val_is_recorded(self):
        path = [("d%02d" % i, 90.0 if i < 3 else 105.0) for i in range(40)]
        ex = X.decide_exits(path, self.zone, self.no_break, False, **self.ix)
        self.assertEqual(ex["entry_below_val"], 1)
        self.assertEqual(ex["E4"], (2, "構造破綻"))       # 実質3日で撤退（申し送り2）

    def test_no_zone_disables_branch1_only(self):
        ex = X.decide_exits(self.path, None, self.broke, False, **self.ix)
        self.assertEqual(ex["zone_ok"], 0)
        self.assertIsNone(ex["struct_i"])
        self.assertEqual(ex["E4"], (16, "テーゼ破綻"))


if __name__ == "__main__":
    unittest.main()
