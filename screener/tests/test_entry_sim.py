"""シャドウF（入口側）の採否フックと、v2 の挙動が変わらないことを固定する。"""
import sqlite3
import unittest

from screener.report import backtest_v2 as V2


def _trade(ticker, day, score, ret, exit_date):
    return {"ticker": ticker, "entry_date": day, "event_date": day, "score": score,
            "ret": ret, "entry_n": 15, "exit_k": 2, "exit_date": exit_date}


class _Con:
    """market_index だけを持つ最小の接続。"""

    def __init__(self, dates, tickers=()):
        self.con = sqlite3.connect(":memory:")
        self.con.execute("CREATE TABLE market_index (date TEXT, close REAL)")
        self.con.executemany("INSERT INTO market_index VALUES (?,?)", dates)
        # simulate は exit 日を価格系列から引く（V1._exit_date）
        self.con.execute("CREATE TABLE daily_prices (ticker TEXT, date TEXT, close REAL)")
        self.con.executemany(
            "INSERT INTO daily_prices VALUES (?,?,?)",
            [(t, d, c) for t in tickers for d, c in dates])

    def execute(self, *a):
        return self.con.execute(*a)


class TestAdmitHook(unittest.TestCase):
    def setUp(self):
        # 市場フィルターを通すために 200本以上の上昇系列を置く
        dates = [("2026-01-%02d" % (i + 1) if i < 9 else "2026-%02d-%02d"
                  % (1 + (i // 28), (i % 28) + 1), 100.0 + i) for i in range(260)]
        self.tickers = ["1111", "2222"] + ["%04d" % i for i in range(20)]
        self.con = _Con(dates, self.tickers)
        self.trades = [_trade("1111", "2026-08-01", 0.5, 0.10, "2026-08-05"),
                       _trade("2222", "2026-08-01", 0.4, -0.10, "2026-08-05")]

    def _sim(self, admit=None):
        return V2.simulate(self.trades, self.con, 15, 2, use_filter=False, admit=admit)

    def test_default_is_unchanged_v2(self):
        sim = self._sim()
        self.assertEqual(sim["n_taken"], 2)
        self.assertEqual(sim["skipped_admit"], 0)

    def test_admit_can_skip(self):
        sim = self._sim(admit=lambda t: None if t["ticker"] == "2222" else V2.POS_FRACTION)
        self.assertEqual([t["ticker"] for t in sim["trades"]], ["1111"])
        self.assertEqual(sim["skipped_admit"], 1)

    def test_skipped_candidate_frees_the_slot(self):
        # 枠は「建てた数」で埋まる。見送った候補は枠を消費しない
        many = [_trade("%04d" % i, "2026-08-01", 0.5, 0.01, "2026-08-05")
                for i in range(20)]
        sim = V2.simulate(many, self.con, 15, 2, use_filter=False,
                          admit=lambda t: None if int(t["ticker"]) < 5 else V2.POS_FRACTION)
        self.assertEqual(sim["n_taken"], V2.MAX_POSITIONS)
        self.assertEqual([t["ticker"] for t in sim["trades"]][0], "0005")

    def test_sizing_changes_allocation_not_count(self):
        full = self._sim(admit=lambda t: V2.POS_FRACTION)
        half = self._sim(admit=lambda t: 0.05)
        self.assertEqual(full["n_taken"], half["n_taken"])
        self.assertNotEqual(full["nav_final"], half["nav_final"])


class TestVerdict(unittest.TestCase):
    """サイジングだけの変種は期待値が動かない。0 を「マイナス一致」と書かない。"""

    def test_zero_difference(self):
        from screener.report import entry_sim as F
        self.assertIn("差なし", F.verdict([0.0, 0.0, 0.0]))

    def test_signs(self):
        from screener.report import entry_sim as F
        self.assertIn("プラス", F.verdict([0.01, 0.02, 0.003]))
        self.assertIn("マイナス", F.verdict([-0.01, -0.02, -0.003]))
        self.assertIn("符号不定", F.verdict([0.01, -0.02, 0.003]))
        self.assertIn("3本", F.verdict([0.01]))


if __name__ == "__main__":
    unittest.main()
