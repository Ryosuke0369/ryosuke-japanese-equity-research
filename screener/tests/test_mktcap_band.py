"""P3-4 時価総額帯（docs/backtest_acceptance_criteria.md）。

境界は各時点のユニバース条件（min <= mktcap <= max）に揃える。各帯が
ちょうど「その拡大で増えた分」になっていることと、帯の判定に
エントリー当日以降の時価総額を使わないことを固定する。
"""
import sqlite3
import unittest

from screener.report import backtest_eval as V1


class TestMktcapBand(unittest.TestCase):
    def test_boundaries_match_universe_history(self):
        cases = [
            (4999, V1.BAND_OUT), (5000, "50-600"), (60000, "50-600"),
            (60001, "600-1000"), (100000, "600-1000"),
            (100001, "1000-3000"), (300000, "1000-3000"),
            (300001, V1.BAND_OUT), (None, V1.BAND_UNKNOWN),
        ]
        for mc, want in cases:
            with self.subTest(mktcap=mc):
                self.assertEqual(V1.mktcap_band(mc), want)

    def test_top_band_matches_universe_rules(self):
        from screener import common as C
        size = C.load_yaml("universe_rules.yaml")["size"]
        self.assertEqual(V1.MKTCAP_BANDS[0][1], size["mktcap_min_mn"])
        self.assertEqual(V1.MKTCAP_BANDS[-1][2], size["mktcap_max_mn"])


class TestMktcapAtEntry(unittest.TestCase):
    def setUp(self):
        self.con = sqlite3.connect(":memory:")
        self.con.execute("CREATE TABLE prices (code TEXT, date TEXT, mktcap REAL)")
        self.con.executemany("INSERT INTO prices VALUES (?,?,?)", [
            ("1111", "2023-01-04", 50000),
            ("1111", "2023-01-05", None),       # 欠測日は飛ばす
            ("1111", "2023-01-06", 150000),     # エントリー当日: 使わない
        ])

    def test_uses_prior_day_only(self):
        self.assertEqual(V1.mktcap_at_entry(self.con, "1111", "2023-01-06"), 50000)

    def test_unknown_when_no_history(self):
        self.assertIsNone(V1.mktcap_at_entry(self.con, "1111", "2023-01-04"))
        self.assertIsNone(V1.mktcap_at_entry(self.con, "9999", "2023-01-06"))

    def test_tag_trades(self):
        trades = [{"ticker": "1111", "entry_date": "2023-01-06"},
                  {"ticker": "9999", "entry_date": "2023-01-06"}]
        V1.tag_mktcap_bands(trades, self.con)
        self.assertEqual(trades[0]["mktcap_band"], "50-600")
        self.assertEqual(trades[1]["mktcap_band"], V1.BAND_UNKNOWN)


if __name__ == "__main__":
    unittest.main()
