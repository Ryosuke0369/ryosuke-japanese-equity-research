"""シャドウE フェーズA（backtest_acceptance_criteria.md E-A）のイベント反応記録。

分類・T日・窓の算術・分割をまたぐ連鎖・Point-in-Time（T+0 以降を書き換えても
特徴量が変わらない）を固定する。分類と T日の例は fixtures/technical_event_20260919.json。
"""
import json
import os
import unittest

from screener.technical import event_response as E

FIX = os.path.join(os.path.dirname(__file__), "fixtures", "technical_event_20260919.json")


def _fixture():
    with open(FIX, encoding="utf-8") as fh:
        return json.load(fh)


class TestClassify(unittest.TestCase):
    def test_fixture_titles(self):
        for c in _fixture()["classify"]:
            with self.subTest(title=c["title"]):
                self.assertEqual(E.classify(c["subtype"], c["title"]), c["want"])

    def test_order_rule_needs_title_only(self):
        # 決算系の subtype が付いた行は受注語を含んでも受注関連にしない
        self.assertEqual(E.classify("決算説明資料", "受注状況のご説明"), E.TYPE_OTHER)


class TestTiming(unittest.TestCase):
    def test_fixture_cases(self):
        fx = _fixture()
        cal = fx["timing_calendar"]
        for c in fx["timing"]:
            with self.subTest(case=c):
                self.assertEqual(list(E.timing_and_t0(c["date"], c["time"], cal)), c["want"])


class TestBuildEvents(unittest.TestCase):
    CAL = ["2026-08-06", "2026-08-07", "2026-08-10", "2026-08-11"]

    def test_dedup_keeps_earliest_and_marks_concurrent(self):
        raw = [
            {"code": "1111", "date": "2026-08-06", "time": "16:00", "subtype": "決算短信",
             "title": "2027年３月期 第１四半期決算短信〔日本基準〕（連結）", "source": "filings"},
            {"code": "1111", "date": "2026-08-06", "time": "15:30", "subtype": "決算短信",
             "title": "2027年３月期 第１四半期決算短信〔日本基準〕（連結）", "source": "filings"},
            {"code": "1111", "date": "2026-08-06", "time": "15:30", "subtype": "業績予想修正",
             "title": "業績予想の修正に関するお知らせ", "source": "filings"},
            {"code": "1111", "date": "2026-08-06", "time": "15:30", "subtype": "決算説明資料",
             "title": "決算説明資料", "source": "filings"},
            {"code": "2222", "date": "2026-08-07", "time": "10:00", "subtype": "配当予想修正",
             "title": "配当予想の修正に関するお知らせ", "source": "filings"},
        ]
        events, other = E.build_events(raw, self.CAL)
        self.assertEqual(len(events), 3)
        q = [e for e in events if e["type"] == E.TYPE_Q][0]
        self.assertEqual((q["time"], q["timing"], q["t0"]), ("15:30", "引け後", "2026-08-07"))
        self.assertEqual(q["concurrent"], E.TYPE_REV)
        d = [e for e in events if e["type"] == E.TYPE_DIV][0]
        self.assertEqual((d["t0"], d["concurrent"]), ("2026-08-07", ""))
        self.assertEqual(other["決算説明資料"], 1)


class TestSeries(unittest.TestCase):
    def test_split_does_not_break_chain(self):
        cal = ["d0", "d1", "d2", "d3"]
        idx = {d: i for i, d in enumerate(cal)}
        rows = [("d0", 200, 10, 2000, 1.0), ("d1", 200, 10, 2000, 1.0),
                ("d2", 100, 20, 2000, 0.5), ("d3", 100, 20, 2000, 1.0)]
        s = E.build_series(rows, idx, len(cal))
        self.assertEqual(s["px"], [1.0, 1.0, 1.0, 1.0])
        self.assertEqual(s["vol"], [10, 10, 10, 10])
        self.assertEqual(s["close"], [200, 200, 100, 100])

    def test_no_trade_day_and_after_last_row(self):
        cal = ["d0", "d1", "d2", "d3", "d4"]
        idx = {d: i for i, d in enumerate(cal)}
        rows = [("d1", 100, 5, 500, 1.0), ("d3", 110, 5, 550, 1.0)]
        s = E.build_series(rows, idx, len(cal))
        self.assertEqual(s["px"][0], None)          # 上場前
        self.assertEqual(s["px"][2], 1.0)           # 売買なし: 価格は直前値
        self.assertEqual(s["vol"][2], 0.0)          # 出来高は 0
        self.assertAlmostEqual(s["px"][3], 1.1)
        self.assertEqual(s["px"][4], None)          # 最終行より後は捏造しない


def _synthetic(idx0=40, n=61):
    cal = ["D%03d" % i for i in range(n)]
    close = []
    for i in range(n):
        if i <= idx0 - 16:
            close.append(100.0)
        elif i <= idx0 - 1:
            close.append(110.0)                     # ① +10%
        elif i <= idx0 + 1:
            close.append(121.0)                     # ② +10%
        else:
            close.append(133.1)                     # ③ +10%
    vol = [1000.0 if idx0 - 36 <= i <= idx0 - 17 else 3000.0 for i in range(n)]
    rows = [(cal[i], close[i], vol[i], close[i] * vol[i], 1.0) for i in range(n)]
    idx = {d: i for i, d in enumerate(cal)}
    return cal, idx, rows


class TestWindow(unittest.TestCase):
    def setUp(self):
        self.idx0 = 40
        self.cal, self.idx, self.rows = _synthetic(self.idx0)
        n = len(self.cal)
        self.flat = [1000.0] * n
        self.univ = [1.0] * n
        self.ev = {"event_id": "1111_x_D040", "code": "1111", "type": E.TYPE_Q,
                   "timing": "引け後", "t0": self.cal[self.idx0], "concurrent": ""}

    def _measure(self, rows):
        panel = {"1111": E.build_series(rows, self.idx, len(self.cal))}
        ev_rows, long_rows = E.measure([dict(self.ev)], panel, self.flat, self.univ, self.cal)
        return ev_rows[0], long_rows

    def test_intervals(self):
        r, long_rows = self._measure(self.rows)
        self.assertEqual(r["window_ok"], 1)
        self.assertAlmostEqual(r["f_pre_rel_topix"], 0.10)
        self.assertAlmostEqual(r["y_init_rel_topix"], 0.10)
        self.assertAlmostEqual(r["y_drift_rel_topix"], 0.10)
        self.assertAlmostEqual(r["f_vol_base"], 1000.0)
        self.assertAlmostEqual(r["f_vol_ratio_tm1"], 3.0)
        self.assertEqual([x["offset"] for x in long_rows], list(range(-16, 21)))
        self.assertEqual({x["role"] for x in long_rows if x["offset"] < 0}, {"feature"})
        self.assertEqual({x["role"] for x in long_rows if x["offset"] >= 0}, {"outcome"})

    def test_relative_to_benchmark(self):
        n = len(self.cal)
        bm = [1000.0 if i <= self.idx0 - 16 else 1050.0 for i in range(n)]   # 窓の中で +5%
        panel = {"1111": E.build_series(self.rows, self.idx, n)}
        ev_rows, _ = E.measure([dict(self.ev)], panel, bm, self.univ, self.cal)
        self.assertAlmostEqual(ev_rows[0]["f_pre_rel_topix"], 0.05)
        self.assertAlmostEqual(ev_rows[0]["y_init_rel_topix"], 0.10)

    def test_missing_t20_leaves_drift_empty(self):
        cal, idx, rows = _synthetic(self.idx0, n=55)                        # T+14 まで
        panel = {"1111": E.build_series(rows, idx, len(cal))}
        ev_rows, _ = E.measure([dict(self.ev)], panel, [1000.0] * 55, [1.0] * 55, cal)
        self.assertIsNone(ev_rows[0]["y_drift_rel_topix"])
        self.assertIsNotNone(ev_rows[0]["y_drift10_rel_topix"])

    def test_features_do_not_see_t0_or_later(self):
        base, _ = self._measure(self.rows)
        tampered = [(d, c * (3.0 if self.idx[d] >= self.idx0 else 1.0),
                     v * (50.0 if self.idx[d] >= self.idx0 else 1.0), t, f)
                    for d, c, v, t, f in self.rows]
        other, _ = self._measure(tampered)
        feats = [k for k in E.EVENT_COLS if k.startswith("f_")]
        self.assertTrue(feats)
        for k in feats:
            with self.subTest(feature=k):
                self.assertEqual(base[k], other[k])
        self.assertNotEqual(base["y_init_rel_topix"], other["y_init_rel_topix"])


class TestDescribe(unittest.TestCase):
    def test_quartiles_and_cluster_t(self):
        pairs = [(0.01, "a"), (0.02, "a"), (0.03, "b"), (0.04, "b"), (None, "c")]
        st = E.describe(pairs)
        self.assertEqual(st["n"], 4)
        self.assertAlmostEqual(st["median"], 0.025)
        self.assertAlmostEqual(st["q1"], 0.0175)
        self.assertEqual(st["n_dates"], 2)
        self.assertEqual(st["pos"], 1.0)


if __name__ == "__main__":
    unittest.main()
