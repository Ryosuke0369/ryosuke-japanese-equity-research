"""screener/tests/test_prices_backfill.py — 契約範囲外の日をまたいで続行できること。

このファイルが守る規則: **「データが無い」と「契約範囲外」と「推測で埋める」は
三つとも別物**。J-Quants Light は5年ローリングなので、日付が変わるだけで窓の
古い端が1営業日落ちる。2026-09-01 00:00:57、既定の --from 2021-08-31 が窓から
落ちた直後に初日で 400 を踏み、株価バックフィルは1行も書かずに1秒で終了した。

400 を握りつぶして空で埋めるのは論外。だが API が message で契約範囲を明示して
いるのに、それを読まずに全体を止めるのも同じくらい高くつく —— 一晩まるごと
無駄になった。読めたら繰り上げて続行、読めなければ止まる、が正しい分岐。
"""
from __future__ import annotations

import os
import sys
import tempfile
import unittest
from datetime import date

sys.path.insert(0, os.path.dirname(os.path.dirname(
    os.path.dirname(os.path.abspath(__file__)))))

from screener import common as C
from screener.fetch import jquants_prices as P


class _Throttle:
    min_interval = 1.3


class _Fetcher:
    """jq_get を差し替えるので HTTP は踏まない。ログが触る属性だけ持つ。"""
    throttle = _Throttle()
    n_requests = 0


def _bar(code="1301", close=100.0):
    return {"Code": code, "C": close, "Vo": 1000, "Va": 100000,
            "O": 99.0, "H": 101.0, "L": 98.0,
            "AdjFactor": 1.0, "AdjC": close, "AdjVo": 1000, "MktCap": 1e9}


class CoveredRangeTest(unittest.TestCase):
    def test_reads_the_real_400_message(self):
        msg = ("/equities/bars/daily: HTTP 400 — パラメータ誤り (message: Your "
               "subscription covers the following dates: 2021-09-01 ~ . If you "
               "want more data, please check other plans:https://example.invalid)")
        self.assertEqual(P.covered_range(msg), (date(2021, 9, 1), None))

    def test_reads_both_ends(self):
        self.assertEqual(
            P.covered_range("covers the following dates: 2021-09-01 ~ 2026-06-08"),
            (date(2021, 9, 1), date(2026, 6, 8)))

    def test_unreadable_message_is_not_guessed(self):
        self.assertIsNone(P.covered_range("HTTP 400 — 何か別の理由"))
        self.assertIsNone(P.covered_range(""))


class BackfillTest(unittest.TestCase):
    def setUp(self):
        fd, self.db = tempfile.mkstemp(suffix=".db")
        os.close(fd)
        self.con = C.init_db(self.db)
        self._real = P.jq_get

    def tearDown(self):
        P.jq_get = self._real
        self.con.close()
        for s in ("", "-wal", "-shm"):
            try:
                os.remove(self.db + s)
            except OSError:
                pass

    def _install(self, fn):
        P.jq_get = fn

    def test_out_of_window_days_are_skipped_not_fatal(self):
        """窓から落ちた初日で全体を止めない。範囲内の日は最後まで取る。"""
        lo = date(2026, 1, 5)
        seen = []

        def fake(fetcher, path, **kw):
            d = date(int(kw["date"][:4]), int(kw["date"][4:6]), int(kw["date"][6:]))
            if d < lo:
                raise RuntimeError(
                    "/equities/bars/daily: HTTP 400 — パラメータ誤り (message: "
                    "Your subscription covers the following dates: "
                    f"{lo.isoformat()} ~ .)")
            seen.append(d)
            return [_bar()]

        self._install(fake)
        r = P.backfill_prices(self.con, _Fetcher(), date(2026, 1, 1),
                              date(2026, 1, 9), None)
        # 1/1,1/2 が範囲外、1/5..1/9 の5営業日が対象。
        self.assertEqual(r["skipped"], 2)
        self.assertEqual(r["days"], 5, "範囲外の初日で全体が止まっている")
        self.assertEqual(seen, [date(2026, 1, 5), date(2026, 1, 6), date(2026, 1, 7),
                                date(2026, 1, 8), date(2026, 1, 9)])
        n = self.con.execute("SELECT COUNT(*) FROM prices").fetchone()[0]
        self.assertEqual(n, 5)

    def test_unreadable_400_still_stops(self):
        """範囲が読めない 400 は従来どおり中断する（推測で埋めない）。"""
        def fake(fetcher, path, **kw):
            raise RuntimeError("/equities/bars/daily: HTTP 400 — 理由不明")

        self._install(fake)
        r = P.backfill_prices(self.con, _Fetcher(), date(2026, 1, 1),
                              date(2026, 1, 9), None)
        self.assertEqual(r["days"], 0)
        self.assertEqual(r["skipped"], 0)
        self.assertEqual(self.con.execute("SELECT COUNT(*) FROM prices").fetchone()[0], 0)

    def test_past_the_far_end_stops(self):
        """後端を越えた側では打ち切る（繰り上げ先が無い）。"""
        hi = date(2026, 1, 6)

        def fake(fetcher, path, **kw):
            d = date(int(kw["date"][:4]), int(kw["date"][4:6]), int(kw["date"][6:]))
            if d > hi:
                raise RuntimeError(
                    "/equities/bars/daily: HTTP 400 (message: Your subscription "
                    f"covers the following dates: 2021-09-01 ~ {hi.isoformat()} .)")
            return [_bar()]

        self._install(fake)
        r = P.backfill_prices(self.con, _Fetcher(), date(2026, 1, 1),
                              date(2026, 1, 9), None)
        self.assertEqual(r["days"], 4)   # 1/1,1/2,1/5,1/6
        self.assertEqual(r["skipped"], 0)

    def test_rows_with_null_adj_close_are_refetched(self):
        """行があるだけでは取得済みにしない。新しい列が NULL の日は取り直す。

        2026-08-31 の旧ジョブが close/volume だけを書いた 235 日を、
        全項目版が拾い直せるかどうかがここに懸かっている。
        """
        self.con.execute(
            "INSERT INTO prices (code, date, close, volume) VALUES (?,?,?,?)",
            ("1301", "2026-01-05", 100.0, 1000))
        self.con.commit()
        seen = []

        def fake(fetcher, path, **kw):
            seen.append(kw["date"])
            return [_bar()]

        self._install(fake)
        P.backfill_prices(self.con, _Fetcher(), date(2026, 1, 5), date(2026, 1, 5), None)
        self.assertEqual(seen, ["20260105"], "adj_close が NULL の日を飛ばしている")
        row = self.con.execute(
            "SELECT adj_close, open FROM prices WHERE date='2026-01-05'").fetchone()
        self.assertIsNotNone(row[0])
        self.assertIsNotNone(row[1])

    def test_fully_fetched_day_is_skipped(self):
        """全項目が埋まっている日は二度と叩かない（差分取得が効いていること）。"""
        self.con.execute(
            "INSERT INTO prices (code, date, close, volume, open, high, low, "
            " adj_factor, adj_close, adj_volume, mktcap) "
            "VALUES ('1301','2026-01-05',100,1000,99,101,98,1.0,100,1000,1e9)")
        self.con.commit()
        seen = []

        def fake(fetcher, path, **kw):
            seen.append(kw["date"])
            return [_bar()]

        self._install(fake)
        P.backfill_prices(self.con, _Fetcher(), date(2026, 1, 5), date(2026, 1, 5), None)
        self.assertEqual(seen, [], "取得済みの日を取り直している")


if __name__ == "__main__":
    unittest.main()
