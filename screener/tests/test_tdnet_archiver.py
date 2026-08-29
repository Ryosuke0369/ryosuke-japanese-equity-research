"""screener/tests/test_tdnet_archiver.py

Covers the parts of §2-1 that are easy to get wrong and impossible to notice
later: the classification filter, and the promise that a day which fails is
RECORDED rather than skipped. TDnet keeps ~1 month, so "we thought we had it"
is the expensive failure.

    python -m screener.tests.test_tdnet_archiver
    python -m unittest discover -s screener/tests -t .
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
from screener.fetch import tdnet_archiver as T


class TestClassify(unittest.TestCase):
    def test_tanshin(self):
        for title in (
            "2027年３月期第１四半期決算短信〔日本基準〕(連結)",
            "2026年12月期 中間決算短信〔日本基準〕（非連結）",
            "2026年７月期　決算短信〔ＩＦＲＳ〕（連結）",
            "（訂正）「2026年３月期 決算短信」の一部訂正について",
        ):
            self.assertEqual(T.classify(title)[0], "決算短信", title)
            self.assertEqual(T.classify(title)[1], "短信", title)

    def test_forecast_revision(self):
        for title in (
            "2026年７月期通期業績予想の修正に関するお知らせ",
            "業績予想の修正に関するお知らせ",
            "通期業績見通しの変更について",
        ):
            self.assertEqual(T.classify(title)[0], "業績予想修正", title)

    def test_dividend_revision(self):
        for title in (
            "2026年12月期の期末配当予想の修正（無配）に関するお知らせ",
            "配当予想の修正に関するお知らせ",
        ):
            self.assertEqual(T.classify(title)[0], "配当予想修正", title)

    def test_out_of_scope(self):
        # 「約束」(MOU/資本提携/自己株取得) は §8 のとおりシグナル対象外。
        for title in (
            "SBIホールディングスとの資本業務提携契約の締結等に関する補足説明資料",
            "自己株式の取得中止及び取得状況に関するお知らせ",
            "代表取締役の異動に関するお知らせ",
        ):
            self.assertEqual(T.classify(title), (None, None), title)


class TestListParsing(unittest.TestCase):
    HTML = """
    <html><body>
      <div id="pager-box-top">1～2件 / 全2件</div>
      <table id="main-list-table">
        <tr>
          <td class="kjTime">15:00</td><td class="kjCode">61180</td>
          <td class="kjName">アイダ</td>
          <td class="kjTitle"><a href="140120260828527542.pdf">2027年３月期第１四半期決算短信</a></td>
          <td class="kjXbrl"><a href="081220260828527542.zip">XBRL</a></td>
          <td class="kjPlace">東</td><td class="kjHistroy"></td>
        </tr>
        <tr>
          <td class="kjTime">15:30</td><td class="kjCode">398A0</td>
          <td class="kjName">テスト</td>
          <td class="kjTitle"><a href="140120260828527885.pdf">代表取締役の異動</a></td>
          <td class="kjXbrl"></td>
          <td class="kjPlace">東</td><td class="kjHistroy"></td>
        </tr>
      </table>
    </body></html>
    """

    def test_rows_and_total(self):
        rows, total = T.parse_list_page(self.HTML)
        self.assertEqual(total, 2)
        self.assertEqual(len(rows), 2)
        self.assertEqual(rows[0]["zip"], "081220260828527542.zip")
        self.assertIsNone(rows[1]["zip"])

    def test_empty_page_is_not_an_error(self):
        rows, total = T.parse_list_page("<html><body></body></html>")
        self.assertEqual((rows, total), ([], 0))


class TestCodeNormalisation(unittest.TestCase):
    def test_five_char_padding_is_stripped(self):
        self.assertEqual(C.normalise_code("61180"), "6118")
        self.assertEqual(C.normalise_code("398A0"), "398A")

    def test_four_char_left_alone(self):
        self.assertEqual(C.normalise_code("6118"), "6118")
        self.assertEqual(C.normalise_code("285A"), "285A")

    def test_trailing_zero_only_stripped_at_length_five(self):
        # 7203 must not become 720; a genuine 5-char code ending in a non-zero
        # (there are none today, but the rule must not silently mangle one).
        self.assertEqual(C.normalise_code("7203"), "7203")
        self.assertEqual(C.normalise_code("12345"), "12345")


class TestMissingDayRecording(unittest.TestCase):
    """The §2-1 requirement: 取得失敗日はDBに欠損記録(黙って飛ばさない)."""

    def setUp(self):
        fd, self.db = tempfile.mkstemp(suffix=".db")
        os.close(fd)
        self.con = C.init_db(self.db)

    def tearDown(self):
        self.con.close()
        for suffix in ("", "-wal", "-shm"):
            try:
                os.remove(self.db + suffix)
            except OSError:
                pass

    def test_failed_day_is_recorded_and_still_counts_as_missing(self):
        run_id = C.start_run(self.con, "tdnet", "2026-08-27")
        C.finish_run(self.con, run_id, "failed", error="HTTP 503")

        row = self.con.execute(
            "SELECT status, error, attempt FROM fetch_runs WHERE id=?", (run_id,)
        ).fetchone()
        self.assertEqual(row["status"], "failed")
        self.assertEqual(row["error"], "HTTP 503")

        missing = C.missing_days(self.con, "tdnet",
                                 date(2026, 8, 27), date(2026, 8, 27))
        self.assertEqual(missing, ["2026-08-27"],
                         "a failed day must still be reported as missing")

    def test_retry_increments_attempt_and_clears_the_gap(self):
        C.finish_run(self.con, C.start_run(self.con, "tdnet", "2026-08-27"),
                     "failed", error="HTTP 503")
        rid2 = C.start_run(self.con, "tdnet", "2026-08-27")
        C.finish_run(self.con, rid2, "ok", n_listed=167, n_target=9, n_saved=9)

        self.assertEqual(
            self.con.execute("SELECT attempt FROM fetch_runs WHERE id=?",
                             (rid2,)).fetchone()["attempt"], 2)
        self.assertEqual(
            C.missing_days(self.con, "tdnet", date(2026, 8, 27), date(2026, 8, 27)),
            [], "a successful retry must close the gap")

    def test_empty_day_is_covered_not_missing(self):
        """A holiday returns zero rows. That is 'empty', not a failure — if it
        counted as missing the backfill would re-fetch holidays for ever."""
        C.finish_run(self.con, C.start_run(self.con, "tdnet", "2026-08-28"),
                     "empty", n_listed=0)
        self.assertEqual(
            C.missing_days(self.con, "tdnet", date(2026, 8, 28), date(2026, 8, 28)),
            [])

    def test_weekend_never_reported_as_missing(self):
        # 2026-08-29 / 30 are Sat / Sun.
        self.assertEqual(
            C.missing_days(self.con, "tdnet", date(2026, 8, 29), date(2026, 8, 30)),
            [])

    def test_untouched_weekday_is_missing(self):
        self.assertEqual(
            C.missing_days(self.con, "tdnet", date(2026, 8, 26), date(2026, 8, 26)),
            ["2026-08-26"])


class TestArchiveDayFailurePath(unittest.TestCase):
    """End-to-end on the real archive_day(): a listing that cannot be fetched
    must leave a 'failed' row, not an exception and not silence."""

    def setUp(self):
        fd, self.db = tempfile.mkstemp(suffix=".db")
        os.close(fd)
        self.con = C.init_db(self.db)
        self._url = T.LIST_URL

    def tearDown(self):
        T.LIST_URL = self._url
        self.con.close()
        for suffix in ("", "-wal", "-shm"):
            try:
                os.remove(self.db + suffix)
            except OSError:
                pass

    def test_unreachable_listing_records_failed(self):
        # A host that cannot resolve: the same shape as TDnet being down.
        T.LIST_URL = "https://tdnet.invalid.example/I_list_{page:03d}_{ymd}.html"
        fetcher = C.Fetcher(min_interval=0.0, retries=1, timeout=5)
        res = T.archive_day(self.con, fetcher, date(2026, 8, 27))

        self.assertEqual(res["status"], "failed")
        row = self.con.execute(
            "SELECT status, error FROM fetch_runs WHERE target_date='2026-08-27'"
        ).fetchone()
        self.assertEqual(row["status"], "failed")
        self.assertTrue(row["error"], "the failure reason must be stored")
        self.assertEqual(
            C.missing_days(self.con, "tdnet", date(2026, 8, 27), date(2026, 8, 27)),
            ["2026-08-27"],
            "a day that failed must still show up for the backfill to retry")

    def test_empty_listing_records_empty(self):
        class _Empty:
            status_code = 200
            content = b"<html><body></body></html>"

        fetcher = C.Fetcher(min_interval=0.0, retries=1, timeout=5)
        fetcher.get = lambda *a, **k: _Empty()
        res = T.archive_day(self.con, fetcher, date(2026, 8, 27))
        self.assertEqual(res["status"], "empty")
        self.assertEqual(
            C.missing_days(self.con, "tdnet", date(2026, 8, 27), date(2026, 8, 27)),
            [])


class TestBusinessDays(unittest.TestCase):
    def test_from_saturday_walks_back_to_weekdays(self):
        got = C.business_days_back(3, end=date(2026, 8, 29))   # Saturday
        self.assertEqual([d.isoformat() for d in got],
                         ["2026-08-26", "2026-08-27", "2026-08-28"])

    def test_from_wednesday_includes_today(self):
        got = C.business_days_back(2, end=date(2026, 8, 26))
        self.assertEqual([d.isoformat() for d in got],
                         ["2026-08-25", "2026-08-26"])


if __name__ == "__main__":
    unittest.main(verbosity=2)
