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


class _FakeListing:
    """ページ1だけ HTML を返し、2ページ目以降は空を返す一覧のフェイク。"""

    def __init__(self, html):
        self.html = html.encode("utf-8")

    def get(self, url, **kw):
        page1 = "_001_" in url
        body = self.html if page1 else b"<html><body></body></html>"

        class R:
            status_code = 200
            content = body
        return R()

    def download(self, url, dest):
        raise RuntimeError("download failed (test)")


def _listing(total, titles):
    trs = "".join(
        '<tr><td>15:00</td><td>6118%d</td><td>X</td>'
        '<td><a href="1401202608275%05d.pdf">%s</a></td><td></td><td>東</td><td></td></tr>'
        % (i, i, t) for i, t in enumerate(titles))
    return ('<html><body><div id="pager-box-top">1～%d件 / 全%d件</div>'
            '<table id="main-list-table">%s</table></body></html>' % (len(titles), total, trs))


class TestPostconditionGate(unittest.TestCase):
    """2026-09-13: 部分取得を ok と記録しない後条件ゲート。"""

    OUT = ["代表取締役の異動に関するお知らせ", "自己株式の取得状況に関するお知らせ"]

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

    def _missing(self, d):
        return C.missing_days(self.con, "tdnet", d, d)

    def test_総件数と読めた行数が違えばincomplete(self):
        from datetime import datetime
        d = date(2026, 8, 27)
        res = T.archive_day(self.con, _FakeListing(_listing(3, self.OUT)), d,
                            now=datetime(2026, 9, 13, 12))
        self.assertEqual(res["status"], "incomplete")
        self.assertEqual(self._missing(d), ["2026-08-27"])

    def test_対象日の終了前はprovisionalで欠損扱い(self):
        from datetime import datetime
        d = date(2026, 9, 11)
        res = T.archive_day(self.con, _FakeListing(_listing(2, self.OUT)), d,
                            now=datetime(2026, 9, 11, 10, 50))   # 9/11 の朝の実例
        self.assertEqual(res["status"], "provisional")
        self.assertEqual(self._missing(d), ["2026-09-11"])

    def test_対象日の終了後で件数一致ならok(self):
        from datetime import datetime
        d = date(2026, 9, 11)
        res = T.archive_day(self.con, _FakeListing(_listing(2, self.OUT)), d,
                            now=datetime(2026, 9, 12, 0, 1))
        self.assertEqual(res["status"], "ok")
        self.assertEqual(self._missing(d), [])

    def test_当日朝の0件はemptyではない(self):
        from datetime import datetime
        d = date(2026, 9, 3)
        res = T.archive_day(self.con, _FakeListing("<html><body></body></html>"), d,
                            now=datetime(2026, 9, 3, 0, 31))
        self.assertEqual(res["status"], "provisional")
        self.assertEqual(self._missing(d), ["2026-09-03"])

    def test_対象書類が保存できなければpartialで欠損扱い(self):
        from datetime import datetime
        d = date(2026, 8, 27)
        html = _listing(1, ["2027年３月期第１四半期決算短信〔日本基準〕(連結)"])
        res = T.archive_day(self.con, _FakeListing(html), d, now=datetime(2026, 9, 13))
        self.assertEqual(res["status"], "partial")
        self.assertEqual(self._missing(d), ["2026-08-27"],
                         "partial は covered ではない（2026-09-13 変更）")


class TestTitlesCode(unittest.TestCase):
    """disclosure_titles.code は一覧の code_raw から取る（2026-09-13 修正）。"""

    def setUp(self):
        fd, self.db = tempfile.mkstemp(suffix=".db")
        os.close(fd)
        self.con = C.init_db(self.db)
        from screener.fetch import tdnet_titles as TT
        self.TT = TT
        self.con.executescript(TT.DDL)
        self._orig = T.fetch_day_index

    def tearDown(self):
        T.fetch_day_index = self._orig
        self.con.close()
        for suffix in ("", "-wal", "-shm"):
            try:
                os.remove(self.db + suffix)
            except OSError:
                pass

    def test_codeが保存される(self):
        T.fetch_day_index = lambda f, d: ([{
            "time": "15:00", "code_raw": "34410", "name": "山王",
            "title": "2026年7月期 決算発表日のお知らせ", "pdf": "140120260901000001.pdf",
            "zip": None, "place": "東"}], 1)
        self.TT.archive_titles(self.con, None, date(2026, 9, 1))
        row = self.con.execute("SELECT code, doc_name FROM disclosure_titles").fetchone()
        self.assertEqual(row["code"], "3441")
        self.assertEqual(row["doc_name"], "140120260901000001.pdf")

    def test_旧行のcodeを会社名込みで再導出し_曖昧なら埋めない(self):
        c = self.con
        c.execute("INSERT INTO disclosure_titles (code, date, time, title, doc_name) "
                  "VALUES (NULL,'2026-08-20','15:00','業績予想の修正に関するお知らせ','山王')")
        c.execute("INSERT INTO disclosure_titles (code, date, time, title, doc_name) "
                  "VALUES (NULL,'2026-08-20','15:00','業績予想の修正に関するお知らせ','同名社')")
        T.fetch_day_index = lambda f, d: ([
            {"time": "15:00", "code_raw": "34410", "name": "山王",
             "title": "業績予想の修正に関するお知らせ", "pdf": "a.pdf"},
            {"time": "15:00", "code_raw": "11110", "name": "同名社",
             "title": "業績予想の修正に関するお知らせ", "pdf": "b.pdf"},
            {"time": "15:00", "code_raw": "22220", "name": "同名社",
             "title": "業績予想の修正に関するお知らせ", "pdf": "c.pdf"}], 3)
        st = self.TT.rederive_codes(c, None)
        self.assertEqual(st["updated"], 1)
        self.assertEqual(st["ambiguous"], 1)
        self.assertEqual(c.execute("SELECT code FROM disclosure_titles WHERE doc_name='a.pdf'")
                         .fetchone()["code"], "3441")


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
