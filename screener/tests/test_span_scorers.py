"""tests/test_span_scorers.py — span-matched スコアラーの不変条件。

守りたいのは3つ。
1. **異なる長さの期を比べない。** span-matched の存在理由そのもの。
2. **前年ペアが無ければ黙って別の期で代用しない。** available=False で返す。
3. **古い期の証拠を今の証拠として出さない。**（MAX_STALE_Q）
"""
import os
import sqlite3
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from screener.projection import span_matched as SM     # noqa: E402
from screener.signals import freshness as FR           # noqa: E402
from screener.signals import span_scorers as SP        # noqa: E402

DDL = """
CREATE TABLE quarterly_standalone_all (
  id INTEGER PRIMARY KEY AUTOINCREMENT, ticker TEXT, fiscal_year INTEGER,
  quarter_type TEXT, period_end TEXT, span_q INTEGER, period_start TEXT,
  sales REAL, operating_profit REAL, gross_profit REAL, operating_cf REAL,
  is_valid INTEGER DEFAULT 1, invalid_reason TEXT, generation INTEGER DEFAULT 1);
CREATE TABLE balance_sheet_items (
  id INTEGER PRIMARY KEY AUTOINCREMENT, ticker TEXT, period_end TEXT,
  item_key TEXT, value REAL, generation INTEGER DEFAULT 1);
CREATE TABLE pl_adjustments (
  id INTEGER PRIMARY KEY AUTOINCREMENT, ticker TEXT, period_end TEXT,
  item_key TEXT, amount REAL, note TEXT, generation INTEGER DEFAULT 1);
CREATE TABLE company_forecasts (
  id INTEGER PRIMARY KEY AUTOINCREMENT, ticker TEXT, fiscal_year INTEGER,
  forecast_sales REAL, forecast_op REAL, source_date TEXT);
CREATE TABLE filings (
  id INTEGER PRIMARY KEY AUTOINCREMENT, ticker TEXT, filing_date TEXT,
  period_end TEXT, quarter_type TEXT, fiscal_year INTEGER, source TEXT,
  pdf_path TEXT, xbrl_path TEXT, doc_id TEXT, generation INTEGER DEFAULT 1);
"""


def q(con, ticker, pe, span, start, sales, op=None, cf=None, valid=1):
    fy = int(pe[2:6])
    con.execute(
        "INSERT INTO quarterly_standalone_all (ticker, fiscal_year, quarter_type,"
        " period_end, span_q, period_start, sales, operating_profit, operating_cf,"
        " is_valid) VALUES (?,?,?,?,?,?,?,?,?,?)",
        (ticker, fy, pe[-2:], pe, span, start, sales, op, cf, valid))
    con.execute("INSERT INTO filings (ticker, filing_date, period_end, fiscal_year,"
                " generation) VALUES (?,?,?,?,1)",
                (ticker, "%04d-01-01" % (fy + 1), pe, fy))


def bs(con, ticker, pe, key, value):
    con.execute("INSERT INTO balance_sheet_items (ticker, period_end, item_key, value)"
                " VALUES (?,?,?,?)", (ticker, pe, key, value))


class Base(unittest.TestCase):
    def setUp(self):
        self.con = sqlite3.connect(":memory:")
        self.con.row_factory = sqlite3.Row
        self.con.executescript(DDL)
        SM.clear_cache()

    def tearDown(self):
        SM.clear_cache()


class TestSpanMatching(Base):
    def test_異なるspanは組にしない(self):
        """半期(span=2)を四半期(span=1)の前年と比べてはいけない。"""
        c = self.con
        q(c, "T", "FY2025-Q2", 1, "FY2025-Q2", 100)
        q(c, "T", "FY2026-Q2", 2, "FY2026-Q1", 210)
        peer = SM.find_yoy_peer(c, "T", "FY2026-Q2", 2)
        self.assertIsNone(peer, "span=2 の前年に span=1 を返してはいけない")

    def test_同じspanなら組にする(self):
        c = self.con
        q(c, "T", "FY2025-Q2", 2, "FY2025-Q1", 190)
        q(c, "T", "FY2026-Q2", 2, "FY2026-Q1", 210)
        peer = SM.find_yoy_peer(c, "T", "FY2026-Q2", 2)
        self.assertIsNotNone(peer)
        self.assertEqual(peer["period_end"], "FY2025-Q2")

    def test_無効な期は比較相手にしない(self):
        c = self.con
        q(c, "T", "FY2025-Q2", 2, "FY2025-Q1", 190, valid=0)
        q(c, "T", "FY2026-Q2", 2, "FY2026-Q1", 210)
        self.assertIsNone(SM.find_yoy_peer(c, "T", "FY2026-Q2", 2))

    def test_四半期が一度も無い銘柄はspan_matched(self):
        """半期・通期しか出さない銘柄を quarter モードに落とすと、
        span>=2 の行が全部捨てられて永久に見えなくなる。"""
        c = self.con
        for y in (2024, 2025, 2026):
            q(c, "T", "FY%d-Q2" % y, 2, "FY%d-Q1" % y, 100)
            q(c, "T", "FY%d-Q4" % y, 2, "FY%d-Q3" % y, 110)
        self.assertIsNotNone(SM.switch_point(c, "T"))
        self.assertEqual(SM.mode_for(c, "T", "FY2026-Q2"), "span_matched")
        self.assertTrue(SM.yoy_series(c, "T", "sales"))

    def test_履歴が短いだけでは切り替えない(self):
        """最初の四半期より前の空白は「切れ目」ではない。"""
        c = self.con
        for y in (2025, 2026):
            for i in (1, 2, 3, 4):
                q(c, "T", "FY%d-Q%d" % (y, i), 1, "FY%d-Q%d" % (y, i), 100)
        self.assertIsNone(SM.switch_point(c, "T"))


class TestS1(Base):
    def _setup(self, span):
        c = self.con
        start = lambda y: "FY%d-Q%d" % (y, 3 - span)      # noqa: E731
        for y in (2025, 2026):
            q(c, "T", "FY%d-Q2" % y, span, start(y), 1000.0)
            bs(c, "T", "FY%d-Q2" % y, "accounts_receivable", 500.0 if y == 2025 else 400.0)
        return c

    def test_span2でも前年と組めばDSOが出る(self):
        c = self._setup(2)
        r = SP.s1_dso(c, "T")
        self.assertTrue(r["available"], r["evidence"])
        self.assertEqual(r["span_q"], 2)
        self.assertEqual(r["peer_period"], "FY2025-Q2")
        # AR が 500→400 で売上同額なので DSO は 20% 改善 → 満点
        self.assertAlmostEqual(r["score"], 1.0, places=3)

    def test_日数はspanに比例する(self):
        """span=2 の DSO は 6ヶ月ぶんの日数で出す。比率は変わらない。"""
        r1 = SP.s1_dso(self._setup(1), "T")
        SM.clear_cache()
        self.setUp()
        r2 = SP.s1_dso(self._setup(2), "T")
        self.assertAlmostEqual(r1["score"], r2["score"], places=6)
        self.assertAlmostEqual(r2["details"]["dso_now"],
                               r1["details"]["dso_now"] * 2, places=3)

    def test_古すぎる組は返さない(self):
        c = self.con
        for y in (2020, 2021):
            q(c, "T", "FY%d-Q2" % y, 1, "FY%d-Q2" % y, 1000.0)
            bs(c, "T", "FY%d-Q2" % y, "accounts_receivable", 500.0)
        q(c, "T", "FY2026-Q2", 1, "FY2026-Q2", 1000.0)   # 直近開示だけ AR 無し
        r = SP.s1_dso(c, "T")
        self.assertFalse(r["available"])
        self.assertIn("古い", r["evidence"])


class TestS5(Base):
    def test_半期1本でも進捗率が出る(self):
        """四半期が消えても、半期そのものが期首からの累計。"""
        c = self.con
        for y in (2023, 2024):                 # 季節性: 上期 40%
            for i, s in ((1, 200), (2, 200), (3, 300), (4, 300)):
                q(c, "T", "FY%d-Q%d" % (y, i), 1, "FY%d-Q%d" % (y, i), s)
            q(c, "T", "FY%d-Q4" % y, 4, "FY%d-Q1" % y, 1000)
        q(c, "T", "FY2025-Q2", 2, "FY2025-Q1", 500, op=50)
        c.execute("INSERT INTO company_forecasts (ticker, fiscal_year, forecast_sales,"
                  " forecast_op, source_date) VALUES ('T',2025,1000,100,'2025-05-01')")
        r = SP.s5_progress(c, "T")
        self.assertTrue(r["available"], r["evidence"])
        self.assertEqual(r["span_q"], 2)
        self.assertAlmostEqual(r["details"]["expected"], 0.40, places=6)
        self.assertAlmostEqual(r["details"]["progress_sales"], 0.50, places=6)

    def test_欠けた四半期をゼロとして足さない(self):
        """Q1 が無いのに Q2 単独を『上期累計』として扱うと、進捗率が
        半分に見えて最大の売りシグナルが出る（別枠の実挙動）。
        組み立てられないなら available=False が正しい。"""
        c = self.con
        q(c, "T", "FY2025-Q2", 1, "FY2025-Q2", 200)      # Q1 が無い
        c.execute("INSERT INTO company_forecasts (ticker, fiscal_year, forecast_sales,"
                  " forecast_op, source_date) VALUES ('T',2025,1000,100,'2025-05-01')")
        r = SP.s5_progress(c, "T")
        self.assertFalse(r["available"], r["evidence"])

    def test_通期発表済みは対象外(self):
        c = self.con
        q(c, "T", "FY2025-Q4", 4, "FY2025-Q1", 1000, op=100)
        c.execute("INSERT INTO company_forecasts (ticker, fiscal_year, forecast_sales,"
                  " forecast_op, source_date) VALUES ('T',2025,1000,100,'2025-05-01')")
        self.assertFalse(SP.s5_progress(c, "T")["available"])


class TestS4(Base):
    def test_CFと営業利益は同じ期間で比べる(self):
        """span 合計が4になるまで後ろから取る。CFが2四半期・OPが4四半期
        という食い違い（別枠の位置ベース組み立て）を起こさない。"""
        c = self.con
        q(c, "T", "FY2026-Q2", 2, "FY2026-Q1", 500, op=50, cf=60)
        q(c, "T", "FY2026-Q4", 2, "FY2026-Q3", 500, op=50, cf=90)
        r = SP.s4_cash_quality(c, "T")
        self.assertTrue(r["available"], r["evidence"])
        self.assertAlmostEqual(r["details"]["cf_op_ratio"], 150.0 / 100.0, places=6)

    def test_1年に満たなければ返さない(self):
        c = self.con
        q(c, "T", "FY2026-Q2", 2, "FY2026-Q1", 500, op=50, cf=60)
        r = SP.s4_cash_quality(c, "T")
        self.assertFalse(r["available"])
        self.assertIn("1年", r["evidence"])


class TestPIT(Base):
    def test_as_of以前に出ていない期は見えない(self):
        c = self.con
        for y in (2025, 2026):
            q(c, "T", "FY%d-Q2" % y, 2, "FY%d-Q1" % y, 1000.0)
            bs(c, "T", "FY%d-Q2" % y, "accounts_receivable", 500.0 if y == 2025 else 400.0)
        # filings は FY+1 年の 1/1 に入る。2026 期は 2027-01-01 開示。
        self.assertTrue(SP.s1_dso(c, "T", as_of="2027-06-01")["available"])
        self.assertFalse(SP.s1_dso(c, "T", as_of="2026-06-01")["available"])


if __name__ == "__main__":
    unittest.main(verbosity=2)


class TestRuleB(Base):
    """規則B（期ごとに最細粒度・常に同 span 同士）の不変条件。設計書 §9。"""

    def test_期ごとに最も細かい粒度を採る(self):
        """同じ期に span=1 と span=2 の両方があれば span=1 を使う。
        細かいほうが情報が多く、前年も同じ粒度で取れる見込みが高い。"""
        c = self.con
        for y in (2025, 2026):
            q(c, "T", "FY%d-Q2" % y, 1, "FY%d-Q2" % y, 100)
            q(c, "T", "FY%d-Q2" % y, 2, "FY%d-Q1" % y, 190)
        ser = SM.yoy_series(c, "T", "sales")
        self.assertEqual([r["span_q"] for r in ser], [1])
        self.assertEqual(ser[0]["mode"], "quarter")

    def test_通期は四半期系列に混ぜない(self):
        """span=4 を混ぜると『前四半期比』の意味が壊れる。"""
        c = self.con
        for y in (2025, 2026):
            q(c, "T", "FY%d-Q4" % y, 4, "FY%d-Q1" % y, 1000)
        self.assertEqual(SM.yoy_series(c, "T", "sales"), [])

    def test_粒度が混ざってもペアは常に同spanになる(self):
        """規則Bは時系列に粒度の混在を許す。**許すのは混在であって、
        異なる長さ同士の比較ではない。**"""
        c = self.con
        for y in (2025, 2026):
            q(c, "T", "FY%d-Q1" % y, 1, "FY%d-Q1" % y, 50)
            q(c, "T", "FY%d-Q4" % y, 2, "FY%d-Q3" % y, 120)
        ser = SM.yoy_series(c, "T", "sales")
        self.assertEqual(sorted(r["span_q"] for r in ser), [1, 2])
        for r in ser:
            peer = SM.find_yoy_peer(c, "T", r["peer_period"], r["span_q"])
            self.assertEqual(SM.parse_pe(r["period_end"])[1],
                             SM.parse_pe(r["peer_period"])[1],
                             "同じ四半期位置どうしで組むこと")

    def test_傾き系は同一spanの連鎖を要求する_S4(self):
        """S4 は1年ぶんを積み上げる。**span の合計が4でも期間が重なったら
        足してはいけない。** 同じ期に span=1 と span=2 が並ぶのは普通。"""
        c = self.con
        q(c, "T", "FY2026-Q2", 1, "FY2026-Q2", 250, op=25, cf=30)
        q(c, "T", "FY2026-Q2", 2, "FY2026-Q1", 500, op=50, cf=60)
        q(c, "T", "FY2026-Q4", 2, "FY2026-Q3", 500, op=50, cf=90)
        r = SP.s4_cash_quality(c, "T")
        self.assertTrue(r["available"], r["evidence"])
        # 正: 上期(2)+下期(2)。誤: 下期(2)+上期(2)+Q2単独(1) は span 合計5で
        # 弾かれるが、下期(2)+Q2単独(1)+... のように重なる拾い方も禁じる。
        self.assertAlmostEqual(r["details"]["cf_op_ratio"], 150.0 / 100.0, places=6)

    def test_傾き系は同一spanの連鎖を要求する_S5(self):
        """S5 の期首累計は隙間も重複も作らない。Q1単独と上期が並ぶとき
        両方足すと Q1 を二重に数える。"""
        c = self.con
        for y in (2023, 2024):
            for i, sv in ((1, 200), (2, 200), (3, 300), (4, 300)):
                q(c, "T", "FY%d-Q%d" % (y, i), 1, "FY%d-Q%d" % (y, i), sv)
            q(c, "T", "FY%d-Q4" % y, 4, "FY%d-Q1" % y, 1000)
        q(c, "T", "FY2025-Q1", 1, "FY2025-Q1", 200, op=20)
        q(c, "T", "FY2025-Q2", 2, "FY2025-Q1", 500, op=50)   # Q1 を含む上期
        c.execute("INSERT INTO company_forecasts (ticker, fiscal_year, forecast_sales,"
                  " forecast_op, source_date) VALUES ('T',2025,1000,100,'2025-05-01')")
        ytd = SP._ytd(c, "T", 2025, None)
        self.assertEqual(ytd[1], 2, "Q1〜Q2 を覆うこと")
        self.assertEqual(ytd[2], 500.0, "Q1 を二重に数えないこと（700 は誤り）")

    def test_連鎖が切れたら伸ばさない(self):
        """Q1 が無いまま Q2 単独を『上期累計』として扱わない。"""
        c = self.con
        q(c, "T", "FY2025-Q2", 1, "FY2025-Q2", 200)
        self.assertIsNone(SP._ytd(c, "T", 2025, None))


class TestEvidenceDoc(Base):
    """原文リンクは**根拠期の書類**を指す。最新の開示ではない。"""

    def test_根拠期の書類を返す(self):
        c = self.con
        for y in (2025, 2026):
            q(c, "T", "FY%d-Q2" % y, 2, "FY%d-Q1" % y, 1000.0)
            bs(c, "T", "FY%d-Q2" % y, "accounts_receivable", 500.0 if y == 2025 else 400.0)
        # 根拠期より新しい開示（通期）を足しても、リンクはそちらに動かない
        c.execute("INSERT INTO filings (ticker, filing_date, period_end, fiscal_year,"
                  " generation) VALUES ('T','2027-05-01','FY2026-Q4',2026,1)")
        c.execute("UPDATE filings SET doc_id='S100AAAA', source='edinet' "
                  "WHERE period_end='FY2026-Q2'")
        c.execute("UPDATE filings SET doc_id='S100ZZZZ', source='edinet' "
                  "WHERE period_end='FY2026-Q4'")
        r = SP.s1_dso(c, "T")
        self.assertTrue(r["available"], r["evidence"])
        self.assertEqual(r["period"], "FY2026-Q2")
        self.assertEqual(r["doc_id"], "S100AAAA",
                         "最新開示(FY2026-Q4)ではなく根拠期の書類を指すこと")

    def test_複数期スコアは最新根拠期と期間注記を持つ(self):
        c = self.con
        q(c, "T", "FY2026-Q2", 2, "FY2026-Q1", 500, op=50, cf=60)
        q(c, "T", "FY2026-Q4", 2, "FY2026-Q3", 500, op=50, cf=90)
        c.execute("UPDATE filings SET doc_id='S100Q2', source='edinet' "
                  "WHERE period_end='FY2026-Q2'")
        c.execute("UPDATE filings SET doc_id='S100Q4', source='edinet' "
                  "WHERE period_end='FY2026-Q4'")
        r = SP.s4_cash_quality(c, "T")
        self.assertEqual(r["doc_id"], "S100Q4", "最新の根拠期に飛ばすこと")
        self.assertIn("FY2026-Q2", r["period_note"])
        self.assertIn("FY2026-Q4", r["period_note"])


class TestFreshness(Base):
    """シグナル鮮度（stale）。**表示のみ・スコア不変。**

    実例（2026-09-02 に人間が原文精読して発見）:
      3565 アセンテックの契約負債シグナルは根拠期 FY2025-Q2（開示 2024-09-11）。
      最新の FY2026-Q2 まで開示があるのに、約2年前の証拠で第2位に載っていた。
    """

    def _pair(self, y_cur, y_prev, span=2):
        c = self.con
        for y in (y_prev, y_cur):
            q(c, "T", "FY%d-Q2" % y, span, "FY%d-Q1" % y, 1000.0)
            bs(c, "T", "FY%d-Q2" % y, "accounts_receivable",
               500.0 if y == y_prev else 400.0)
        return c

    def test_根拠期が最新なら鮮度は新しい(self):
        """1433 の形: basis == latest → stale=0。"""
        c = self._pair(2027, 2026)
        r = SP.s1_dso(c, "T")
        self.assertTrue(r["available"])
        scores = {"S1": r}
        FR.annotate(scores, c, "T")
        self.assertEqual(scores["S1"]["stale"], 0)
        self.assertEqual(scores["S1"]["stale_lag"], 0)
        flag, lag, _ = FR.summarize(scores, {"S1"})
        self.assertEqual((flag, lag), (0, 0))

    def test_前年H1を合成すれば当期と組める(self):
        """3565 の形: 当期 FY2026-Q2(span=2) の前年 H1 は Q1・Q2 の
        四半期2本でしか無い。**足せば H1 そのもの**なので合成して組む。
        合成前は1年遡って stale=1 になっていた（2026-09-02 の実測）。"""
        c = self.con
        for y in (2024, 2025):
            for i in (1, 2):
                q(c, "T", "FY%d-Q%d" % (y, i), 1, "FY%d-Q%d" % (y, i), 500.0)
                bs(c, "T", "FY%d-Q%d" % (y, i), "accounts_receivable",
                   500.0 if y == 2024 else 400.0)
        q(c, "T", "FY2026-Q2", 2, "FY2026-Q1", 1100.0)     # 当期は半期のみ
        bs(c, "T", "FY2026-Q2", "accounts_receivable", 380.0)
        bs(c, "T", "FY2025-Q2", "accounts_receivable", 400.0)
        r = SP.s1_dso(c, "T")
        self.assertTrue(r["available"], r["evidence"])
        self.assertEqual(r["period"], "FY2026-Q2", "当期で組めること")
        self.assertEqual(r["peer_period"], "FY2025-Q2")
        self.assertEqual(r["peer_synthesized"], 1, "合成だと明示すること")
        self.assertIn("合成", r["evidence"])
        scores = {"S1": r}
        FR.annotate(scores, c, "T")
        self.assertEqual(scores["S1"]["stale"], 0)

    def test_合成できないときは遡ってstaleになる(self):
        """前年 H1 に隙間があれば**作らない**。欠けた四半期をゼロとして
        足すのは S5 で既に踏んだ誤り。作れないなら遡り、古いと明示する。"""
        c = self.con
        for y in (2024, 2025):
            for i in (1, 2):
                q(c, "T", "FY%d-Q%d" % (y, i), 1, "FY%d-Q%d" % (y, i), 500.0)
                bs(c, "T", "FY%d-Q%d" % (y, i), "accounts_receivable",
                   500.0 if y == 2024 else 400.0)
        c.execute("DELETE FROM quarterly_standalone_all "
                  "WHERE ticker='T' AND period_end='FY2025-Q1'")   # 隙間を作る
        q(c, "T", "FY2026-Q2", 2, "FY2026-Q1", 1100.0)
        bs(c, "T", "FY2026-Q2", "accounts_receivable", 380.0)
        r = SP.s1_dso(c, "T")
        self.assertTrue(r["available"], r["evidence"])
        self.assertEqual(r["period"], "FY2025-Q2", "組める最後の期まで遡ること")
        scores = {"S1": r}
        latest = FR.annotate(scores, c, "T")
        self.assertEqual(latest, "FY2026-Q2")
        self.assertEqual(scores["S1"]["stale"], 1)
        self.assertEqual(scores["S1"]["stale_lag"], 4)

    def test_根拠期を持たないシグナルは判定不能(self):
        """別枠の位置ベース比較は期が特定できない。**0（新しい）と混ぜない。**"""
        scores = {"S1": {"available": True, "score": 0.5, "period": None}}
        c = self.con
        q(c, "T", "FY2026-Q2", 2, "FY2026-Q1", 1000.0)
        FR.annotate(scores, c, "T")
        self.assertIsNone(scores["S1"]["stale"])
        flag, lag, detail = FR.summarize(scores, {"S1"})
        self.assertEqual(flag, 0)
        self.assertIn("判定不能", detail)

    def test_スコアは変わらない(self):
        """鮮度は表示のみ。annotate がスコアに触らないこと。"""
        c = self._pair(2026, 2025)
        r = SP.s1_dso(c, "T")
        before = r["score"]
        FR.annotate({"S1": r}, c, "T")
        self.assertEqual(r["score"], before)


class TestLagDefinition(unittest.TestCase):
    """lag の定義を固定する（タスクD）。

    lag = 期インデックスの差（四半期数）。span の長さは見ない。
    """

    def test_1年前は4(self):
        self.assertEqual(FR.lag_quarters("FY2025-Q2", "FY2026-Q2"), 4)

    def test_FY2025Q2とFY2027Q1は7(self):
        """当初 fixture は 4 とされていたが算術誤り。正解は 7。"""
        self.assertEqual(FR.lag_quarters("FY2025-Q2", "FY2027-Q1"), 7)

    def test_同一期は0(self):
        self.assertEqual(FR.lag_quarters("FY2026-Q4", "FY2026-Q4"), 0)

    def test_読めなければNone(self):
        self.assertIsNone(FR.lag_quarters("FY2026-Q4", "こわれた"))
