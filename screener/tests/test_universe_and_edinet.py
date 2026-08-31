"""screener/tests/test_universe_and_edinet.py — P1 (仕様書 §2-2 / §2-3).

The rule this file exists to protect: **"does not meet the condition" and
"we have not fetched the data yet" must never collapse into the same state.**
If a company with no market-cap data is recorded as excluded, the day J-Quants
starts working nobody can tell how many companies were newly admitted versus
how many were wrongly dropped all along.
"""
from __future__ import annotations

import os
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(
    os.path.dirname(os.path.abspath(__file__)))))

from screener import common as C
from screener.fetch import jquants_universe as U
from screener.fetch import edinet_bulk as E


class _DbCase(unittest.TestCase):
    def setUp(self):
        fd, self.db = tempfile.mkstemp(suffix=".db")
        os.close(fd)
        self.con = C.init_db(self.db)

    def tearDown(self):
        self.con.close()
        for s in ("", "-wal", "-shm"):
            try:
                os.remove(self.db + s)
            except OSError:
                pass

    def add(self, code, name="X", market="プライム（内国株式）", sector="機械",
            mktcap=None, adv20=None):
        U.upsert_company(self.con, code, name, market, sector, sector17=None,
                         scale=None, mktcap=mktcap, adv20=adv20, source="test")
        self.con.commit()


class TestUniverseRules(_DbCase):
    def test_in_universe(self):
        self.add("1111", mktcap=20000, adv20=80)          # 200億円 / 8千万円
        r = U.apply_universe_rules(self.con)
        self.assertEqual(r["universe"], 1)
        row = self.con.execute("SELECT universe_flag, exclude_reason FROM companies "
                               "WHERE code='1111'").fetchone()
        self.assertEqual(row["universe_flag"], 1)
        self.assertIsNone(row["exclude_reason"])

    def test_too_big_and_too_small(self):
        # 閾値は universe_rules.yaml が正。ここに数字を書き写すと、上限を
        # 動かすたびにテストが「仕様変更」ではなく「破損」として落ちる
        # (2026-08-31 の 600億->1,000億 で実際に落ちた)。yaml から境界を取る。
        size = C.load_yaml("universe_rules.yaml")["size"]
        self.add("1111", mktcap=size["mktcap_min_mn"] - 1, adv20=80)
        self.add("2222", mktcap=size["mktcap_max_mn"] + 1, adv20=80)
        U.apply_universe_rules(self.con)
        for code in ("1111", "2222"):
            row = self.con.execute("SELECT universe_flag, exclude_reason "
                                   "FROM companies WHERE code=?", (code,)).fetchone()
            self.assertEqual(row["universe_flag"], 0)
            self.assertEqual(row["exclude_reason"], "時価総額レンジ外")

    def test_illiquid_excluded(self):
        """4192 スパイダープラスの教訓の機械化: 出来ても出られない銘柄を外す。"""
        self.add("1111", mktcap=20000, adv20=12)          # 1,200万円 < 3,000万円
        U.apply_universe_rules(self.con)
        row = self.con.execute("SELECT universe_flag, exclude_reason FROM companies "
                               "WHERE code='1111'").fetchone()
        self.assertEqual(row["universe_flag"], 0)
        self.assertEqual(row["exclude_reason"], "流動性不足")

    def test_banks_insurers_securities_excluded(self):
        for code, sector in (("8306", "銀行業"), ("8750", "保険業"),
                             ("8601", "証券、商品先物取引業")):
            self.add(code, sector=sector, mktcap=20000, adv20=80)
        U.apply_universe_rules(self.con)
        for code in ("8306", "8750", "8601"):
            row = self.con.execute("SELECT universe_flag, exclude_reason "
                                   "FROM companies WHERE code=?", (code,)).fetchone()
            self.assertEqual(row["universe_flag"], 0)
            self.assertIn("業種除外", row["exclude_reason"])

    def test_reit_and_etf_excluded_by_market(self):
        self.add("1306", market="ETF・ETN", sector=None, mktcap=20000, adv20=800)
        self.add("8951", market="REIT・ベンチャーファンド・カントリーファンド・インフラファンド",
                 sector=None, mktcap=20000, adv20=800)
        U.apply_universe_rules(self.con)
        for code in ("1306", "8951"):
            row = self.con.execute("SELECT universe_flag, exclude_reason "
                                   "FROM companies WHERE code=?", (code,)).fetchone()
            self.assertEqual(row["universe_flag"], 0)
            self.assertIn("市場区分除外", row["exclude_reason"])

    def test_v2_product_categories_are_excluded_by_market(self):
        """J-Quants V2 の形。MktNm は「プライム」等の純粋な市場名になり、
        REIT/ETF の別は ProdCat へ移った。_market_label が商品種別名を
        市場名へ戻すので、除外ルールは市場名のままで効き続ける。
        戻し忘れると REIT/ETF が黙ってユニバースに入る —— それを禁じる。"""
        cases = {
            "1306": {"MktNm": "その他", "ProdCat": "014"},   # ETF
            "8951": {"MktNm": "その他", "ProdCat": "013"},   # REIT
            "8963": {"MktNm": "プライム", "ProdCat": "012"},  # 優先出資証券
        }
        for code, row in cases.items():
            self.add(code, market=U._market_label(row), sector=None,
                     mktcap=20000, adv20=800)
        U.apply_universe_rules(self.con)
        for code in cases:
            row = self.con.execute("SELECT universe_flag, exclude_reason "
                                   "FROM companies WHERE code=?", (code,)).fetchone()
            self.assertEqual(row["universe_flag"], 0, code)
            self.assertIn("市場区分除外", row["exclude_reason"], code)

    def test_v2_domestic_stock_keeps_a_plain_market_name(self):
        """内国株券は商品種別を併記しない。併記すると『プライム(内国株券)』が
        除外語に一致する事故が将来起きうるし、V1 と表示が変わって読みにくい。"""
        self.assertEqual(U._market_label({"MktNm": "プライム", "ProdCat": "011"}),
                         "プライム")
        self.assertEqual(U._market_label({"MktNm": "グロース", "ProdCat": "011"}),
                         "グロース")

    def test_pro_market_is_excluded_in_either_spelling(self):
        """'PRO Market' と書いてあった頃は、実データの 'TOKYO PRO MARKET' と
        大小が合わず一度も一致していなかった。DB には過去の取り込み由来で
        両方の綴りが実在するので、どちらでも除外に落ちること。
        取りこぼすと「除外」ではなく「未判定」に化けるのが一番まずい。"""
        self.add("9999", market="TOKYO PRO MARKET", sector=None,
                 mktcap=20000, adv20=800)
        self.add("9998", market="PRO Market", sector=None,
                 mktcap=20000, adv20=800)
        U.apply_universe_rules(self.con)
        for code in ("9999", "9998"):
            row = self.con.execute("SELECT universe_flag, exclude_reason "
                                   "FROM companies WHERE code=?", (code,)).fetchone()
            self.assertEqual(row["universe_flag"], 0, code)
            self.assertIn("市場区分除外", row["exclude_reason"], code)

    def test_missing_data_is_pending_not_excluded(self):
        """核心。時価総額も売買代金も無い会社は『判定していない』であって
        『条件を満たさない』ではない。"""
        self.add("1111", mktcap=None, adv20=None)
        r = U.apply_universe_rules(self.con)
        row = self.con.execute("SELECT universe_flag, exclude_reason FROM companies "
                               "WHERE code='1111'").fetchone()
        self.assertEqual(row["universe_flag"], 0)
        self.assertEqual(row["exclude_reason"], "規模・流動性データ未取得")
        self.assertEqual(r["pending"], 1)
        self.assertEqual(r["excluded"], 0,
                         "未取得は excluded に数えてはいけない")

    def test_sector_exclusion_beats_missing_data(self):
        # 銀行はデータの有無に関係なく除外。pending に混ぜない。
        self.add("8306", sector="銀行業", mktcap=None, adv20=None)
        r = U.apply_universe_rules(self.con)
        self.assertEqual(r["excluded"], 1)
        self.assertEqual(r["pending"], 0)


class TestUpsertDoesNotBlankFields(_DbCase):
    def test_later_source_without_a_field_keeps_the_old_value(self):
        """TDnetアーカイバは社名だけの行を作る。あとから来たJ-Quants/JPXが
        持っていない列で既存の値を消してはいけない。"""
        self.add("6118", name="アイダエンジニアリング", market=None, sector=None)
        U.upsert_company(self.con, "6118", None, "プライム（内国株式）", "機械",
                         sector17=None, scale=None, mktcap=None, adv20=None,
                         source="jpx")
        self.con.commit()
        row = self.con.execute("SELECT name, market, sector FROM companies "
                               "WHERE code='6118'").fetchone()
        self.assertEqual(row["name"], "アイダエンジニアリング")
        self.assertEqual(row["market"], "プライム（内国株式）")
        self.assertEqual(row["sector"], "機械")


class TestEdinetSelection(_DbCase):
    def test_doc_types_cover_annual_and_semiannual_including_corrections(self):
        self.assertEqual(E.DOC_TYPES["120"], "有報")
        self.assertEqual(E.DOC_TYPES["130"], "有報")     # 訂正
        self.assertEqual(E.DOC_TYPES["160"], "半期")
        self.assertEqual(E.DOC_TYPES["170"], "半期")     # 訂正
        # 四半期報告書(140/150)は 2024-04-01 以後に開始する四半期から廃止されたが、
        # **過去データの遡及では現役の制度**。実地確認で 2022-11-14 に996件、
        # 2024-02-14 に907件が存在し 2024-05-15 を最後に消滅する。
        # TDnet は約40日しか保持せず短信では遡れないので、FY2023〜FY2024 の
        # 連続四半期はこの書類種別からしか作れない(2026-08-31)。
        self.assertEqual(E.DOC_TYPES["140"], "四半期")
        self.assertEqual(E.DOC_TYPES["150"], "四半期")

    def test_validation_eight(self):
        self.assertEqual(len(E.VALIDATION_CODES), 8)
        for c in ("285A", "2962", "3110", "278A", "3905", "6217", "4192", "6855"):
            self.assertIn(c, E.VALIDATION_CODES)

    def test_trial_codes_always_include_the_validation_eight(self):
        for i in range(30):
            self.add(f"{9000 + i}")
        codes = E.trial_codes(self.con, extra=5)
        for c in E.VALIDATION_CODES:
            self.assertIn(c, codes, "検証8銘柄は必ず試走に含める")
        self.assertGreaterEqual(len(codes), 8)

    def test_universe_codes_do_not_depend_on_the_market_name_spelling(self):
        """J-Quants V2 で市場区分名から『（内国株式）』が消えた。市場名を
        部分一致で見る条件は 1,093 社を 8 社まで取りこぼしていた ——
        取得対象の抽出は exclude_reason だけで決める。"""
        self.add("7203", market="プライム", mktcap=20000, adv20=800)      # V2 表記
        self.add("6118", market="プライム（内国株式）", mktcap=20000, adv20=800)  # V1 表記
        U.apply_universe_rules(self.con)
        codes = E.universe_codes(self.con)
        self.assertIn("7203", codes, "V2 表記の銘柄が対象から漏れている")
        self.assertIn("6118", codes, "V1 表記の銘柄が対象から漏れている")

    def test_universe_codes_include_the_validation_eight_even_when_excluded(self):
        """検証8銘柄はユニバースの部分集合ではない(5銘柄は時価総額上限超え)。
        和集合を取らないと 仕様書 §6 の検証データが欠ける。"""
        self.add("285A", market="プライム", mktcap=42_690_947, adv20=2_234_589)
        U.apply_universe_rules(self.con)
        row = self.con.execute("SELECT universe_flag FROM companies "
                               "WHERE code='285A'").fetchone()
        self.assertEqual(row["universe_flag"], 0, "前提: 285A はユニバース外")
        self.assertIn("285A", E.universe_codes(self.con))

    def test_trial_codes_are_deterministic(self):
        for i in range(30):
            self.add(f"{9000 + i}")
        self.assertEqual(E.trial_codes(self.con, extra=10),
                         E.trial_codes(self.con, extra=10),
                         "試走の対象がランで変わると本番の見積りにならない")

    def _pending(self, code, doc_id, ok=0):
        """ok=1 は取得済み。path も入れる —— download_pending の pending 判定は
        `xbrl_ok=0 OR path IS NULL` なので、path が空だと取得済みにならない。"""
        self.con.execute(
            "INSERT INTO filings (code, date, type, source, doc_id, xbrl_ok, path) "
            "VALUES (?,?,?,?,?,?,?)",
            (code, "2026-06-01", "有報", "edinet", doc_id, ok,
             f"raw/edinet/2026-06-01/{doc_id}.zip" if ok else None))
        self.con.commit()

    def test_download_is_filtered_by_codes_not_the_index(self):
        """索引は全上場銘柄を持ち、絞り込みはダウンロード側で行う。
        索引時に絞ると、あとでユニバースを広げたとき『走査済みの日』に載って
        いる新規銘柄の書類が永久に入らなくなる。"""
        self._pending("2962", "S100AAAA")      # ユニバース内
        self._pending("9999", "S100BBBB")      # 対象外
        calls = []

        class _F:
            def download(self, url, dest):
                calls.append(url)
                raise RuntimeError("ネットワークは踏まない")

        E.download_pending(self.con, _F(), codes={"2962"})
        self.assertEqual(len(calls), 1, "対象外の銘柄まで落としにいっている")
        self.assertIn("S100AAAA", calls[0])

    def test_widening_the_universe_does_not_refetch_what_is_already_there(self):
        """ユニバースを広げても既取得分は再取得しない。差分だけが対象になる
        —— これが『既取得分を無効にせず差分のみ追加取得』の実体。"""
        self._pending("2962", "S100AAAA", ok=1)   # 取得済み
        self._pending("278A", "S100CCCC", ok=0)   # 上限拡大で新たに対象化
        calls = []

        class _F:
            def download(self, url, dest):
                calls.append(url)
                raise RuntimeError("ネットワークは踏まない")

        E.download_pending(self.con, _F(), codes={"2962", "278A"})
        self.assertEqual(len(calls), 1, "取得済みを取り直している")
        self.assertIn("S100CCCC", calls[0])

    def test_weekday_sweep_skips_weekends(self):
        from datetime import date
        days = list(E.weekdays(date(2026, 8, 24), date(2026, 8, 30)))
        self.assertEqual([d.isoformat() for d in days],
                         ["2026-08-24", "2026-08-25", "2026-08-26",
                          "2026-08-27", "2026-08-28"])


if __name__ == "__main__":
    unittest.main(verbosity=2)


class TestEdinetXbrlParsing(unittest.TestCase):
    """EDINET(有報/半期)の iXBRL を短信と同じ経路で読むための語彙。

    かつて parse_archive は source='tdnet' 固定で、EDINET を落としても永久に
    解析されなかった。unknown_tags が空でも「未マップが無い」のではなく
    「一度も見ていない」だけ、という状態だった(2026-08-31 修正)。
    """

    def setUp(self):
        from screener.extract import xbrl_parser as X
        self.X = X
        self.m = X.load_mapping()

    def test_edinet_year_vocabulary(self):
        """短信は Prior/Prior2、EDINET は Prior1..Prior4 と数字を必ず付ける。
        有報の主要財務データは5期分載るので Prior4 まで実在する。"""
        for ctx, rel in (("CurrentYearDuration", "current"),
                         ("Prior1YearInstant", "prior"),
                         ("Prior2YearDuration", "prior2"),
                         ("Prior3YearInstant", "prior3"),
                         ("Prior4YearDuration", "prior4")):
            self.assertEqual(self.m.parse_context(ctx, "edinet")["year_rel"], rel, ctx)

    def test_interim_contexts_are_half_year(self):
        """半期報告書の Interim* は期首からの2四半期累計 -> q_no=2。
        member ではなくコンテキスト名そのものが四半期を表す。"""
        for ctx in ("InterimDuration", "InterimInstant"):
            d = self.m.parse_context(ctx, "edinet")
            self.assertEqual(d["q_no"], 2, ctx)
            self.assertEqual(d["year_rel"], "current", ctx)
        d = self.m.parse_context("Prior1InterimDuration", "edinet")
        self.assertEqual((d["year_rel"], d["q_no"]), ("prior", 2))

    def test_ytd_context_does_not_claim_a_quarter(self):
        """*YTDDuration から四半期を決めてはいけない。EDINET半期では2四半期
        累計だが、TDnet短信は Q1〜Q4 のどれでも CurrentYTDDuration を使う。
        ここで 2 を返すと短信の Q3 が q_no=2 として積まれ、単独値(累計の差分)が
        丸ごと狂う —— 2026-08-31 に実際に埋め込んで踏んだ。"""
        for src in ("edinet", "tdnet"):
            d = self.m.parse_context("CurrentYTDDuration", src)
            self.assertIsNone(d["q_no"], f"{src}: YTD が四半期を主張している")
            self.assertEqual(d["year_rel"], "current")

    def test_accumulated_quarter_context_gives_the_quarter(self):
        """短信サマリーの AccumulatedQ<n> は四半期を明示する唯一の手がかり。"""
        for n in (1, 2, 3):
            d = self.m.parse_context(
                f"CurrentAccumulatedQ{n}Duration_ConsolidatedMember_ResultMember",
                "tdnet")
            self.assertEqual(d["q_no"], n)
            self.assertEqual(d["year_rel"], "current")
        d = self.m.parse_context("PriorAccumulatedQ3Duration_ConsolidatedMember_"
                                 "ResultMember", "tdnet")
        self.assertEqual((d["year_rel"], d["q_no"]), ("prior", 3))

    def test_edinet_consolidated_context_has_no_member(self):
        """EDINET は連結に member を付けず、単体だけ NonConsolidatedMember が
        付く。member 無しを連結と補わないと、連結の数字が consolidation=None に
        なって prefer_consolidation の優先が効かなくなる。"""
        self.assertEqual(
            self.m.parse_context("CurrentYearDuration", "edinet")["consolidation"],
            "consolidated")
        self.assertEqual(
            self.m.parse_context("CurrentYearInstant_NonConsolidatedMember",
                                 "edinet")["consolidation"], "nonconsolidated")
        # 短信は両方に member が付くので、この補完を適用してはいけない
        self.assertIsNone(
            self.m.parse_context("CurrentYearDuration", "tdnet")["consolidation"])

    def test_period_label_uses_the_dei_fiscal_year_end(self):
        """12月期・11月期の会社は提出年と会計年度がずれる。EDINET は DEI に
        決算期末日を持っているので、開示日推定ではなくそれを使う。"""
        row = {"date": "2026-02-24", "code": "6217"}
        # 11月期。2026-02 提出だが当期は FY2025。
        self.assertEqual(
            self.X.period_label(row, {"year_rel": "current"}, "2025-11-30"), "FY2025")
        self.assertEqual(
            self.X.period_label(row, {"year_rel": "prior2"}, "2025-11-30"), "FY2023")
        # fy_end が無ければ従来どおり開示日から推定する(短信の経路)
        self.assertEqual(self.X.period_label(row, {"year_rel": "current"}), "FY2026")

    def test_fiscal_year_comes_from_dei_not_the_filing_date(self):
        """会計年度を提出日から推定してはいけない。3月期の会社が8月に出す
        第1四半期短信は「2027年3月期」で、提出年(2026)ではない。

        提出日推定だと FY2026 とラベルされ、EDINET由来の FY2026(上期・下期)と
        衝突して**別の会計年度どうしを引き算する**。3905 で粗利率115.3%という
        あり得ない値が出て発覚した(2026-08-31)。短信も有報も DEI に期末日を
        持っているので、推定する理由が最初から無かった。
        """
        row = {"date": "2026-08-14", "code": "3905"}
        self.assertEqual(
            self.X.period_label(row, {"year_rel": "current"}, "2027-03-31"), "FY2027")
        self.assertEqual(
            self.X.period_label(row, {"year_rel": "prior"}, "2027-03-31"), "FY2026")
        # DEI が無いときだけ提出日にフォールバックする
        self.assertEqual(self.X.period_label(row, {"year_rel": "current"}), "FY2026")

    def test_fy_end_is_read_from_tdnet_too(self):
        """fy_end の抽出を EDINET 限定にしていたのが上のバグの原因。
        短信は tse-ed-t:FiscalYearEnd と jpdei_cor:CurrentFiscalYearEndDateDEI の
        両方を持つ。どちらの綴りでも拾えること。"""
        for tag in ("jpdei_cor:CurrentFiscalYearEndDateDEI", "tse-ed-t:FiscalYearEnd"):
            self.assertTrue(
                any(tag.endswith(k) or tag == k for k in self.X._FY_END_TAGS),
                f"{tag} を拾えない")

    def test_only_publicdoc_ixbrl_is_read(self):
        """EDINET は `_ixbrl.htm`(アンダースコア)。AuditDoc は監査報告書で
        財務数値を持たず、読むと監査文言が unknown を無意味に膨らませる。"""
        names = ["XBRL/PublicDoc/0101010_honbun_x_ixbrl.htm",
                 "XBRL/AuditDoc/jpaud-aai_ixbrl.htm",
                 "XBRL/PublicDoc/x.xsd"]
        got = self.X._ixbrl_members(names, "edinet")
        self.assertEqual([n for n, _ in got],
                         ["XBRL/PublicDoc/0101010_honbun_x_ixbrl.htm"])
        # 短信はハイフン区切りで、Summary と Attachment を区別する
        td = ["XBRL/Summary/a-ixbrl.htm", "XBRL/Attachment/b-ixbrl.htm"]
        self.assertEqual([p for _, p in self.X._ixbrl_members(td, "tdnet")],
                         ["summary", "attachment"])


class TestEdinetCoverageReport(_DbCase):
    """索引をユニバース非依存にした結果、「未取得」の多くは設計どおりの
    対象外になった。これを失敗として数えると本物の失敗が埋もれる。"""

    def test_out_of_scope_documents_are_not_counted_as_failures(self):
        from screener.report import edinet_coverage as R
        self.add("2962", mktcap=20000, adv20=800)      # ユニバース内
        self.add("9999", mktcap=20000, adv20=1)        # 流動性で除外
        U.apply_universe_rules(self.con)
        for code, doc, ok in (("2962", "S1", 1), ("9999", "S2", 0)):
            self.con.execute(
                "INSERT INTO filings (code, date, type, source, doc_id, xbrl_ok, path) "
                "VALUES (?,?,?,?,?,?,?)",
                (code, "2026-06-01", "有報", "edinet", doc, ok,
                 "raw/x.zip" if ok else None))
        self.con.commit()
        logged = []
        orig, C.log = C.log, lambda m: logged.append(str(m))
        try:
            R.failures(self.con, 10)
        finally:
            C.log = orig
        joined = "\n".join(logged)
        self.assertIn("取得対象なのに XBRL が取れていない書類: 0 件", joined)
        self.assertIn("失敗ではない", joined)


class TestSegmentDimension(_DbCase):
    """docs/segment_dimension_design.md —— 次元付きファクトを軸を分けて持つ。"""

    def setUp(self):
        super().setUp()
        from screener.extract import xbrl_parser as X
        self.X = X

    def test_member_classification_separates_totals_from_individual_segments(self):
        """報告セグメント計を個別セグメントと同じ軸に置くと、セグメント別集計で
        全社ぶんがもう一度足される。軸を分けることでそれを防ぐ。"""
        f = self.X.classify_member
        self.assertEqual(f("jpcrp030000-asr_E02121-000JapanReportableSegmentsMember"),
                         ("segment", "Japan"))
        self.assertEqual(f("ReportableSegmentsMember")[0], "segment_total")
        self.assertEqual(f("TotalOfReportableSegmentsAndOthersMember")[0], "segment_total")
        self.assertEqual(f("ReconcilingItemsMember")[0], "adjustment")
        self.assertEqual(f("RetainedEarningsMember")[0], "equity_component")
        self.assertEqual(f("No1MajorShareholdersMember")[0], "other")

    def test_segment_change_invalidates_but_does_not_delete(self):
        """区分変更で時系列が切れた期は valid_flag=0。行は消さない ——
        『区分が変わった』と『まだ判定していない』を混同しないため。"""
        rows = [("FY2024", "Japan"), ("FY2024", "Philippines"),
                ("FY2025", "Japan"), ("FY2025", "Asia")]   # Asia が新設
        for i, (period, member) in enumerate(rows):
            # financials_dim.filing_id は filings への外部キー。親行を作る。
            cur = self.con.execute(
                "INSERT INTO filings (code, date, type, source, doc_id) "
                "VALUES (?,?,?,?,?)",
                ("3441", "2026-06-01", "有報", "edinet", f"D{i}"))
            self.con.execute(
                "INSERT INTO financials_dim (filing_id, code, period, item, axis, "
                " member, member_raw, value, context_ref) "
                "VALUES (?,?,?,?,?,?,?,?,?)",
                (cur.lastrowid, "3441", period, "revenue", "segment", member,
                 member + "Member", 100.0, f"ctx{i}"))
        self.con.commit()
        self.X.flag_segment_changes(self.con)
        got = {(r["period"], r["member"]): r["valid_flag"] for r in self.con.execute(
            "SELECT period, member, valid_flag FROM financials_dim")}
        self.assertEqual(got[("FY2025", "Asia")], 0, "新設セグメントが有効のまま")
        self.assertEqual(got[("FY2025", "Japan")], 1, "継続セグメントを無効にしている")
        self.assertEqual(got[("FY2024", "Japan")], 1, "最初の期は比較相手が無い")
        self.assertEqual(self.con.execute(
            "SELECT COUNT(*) c FROM financials_dim").fetchone()["c"], 4,
            "行を消している")
