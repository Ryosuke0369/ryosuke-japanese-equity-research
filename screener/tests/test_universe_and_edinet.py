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
        # (2026-08-30 の 600億->1,000億 で実際に落ちた)。yaml から境界を取る。
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
        # 四半期報告書(140)は制度廃止済み。仕様書 §2-2 のとおり対象外。
        self.assertNotIn("140", E.DOC_TYPES)

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
