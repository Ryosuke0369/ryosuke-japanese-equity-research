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
        self.add("1111", mktcap=4000, adv20=80)           # 40億 < 50億
        self.add("2222", mktcap=90000, adv20=80)          # 900億 > 600億
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

    def test_trial_codes_are_deterministic(self):
        for i in range(30):
            self.add(f"{9000 + i}")
        self.assertEqual(E.trial_codes(self.con, extra=10),
                         E.trial_codes(self.con, extra=10),
                         "試走の対象がランで変わると本番の見積りにならない")

    def test_weekday_sweep_skips_weekends(self):
        from datetime import date
        days = list(E.weekdays(date(2026, 8, 24), date(2026, 8, 30)))
        self.assertEqual([d.isoformat() for d in days],
                         ["2026-08-24", "2026-08-25", "2026-08-26",
                          "2026-08-27", "2026-08-28"])


if __name__ == "__main__":
    unittest.main(verbosity=2)
