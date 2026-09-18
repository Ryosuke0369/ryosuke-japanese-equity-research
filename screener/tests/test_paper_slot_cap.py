"""screener/tests/test_paper_slot_cap.py — 建玉枠の不変条件。

守るのは1つ: **どの週を回しても同時保有が上限を超えない。**

2026-09-02、初回凍結で上限10に対し13件建った。週次バッチは1回の実行で
月〜金の複数日ぶんを決めるが、DBへの書き込みはループの後なので、
DBだけを見ると「9/15 に4件建てた」ことが 9/17 の判定から見えず、
各日が「枠は10空いている」と誤認していた。

事前登録(v2-1)は「同時保有10枠をハード制約」と定めている。これは
**値の変更ではなく、値を正しく適用できていない実装の修正**である。
"""
from __future__ import annotations

import os
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(
    os.path.dirname(os.path.abspath(__file__)))))

from screener import common as C
from screener.report import paper_weekly as PW


class _Rows(list):
    """sqlite3 のカーソル互換。呼び出し側が fetchall() を使うため。"""

    def fetchall(self):
        return list(self)

    def fetchone(self):
        return self[0] if self else None


class _FakeProjection:
    """投影DBの代わり。市場フィルターとセクターだけ答える。"""

    def __init__(self, days, index_close, sectors):
        self._days = days
        self._idx = index_close
        self._sectors = sectors

    def execute(self, sql, args=()):
        if "market_index" in sql:
            return _Rows(zip(self._days, self._idx))
        if "FROM universe" in sql:
            return _Rows(self._sectors.items())
        if "daily_prices" in sql:
            return _Rows((d,) for d in self._days)
        return _Rows()


class SlotCapTest(unittest.TestCase):
    def setUp(self):
        fd, self.db = tempfile.mkstemp(suffix=".db")
        os.close(fd)
        self.con = C.init_db(self.db)
        # 200本以上ないと市場フィルターが常に False になる
        self.days = ["2026-%02d-%02d" % (m, d)
                     for m in range(1, 10) for d in range(1, 29)][:260]
        self.idx = [1000.0 + i for i in range(len(self.days))]   # 単調上昇=フィルター通過
        self.sectors = {}

    def tearDown(self):
        self.con.close()
        for s in ("", "-wal", "-shm"):
            try:
                os.remove(self.db + s)
            except OSError:
                pass

    def _scored(self, n_per_day, days, score=0.7, n_avail=5):
        out = []
        for d in days:
            for i in range(n_per_day):
                code = "%04d" % (1000 + len(out))
                self.sectors[code] = "セクター%d" % (i % 4)
                out.append({
                    "code": code, "event_date": "2026-10-01", "entry_date": d,
                    "period_end": "FY2027-Q2", "quarter_type": "2Q",
                    "evidence_score": score,
                    "scores": {"S%d" % k: {"score": score, "available": k <= n_avail}
                               for k in range(1, 9)},
                })
        return out

    def _pcon(self):
        return _FakeProjection(self.days, self.idx, self.sectors)

    def test_v2_never_exceeds_ten_open_positions(self):
        """複数日にまたがる1回の実行でも、建玉は10を超えない。"""
        week = ["2026-09-15", "2026-09-16", "2026-09-17", "2026-09-18"]
        scored = self._scored(6, week)              # 4日 x 6件 = 24候補
        frozen, entries = PW.decide(self._pcon(), self.con, scored, dry_run=True)
        self.assertLessEqual(len(entries), PW.MAX_POSITIONS,
                             "枠を超えて建玉している（日跨ぎの累積が効いていない）")
        self.assertEqual(len(entries), PW.MAX_POSITIONS,
                         "候補は十分あるのに枠を使い切っていない")
        self.assertTrue(any(r["decision"] == "skip_full" for r in frozen),
                        "枠満杯のスキップが記録されていない")

    def test_shadow_b_respects_exposure_cap(self):
        """シャドウBの合計エクスポージャは 100% を超えない。"""
        week = ["2026-09-15", "2026-09-16", "2026-09-17", "2026-09-18"]
        scored = self._scored(6, week, score=0.7)   # 0.7 -> 15% ずつ
        frozen, entries = PW.decide_variant(self._pcon(), self.con, scored,
                                            "B", dry_run=True)
        expo = sum(e["size"] for e in entries)
        self.assertLessEqual(expo, PW.VB_MAX_EXPOSURE + 1e-9,
                             "エクスポージャ上限を超えている: %.2f" % expo)
        self.assertLessEqual(len(entries), PW.MAX_POSITIONS)

    # ------------------------------------------------ シャドウD（§35b / P3-2）
    def test_shadow_d_縮小は採用本数が多い方を上に持ち上げる(self):
        """素のスコアが同じなら、採用本数が多い銘柄が上位に来る。

        現行（k=0）は採用1本も5本も同じ 0.7 で並び、順位はコード順の
        同値処理で決まっていた。縮小はそこを分ける。
        """
        week = ["2026-09-15"]
        scored = self._scored(2, week, score=0.7, n_avail=1)    # 採用1本
        thick = self._scored(2, week, score=0.7, n_avail=4)     # 採用4本
        for e in thick:
            e["code"] = "9" + e["code"][1:]                     # コード順では後ろ
            self.sectors[e["code"]] = "セクターX"
        frozen, entries = PW.decide_variant(self._pcon(), self.con,
                                            scored + thick, "D", dry_run=True)
        got = [e["code"] for e in entries]
        self.assertTrue(got[0].startswith("9"),
                        "採用本数の多い銘柄が先頭に来ていない: %r" % got)

    def test_shadow_d_縮小後に閾値を割ったらスキップする(self):
        """0.15 の採用1本は縮小で 0.075 になり、閾値 0.10 を割る。"""
        week = ["2026-09-15"]
        scored = self._scored(3, week, score=0.15, n_avail=1)
        frozen, entries = PW.decide_variant(self._pcon(), self.con, scored,
                                            "D", dry_run=True)
        self.assertEqual(entries, [], "縮小後に閾値を割った候補が建玉になっている")
        self.assertTrue(all(r["decision"] == "skip_score" for r in frozen))
        # 同じ候補を v2 と同じ扱い（C は閾値だけ違う）にすると通ることを対照で示す
        _f, e_b = PW.decide_variant(self._pcon(), self.con, scored, "B",
                                    dry_run=True)
        self.assertTrue(e_b, "対照（B）でも建たないなら、この検証は意味がない")

    def test_shadow_d_素のスコアをsnapshotに残す(self):
        """**縮小後の値で素のスコアを上書きしない。** 横比較ができなくなる。"""
        week = ["2026-09-15"]
        scored = self._scored(1, week, score=0.8, n_avail=2)
        frozen, _e = PW.decide_variant(self._pcon(), self.con, scored, "D",
                                       dry_run=True)
        r = frozen[0]
        self.assertAlmostEqual(r["evidence_score"], 0.8)
        self.assertEqual(r["n_available"], 2)
        self.assertIn("adj=0.533", r["note"])       # 0.8 * 2/3

    def test_vd_adjust_の倍率(self):
        self.assertAlmostEqual(PW.vd_adjust(1.0, 1), 0.500, places=3)
        self.assertAlmostEqual(PW.vd_adjust(1.0, 2), 0.667, places=3)
        self.assertAlmostEqual(PW.vd_adjust(1.0, 3), 0.750, places=3)
        self.assertAlmostEqual(PW.vd_adjust(1.0, 4), 0.800, places=3)
        # 負のスコアも中立(0)へ寄る。片側だけ縮めると符号で扱いが変わる
        self.assertAlmostEqual(PW.vd_adjust(-1.0, 1), -0.500, places=3)
        # **「評価できていない」を「評価して0点」に化けさせない**
        self.assertIsNone(PW.vd_adjust(None, 3))
        self.assertIsNone(PW.vd_adjust(0.5, 0))

    def test_v3_respects_sector_cap_across_days(self):
        """同一セクターの同時保有上限も日を跨いで効く。"""
        week = ["2026-09-15", "2026-09-16", "2026-09-17"]
        scored = []
        for d in week:
            for i in range(4):
                code = "%04d" % (2000 + len(scored))
                self.sectors[code] = "同一セクター"
                scored.append({
                    "code": code, "event_date": "2026-10-01", "entry_date": d,
                    "period_end": "FY2027-Q2", "quarter_type": "2Q",
                    "evidence_score": 0.7,
                    "scores": {"S%d" % k: {"score": 0.7, "available": True}
                               for k in range(1, 9)},
                })
        frozen, entries = PW.decide_v3(self._pcon(), self.con, scored, dry_run=True)
        self.assertLessEqual(len(entries), PW.V3_MAX_PER_SECTOR,
                             "同一セクター上限が日跨ぎで効いていない")

    def test_existing_open_positions_reduce_free_slots(self):
        """DBに既に建玉があれば、その分だけ枠が減る。"""
        for i in range(8):
            self.con.execute(
                "INSERT INTO paper_trades (code, entry_date, position_size, status) "
                "VALUES (?,?,?, 'open')", ("%04d" % (3000 + i), "2026-09-01", 0.1))
        self.con.commit()
        scored = self._scored(6, ["2026-09-15", "2026-09-16"])
        _f, entries = PW.decide(self._pcon(), self.con, scored, dry_run=True)
        self.assertEqual(len(entries), 2, "既存8件に対して2枠しか空いていないはず")


if __name__ == "__main__":
    unittest.main()
