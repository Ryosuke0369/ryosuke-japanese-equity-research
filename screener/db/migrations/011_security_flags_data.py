"""011 — security_flags の実データ投入（Kimi 側で一次ソース確認済み）。

009 では JPX ページに到達できず指定日・URL を NULL にしていた。
**確認できたので埋める。推測ではなく一次ソース由来。**

flag_type を1つ増やす。

  special_alert        特別注意銘柄。**主出力から除外する。**
                       「計上済み数字＝証拠」の前提が取引所に否定されている。
  listing_maintenance  上場維持基準の未適合（改善期間中）。**除外しない。**
                       数字の信頼性の話ではなく、上場そのものの継続性の話。
                       証拠としての数字は生きているので、表示フラグに留める。

    python -m screener.db.migrations.011_security_flags_data
"""
from __future__ import annotations

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))))))

from screener import common as C

ROWS = [
    ("4813", "special_alert", "2025-08-27", None,
     "https://www.jpx.co.jp/news/1023/20250826-11.html",
     "適時開示規定違反（過年度不適切会計、2018年1月期〜）。"
     "上場契約違約金 4,800万円。"),
    ("4813", "listing_maintenance", "2026-04-30", None,
     None,
     "流通株式時価総額 100億円未達（2026-01-31時点）。改善期間入りの開示日。"
     "2027-01-31 時点で未適合なら監理銘柄、2027-08-01 上場廃止の可能性。"),
]


def main(argv=None) -> int:
    sys.stdout.reconfigure(encoding="utf-8")
    con = C.init_db()
    n_ins = n_upd = 0
    for code, ftype, since, until, url, note in ROWS:
        old = con.execute(
            "SELECT since_date FROM security_flags WHERE code=? AND flag_type=?",
            (code, ftype)).fetchone()
        if old is None:
            con.execute(
                "INSERT INTO security_flags (code, flag_type, since_date,"
                " until_date, source_url, note) VALUES (?,?,?,?,?,?)",
                (code, ftype, since, until, url, note))
            n_ins += 1
        else:
            # 009 で指定日 NULL のまま入れた行を、一次ソースで確認した値に直す。
            con.execute(
                "UPDATE security_flags SET since_date=?, until_date=?,"
                " source_url=?, note=? WHERE code=? AND flag_type=? "
                "AND since_date IS ?",
                (since, until, url, note, code, ftype, old[0]))
            n_upd += 1
    con.commit()
    C.log("  追加 %d / 更新 %d" % (n_ins, n_upd))
    for r in con.execute("SELECT code, flag_type, since_date, source_url, note "
                         "FROM security_flags ORDER BY code, flag_type"):
        C.log("  %s %-20s since=%s  %s" % (r[0], r[1], r[2] or "未確認",
                                           (r[4] or "")[:44]))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
