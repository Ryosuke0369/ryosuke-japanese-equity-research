"""008_presentation_doc_kind.py — 説明会資料に doc_kind を足し、世代を振り直す。

2026-09-02 の初回収集で2つの誤りが判明した。

1. **世代が種類を跨いでいた**: 同じ期に本編・サマリ版・書き起こし・質疑応答が
   並行して出るのが普通で、それらを1本の世代列に並べると「最新世代を使う」が
   **訂正版ではなく単に後から出た別文書**を指してしまう。
2. **世代順が処理順だった**: 訂正版が原本より前の世代になっていた
   （2477: gen1=08-24訂正 / gen2=08-20原本）。

対応: (code, period_label, doc_kind) で世代を振り、順序は disclosed_date 昇順、
同日は訂正を必ず後置する。

    python -m screener.db.migrations.008_presentation_doc_kind
"""
from __future__ import annotations

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))))))

from screener import common as C
from screener.extract.presentation_text import classify_doc_kind


def main(argv=None) -> int:
    sys.stdout.reconfigure(encoding="utf-8")
    con = C.init_db()
    have = {r[1] for r in con.execute("PRAGMA table_info(presentation_materials)")}
    if "doc_kind" not in have:
        con.execute("ALTER TABLE presentation_materials ADD COLUMN doc_kind TEXT")
        C.log("  doc_kind 列を追加")
    rows = con.execute(
        "SELECT filing_id, title FROM presentation_materials").fetchall()
    for r in rows:
        con.execute("UPDATE presentation_materials SET doc_kind=? WHERE filing_id=?",
                    (classify_doc_kind(r["title"]), r["filing_id"]))
    con.commit()
    C.log("  doc_kind を %d 件に付与" % len(rows))

    # 世代を振り直す。disclosed_date 昇順、同日は訂正を後置。
    groups = con.execute(
        "SELECT code, period_label, doc_kind FROM presentation_materials "
        "WHERE period_label IS NOT NULL GROUP BY code, period_label, doc_kind"
    ).fetchall()
    n = 0
    for g in groups:
        items = con.execute(
            "SELECT filing_id FROM presentation_materials "
            "WHERE code=? AND period_label IS ? AND doc_kind IS ? "
            "ORDER BY disclosed_date, is_correction, filing_id",
            (g["code"], g["period_label"], g["doc_kind"])).fetchall()
        for i, it in enumerate(items, 1):
            con.execute("UPDATE presentation_materials SET generation=? "
                        "WHERE filing_id=?", (i, it["filing_id"]))
            n += 1
    con.commit()
    C.log("  世代を振り直し: %d 件 / %d グループ" % (n, len(groups)))

    C.log("=== doc_kind の分布 ===")
    for r in con.execute("SELECT doc_kind, COUNT(*) n FROM presentation_materials "
                         "GROUP BY doc_kind ORDER BY n DESC"):
        C.log("  %-10s %3d" % (r["doc_kind"], r["n"]))
    unk = con.execute("SELECT COUNT(*) c FROM presentation_materials "
                      "WHERE doc_kind='unknown'").fetchone()["c"]
    tot = con.execute("SELECT COUNT(*) c FROM presentation_materials").fetchone()["c"]
    C.log("  unknown 率: %.1f%%（高いなら分類器を見直す）" % (unk / max(tot, 1) * 100))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
