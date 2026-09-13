"""004_edinet_quarter_from_dei.py — EDINET書類の q_no を DEI から埋め直す。

なぜ全件再パースではなく UPDATE なのか
--------------------------------------
q_no は store_filing が「文脈から四半期が決まらなかった行」に配るだけの列で、
行の採否にも期ラベルにも効かない。つまり値を後から入れても、再パースした
結果と同じものになる。18,827本の zip を丸ごと解析し直すと1.5時間かかるが、
DEI ヘッダだけ読めば数分で済む。

何を直すのか
------------
filing_quarter() は短信の `CurrentAccumulatedQ<n>Duration` しか見ていない。
EDINET の四半期報告書は文脈を `CurrentYTDDuration` としか書かないので、
2026-09-01 時点で subtype 140 の 70.5% / 160 の 38.5% が q_no NULL だった。
四半期が分からなければ「当期累計−前四半期累計」が作れないので、
四半期報告書7,950本が丸ごと死んでいた（実測: 5四半期以上連続する銘柄が1社）。

    python -m screener.db.migrations.004_edinet_quarter_from_dei
    python -m screener.db.migrations.004_edinet_quarter_from_dei --dry-run
"""
from __future__ import annotations

import argparse
import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.dirname(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))))))

from screener import common as C
from screener.extract.xbrl_parser import dei_quarter_from_zip


def main(argv=None) -> int:
    sys.stdout.reconfigure(encoding="utf-8")
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--dry-run", action="store_true")
    p.add_argument("--lock-wait", type=float, default=3600.0)
    a = p.parse_args(argv)

    con = C.init_db()
    rows = con.execute(
        "SELECT id, subtype, xbrl_path FROM filings "
        "WHERE source='edinet' AND xbrl_ok=1 AND xbrl_path IS NOT NULL "
        "AND id IN (SELECT DISTINCT filing_id FROM financials_cum "
        "           WHERE q_no IS NULL) ORDER BY id").fetchall()
    C.log(f"q_no が NULL の行を持つ EDINET 書類: {len(rows):,} 本")

    t0 = time.time()
    stats = {"read": 0, "no_dei": 0, "cum": 0, "dim": 0}
    for i, r in enumerate(rows, 1):
        q = dei_quarter_from_zip(C.full_path(r["xbrl_path"]))
        if q is None:
            stats["no_dei"] += 1
            continue
        stats["read"] += 1
        if not a.dry_run:
            stats["cum"] += con.execute(
                "UPDATE financials_cum SET q_no=? WHERE filing_id=? AND q_no IS NULL",
                (q, r["id"])).rowcount
            stats["dim"] += con.execute(
                "UPDATE financials_dim SET q_no=? WHERE filing_id=? AND q_no IS NULL",
                (q, r["id"])).rowcount
        if i % 2000 == 0:
            if not a.dry_run:
                con.commit()
            C.log(f"  [{i}/{len(rows)}] {time.time()-t0:.0f}s "
                  f"cum {stats['cum']:,} 行 / dim {stats['dim']:,} 行")
    if not a.dry_run:
        con.commit()
    C.log(f"DEI が読めた書類 {stats['read']:,} / 読めなかった {stats['no_dei']:,}")
    C.log(f"更新 financials_cum {stats['cum']:,} 行 / financials_dim {stats['dim']:,} 行"
          + ("  (dry-run: 実際には書いていない)" if a.dry_run else ""))
    left = con.execute("SELECT COUNT(*) c FROM financials_cum fc "
                       "JOIN filings f ON f.id=fc.filing_id "
                       "WHERE f.source='edinet' AND fc.q_no IS NULL").fetchone()["c"]
    C.log(f"残る q_no NULL (edinet): {left:,} 行")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
