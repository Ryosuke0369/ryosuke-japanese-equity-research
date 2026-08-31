"""screener/report/summary_vs_statement.py — 「主要な経営指標等の推移」と
財務諸表本体が同じ数字を言っているかの突合。

なぜ DB を見ずに zip を読み直すか
--------------------------------
financials_cum の主キーは (filing_id, item, context_ref)。
jppfs_cor:NetSales と jpcrp_cor:NetSalesSummaryOfBusinessResults は同じ
item(revenue) の同じ文脈(CurrentYearDuration)に落ちるので、INSERT OR REPLACE
で**片方が消えてから**でないと DB には現れない。消える前に比べる必要がある
ので、ここだけは zip を直接読む。

出すもの
  - 重複件数（両方が同じ item×文脈を持つ組）
  - 一致件数／不一致件数と、不一致の実例
  - Summary にしか無い期の件数（= 取り込みで純増する時系列の量）

Usage
    python -m screener.report.summary_vs_statement --limit 300
"""
from __future__ import annotations

import argparse
import os
import sys
from collections import defaultdict

try:
    from screener import common as C
except ImportError:                                     # pragma: no cover
    sys.path.insert(0, os.path.dirname(os.path.dirname(
        os.path.dirname(os.path.abspath(__file__)))))
    from screener import common as C

from screener.extract import xbrl_parser as X

SUFFIX = "SummaryOfBusinessResults"


def compare(con, limit: int, tol: float) -> dict:
    mapping = X.load_mapping()
    rows = con.execute(
        "SELECT id, code, date, xbrl_path FROM filings "
        "WHERE source='edinet' AND type='有報' AND xbrl_ok=1 "
        "AND xbrl_path IS NOT NULL ORDER BY date DESC LIMIT ?", (limit,)).fetchall()
    C.log(f"突合対象: {len(rows)} 件の有報")

    stat = {"filings": 0, "dup": 0, "same": 0, "diff": 0, "summary_only": 0}
    examples: list[str] = []
    by_item = defaultdict(lambda: {"dup": 0, "diff": 0, "only": 0})

    for r in rows:
        path = C.full_path(r["xbrl_path"])
        if not os.path.exists(path):
            continue
        try:
            facts = X.facts_from_zip(path, "edinet")
        except Exception:
            continue
        stat["filings"] += 1
        stmt: dict[tuple, float] = {}
        summ: dict[tuple, float] = {}
        for f in facts:
            if f["value"] is None:
                continue
            item = mapping.item_for(f["tag"])
            if item is None:
                continue
            # 内訳文脈は見出しではないので突合対象から外す（本体パーサと同じ規則）
            dims = mapping.parse_context(f["context"], "edinet")
            if any(p.endswith("Member") for p in dims["rest"]):
                continue
            key = (item, f["context"])
            (summ if f["tag"].endswith(SUFFIX) else stmt)[key] = f["value"]

        for key, v in summ.items():
            if key in stmt:
                stat["dup"] += 1
                by_item[key[0]]["dup"] += 1
                a, b = stmt[key], v
                scale = max(abs(a), abs(b), 1.0)
                if abs(a - b) / scale <= tol:
                    stat["same"] += 1
                else:
                    stat["diff"] += 1
                    by_item[key[0]]["diff"] += 1
                    if len(examples) < 15:
                        examples.append(
                            f"  {r['code']:<6} {r['date']} {key[0]:<16} "
                            f"{key[1]:<34} 本体={a:>18,.0f} 推移={b:>18,.0f}")
            else:
                stat["summary_only"] += 1
                by_item[key[0]]["only"] += 1

    C.log(f"読めた有報: {stat['filings']} 件")
    C.log(f"重複(同じ item×文脈を両方が持つ): {stat['dup']:,} 組")
    if stat["dup"]:
        C.log(f"  値が一致 : {stat['same']:,} 組 ({stat['same']/stat['dup']:.3%})")
        C.log(f"  値が不一致: {stat['diff']:,} 組 ({stat['diff']/stat['dup']:.3%})"
              f"   ← 許容誤差 {tol:.2%}")
    C.log(f"「推移」にしかない期: {stat['summary_only']:,} 件 "
          f"<- 取り込みで純増する時系列")
    C.log("項目別:")
    for item in sorted(by_item):
        d = by_item[item]
        C.log(f"  {item:<18} 重複 {d['dup']:>6,} / 不一致 {d['diff']:>5,} "
              f"/ 推移のみ {d['only']:>6,}")
    if examples:
        C.log("不一致の実例:")
        for e in examples:
            C.log(e)
    return stat


def main(argv=None) -> int:
    sys.stdout.reconfigure(encoding="utf-8")
    p = argparse.ArgumentParser(description="「主要な経営指標等の推移」と本体の突合")
    p.add_argument("--limit", type=int, default=300, help="突合する有報の件数")
    p.add_argument("--tol", type=float, default=0.005,
                   help="相対許容誤差。推移は百万円単位に丸められることがある")
    a = p.parse_args(argv)
    compare(C.connect(), a.limit, a.tol)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
