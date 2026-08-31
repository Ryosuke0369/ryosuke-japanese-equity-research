"""screener/report/edinet_coverage.py — EDINET一括取得の完了レポート。

夜間ランの後に「で、何が取れて何が取れなかったのか」を一画面で出す。
出すのは3つ:

  1. カバレッジ率 —— 取得対象(ユニバース候補 ∪ 検証8銘柄)のうち、有報/半期が
     1件でも取れた会社の割合。分母を「上場全社」にすると常に低く見えて意味が
     無いので、分母は必ず取得対象にする。
  2. 失敗一覧 —— 索引には載ったのに XBRL が落ちていない書類、および
     fetch_runs に failed で残っている実行。「0件だった」と「落とせなかった」を
     混ぜない。
  3. unknownタグ頻度 —— マッピングできなかったタグの上位。次に
     account_mapping.yaml へ足すべき候補がこれ。

Usage
    python -m screener.report.edinet_coverage
    python -m screener.report.edinet_coverage --unknown-top 40 --fail-limit 50
"""
from __future__ import annotations

import argparse
import os
import sys

try:
    from screener import common as C
except ImportError:                                     # pragma: no cover
    sys.path.insert(0, os.path.dirname(os.path.dirname(
        os.path.dirname(os.path.abspath(__file__)))))
    from screener import common as C

from screener.fetch import edinet_bulk as E


def coverage(con) -> dict:
    targets = E.universe_codes(con)
    rows = con.execute(
        "SELECT code, COUNT(*) n, SUM(COALESCE(xbrl_ok, 0)) ok "
        "FROM filings WHERE source='edinet' GROUP BY code").fetchall()
    got = {r["code"]: r for r in rows}

    indexed = sum(1 for c in targets if c in got)
    downloaded = sum(1 for c in targets if (got.get(c) or {})
                     and got[c]["ok"] and got[c]["ok"] > 0)
    n = max(len(targets), 1)
    C.log(f"取得対象: {len(targets)} 社 (ユニバース候補 ∪ 検証8銘柄)")
    C.log(f"  索引に載った会社   {indexed:>5} 社 ({indexed / n:6.1%})")
    C.log(f"  XBRLが取れた会社   {downloaded:>5} 社 ({downloaded / n:6.1%})  <- カバレッジ率")
    C.log(f"  1件も無い会社      {len(targets) - indexed:>5} 社")

    tot = con.execute("SELECT COUNT(*) c, SUM(COALESCE(xbrl_ok,0)) d "
                      "FROM filings WHERE source='edinet'").fetchone()
    C.log(f"書類ベース: 索引 {tot['c'] or 0} 件 / 取得済 {tot['d'] or 0} 件")
    for r in con.execute("SELECT type, COUNT(*) c, SUM(COALESCE(xbrl_ok,0)) d "
                         "FROM filings WHERE source='edinet' "
                         "GROUP BY type ORDER BY c DESC"):
        C.log(f"  {str(r['type']):<8} 索引 {r['c']:>6} / 取得 {r['d'] or 0:>6}")

    C.log("検証8銘柄 (仕様書 §6) —— ユニバース外でも必ず取得対象に入れている:")
    for code in E.VALIDATION_CODES:
        r = con.execute(
            "SELECT COUNT(*) c, SUM(COALESCE(xbrl_ok,0)) d, MIN(date) f, MAX(date) l "
            "FROM filings WHERE source='edinet' AND code=?", (code,)).fetchone()
        nm = con.execute("SELECT name, universe_flag FROM companies WHERE code=?",
                         (code,)).fetchone()
        flag = "" if (nm and nm["universe_flag"]) else " [ユニバース外]"
        C.log(f"  {code:<5} {str(nm['name'] if nm else '?')[:16]:<16} "
              f"索引={r['c'] or 0:<3} 取得={r['d'] or 0:<3} "
              f"{r['f'] or '-'} .. {r['l'] or '-'}{flag}")
    return {"targets": len(targets), "indexed": indexed, "downloaded": downloaded}


def failures(con, limit: int) -> None:
    C.log("--- 失敗一覧 ---")
    rows = con.execute(
        "SELECT code, date, type, path FROM filings "
        "WHERE source='edinet' AND COALESCE(xbrl_ok,0)=0 "
        "ORDER BY date DESC, code LIMIT ?", (limit,)).fetchall()
    n = con.execute("SELECT COUNT(*) c FROM filings WHERE source='edinet' "
                    "AND COALESCE(xbrl_ok,0)=0").fetchone()["c"]
    C.log(f"索引済みだが XBRL 未取得の書類: {n} 件"
          + (f" (先頭 {limit} 件を表示)" if n > limit else ""))
    for r in rows:
        C.log(f"  {r['code']:<6} {r['date']} {str(r['type']):<6} {r['path'] or ''}")

    runs = con.execute(
        "SELECT source, target_date, status, attempt, error FROM fetch_runs "
        "WHERE status<>'ok' AND source LIKE '%edinet%' "
        "ORDER BY id DESC LIMIT ?", (limit,)).fetchall()
    C.log(f"fetch_runs で ok でない実行: {len(runs)} 件")
    for r in runs:
        C.log(f"  {r['target_date']} {r['status']} attempt={r['attempt']} "
              f"{str(r['error'] or '')[:120]}")


def unknown_tags(con, top: int) -> None:
    C.log("--- unknown タグ頻度 (account_mapping.yaml に足す候補) ---")
    rows = con.execute(
        "SELECT source, tag, n, n_filings, sample_value FROM unknown_tags "
        "ORDER BY n DESC LIMIT ?", (top,)).fetchall()
    if not rows:
        C.log("  (unknown_tags は空)")
    for r in rows:
        C.log(f"  {r['n']:>7} 回 / {r['n_filings']:>5} 書類  "
              f"{str(r['source']):<14} {r['tag'][:60]}")
    # EDINET の XBRL はまだ解析経路が無い。空のときに「未マップが無い」と
    # 誤読されると、対応漏れが見えなくなるので必ず断っておく。
    src = con.execute("SELECT DISTINCT source FROM unknown_tags").fetchall()
    srcs = {r["source"] for r in src}
    if not any(str(s).startswith("edinet") for s in srcs):
        C.log("  注: 集計は TDnet 短信のみ。xbrl_parser.parse_archive は "
              "source='tdnet' に固定されており、EDINET の有報/半期を解析する"
              "経路がまだ無い。上の頻度に EDINET 分は含まれていない。")


def main(argv=None) -> int:
    sys.stdout.reconfigure(encoding="utf-8")
    p = argparse.ArgumentParser(description="EDINET一括取得の完了レポート")
    p.add_argument("--unknown-top", type=int, default=25)
    p.add_argument("--fail-limit", type=int, default=30)
    a = p.parse_args(argv)

    con = C.connect()
    coverage(con)
    failures(con, a.fail_limit)
    unknown_tags(con, a.unknown_top)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
