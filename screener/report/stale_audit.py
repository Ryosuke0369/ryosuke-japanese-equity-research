"""screener/report/stale_audit.py — 過去の推薦がどれだけ古い証拠に依存していたか。

タスク1のシャドウ計測。**スコアは一切変えない。**測るだけ。

対象
----
1. 保存済みの週次CSV（`--csv` で列挙）。`evidence_docs` 列があれば
   根拠期を読み、無ければ再計算する。
2. `forecast_snapshots` / `shadow_snapshots` に凍結された (as_of, code)。
   凍結時のスコアは旧実装（根拠期を持たない）なので、**当時の as_of で
   PIT 再計算**して鮮度を測る。当時見えていた期だけを使うので、
   「今から見れば古い」ではなく「**当時すでに古かった**」が出る。
"""
from __future__ import annotations

import argparse
import csv
import io
import os
import sqlite3
import sys
from datetime import date

try:
    from screener import common as C
except ImportError:                                     # pragma: no cover
    sys.path.insert(0, os.path.dirname(os.path.dirname(
        os.path.dirname(os.path.abspath(__file__)))))
    from screener import common as C

from screener.report import backtest_eval as V1
from screener.signals import freshness as FR
from screener.signals.span_runner import score_ticker


def _scorers():
    root = str(V1._external_root())
    if root not in sys.path:
        sys.path.insert(0, root)
    from module_b.run_scorers import SCORERS_ALL
    return SCORERS_ALL


def measure(pcon, pairs):
    """(as_of, code) の並びについて鮮度を測る。行ごとの dict を返す。"""
    scorers = _scorers()
    out = []
    for as_of, code in pairs:
        try:
            # 過去の推薦を当時の条件で遡及計測する道具なので、2026-09-13 の既定切替に追随させない。
            res = score_ticker(pcon, code, scorers, as_of=as_of, policy="prefer_span")
        except Exception as e:
            out.append({"as_of": as_of, "code": code, "error": str(e)})
            continue
        scores = {k: v for k, v in res.items()
                  if k not in ("evidence_score", "_source")}
        fired = {k for k, v in scores.items()
                 if isinstance(v, dict) and v.get("available")
                 and (v.get("score") or 0) > 0}
        FR.annotate(scores, pcon, code, as_of)
        flag, lag, detail = FR.summarize(scores, fired)
        out.append({"as_of": as_of, "code": code, "n_fired": len(fired),
                    "stale_flag": flag, "stale_lag": lag, "detail": detail,
                    "score": res.get("evidence_score")})
    return out


def report(rows, label):
    n = len(rows)
    err = sum(1 for r in rows if r.get("error"))
    fired = [r for r in rows if not r.get("error") and r["n_fired"] > 0]
    stale = [r for r in fired if r["stale_flag"]]
    undet = [r for r in fired if not r["stale_flag"] and "判定不能" in r["detail"]]
    C.log("")
    C.log("== %s: %d 件（例外 %d）" % (label, n, err))
    if not fired:
        C.log("   発火したシグナルを持つ行が無い")
        return
    C.log("   発火あり %d 件 / うち **古い証拠に依存 %d 件 (%.1f%%)**"
          % (len(fired), len(stale), len(stale) * 100.0 / len(fired)))
    C.log("   うち根拠期を持たない（判定不能）を含む行 %d 件" % len(undet))
    if stale:
        by = {}
        for r in stale:
            by[r["stale_lag"]] = by.get(r["stale_lag"], 0) + 1
        C.log("   遅れの分布（期）: %s"
              % ", ".join("%d期 %d件" % kv for kv in sorted(by.items())))
        C.log("   最も古い例:")
        for r in sorted(stale, key=lambda x: -x["stale_lag"])[:5]:
            C.log("     %s %s  %s" % (r["as_of"], r["code"], r["detail"][:100]))


def from_csv(path):
    """週次CSV から (as_of, code) を作る。as_of は est_date ではなく生成日。"""
    rows = list(csv.DictReader(io.open(path, encoding="utf-8-sig")))
    base = os.path.basename(path)
    digits = "".join(ch for ch in base if ch.isdigit())[:8]
    as_of = ("%s-%s-%s" % (digits[:4], digits[4:6], digits[6:8])
             if len(digits) == 8 else date.today().isoformat())
    return [(as_of, r["code"]) for r in rows if r.get("code")]


def main(argv=None):
    sys.stdout.reconfigure(encoding="utf-8")
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--csv", nargs="*", default=[], help="保存済み週次CSV")
    p.add_argument("--snapshots", action="store_true",
                   help="forecast_snapshots / shadow_snapshots も測る")
    a = p.parse_args(argv)

    pdb = os.path.join(C.DATA_DIR, "projection.db")
    pcon = sqlite3.connect("file:%s?mode=ro" % pdb.replace("\\", "/"), uri=True)
    pcon.row_factory = sqlite3.Row

    for path in a.csv:
        if not os.path.exists(path):
            C.log("! 見つからない: %s" % path)
            continue
        report(measure(pcon, from_csv(path)), os.path.basename(path))

    if a.snapshots:
        mcon = C.init_db()
        for tbl in ("forecast_snapshots", "shadow_snapshots"):
            try:
                pairs = [(r[0], r[1]) for r in mcon.execute(
                    "SELECT DISTINCT as_of, code FROM %s ORDER BY as_of, code" % tbl)]
            except sqlite3.OperationalError as e:
                C.log("! %s: %s" % (tbl, e))
                continue
            report(measure(pcon, pairs), tbl)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
