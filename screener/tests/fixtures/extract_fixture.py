"""screener/tests/fixtures/extract_fixture.py — 実DBから回帰フィクスチャを切り出す。

2026-09-13 の `false_positive_20260913.json` は手作業で作られ、作り方が
残っていなかった。同じ形のフィクスチャを増やすたびに手で SQL を書くのは
再現できないので、切り出しをスクリプトにする。**銘柄コードは引数で渡す**
（CLAUDE.md「銘柄コード入りスクリプトを作らない」）。

    python -m screener.tests.fixtures.extract_fixture \
        --codes 3441 6838 --as-of 2026-09-18 \
        --out screener/tests/fixtures/signal_design_20260918.json \
        --readme "..."
"""
from __future__ import annotations

import argparse
import json
import os
import sqlite3
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(
    os.path.dirname(os.path.abspath(__file__)))))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

from screener import common as C                             # noqa: E402

PROJECTION_TABLES = ("universe", "filings", "quarterly_standalone",
                     "quarterly_standalone_all", "balance_sheet_items",
                     "pl_adjustments", "company_forecasts", "evidence_events")
# 本体DBからも持ってくる（S13 の四半期推移テストに要る）。
MAIN_TABLES = {"s13_orders": "code", "filings": "code",
               "financials_cum": "code", "disclosure_flags": "code"}


def _dump(con, table, key, codes):
    try:
        cols = [r[1] for r in con.execute("PRAGMA table_info(%s)" % table)]
    except sqlite3.OperationalError:
        return []
    if not cols:
        return []
    if key not in cols:
        return []
    q = "SELECT %s FROM %s WHERE %s IN (%s)" % (
        ",".join(cols), table, key, ",".join("?" * len(codes)))
    return [dict(zip(cols, r)) for r in con.execute(q, codes)]


def main(argv=None):
    sys.stdout.reconfigure(encoding="utf-8")
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--codes", nargs="+", required=True)
    p.add_argument("--as-of", required=True)
    p.add_argument("--out", required=True)
    p.add_argument("--readme", default="")
    p.add_argument("--projection-db",
                   default=os.path.join(C.DATA_DIR, "projection.db"))
    p.add_argument("--main-db", default=C.DB_PATH)
    a = p.parse_args(argv)

    pcon = sqlite3.connect(a.projection_db)
    mcon = sqlite3.connect(a.main_db)
    out = {"_README": a.readme, "as_of": a.as_of,
           "projection": {}, "main": {}}
    for t in PROJECTION_TABLES:
        out["projection"][t] = _dump(pcon, t, "ticker", a.codes)
    for t, key in MAIN_TABLES.items():
        out["main"][t] = _dump(mcon, t, key, a.codes)
    # 本体 filings の id は s13_orders.filing_id と突き合わせるので残す。
    os.makedirs(os.path.dirname(os.path.abspath(a.out)), exist_ok=True)
    with open(a.out, "w", encoding="utf-8") as fh:
        json.dump(out, fh, ensure_ascii=False, indent=1)
    for sec in ("projection", "main"):
        print(sec, {k: len(v) for k, v in out[sec].items()})
    print("→", a.out)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
