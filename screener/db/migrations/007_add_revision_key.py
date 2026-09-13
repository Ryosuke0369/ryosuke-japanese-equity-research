"""007_add_revision_key.py — 凍結レコードに revision を入れ、訂正版を追記できるようにする。

追記専用を守りながら訂正するには、「旧版を残したまま新版を入れられる」
必要がある。現状の UNIQUE(as_of, code, event_date) は旧版が鍵を占有するので
新版が入らない。**revision を鍵に含める**。

  revision=1  最初の凍結（枠制約バグを含む）
  revision=2  同一 as_of での再凍結（正しい枠計算）

SQLite は UNIQUE を後から変更できないので、テーブルを作り直して移送する。

    python -m screener.db.migrations.007_add_revision_key
"""
from __future__ import annotations

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))))))

from screener import common as C

SPECS = {
    "forecast_snapshots": ("as_of, code, event_date", "snapshot_id"),
    "paper_trades": ("code, entry_date, event_date", "trade_id"),
    "shadow_snapshots": ("variant, as_of, code, event_date", "snapshot_id"),
    "shadow_trades": ("variant, code, entry_date, event_date", "trade_id"),
}


def main(argv=None) -> int:
    sys.stdout.reconfigure(encoding="utf-8")
    con = C.init_db()
    con.execute("PRAGMA foreign_keys = OFF")
    for table, (keycols, pk) in SPECS.items():
        cols = [r[1] for r in con.execute("PRAGMA table_info(%s)" % table)]
        if "revision" in cols:
            C.log("  %s: revision は既にある" % table)
            continue
        # 旧定義を取り、UNIQUE 句に revision を足した新定義を作る
        ddl = con.execute("SELECT sql FROM sqlite_master WHERE name=?",
                          (table,)).fetchone()[0]
        old_u = "UNIQUE(%s)" % keycols
        assert old_u in ddl, "UNIQUE 句が見つからない: %s" % table
        # 列定義は UNIQUE 句より前に置く。UNIQUE の後ろに足すと SQLite の
        # 構文順序に反し「no such column: revision」で作成に失敗する。
        new_ddl = ddl.replace(table, table + "_new", 1).replace(
            old_u,
            "revision INTEGER NOT NULL DEFAULT 1,\n    " + old_u[:-1] + ", revision)")
        con.execute(new_ddl)
        collist = ", ".join(cols)
        con.execute("INSERT INTO %s_new (%s, revision) SELECT %s, 1 FROM %s"
                    % (table, collist, collist, table))
        n = con.execute("SELECT COUNT(*) FROM %s_new" % table).fetchone()[0]
        con.execute("DROP TABLE %s" % table)
        con.execute("ALTER TABLE %s_new RENAME TO %s" % (table, table))
        C.log("  %s: 作り直して %d 行を移送（revision=1）" % (table, n))
    con.commit()
    con.execute("PRAGMA foreign_keys = ON")
    C.log("integrity_check: %s" % con.execute("PRAGMA integrity_check").fetchone()[0])
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
