"""010 — s12_evidence に paragraph_class を足す（追記のみ）。

マクロ経済の定型文を Tier の加点根拠から外した（2026-09-02）。
**外したことを後から監査できるように**、根拠ごとに段落の種類
(macro / industry / company) を残す。列が無いと「なぜこの語が
拾われなかったのか」を人が確かめられない。

    python -m screener.db.migrations.010_s12_paragraph_class
"""
from __future__ import annotations

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))))))

from screener import common as C


def main(argv=None) -> int:
    sys.stdout.reconfigure(encoding="utf-8")
    con = C.init_db()
    have = {r[1] for r in con.execute("PRAGMA table_info(s12_evidence)")}
    if "paragraph_class" in have:
        C.log("  paragraph_class は既にある")
        return 0
    con.execute("ALTER TABLE s12_evidence ADD COLUMN paragraph_class TEXT")
    con.commit()
    C.log("  paragraph_class 列を追加（既存行は NULL = 再計算前）")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
