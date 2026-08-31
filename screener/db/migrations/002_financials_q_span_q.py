r"""002 - financials_q に span_q を足す。

`CREATE TABLE IF NOT EXISTS` は既存テーブルに列を足さない。schema.sql に
span_q を書いても、既に financials_q を持っている DB には反映されないので、
ALTER TABLE で埋める。

span_q は「この単独値が何四半期ぶんか」。1 が真の四半期単独値。短信が揃って
いない会社は EDINET の有報(q4累計)と半期(q2累計)しか無く、その差分は
「下期6ヶ月」になる。それを四半期と名乗らせると、3ヶ月と6ヶ月が同じ土俵に
載って傾きの大きさが二重になる —— 単独値の意味そのものが壊れる。

既存行は span_q=1 で埋める。2026-08-31 時点で financials_q は空なので実害は
無いが、行がある DB で走らせても壊さないように既定値を入れる。

    python -m screener.db.migrations.002_financials_q_span_q [--apply]

--apply 無しは報告のみ(ドライラン)。何度走らせても安全。
"""
from __future__ import annotations

import argparse
import os
import sys

try:
    from screener import common as C
except ImportError:                                     # pragma: no cover
    sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(
        os.path.dirname(os.path.abspath(__file__))))))
    from screener import common as C


def has_column(con, table: str, column: str) -> bool:
    return any(r[1] == column for r in con.execute(f"PRAGMA table_info({table})"))


def main(argv=None) -> int:
    sys.stdout.reconfigure(encoding="utf-8")
    p = argparse.ArgumentParser(description="financials_q に span_q を追加")
    p.add_argument("--apply", action="store_true")
    a = p.parse_args(argv)

    con = C.connect()
    if not has_column(con, "financials_q", "code"):
        C.log("financials_q が無い。init_db が schema.sql から作るので何もしない")
        return 0
    if has_column(con, "financials_q", "span_q"):
        C.log("span_q は既にある。何もしない")
        return 0

    n = con.execute("SELECT COUNT(*) c FROM financials_q").fetchone()["c"]
    C.log(f"financials_q {n} 行に span_q(既定1)を追加する")
    if not a.apply:
        C.log("ドライラン。実際に変更するには --apply を付ける")
        return 0
    con.execute("ALTER TABLE financials_q ADD COLUMN span_q INTEGER DEFAULT 1")
    con.execute("UPDATE financials_q SET span_q = 1 WHERE span_q IS NULL")
    con.commit()
    C.log("span_q を追加した")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
