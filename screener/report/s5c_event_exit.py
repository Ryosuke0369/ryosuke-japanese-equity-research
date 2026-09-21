"""screener/report/s5c_event_exit.py — シャドウL: 「修正が出たら降りる」出口。

定義の正本は docs/backtest_acceptance_criteria.md「シャドウL」（2026-09-21・測定前に確定）。
**このスクリプトは定義を実装するだけで、N も判定基準も決めない。**

    入口   : R ≥ 1.30（§55 と同じ集合。`s5c_trades.csv` の fired=1 をそのまま使う）
    出口   : 修正が出たら**その翌営業日の始値**／出なければ **N 営業日後の終値**
    N      : 5 / 10 / 20（この3点だけ）
    判定   : 判定日クラスタ平均 > 0 かつ t ≥ 2

**S5c は 0点・表示のみ。この結果でスコアには接続しない**（2026-09-21 ユーザー判断）。

    python -m screener.report.s5c_event_exit
"""
from __future__ import annotations

import argparse
import bisect
import csv
import os
import sqlite3
import statistics
import sys
from collections import Counter, defaultdict

try:
    from screener import common as C
except ImportError:                                     # pragma: no cover
    sys.path.insert(0, os.path.dirname(os.path.dirname(
        os.path.dirname(os.path.abspath(__file__)))))
    from screener import common as C

from screener.report import s5c_power as W
from screener.report import s5c_trade as K

N_DAYS = (5, 10, 20)            # L-2。この3点だけ
COST = K.COST                   # 往復 0.4%
MIN_N = 100


def load_entries(path):
    """§55 の建玉 CSV から R ≥ 1.30 の入口だけを読む。"""
    out = []
    with open(path, encoding="utf-8-sig") as fh:
        for r in csv.DictReader(fh):
            if r["fired"] == "1":
                out.append({"as_of": r["as_of"], "code": r["code"], "r": float(r["r"])})
    return out


def price_panel(con, codes):
    """code -> {date: (始値, 終値)}（分割調整済み）。"""
    px = defaultdict(dict)
    for code, d, o, c, ac in con.execute(
            "SELECT code, date, open, close, adj_close FROM prices "
            "WHERE adj_close IS NOT NULL AND close IS NOT NULL AND open IS NOT NULL"):
        if code in codes and c:
            px[code][d] = (o * (ac / c), ac)
    return px


def first_revision(evs, code, after, until):
    """(after, until] の最初の修正イベント。"""
    for d, direction, _mag in evs.get(code, ()):
        if after < d <= until:
            return d, direction
    return None, None


def simulate(entries, evs, px, idx, cal, n_days, up_only=False):
    """L-2 の出口。戻り値は建玉の列と、出口の内訳。"""
    rows, why = [], Counter()
    for e in entries:
        s = px.get(e["code"])
        if not s:
            why["価格が無い"] += 1
            continue
        i0 = bisect.bisect_left(cal, e["as_of"])
        if i0 >= len(cal) or cal[i0] != e["as_of"] or e["as_of"] not in s:
            why["エントリー日の価格が無い"] += 1
            continue
        hard_i = min(i0 + n_days, len(cal) - 1)
        hard_date = cal[hard_i]
        p0 = s[e["as_of"]][1]                       # 判定日の終値で建てる
        rev_date, rev_dir = first_revision(evs, e["code"], e["as_of"], hard_date)
        if up_only and rev_dir == "down":
            rev_date = None                          # 副: down では降りない
        exit_date = exit_px = None
        if rev_date:
            j = bisect.bisect_right(cal, rev_date)   # 開示の翌営業日
            for k in range(j, min(j + 5, len(cal))):
                if cal[k] in s:
                    exit_date, exit_px = cal[k], s[cal[k]][0]   # 始値
                    break
            why["修正で降りた"] += 1 if exit_date else 0
        if exit_date is None:
            for k in range(hard_i, max(i0, hard_i - 5) - 1, -1):
                if cal[k] in s:
                    exit_date, exit_px = cal[k], s[cal[k]][1]   # 終値
                    break
            why["N日で降りた"] += 1 if exit_date else 0
        if exit_date is None or not p0:
            why["出口の価格が無い"] += 1
            continue
        gross = exit_px / p0 - 1
        i_a, i_b = idx.get(e["as_of"]), idx.get(exit_date)
        bench = (i_b / i_a - 1) if (i_a and i_b) else None
        rows.append({"as_of": e["as_of"], "code": e["code"], "r": e["r"],
                     "exit_date": exit_date, "exit_by": "修正" if rev_date else "N日",
                     "rev_dir": rev_dir or "", "hold_days": cal.index(exit_date) - i0,
                     "ret_gross": gross, "ret_net": gross - COST, "bench": bench,
                     "excess": None if bench is None else (gross - COST) - bench})
    return rows, why


def verdict(s):
    if not s or s["cl_t"] is None:
        return "判定不能"
    return ("**満たす**" if (s["cl_mean"] > 0 and s["cl_t"] >= 2)
            else "満たさない")


def render(results, n_entries):
    L = ["# シャドウL: 「修正が出たら降りる」出口（記録専用・0点）", ""]
    L.append("- 入口は §55 と同じ **R ≥ 1.30 の %s 観測**（判定日の終値で建てる・往復 %.1f%%）"
             % (f"{n_entries:,}", 100 * COST))
    L.append("- 出口: 修正が出たら**翌営業日の始値** / 出なければ **N 営業日後の終値**")
    L.append("")
    L.append("## 主判定（up / down いずれの修正でも降りる）")
    L.append("")
    L.append(K.HDR)
    for n in N_DAYS:
        L.append(K._line(results[("main", n)]["sum"]))
    L.append("")
    L.append("| N | 判定（クラスタ平均>0 かつ t≥2） | 修正で降りた | N日で降りた | 平均保有日数 |")
    L.append("|---|---|---|---|---|")
    for n in N_DAYS:
        res = results[("main", n)]
        why, rows = res["why"], res["rows"]
        hold = statistics.mean([r["hold_days"] for r in rows]) if rows else 0
        L.append("| %d | %s | %d | %d | %.1f |"
                 % (n, verdict(res["sum"]), why["修正で降りた"], why["N日で降りた"], hold))
    L.append("")
    L.append("## 副（記録のみ・up のときだけ降りる。down は N まで持つ）")
    L.append("")
    L.append(K.HDR)
    for n in N_DAYS:
        L.append(K._line(results[("up_only", n)]["sum"]))
    L.append("")
    L.append("## 修正で降りた建玉だけ（事後・主判定には使わない）")
    L.append("")
    L.append(K.HDR)
    for n in N_DAYS:
        sub = [r for r in results[("main", n)]["rows"] if r["exit_by"] == "修正"]
        L.append(K._line(K.summarise("N=%d 修正で降りた" % n, sub)))
    L.append("")
    L.append("## L-4 の結論")
    L.append("")
    ok = [n for n in N_DAYS if verdict(results[("main", n)]["sum"]) == "**満たす**"]
    if ok:
        L.append("N = %s が基準を満たした。**それでもスコアには接続しない**（2026-09-21 のユーザー判断）。"
                 % ", ".join(str(n) for n in ok))
    else:
        L.append("**3点とも基準を満たさない。**")
        L.append("")
        L.append("固定1か月（§55）に続き、事象で降りる形でも基準を満たさなかった。"
                 "登録した扱いのとおり、**修正先回り型は現行データでは成立しない**ものとして確定させる。")
    return "\n".join(L) + "\n"


def main(argv=None) -> int:
    sys.stdout.reconfigure(encoding="utf-8")
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--db", default=C.DB_PATH)
    p.add_argument("--entries", default=os.path.join(C.DATA_DIR, "s5c_trades.csv"))
    p.add_argument("--out", default=os.path.join(C.DATA_DIR, "s5c_event_exit_report.md"))
    p.add_argument("--csv-out", default=os.path.join(C.DATA_DIR, "s5c_event_exit.csv"))
    a = p.parse_args(argv)

    con = sqlite3.connect("file:%s?mode=ro" % a.db.replace("\\", "/"), uri=True)
    entries = load_entries(a.entries)
    codes = {e["code"] for e in entries}
    C.log("シャドウL: 入口 %d（銘柄 %d）" % (len(entries), len(codes)))
    cal = [r[0] for r in con.execute("SELECT date FROM market_index ORDER BY date")]
    evs, _sk = W.revision_events(con)
    px = price_panel(con, codes)
    idx = K.index_map(con)

    results, all_rows = {}, []
    for mode in ("main", "up_only"):
        for n in N_DAYS:
            rows, why = simulate(entries, evs, px, idx, cal, n, up_only=(mode == "up_only"))
            results[(mode, n)] = {"rows": rows, "why": why,
                                  "sum": K.summarise("N=%d%s" % (n, "（up のみ）" if mode == "up_only" else ""), rows)}
            C.log("  %-8s N=%-2d 建玉 %d / 修正で降りた %d / クラスタ平均 %s"
                  % (mode, n, len(rows), why["修正で降りた"],
                     "–" if not results[(mode, n)]["sum"]
                     else "%+.2f%%" % (100 * results[(mode, n)]["sum"]["cl_mean"])))
            if mode == "main":
                for r in rows:
                    all_rows.append(dict(r, n_days=n))
    with open(a.csv_out, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.DictWriter(fh, fieldnames=list(all_rows[0].keys()))
        w.writeheader()
        w.writerows(all_rows)
    text = render(results, len(entries))
    with open(a.out, "w", encoding="utf-8") as fh:
        fh.write(text)
    print(text)
    C.log("出力: %s / %s" % (a.out, a.csv_out))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
