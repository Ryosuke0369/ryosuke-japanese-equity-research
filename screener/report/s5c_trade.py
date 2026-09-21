"""screener/report/s5c_trade.py — シャドウK: R ≥ 1.30 は「取れる」のか。

定義の正本は docs/backtest_acceptance_criteria.md「シャドウK」（2026-09-21・測定前に確定）。
**このスクリプトは定義を実装するだけで、建て方・出口・判定基準を決めない。**

  判定日の終値で建て、**次の判定日の終値**で出る（固定・約1か月）。往復コスト 0.4%。
  主判定は **判定日クラスタ補正後の平均リターンと t**（K-3 / K-4）。

**S5c は 0点・表示のみ。** この測定の結果でスコアに接続するかはユーザーが判断する。

    python -m screener.report.s5c_trade
"""
from __future__ import annotations

import argparse
import csv
import math
import os
import sqlite3
import statistics
import sys
from collections import defaultdict
from datetime import date, timedelta

try:
    from screener import common as C
except ImportError:                                     # pragma: no cover
    sys.path.insert(0, os.path.dirname(os.path.dirname(
        os.path.dirname(os.path.abspath(__file__)))))
    from screener import common as C

from screener.report import s5c_power as W
from screener.signals import progress_history as P

COST = 0.004                 # 往復（K-1）
FIRE = P.FIRE_R              # 1.30
BANDS = ((1.3, 1.5, "1.3-1.5"), (1.5, 2.0, "1.5-2.0"), (2.0, None, "2.0+"))
MIN_N = 100


def price_map(con, codes=None):
    """code -> {date: 調整後終値}。"""
    px = defaultdict(dict)
    for code, d, c in con.execute(
            "SELECT code, date, adj_close FROM prices WHERE adj_close IS NOT NULL"):
        if codes is None or code in codes:
            px[code][d] = c
    return px


def index_map(con):
    return dict(con.execute("SELECT date, close FROM market_index"))


def build_trades(con, dates, codes, evs, px, idx):
    """判定日ごとに R を計算し、次の判定日までのリターンを付ける。"""
    rows, miss_price, no_r = [], 0, 0
    for i, as_of in enumerate(dates[:-1]):
        nxt = dates[i + 1]
        horizon = min((date.fromisoformat(as_of) + timedelta(days=W.WINDOW_DAYS)).isoformat(),
                      W.LAST_DATA)
        for code in codes:
            r = P.evaluate(con, code, as_of)
            if not r["available"]:
                no_r += 1
                continue
            p0 = px.get(code, {}).get(as_of)
            p1 = px.get(code, {}).get(nxt)
            if not p0 or not p1:
                miss_price += 1
                continue
            gross = p1 / p0 - 1
            i0, i1 = idx.get(as_of), idx.get(nxt)
            bench = (i1 / i0 - 1) if (i0 and i1) else None
            # 保有期間内に up の修正が出たか（**事後の切り分け**・K-2 の5）
            up_in_hold = "none"
            for d, direction, _mag in evs.get(code, ()):
                if as_of < d <= nxt:
                    up_in_hold = direction
                    break
            rows.append({
                "as_of": as_of, "exit_date": nxt, "code": code, "r": r["r"],
                "period": r["period"], "median_progress": r["median_progress"],
                "fired": int(r["fired"]), "ret_gross": gross, "ret_net": gross - COST,
                "bench": bench,
                "excess": None if bench is None else (gross - COST) - bench,
                "rev_in_hold": up_in_hold,
                # 90日窓の結果（§52 と同じ定義。参考として残す）
                "rev_in_90d": W.outcome(evs, code, as_of, horizon),
            })
    return rows, miss_price, no_r


# ------------------------------------------------------------------ 集計
def _t(xs):
    if len(xs) < 2:
        return None
    sd = statistics.stdev(xs)
    return statistics.mean(xs) / (sd / math.sqrt(len(xs))) if sd > 0 else None


def clustered(rows, key="ret_net"):
    """判定日ごとの平均を1点とし、その平均と t（K-3 の主判定）。"""
    by = defaultdict(list)
    for r in rows:
        if r[key] is not None:
            by[r["as_of"]].append(r[key])
    ms = [statistics.mean(v) for v in by.values()]
    if not ms:
        return None, None, 0
    return statistics.mean(ms), _t(ms), len(ms)


def mdd(rows):
    """等金額・月次の資産曲線（各判定日の平均リターンを複利でつなぐ）。"""
    by = defaultdict(list)
    for r in rows:
        by[r["as_of"]].append(r["ret_net"])
    nav = peak = 1.0
    worst = 0.0
    for d in sorted(by):
        nav *= 1 + statistics.mean(by[d])
        peak = max(peak, nav)
        worst = min(worst, nav / peak - 1)
    return worst, nav


def summarise(label, rows):
    if not rows:
        return None
    net = [r["ret_net"] for r in rows]
    exc = [r["excess"] for r in rows if r["excess"] is not None]
    m, t, nd = clustered(rows)
    dd, nav = mdd(rows)
    return {"label": label, "n": len(rows), "codes": len({r["code"] for r in rows}),
            "mean": statistics.mean(net), "median": statistics.median(net),
            "win": sum(x > 0 for x in net) / len(net),
            "excess": statistics.mean(exc) if exc else None,
            "cl_mean": m, "cl_t": t, "n_dates": nd, "mdd": dd, "nav": nav}


def _line(s):
    if not s:
        return "| – | 0 | | | | | | | |"
    note = " ※件数不足" if s["n"] < MIN_N else ""
    return ("| %s | %s%s | %d | %+.2f%% | %+.2f%% | %.1f%% | %s | %s | %+.2f%% | %.3f |"
            % (s["label"], f"{s['n']:,}", note, s["codes"], 100 * s["mean"],
               100 * s["median"], 100 * s["win"],
               "%+.2f%%" % (100 * s["excess"]) if s["excess"] is not None else "–",
               "%+.2f%% / t %s" % (100 * s["cl_mean"],
                                   "–" if s["cl_t"] is None else "%.2f" % s["cl_t"]),
               100 * s["mdd"], s["nav"]))


HDR = ("| 群 | n | 銘柄 | 平均(純) | 中央値 | 勝率 | TOPIX超過 | 判定日クラスタ | MDD | 最終NAV |\n"
       "|---|---|---|---|---|---|---|---|---|---|")


def render(rows, dates, miss_price, no_r, cover):
    L = ["# シャドウK: R ≥ 1.30 のトレード成績（記録専用・0点）", ""]
    L.append("- 判定日 %d（%s 〜 %s）/ 建玉 %s（判定日の終値で建て、次の判定日の終値で出る）"
             % (len(dates) - 1, dates[0], dates[-2], f"{len(rows):,}"))
    L.append("- 往復コスト %.1f%% 控除後。**出口は固定（次の判定日）**" % (100 * COST))
    L.append("- 価格が無くて落とした観測 %s / R が計算できない観測 %s / **株価のカバー率 %.1f%%**"
             % (f"{miss_price:,}", f"{no_r:,}", 100 * cover))
    L.append("")
    L.append("## K-2 群別（主判定は「判定日クラスタ」列）")
    L.append("")
    L.append(HDR)
    fired = [r for r in rows if r["fired"]]
    L.append(_line(summarise("**R ≥ 1.30**", fired)))
    L.append(_line(summarise("R < 1.30", [r for r in rows if not r["fired"]])))
    L.append(_line(summarise("母集団（全体）", rows)))
    L.append("")
    L.append("## R 帯別")
    L.append("")
    L.append(HDR)
    for lo, hi, label in BANDS:
        sub = [r for r in rows if r["r"] >= lo and (hi is None or r["r"] < hi)]
        L.append(_line(summarise(label, sub)))
    L.append("")
    L.append("## 事後の切り分け（K-2 の5・**エントリー時点では分からない**）")
    L.append("")
    L.append(HDR)
    up = [r for r in fired if r["rev_in_hold"] == "up"]
    dn = [r for r in fired if r["rev_in_hold"] == "down"]
    non = [r for r in fired if r["rev_in_hold"] == "none"]
    L.append(_line(summarise("R≥1.30 かつ 保有中に上方修正", up)))
    L.append(_line(summarise("R≥1.30 かつ 保有中に下方修正", dn)))
    L.append(_line(summarise("R≥1.30 かつ 修正なし", non)))
    L.append("")
    if fired:
        L.append("保有1か月のうちに上方修正が出た割合: **%.1f%%**（%d / %d）"
                 % (100 * len(up) / len(fired), len(up), len(fired)))
    L.append("")
    L.append("## 月ごとの建玉数（R ≥ 1.30）")
    L.append("")
    by = defaultdict(int)
    for r in fired:
        by[r["as_of"]] += 1
    counts = sorted(by.values())
    if counts:
        L.append("中央値 %d / 最小 %d / 最大 %d（判定日 %d）"
                 % (statistics.median(counts), counts[0], counts[-1], len(by)))
        L.append("")
        L.append("| 判定日 | 建玉 | 平均(純) |")
        L.append("|---|---|---|")
        for d in sorted(by):
            sub = [r for r in fired if r["as_of"] == d]
            L.append("| %s | %d | %+.2f%% |"
                     % (d, len(sub), 100 * statistics.mean([x["ret_net"] for x in sub])))
    L.append("")
    L.append("## K-4 の判定")
    L.append("")
    s = summarise("R ≥ 1.30", fired)
    if s and s["cl_t"] is not None:
        ok = s["cl_mean"] > 0 and s["cl_t"] >= 2
        L.append("R ≥ 1.30 の判定日クラスタ平均 **%+.2f%% / t %.2f** → **%s**"
                 % (100 * s["cl_mean"], s["cl_t"],
                    "取れている（登録した基準を満たす）" if ok
                    else "**予測力はあるが取れない**（基準を満たさない）"))
    L.append("")
    L.append("**この結果でスコアには接続していない。** 接続の判断はユーザーが行う。")
    return "\n".join(L) + "\n"


def main(argv=None) -> int:
    sys.stdout.reconfigure(encoding="utf-8")
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--db", default=C.DB_PATH)
    p.add_argument("--out", default=os.path.join(C.DATA_DIR, "s5c_trade_report.md"))
    p.add_argument("--csv-out", default=os.path.join(C.DATA_DIR, "s5c_trades.csv"))
    a = p.parse_args(argv)

    con = sqlite3.connect("file:%s?mode=ro" % a.db.replace("\\", "/"), uri=True)
    cal = [r[0] for r in con.execute("SELECT date FROM market_index ORDER BY date")]
    dates = W.judgment_dates(cal)
    # 出口を持たせるため、最後の判定日の次の月初営業日を1つ足す
    after = [d for d in cal if d > dates[-1]]
    if after:
        nxt_month = next((d for d in after if d[:7] != dates[-1][:7]), None)
        if nxt_month:
            dates = dates + [nxt_month]
    codes = [r[0] for r in con.execute(
        "SELECT code FROM companies WHERE universe_flag=1 ORDER BY code")]
    with_px = {r[0] for r in con.execute(
        "SELECT DISTINCT code FROM prices WHERE adj_close IS NOT NULL")}
    cover = len([c for c in codes if c in with_px]) / len(codes)
    C.log("シャドウK: 判定日 %d × 銘柄 %d（株価あり %.1f%%）"
          % (len(dates) - 1, len(codes), 100 * cover))
    evs, _sk = W.revision_events(con)
    px = price_map(con, set(codes))
    rows, miss_price, no_r = build_trades(con, dates, codes, evs, px, index_map(con))
    C.log("建玉 %d" % len(rows))
    with open(a.csv_out, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.DictWriter(fh, fieldnames=list(rows[0].keys()))
        w.writeheader()
        w.writerows(rows)
    text = render(rows, dates, miss_price, no_r, cover)
    with open(a.out, "w", encoding="utf-8") as fh:
        fh.write(text)
    print(text)
    C.log("出力: %s / %s" % (a.out, a.csv_out))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
