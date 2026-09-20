"""screener/report/backtest_v2.py — v2 事前登録（ポジション枠 + 市場フィルター）。

基準と設計の正本は docs/backtest_acceptance_criteria.md の「v2 事前登録」。
**このスクリプトは事前登録を実装するだけで、基準を決めない。**
これが最後のインサンプル判定であり、結果を見てからの再チューニングは禁止。

v1 との違いは2つだけ
--------------------
1. ポジション枠（測定装置の修正）
   v1 は候補トレードを全部「NAV10%固定」で資産曲線に流していた。同日
   エントリーが最大117件あるので、実際には建てられない建玉を前提に
   MDD を出していた。v2 は**同時保有10枠のハード制約**を課し、枠が
   埋まっていれば新規シグナルをスキップする。同日の候補が空き枠を
   超える場合は score 降順（同点は ticker 昇順）で決定的に採用する。

2. 市場フィルター（設計の修正）
   TOPIX が 200日移動平均を下回る間は新規エントリーを停止する。
   移動平均は entry_date の**前日まで**の200本で計算する（当日終値は
   寄り付き時点で未知なので使わない）。200本に満たない期間は
   「上回っていることを確認できない」ものとして建てない。

別枠のコードは1行も変更しない。候補トレードの生成は別枠に任せ、
建玉の可否と資産曲線だけをこちらで組み直す。

    python -m screener.report.backtest_v2
    python -m screener.report.backtest_v2 --quick
"""
from __future__ import annotations

import argparse
import os
import sqlite3
import sys
from collections import defaultdict

try:
    from screener import common as C
except ImportError:                                     # pragma: no cover
    sys.path.insert(0, os.path.dirname(os.path.dirname(
        os.path.dirname(os.path.abspath(__file__)))))
    from screener import common as C

from screener.report import backtest_eval as V1

# ---- v2 事前登録の値。実行後に動かさない ---------------------------------
MAX_POSITIONS = 10
POS_FRACTION = 0.10          # 1建玉あたり NAV の10%
MA_WINDOW = 200              # TOPIX 200日移動平均


def topix_ma_ok(index_dates, index_close, entry_date):
    """entry_date の前日までの200本で移動平均を出し、直近終値が上回るか。

    200本に満たなければ False（上回っていることを確認できない）。
    """
    lo, hi = 0, len(index_dates) - 1
    while lo <= hi:                                      # bisect_left
        mid = (lo + hi) // 2
        if index_dates[mid] < entry_date:
            lo = mid + 1
        else:
            hi = mid - 1
    prior = lo                                           # entry_date 未満の本数
    if prior < MA_WINDOW:
        return False, None
    window = index_close[prior - MA_WINDOW:prior]
    ma = sum(window) / MA_WINDOW
    last = index_close[prior - 1]
    return last > ma, ma


def simulate(trades, con, entry_n, exit_k, use_filter=True, admit=None):
    """枠制約と市場フィルターを課して、実行可能な建玉だけで資産曲線を作る。

    `admit` は入口側の実験（シャドウF）のための任意フック。候補1件を受け取り
    **建てるなら NAV に対する比率、建てないなら None/0** を返す。既定の None は
    v2 そのもの（全件 POS_FRACTION）で、**v2 の数値は1つも変わらない**。
    """
    rows = [t for t in trades if t["entry_n"] == entry_n and t["exit_k"] == exit_k]
    idx = con.execute("SELECT date, close FROM market_index ORDER BY date").fetchall()
    idx_d = [r[0] for r in idx]
    idx_c = [r[1] for r in idx]

    # 候補に exit 日を与える（別枠は exit 日を持たない）
    cand = []
    for t in rows:
        t = dict(t)
        t["exit_date"] = V1._exit_date(con, t)
        t["net"] = t["ret"] - V1.COST_ROUND_TRIP
        cand.append(t)
    # 同日内は score 降順、同点は ticker 昇順で決定的に
    cand.sort(key=lambda t: (t["entry_date"], -float(t["score"]), t["ticker"]))

    by_day = defaultdict(list)
    for t in cand:
        by_day[t["entry_date"]].append(t)

    nav = 1.0
    peak, mdd = 1.0, 0.0
    open_pos = []                # (exit_date, alloc, net, trade)
    taken, skipped_full, skipped_filter, skipped_admit = [], 0, 0, 0
    filter_off_days = set()
    nav_curve = []

    for day in sorted(by_day):
        # まず当日までに満期を迎えた建玉を決済する
        still = []
        for ex, alloc, net, tr in open_pos:
            if ex <= day:
                nav += alloc * net
                peak = max(peak, nav)
                mdd = min(mdd, nav / peak - 1)
                nav_curve.append((ex, nav))
            else:
                still.append((ex, alloc, net, tr))
        open_pos = still

        ok, _ma = topix_ma_ok(idx_d, idx_c, day) if use_filter else (True, None)
        if not ok:
            filter_off_days.add(day)
            skipped_filter += len(by_day[day])
            continue

        free = MAX_POSITIONS - len(open_pos)
        for t in by_day[day]:
            if free <= 0:
                skipped_full += 1
                continue
            frac = POS_FRACTION if admit is None else admit(t)
            if not frac:
                skipped_admit += 1
                continue
            alloc = nav * frac
            open_pos.append((t["exit_date"], alloc, t["net"], t))
            taken.append(t)
            free -= 1

    # exit 日だけで並べる。タプル全体で比較すると同着時に dict 同士の
    # 比較に落ちて TypeError になる。
    for ex, alloc, net, tr in sorted(open_pos, key=lambda x: x[0]):
        nav += alloc * net
        peak = max(peak, nav)
        mdd = min(mdd, nav / peak - 1)
        nav_curve.append((ex, nav))

    return {
        "trades": taken, "n_candidates": len(cand),
        "n_taken": len(taken), "skipped_full": skipped_full,
        "skipped_filter": skipped_filter, "skipped_admit": skipped_admit,
        "filter_off_days": sorted(filter_off_days),
        "nav_final": nav, "max_drawdown": mdd, "nav_curve": nav_curve,
    }


def metrics(sim, con):
    """事前登録 §2 の5項目。リターンは往復コスト控除後。"""
    taken = sim["trades"]
    if not taken:
        return None
    net = [t["net"] for t in taken]
    wins = [r for r in net if r > 0]
    profits = sorted((r for r in net if r > 0), reverse=True)
    total_profit = sum(profits)
    exc = []
    for t in taken:
        ir = V1.index_return(con, t["entry_date"], t["exit_date"])
        if ir is not None:
            exc.append(t["net"] - ir)
    return {
        "n_trades": len(taken),
        "win_rate": len(wins) / len(taken),
        "expectancy_net": sum(net) / len(net),
        "expectancy_gross": sum(t["ret"] for t in taken) / len(taken),
        "excess_vs_topix": (sum(exc) / len(exc)) if exc else None,
        "max_drawdown": sim["max_drawdown"],
        "top5_profit_share": (sum(profits[:5]) / total_profit) if total_profit > 0 else None,
        "nav_final": sim["nav_final"],
    }


def verdict(m):
    checks = [
        ("トレード数 >= %d" % V1.MIN_TRADES, m["n_trades"] >= V1.MIN_TRADES, m["n_trades"]),
        ("期待値 > 0（コスト控除後）", m["expectancy_net"] > 0, m["expectancy_net"]),
        ("TOPIX超過 > 0", m["excess_vs_topix"] is not None and m["excess_vs_topix"] > 0,
         m["excess_vs_topix"]),
        ("最大DD >= -15%", m["max_drawdown"] >= V1.MAX_DD, m["max_drawdown"]),
        ("上位5トレード利益集中 < 30%",
         m["top5_profit_share"] is not None
         and m["top5_profit_share"] < V1.MAX_TOP5_PROFIT_SHARE, m["top5_profit_share"]),
    ]
    return checks, all(ok for _, ok, _ in checks)


def main(argv=None):
    sys.stdout.reconfigure(encoding="utf-8")
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--quick", action="store_true", help="shift=0 のみ")
    a = p.parse_args(argv)

    shifts = (0,) if a.quick else V1.SHIFTS
    store = {}
    for s in shifts:
        C.log("=== v2 バックテスト shift=%+d ===" % s)
        res, trades = V1.run_one(s)
        con = sqlite3.connect("file:%s?mode=ro"
                              % V1._db_path(s).replace("\\", "/"), uri=True)
        sim = simulate(trades, con, V1.MAIN_ENTRY, V1.MAIN_EXIT)
        store[s] = {"sim": sim, "con": con, "trades": trades, "res": res,
                    "m": metrics(sim, con)}
        C.log("  候補 %s / 建玉 %s / 枠満杯でスキップ %s / フィルターでスキップ %s"
              % (format(sim["n_candidates"], ","), format(sim["n_taken"], ","),
                 format(sim["skipped_full"], ","), format(sim["skipped_filter"], ",")))

    base = store[0]
    m = base["m"]
    C.log("")
    C.log("=== v2 主セル T-15 / T+2 ===")
    checks, passed = verdict(m)
    for label, ok, val in checks:
        shown = "%.4f" % val if isinstance(val, float) else str(val)
        C.log("  [%s] %-34s 実測 %s" % ("合" if ok else "否", label, shown))
    if m["n_trades"] < V1.MIN_TRADES:
        C.log("  → 総合判定: 判定不能（トレード数がサンプル要件を割った）")
    else:
        C.log("  → 総合判定: %s" % ("合格" if passed else "不合格"))
    C.log("  勝率 %.1f%% / 最終NAV %.3f" % (m["win_rate"] * 100, m["nav_final"]))

    C.log("")
    C.log("=== 時価総額帯別（P3-4・枠制約下で建てた分・記述統計）===")
    V1.band_report(base["sim"]["trades"], base["con"], V1.MAIN_ENTRY, V1.MAIN_EXIT)

    C.log("")
    C.log("=== 市場フィルターの作動 ===")
    days = base["sim"]["filter_off_days"]
    C.log("  エントリー停止となった日: %d 日 / 停止でスキップしたイベント %s 件"
          % (len(days), format(base["sim"]["skipped_filter"], ",")))
    if days:
        runs, start, prev = [], days[0], days[0]
        for d in days[1:]:
            if (int(d[:4]), int(d[5:7])) != (int(prev[:4]), int(prev[5:7])) \
                    and d[:7] != prev[:7]:
                runs.append((start, prev)); start = d
            prev = d
        runs.append((start, prev))
        for s0, s1 in runs[:12]:
            C.log("    %s .. %s" % (s0, s1))

    C.log("")
    C.log("=== エントリー感応度（v2の枠制約下）===")
    for n in (15, 10, 5):
        sim = simulate(base["trades"], base["con"], n, V1.MAIN_EXIT)
        mm = metrics(sim, base["con"])
        if mm:
            C.log("  T-%-3d 建玉 %5s  勝率 %5.1f%%  期待値(純) %+.4f  MDD %.4f"
                  % (n, format(mm["n_trades"], ","), mm["win_rate"] * 100,
                     mm["expectancy_net"], mm["max_drawdown"]))

    if not a.quick:
        C.log("")
        C.log("=== 発表日 ±3営業日シフト（v2）===")
        wr = []
        for s in V1.SHIFTS:
            mm = store[s]["m"]
            wr.append(mm["win_rate"] * 100)
            C.log("  %+d  建玉 %5s  勝率 %5.1f%%  期待値(純) %+.4f  MDD %.4f"
                  % (s, format(mm["n_trades"], ","), mm["win_rate"] * 100,
                     mm["expectancy_net"], mm["max_drawdown"]))
        allpos = all(store[s]["m"]["expectancy_net"] > 0 for s in V1.SHIFTS)
        C.log("  3本すべて期待値プラス: %s / 勝率変動 %.1f pt"
              % ("はい" if allpos else "いいえ", max(wr) - min(wr)))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
