"""screener/report/entry_sim.py — シャドウF: 受容帯とテーゼ破綻を**入口側**で使う。

定義の正本は docs/backtest_acceptance_criteria.md「シャドウF 事前登録」（2026-09-20・測定前に確定）。
**このスクリプトは事前登録を実装するだけで、パラメータを決めない。**

    v2   v2 そのもの（比較の基準）
    F1a  エントリー日の終値が受容帯下限(VAL)未満なら建てない
    F1b  VAL 未満は NAV 5%（通常の半分）で建てる
    F1c  VAL 以上 VAH 以下のときだけ建てる
    F2a  T−15 時点で直近に開示済みの四半期がテーゼ破綻なら建てない
    F2b  直近2四半期が連続で破綻なら建てない

出口はどの変種も v2 と同じ固定 T+2（E の測定で 4分岐は悪化と出たため、入口の効果だけを見る）。
**判定は「F* − v2 の期待値差の符号が ±3シフト3本で一致するか」のみ。** 合否は出さない。

    python -m screener.report.entry_sim --quick
"""
from __future__ import annotations

import argparse
import csv
import os
import sqlite3
import sys
from collections import Counter

try:
    from screener import common as C
except ImportError:                                     # pragma: no cover
    sys.path.insert(0, os.path.dirname(os.path.dirname(
        os.path.dirname(os.path.abspath(__file__)))))
    from screener import common as C

from screener.report import backtest_eval as V1
from screener.report import backtest_v2 as V2
from screener.report import exit_sim as X
from screener.signals import sales_direction as SD
from screener.signals import span_scorers as SS
from screener.technical import acceptance_zone as AZ

HALF_FRACTION = 0.05            # F1b のサイジング（v2 の半分）。動かさない
VARIANTS = ("v2", "F1a", "F1b", "F1c", "F2a", "F2b")


# ------------------------------------------------------------------ 入口の材料
def zone_at_entry(mcon, pcon, ticker, entry_date, cache):
    """(受容帯, エントリー日の終値)。60日ぶんの OHLCV が無ければ受容帯は None。"""
    key = (ticker, entry_date)
    if key in cache:
        return cache[key]
    bars = X.zone_bars(mcon, ticker, entry_date)
    zone = AZ.value_area(bars) if len(bars) >= AZ.N_DAYS else None
    row = pcon.execute("SELECT close FROM daily_prices WHERE ticker=? AND date=?",
                       (ticker, entry_date)).fetchone()
    cache[key] = (zone, row[0] if row else None)
    return cache[key]


def thesis_streak(pcon, ticker, as_of, n=2, cache={}):
    """as_of 時点で公知の**直近 n 四半期**がそれぞれテーゼ破綻か（新しい順）。

    判定は出口の分岐2 と同じ（売上 or 営業利益が前年同期比 ≤ 0。OR）。
    取れない期は None（判定不能）。
    """
    key = (ticker, as_of, n)
    if key in cache:
        return cache[key]
    vis = SS.visible_periods(pcon, ticker, as_of)
    series = {}
    for item, label in (("sales", "sales"), ("operating_profit", "op")):
        ts = SD.tiles(pcon, ticker, vis, item=item)
        series[label] = ts
    out = []
    for back in range(n):
        vals = []
        for label in ("sales", "op"):
            ts = series[label]
            i = len(ts) - 1 - back
            if i < 0:
                continue
            v = X._yoy_pct(ts, i)
            if v is not None:
                vals.append(v)
        out.append(None if not vals else any(v <= 0 for v in vals))
    cache[key] = out
    return out


def admitters(mcon, pcon):
    """変種ごとの admit 関数。戻り値は NAV 比率（建てないなら None）。"""
    zcache, log = {}, []

    def note(variant, t, reason, frac):
        log.append({"variant": variant, "ticker": t["ticker"],
                    "entry_date": t["entry_date"], "event_date": t["event_date"],
                    "reason": reason, "fraction": frac or 0.0,
                    "ret_if_taken": float(t["ret"]) - V1.COST_ROUND_TRIP})

    def zone_pos(t):
        zone, close = zone_at_entry(mcon, pcon, t["ticker"], t["entry_date"], zcache)
        if zone is None or close is None:
            return "帯なし", None
        if close < zone["val"]:
            return "帯の下", zone
        if close > zone["vah"]:
            return "帯の上", zone
        return "帯の中", zone

    def f1a(t):
        pos, _z = zone_pos(t)
        frac = None if pos == "帯の下" else V2.POS_FRACTION
        note("F1a", t, pos, frac)
        return frac

    def f1b(t):
        pos, _z = zone_pos(t)
        frac = HALF_FRACTION if pos == "帯の下" else V2.POS_FRACTION
        note("F1b", t, pos, frac)
        return frac

    def f1c(t):
        pos, _z = zone_pos(t)
        frac = V2.POS_FRACTION if pos in ("帯の中", "帯なし") else None
        note("F1c", t, pos, frac)
        return frac

    def f2(variant, t, need):
        st = thesis_streak(pcon, t["ticker"], t["entry_date"], n=need)
        if st[0] is None:
            reason, frac = "判定不能", V2.POS_FRACTION      # 判定不能は建てる（F-2）
        elif need == 1:
            broken = bool(st[0])
            reason, frac = ("破綻", None) if broken else ("健全", V2.POS_FRACTION)
        else:
            broken = bool(st[0]) and bool(st[1])
            reason = "2期連続破綻" if broken else ("直近のみ破綻" if st[0] else "健全")
            frac = None if broken else V2.POS_FRACTION
        note(variant, t, reason, frac)
        return frac

    return {
        "v2": None,
        "F1a": f1a, "F1b": f1b, "F1c": f1c,
        "F2a": lambda t: f2("F2a", t, 1),
        "F2b": lambda t: f2("F2b", t, 2),
    }, log


# ------------------------------------------------------------------ 実行
def run_shift(shift, out_dir):
    C.log("=== シャドウF 入口シミュレーション shift=%+d ===" % shift)
    _res, trades = V1.run_one(shift)
    pcon = sqlite3.connect("file:%s?mode=ro" % V1._db_path(shift).replace("\\", "/"), uri=True)
    mcon = sqlite3.connect("file:%s?mode=ro" % C.DB_PATH.replace("\\", "/"), uri=True)
    adm, log = admitters(mcon, pcon)
    out = {}
    for v in VARIANTS:
        sim = V2.simulate(trades, pcon, V1.MAIN_ENTRY, V1.MAIN_EXIT, admit=adm[v])
        m = V2.metrics(sim, pcon)
        m["skipped_admit"] = sim["skipped_admit"]
        m["skipped_full"] = sim["skipped_full"]
        out[v] = m
        C.log("  %-4s 建玉 %4d（入口で見送り %4d / 枠満杯 %4d）期待値 %+.4f"
              % (v, m["n_trades"], sim["skipped_admit"], sim["skipped_full"],
                 m["expectancy_net"]))
    path = os.path.join(out_dir, "entry_sim_decisions_shift%+d.csv" % shift)
    with open(path, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.DictWriter(fh, fieldnames=list(log[0].keys()))
        w.writeheader()
        w.writerows(log)
    return {"shift": shift, "m": out, "log": log, "csv": path}


def render(store):
    L = ["# シャドウF: 受容帯とテーゼ破綻を入口で使う（記述統計・採用しない）", ""]
    for s, st in store.items():
        L.append("## shift %+d" % s)
        L.append("")
        L.append("| 変種 | 建玉 | 勝率 | 期待値(純) | TOPIX超過 | MDD | 上位5集中 | 入口で見送り | v2との差 |")
        L.append("|---|---|---|---|---|---|---|---|---|")
        base = st["m"]["v2"]["expectancy_net"]
        for v in VARIANTS:
            m = st["m"][v]
            L.append("| %s | %d | %.1f%% | %+.4f | %s | %+.4f | %s | %d | %s |"
                     % (v, m["n_trades"], 100 * m["win_rate"], m["expectancy_net"],
                        "%+.4f" % m["excess_vs_topix"] if m["excess_vs_topix"] is not None else "n/a",
                        m["max_drawdown"],
                        "%.1f%%" % (100 * m["top5_profit_share"]) if m["top5_profit_share"] else "n/a",
                        m["skipped_admit"],
                        "–" if v == "v2" else "%+.4f" % (m["expectancy_net"] - base)))
        L.append("")
    L.append("## 判定（F-0: F* − v2 の符号が3本で一致するか）")
    L.append("")
    L.append("| 変種 | " + " | ".join("shift %+d" % s for s in store) + " | 符号一致 |")
    L.append("|---|" + "---|" * (len(store) + 1))
    for v in VARIANTS[1:]:
        ds = [store[s]["m"][v]["expectancy_net"] - store[s]["m"]["v2"]["expectancy_net"]
              for s in store]
        agree = "–（3本そろっていない）" if len(ds) < 3 else (
            "はい（%s）" % ("プラス" if all(d > 0 for d in ds) else "マイナス")
            if len({d > 0 for d in ds}) == 1 else "いいえ（符号不定）")
        L.append("| %s | %s | %s |" % (v, " | ".join("%+.4f" % d for d in ds), agree))
    L.append("")
    L.append("## 入口で見送った候補が、建てていたらどうだったか（F-3）")
    L.append("")
    L.append("| 変種 | 理由 | 見送り件数 | 建てた場合の期待値(純) |")
    L.append("|---|---|---|---|")
    for s, st in store.items():
        if s != 0:
            continue
        by = {}
        for r in st["log"]:
            if r["fraction"]:
                continue
            by.setdefault((r["variant"], r["reason"]), []).append(r["ret_if_taken"])
        for (v, reason), vals in sorted(by.items()):
            L.append("| %s | %s | %d | %+.4f |"
                     % (v, reason, len(vals), sum(vals) / len(vals)))
    L.append("")
    L.append("（shift 0 のみ。見送った候補は枠を次の候補に回すので、"
             "**この期待値がそのまま全体の差にはならない**）")
    return "\n".join(L) + "\n"


def main(argv=None) -> int:
    sys.stdout.reconfigure(encoding="utf-8")
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--quick", action="store_true", help="shift 0 のみ")
    p.add_argument("--out-dir", default=C.DATA_DIR)
    a = p.parse_args(argv)
    shifts = (0,) if a.quick else V1.SHIFTS
    store = {s: run_shift(s, a.out_dir) for s in shifts}
    text = render(store)
    print(text)
    out = os.path.join(a.out_dir, "entry_sim_report.md")
    with open(out, "w", encoding="utf-8") as fh:
        fh.write(text)
    C.log("出力: %s" % out)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
