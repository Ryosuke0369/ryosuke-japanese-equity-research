"""screener/report/exit_sim.py — シャドウE フェーズB: 出口4分岐のインサンプル符号測定。

定義の正本は docs/backtest_acceptance_criteria.md シャドウE B-1〜B-3（2026-09-19 提案、
2026-09-20 ユーザー承認）。**このスクリプトは事前登録を実装するだけで、基準を決めない。**
結果を見て分岐の条件・受容帯の N/B/X/M を動かすことは禁止。

測るもの（v2 と同じ建玉集合に対して、出口だけを差し替える）
------------------------------------------------------------
  E0  v2 そのまま（固定 T+2）
  E1  E0 + 分岐1（構造破綻: 受容帯下限を終値で M 日連続割れ）
  E2  E0 + 分岐2（テーゼ破綻: 開示四半期の売上 or 営業利益が前年割れ → T+1）
  E3  E0 + 分岐3（初動利確: 業績/配当予想の修正が同時開示 → T+1）
  E4  分岐1〜4 の全部
  参考 ドリフト保有（T+20）は**記録のみ**。採否には使わない

**建玉の選択は v2 のまま動かさない。** 出口が変われば枠の空き方も変わるが、
選択まで変えると「同じ建玉集合の比較」でなくなる（ユーザー指示・B-2b）。
資産曲線は出口日を差し替えて組み直す。

価格の扱い
----------
リターンは v2 と同じ `projection.db.daily_prices.close`（= 本体の adj_close）で計算する。
受容帯の高値・安値は本体 `prices` の生値に、同じ行の `adj_close/close` を掛けて
調整後に揃える（分割日に偽の帯ができないように）。

    python -m screener.report.exit_sim --quick     # shift 0 のみ
    python -m screener.report.exit_sim
"""
from __future__ import annotations

import argparse
import csv
import os
import sqlite3
import sys
from collections import Counter, defaultdict

try:
    from screener import common as C
except ImportError:                                     # pragma: no cover
    sys.path.insert(0, os.path.dirname(os.path.dirname(
        os.path.dirname(os.path.abspath(__file__)))))
    from screener import common as C

from screener.report import backtest_eval as V1
from screener.report import backtest_v2 as V2
from screener.signals import sales_direction as SD
from screener.signals import span_scorers as SS
from screener.technical import acceptance_zone as AZ

DRIFT_K = 20                       # 記録専用アームの保有営業日（T+20）
THESIS_EXIT_K = 1                  # テーゼ破綻・初動利確の決済は T+1 の終値（B-1）


def _cal():
    """別枠と同じ営業日カレンダー（祝日込み）。出口日は別枠と同じ数え方で出す。"""
    root = str(V1._external_root())
    if root not in sys.path:
        sys.path.insert(0, root)
    from common.jp_calendar import add_business_days
    return add_business_days


def bd_after(event_date, k):
    """event_date の k 営業日後（別枠 backtest.py と同じ add_business_days）。"""
    from datetime import date
    return _cal()(date.fromisoformat(event_date), k).isoformat()


def idx_at(path, d):
    """path 上で日付 d 以上の最初の位置。無ければ None。"""
    for i, (dd, _c) in enumerate(path):
        if dd >= d:
            return i
    return None


VARIANTS = ("E0", "E1", "E2", "E3", "E4")       # B-3 の判定対象
# 記録専用（B-1 の「ドリフト保有は記録アーム」）。判定には使わない。
#   drift : 全件 T+20 まで保有
#   E5    : テーゼ破綻なら T+1、健全なら T+20 まで保有（分岐2で条件付けた drift）
RECORD_VARIANTS = ("drift", "E5")


# ------------------------------------------------------------------ 価格
def zone_bars(mcon, ticker, entry_date, n=AZ.N_DAYS):
    """エントリー日の**前営業日まで** n 日ぶんの調整後 OHLC + 売買代金。"""
    rows = mcon.execute(
        "SELECT date, close, adj_close, high, low, turnover_value FROM prices "
        "WHERE code=? AND date<? AND adj_close IS NOT NULL AND close IS NOT NULL "
        "AND high IS NOT NULL AND low IS NOT NULL ORDER BY date DESC LIMIT ?",
        (ticker, entry_date, n)).fetchall()
    bars = []
    for d, c, ac, hi, lo, tv in reversed(rows):
        if not c:
            continue
        k = ac / c                                  # その行の調整係数
        bars.append({"date": d, "high": hi * k, "low": lo * k,
                     "close": ac, "turnover": tv or 0.0})
    return bars


# ------------------------------------------------------------------ 分岐2: テーゼ破綻
def _yoy_pct(ts, i):
    """タイル i の前年同期比（run-rate の差で判定。前年が負でも壊れない）。"""
    cur = ts[i]
    peer = SD._yoy_tile(ts, i)
    if peer is None:
        return None
    return cur["run_rate"] - peer["run_rate"]


def thesis_break(pcon, ticker, event_date, cache={}):
    """T+0 に開示された四半期で BS→PL 転換が起きたか。

    OR（ユーザー承認 2026-09-20）: 売上・営業利益のうち**取れたもののどれかが前年割れ**なら破綻。
    AND（両方とも前年割れのときだけ破綻）は比較用に併せて返す。
    どちらも取れなければ判定不能（破綻にしない）。
    """
    key = (ticker, event_date)
    if key in cache:
        return cache[key]
    vis = SS.visible_periods(pcon, ticker, event_date)
    got = {}
    for item, label in (("sales", "sales"), ("operating_profit", "op")):
        ts = SD.tiles(pcon, ticker, vis, item=item)
        got[label] = _yoy_pct(ts, len(ts) - 1) if ts else None
    have = [v for v in got.values() if v is not None]
    res = {
        "sales_yoy": got["sales"], "op_yoy": got["op"],
        "evaluable": bool(have),
        "break_or": bool(have) and any(v <= 0 for v in have),
        "break_and": bool(have) and all(v <= 0 for v in have),
    }
    cache[key] = res
    return res


# ------------------------------------------------------------------ 分岐3: 同時開示の修正
def revision_dates(mcon):
    """(ticker, 開示日) → 業績/配当予想の修正があった集合（TDnet アーカイブ）。"""
    out = set()
    for code, d in mcon.execute(
            "SELECT code, date FROM filings WHERE source='tdnet' "
            "AND subtype IN ('業績予想修正','配当予想修正')"):
        if code:
            out.add((code, d))
    return out


# ------------------------------------------------------------------ 1建玉の出口
def decide_exits(path, zone, thesis, has_revision, base_i, t1_i, drift_i):
    """各バリアントの (出口 index, 分岐名)。index は path 上の位置（0 = エントリー日）。

    出口日は別枠と同じ営業日カレンダーで event_date から数えたもの（呼び出し側が解決）。
    """
    out = {"base_i": base_i, "struct_i": None, "entry_below_val": 0,
           "zone_ok": int(bool(zone)), "thesis": thesis, "revision": int(has_revision),
           "drift_i": drift_i}

    if zone:
        closes = path[:base_i + 1]
        out["entry_below_val"] = int(bool(closes) and closes[0][1] < zone["val"])
        hit = AZ.first_break(zone["val"], closes, AZ.BREAK_DAYS)
        if hit:
            out["struct_i"] = hit[2]

    thesis_i = t1_i if (thesis and thesis["break_or"] and t1_i is not None) else None
    rev_i = t1_i if (has_revision and t1_i is not None) else None

    out["E5"] = ((t1_i, "テーゼ破綻")
                 if (thesis and thesis["break_or"] and t1_i is not None)
                 else (drift_i, "ドリフトT+20"))
    picks = {
        "E0": [(base_i, "基準T+2")],
        "E1": [(out["struct_i"], "構造破綻"), (base_i, "基準T+2")],
        "E2": [(thesis_i, "テーゼ破綻"), (base_i, "基準T+2")],
        "E3": [(rev_i, "初動利確"), (base_i, "基準T+2")],
        "E4": [(out["struct_i"], "構造破綻"), (thesis_i, "テーゼ破綻"),
               (rev_i, "初動利確"), (base_i, "基準T+2")],
    }
    for v, cands in picks.items():
        ok = [(i, name) for i, name in cands if i is not None]
        i = min(x[0] for x in ok)
        # 同じ index に複数該当したら事前登録の判定順（構造→テーゼ→初動→基準）
        name = next(nm for j, nm in ok if j == i)
        out[v] = (i, name)
    return out


# ------------------------------------------------------------------ 集計
def replay(rows, variant, pcon):
    """出口を差し替えて資産曲線を組み直す（v2 と同じ 10枠・NAV10%・実現ベース）。"""
    trades = []
    for r in rows:
        i, name = r["exits"][variant]
        if i >= len(r["path"]):
            i = len(r["path"]) - 1
        d_exit, p_exit = r["path"][i]
        gross = p_exit / r["path"][0][1] - 1
        trades.append({"ticker": r["t"]["ticker"], "entry_date": r["t"]["entry_date"],
                       "exit_date": d_exit, "ret": gross,
                       "net": gross - V1.COST_ROUND_TRIP, "branch": name,
                       "score": r["t"].get("score"), "hold": i})
    nav, peak, mdd = 1.0, 1.0, 0.0
    open_pos = []
    for tr in sorted(trades, key=lambda x: x["entry_date"]):
        still = []
        for ex, alloc, net in open_pos:
            if ex <= tr["entry_date"]:
                nav += alloc * net
                peak = max(peak, nav)
                mdd = min(mdd, nav / peak - 1)
            else:
                still.append((ex, alloc, net))
        open_pos = still
        open_pos.append((tr["exit_date"], nav * V2.POS_FRACTION, tr["net"]))
    for ex, alloc, net in sorted(open_pos, key=lambda x: x[0]):
        nav += alloc * net
        peak = max(peak, nav)
        mdd = min(mdd, nav / peak - 1)

    net = [t["net"] for t in trades]
    profits = sorted((r for r in net if r > 0), reverse=True)
    exc = []
    for t in trades:
        ir = V1.index_return(pcon, t["entry_date"], t["exit_date"])
        if ir is not None:
            exc.append(t["net"] - ir)
    return {
        "variant": variant, "n_trades": len(trades),
        "win_rate": sum(r > 0 for r in net) / len(net) if net else None,
        "expectancy_net": sum(net) / len(net) if net else None,
        "excess_vs_topix": (sum(exc) / len(exc)) if exc else None,
        "max_drawdown": mdd, "nav_final": nav,
        "top5_profit_share": (sum(profits[:5]) / sum(profits)) if profits else None,
        "avg_hold": sum(t["hold"] for t in trades) / len(trades) if trades else None,
        "branches": Counter(t["branch"] for t in trades), "trades": trades,
    }


def build_rows(trades, pcon, mcon, revisions):
    """建玉ごとに価格パス・受容帯・テーゼ・修正の有無を作る。

    出口日は別枠と同じく **event_date からの営業日** で数える（価格系列上の本数ではない）。
    休場・売買なしでその日が系列に無ければ、次に価格のある日に寄せて `snapped` に数える。
    """
    rows, diffs, snapped, dropped = [], [], 0, 0
    for t in trades:
        last = bd_after(t["event_date"], DRIFT_K + 5)
        path = pcon.execute(
            "SELECT date, close FROM daily_prices WHERE ticker=? AND date>=? AND date<=? "
            "ORDER BY date", (t["ticker"], t["entry_date"], last)).fetchall()
        if not path or path[0][0] != t["entry_date"]:
            dropped += 1
            continue
        base_i = idx_at(path, bd_after(t["event_date"], t["exit_k"]))
        if base_i is None:
            dropped += 1
            continue
        if path[base_i][0] != bd_after(t["event_date"], t["exit_k"]):
            snapped += 1
        t1_i = idx_at(path, bd_after(t["event_date"], THESIS_EXIT_K))
        drift_i = idx_at(path, bd_after(t["event_date"], DRIFT_K))
        if drift_i is None:
            drift_i = len(path) - 1
        bars = zone_bars(mcon, t["ticker"], t["entry_date"])
        zone = AZ.value_area(bars) if len(bars) >= AZ.N_DAYS else None
        th = thesis_break(pcon, t["ticker"], t["event_date"])
        has_rev = (t["ticker"], t["event_date"]) in revisions
        ex = decide_exits(path, zone, th, has_rev, base_i, t1_i, drift_i)
        diffs.append(abs((path[base_i][1] / path[0][1] - 1) - float(t["ret"])))
        rows.append({"t": t, "path": path, "zone": zone, "exits": ex})
    return rows, (max(diffs) if diffs else None), snapped, dropped


def branch_counts(rows):
    n_zone = sum(r["exits"]["zone_ok"] for r in rows)
    below = [r for r in rows if r["exits"]["entry_below_val"]]
    struct = [r for r in rows if r["exits"]["struct_i"] is not None]
    th_eval = [r for r in rows if r["exits"]["thesis"]["evaluable"]]
    th_or = [r for r in rows if r["exits"]["thesis"]["break_or"]]
    th_and = [r for r in rows if r["exits"]["thesis"]["break_and"]]
    rev = [r for r in rows if r["exits"]["revision"]]
    return {
        "n": len(rows), "zone_ok": n_zone, "zone_missing": len(rows) - n_zone,
        "entry_below_val": len(below), "struct": len(struct),
        "thesis_evaluable": len(th_eval), "thesis_or": len(th_or),
        "thesis_and": len(th_and), "revision": len(rev),
        "below_rows": below,
    }


def run_shift(shift, out_dir):
    C.log("=== シャドウE 出口シミュレーション shift=%+d ===" % shift)
    _res, trades = V1.run_one(shift)
    pcon = sqlite3.connect("file:%s?mode=ro" % V1._db_path(shift).replace("\\", "/"), uri=True)
    mcon = sqlite3.connect("file:%s?mode=ro" % C.DB_PATH.replace("\\", "/"), uri=True)
    sim = V2.simulate(trades, pcon, V1.MAIN_ENTRY, V1.MAIN_EXIT)
    taken = sim["trades"]
    C.log("  v2 建玉 %d 件（候補 %d）" % (len(taken), sim["n_candidates"]))

    rows, worst, snapped, dropped = build_rows(taken, pcon, mcon, revision_dates(mcon))
    C.log("  価格パスを組めた建玉 %d 件（除外 %d / 出口日を次の営業日に寄せた %d）"
          % (len(rows), dropped, snapped))
    C.log("  E0 の再計算と v2 の ret の最大差 %.2e" % (worst if worst is not None else float("nan")))
    cnt = branch_counts(rows)
    res = {v: replay(rows, v, pcon) for v in VARIANTS}
    drift = replay([dict(r, exits=dict(r["exits"],
                                       Edrift=(r["exits"]["drift_i"], "ドリフトT+20")))
                    for r in rows], "Edrift", pcon)
    rec = {"drift": drift, "E5": replay(rows, "E5", pcon)}

    path = os.path.join(out_dir, "exit_sim_trades_shift%+d.csv" % shift)
    with open(path, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh)
        w.writerow(["ticker", "entry_date", "event_date", "score", "zone_ok", "val",
                    "entry_below_val", "struct_day", "thesis_evaluable", "thesis_break_or",
                    "thesis_break_and", "sales_yoy", "op_yoy", "revision"]
                   + ["%s_branch" % v for v in VARIANTS]
                   + ["%s_net" % v for v in VARIANTS] + ["drift20_net"])
        for i, r in enumerate(rows):
            e, th = r["exits"], r["exits"]["thesis"]
            w.writerow([r["t"]["ticker"], r["t"]["entry_date"], r["t"]["event_date"],
                        r["t"].get("score"), e["zone_ok"],
                        "" if not r["zone"] else round(r["zone"]["val"], 2),
                        e["entry_below_val"],
                        "" if e["struct_i"] is None else e["struct_i"],
                        int(th["evaluable"]), int(th["break_or"]), int(th["break_and"]),
                        "" if th["sales_yoy"] is None else round(th["sales_yoy"], 1),
                        "" if th["op_yoy"] is None else round(th["op_yoy"], 1),
                        e["revision"]]
                       + [res[v]["trades"][i]["branch"] for v in VARIANTS]
                       + [round(res[v]["trades"][i]["net"], 5) for v in VARIANTS]
                       + [round(drift["trades"][i]["net"], 5)])
    return {"shift": shift, "res": res, "drift": drift, "rec": rec, "counts": cnt,
            "worst_diff": worst, "snapped": snapped, "dropped": dropped,
            "csv": path, "rows": rows}


def report(store):
    L = []
    L.append("=== 分岐の作動件数（建玉ベース）===")
    L.append("  shift  建玉  受容帯あり  帯なし  下限割れ入場  構造破綻  テーゼ判定可  破綻OR  破綻AND  修正同時")
    for s, st in store.items():
        c = st["counts"]
        L.append("  %+d    %4d  %9d  %6d  %12d  %8d  %12d  %6d  %7d  %8d"
                 % (s, c["n"], c["zone_ok"], c["zone_missing"], c["entry_below_val"],
                    c["struct"], c["thesis_evaluable"], c["thesis_or"], c["thesis_and"],
                    c["revision"]))
    L.append("")
    L.append("=== バリアント別（往復コスト %.1f%% 控除後）===" % (V1.COST_ROUND_TRIP * 100))
    for s, st in store.items():
        L.append("  shift %+d" % s)
        L.append("    変種  建玉   勝率   期待値(純)   TOPIX超過    MDD      上位5集中  平均保有  出口内訳")
        for v in VARIANTS + RECORD_VARIANTS:
            m = st["res"][v] if v in st["res"] else st["rec"][v]
            br = " ".join("%s%d" % (k, n) for k, n in m["branches"].most_common())
            L.append("    %-5s %4d  %5.1f%%  %+9.4f  %s  %+.4f  %s  %5.1f日  %s"
                     % (v, m["n_trades"], (m["win_rate"] or 0) * 100, m["expectancy_net"],
                        "%+9.4f" % m["excess_vs_topix"] if m["excess_vs_topix"] is not None else "      n/a",
                        m["max_drawdown"],
                        "%7.1f%%" % (m["top5_profit_share"] * 100) if m["top5_profit_share"] else "    n/a",
                        m["avg_hold"], br))
        L.append("")
    L.append("  drift / E5 は**記録専用**（B-1）。判定には使わない。")
    L.append("  E5 = テーゼ破綻なら T+1、健全なら T+20 保有（記録アームを分岐2で条件付けた**事後の切り口**）")
    L.append("")
    L.append("=== 判定（事前登録 B-3: E4 − E0 の期待値差の符号が3本で一致するか）===")
    diffs = {s: st["res"]["E4"]["expectancy_net"] - st["res"]["E0"]["expectancy_net"]
             for s, st in store.items()}
    for s, d in diffs.items():
        L.append("  shift %+d: E4 − E0 = %+.4f" % (s, d))
    if len(diffs) < 3:
        L.append("  → 3本そろっていないので判定しない（--quick）")
    else:
        signs = {d > 0 for d in diffs.values()}
        L.append("  → 符号一致: %s" % ("はい（%s）" % ("プラス" if all(d > 0 for d in diffs.values()) else "マイナス")
                                      if len(signs) == 1 else "いいえ（符号不定）"))
    L.append("")
    L.append("=== 受容帯下限割れエントリーのコスト（申し送り2の判断材料）===")
    for s, st in store.items():
        below = st["counts"]["below_rows"]
        idx = {id(r): i for i, r in enumerate(st["rows"])}
        e4 = st["res"]["E4"]["trades"]
        e0 = st["res"]["E0"]["trades"]
        if not below:
            L.append("  shift %+d: 0 件" % s)
            continue
        cost = len(below) * V1.COST_ROUND_TRIP
        e4n = [e4[idx[id(r)]]["net"] for r in below]
        e0n = [e0[idx[id(r)]]["net"] for r in below]
        L.append("  shift %+d: %d 件 / 往復コスト計 %.2f%%（1建玉0.4%%）/ "
                 "その建玉の期待値 E0 %+.4f → E4 %+.4f"
                 % (s, len(below), cost * 100, sum(e0n) / len(e0n), sum(e4n) / len(e4n)))
    L.append("")
    L.append("=== 判定不能として数えるもの ===")
    for s, st in store.items():
        c = st["counts"]
        L.append("  shift %+d: 受容帯なし %d / テーゼ判定不能 %d / 通期予想の引き下げ: 判定不能（company_forecasts は 2026-07-23 以降）"
                 % (s, c["zone_missing"], c["n"] - c["thesis_evaluable"]))
    return "\n".join(L)


def main(argv=None):
    sys.stdout.reconfigure(encoding="utf-8")
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--quick", action="store_true", help="shift 0 のみ")
    p.add_argument("--out-dir", default=C.DATA_DIR)
    a = p.parse_args(argv)
    shifts = (0,) if a.quick else V1.SHIFTS
    store = {s: run_shift(s, a.out_dir) for s in shifts}
    text = report(store)
    print(text)
    out = os.path.join(a.out_dir, "exit_sim_report.md")
    with open(out, "w", encoding="utf-8") as fh:
        fh.write("# シャドウE フェーズB 出口4分岐（インサンプル・記述統計）\n\n```\n%s\n```\n" % text)
    C.log("出力: %s" % out)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
