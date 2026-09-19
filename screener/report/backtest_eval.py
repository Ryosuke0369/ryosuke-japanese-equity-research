"""screener/report/backtest_eval.py — 事前登録した合格基準でバックテストを判定する。

基準の正本は docs/backtest_acceptance_criteria.md（2026-09-01 事前登録）。
**このスクリプトは基準を実装するだけで、基準を決めない。** 実行後に
閾値やセルを動かして判定し直すことは禁止されている。

別枠との関係
------------
別枠 module_d/backtest.py の**ファイルは1行も変更しない**。ただし
`WINDOW_START` / `WINDOW_END` はモジュール定数で 2024-09-01..2026-08-31 に
固定されており、これは事前登録 §0 が判定に使うと定めた価格窓
（2021-09-01..2026-08-31）と食い違う。実行時に属性を差し替えて窓を合わせ、
そのことを報告に明記する。ファイルは触っていない。

別枠が出さない数字はここで計算する
----------------------------------
別枠の metrics() は勝率・期待値・MDD・シャープを出すが、事前登録が要求する
  - 往復コスト0.4%控除後の期待値
  - TOPIX超過リターン
  - 利益集中度（上位5トレードが総利益に占める割合）
は持っていない。ここで生リターンから計算する。

    python -m screener.report.backtest_eval
    python -m screener.report.backtest_eval --quick     # 主セルのみ・シフト無し
"""
from __future__ import annotations

import argparse
import csv
import os
import sys
from collections import defaultdict
from datetime import date
from pathlib import Path

try:
    from screener import common as C
except ImportError:                                     # pragma: no cover
    sys.path.insert(0, os.path.dirname(os.path.dirname(
        os.path.dirname(os.path.abspath(__file__)))))
    from screener import common as C

# ---- 事前登録の値。ここを実行後に動かさない -------------------------------
MAIN_ENTRY, MAIN_EXIT = 15, 2          # 主セル: T-15 エントリー / T+2 エグジット
COST_ROUND_TRIP = 0.004                # 往復コスト 0.4%
MIN_TRADES = 300
MAX_DD = -0.15
MAX_TOP5_PROFIT_SHARE = 0.30
POS_SIZE = 0.10                        # 資産曲線のポジション比率
WINDOW = (date(2021, 9, 1), date(2026, 8, 31))
SHIFTS = (0, -3, 3)
WIN_RATE_TOLERANCE_PT = 3.0            # シフト間の勝率変動の許容幅

# ---- P3-4 時価総額帯（2026-09-18 事前登録）。境界を結果を見て動かさない ----
# (帯, 下限, 上限) JPY mn。各帯はちょうど「その拡大で増えた分」:
#   50-600 = 初版ユニバース / 600-1000 = 08-31 拡大分 / 1000-3000 = 09-18 拡大分。
# 当時のユニバース条件（min <= mktcap <= max）に揃え、先頭帯だけ下限を含む。
MKTCAP_BANDS = (("50-600", 5000, 60000),
                ("600-1000", 60000, 100000),
                ("1000-3000", 100000, 300000))
BAND_OUT, BAND_UNKNOWN = "範囲外", "不明"


def mktcap_band(mktcap_mn):
    """時価総額（JPY mn）→ 帯名。記録が無ければ「不明」、どの帯にも無ければ「範囲外」。"""
    if mktcap_mn is None:
        return BAND_UNKNOWN
    for i, (name, lo, hi) in enumerate(MKTCAP_BANDS):
        if (lo <= mktcap_mn if i == 0 else lo < mktcap_mn) and mktcap_mn <= hi:
            return name
    return BAND_OUT


def mktcap_at_entry(mcon, ticker, entry_date):
    """entry_date の**前営業日まで**で最新の時価総額（JPY mn）。無ければ None。

    今日の companies.mktcap で分けると、当時の規模ではなく生き残った後の
    規模で層別することになる。寄り付きで建てる以上、当日の値も使わない。
    """
    r = mcon.execute(
        "SELECT mktcap FROM prices WHERE code=? AND date<? AND mktcap IS NOT NULL "
        "ORDER BY date DESC LIMIT 1", (ticker, entry_date)).fetchone()
    return r[0] if r else None


def tag_mktcap_bands(trades, mcon):
    """各トレードに mktcap_mn_at_entry / mktcap_band を付ける（その場で書き換える）。"""
    cache = {}
    for t in trades:
        key = (t["ticker"], t["entry_date"])
        if key not in cache:
            cache[key] = mktcap_at_entry(mcon, *key)
        t["mktcap_mn_at_entry"] = cache[key]
        t["mktcap_band"] = mktcap_band(cache[key])
    return trades


def band_names():
    return [b[0] for b in MKTCAP_BANDS] + [BAND_OUT, BAND_UNKNOWN]


def _main_db():
    """本体DB（prices.mktcap を持つ）を読み取り専用で開く。"""
    import sqlite3
    path = os.path.join(C.DATA_DIR, "screener.db").replace("\\", "/")
    return sqlite3.connect("file:%s?mode=ro" % path, uri=True)


def _external_root():
    for base in (Path(__file__).resolve().parents[2],):
        for cand in base.glob("Downloads/earnings_screener*/earnings_screener"):
            if (cand / "module_d" / "backtest.py").exists():
                return cand
    raise SystemExit("別枠 earnings_screener が見つからない")


def _db_path(shift):
    base = os.path.join(C.DATA_DIR, "projection.db")
    if not shift:
        return base
    root, ext = os.path.splitext(base)
    return "%s_%+d%s" % (root, shift, ext)


def run_one(shift):
    """1本の投影DBでバックテストを回し、トレードと集計を返す。"""
    root = _external_root()
    if str(root) not in sys.path:
        sys.path.insert(0, str(root))
    import module_d.backtest as BT

    # 窓だけ事前登録に合わせる。ファイルは触らない。
    BT.WINDOW_START, BT.WINDOW_END = WINDOW
    out = Path(C.DATA_DIR) / ("backtest_shift%+d" % shift)
    res = BT.run_backtest(db_path=_db_path(shift), out_dir=out)

    trades = []
    with open(out / "backtest_trades.csv", encoding="utf-8-sig") as fh:
        for r in csv.DictReader(fh):
            r["entry_n"] = int(r["entry_n"])
            r["exit_k"] = int(r["exit_k"])
            r["ret"] = float(r["ret"])
            trades.append(r)

    # P3-4: 時価総額帯を付けて記録する。別枠の backtest_trades.csv は触らず、
    # 帯つきの写しを横に置く（拡大前後の効果を後から分けて読めるように）。
    mcon = _main_db()
    try:
        tag_mktcap_bands(trades, mcon)
    finally:
        mcon.close()
    if trades:
        with open(out / "backtest_trades_banded.csv", "w", newline="",
                  encoding="utf-8-sig") as fh:
            w = csv.DictWriter(fh, fieldnames=list(trades[0].keys()))
            w.writeheader()
            w.writerows(trades)
    return res, trades


def index_return(con, d_entry, d_exit):
    """同じ保有期間の TOPIX リターン。両端が無ければ None。"""
    a = con.execute("SELECT close FROM market_index WHERE date=?", (d_entry,)).fetchone()
    b = con.execute("SELECT close FROM market_index WHERE date=?", (d_exit,)).fetchone()
    if not a or not b or not a[0]:
        return None
    return b[0] / a[0] - 1


def max_drawdown(rets_chrono):
    equity, peak, mdd = 1.0, 1.0, 0.0
    for r in rets_chrono:
        equity *= 1 + r * POS_SIZE
        peak = max(peak, equity)
        mdd = min(mdd, equity / peak - 1)
    return mdd


def evaluate(trades, con, entry_n=MAIN_ENTRY, exit_k=MAIN_EXIT):
    """主セルの5項目を計算する。リターンは往復コスト控除後。"""
    sel = [t for t in trades if t["entry_n"] == entry_n and t["exit_k"] == exit_k]
    if not sel:
        return None
    sel.sort(key=lambda t: t["event_date"])
    net = [t["ret"] - COST_ROUND_TRIP for t in sel]
    wins = [r for r in net if r > 0]
    profits = sorted((r for r in net if r > 0), reverse=True)
    total_profit = sum(profits)
    top5 = sum(profits[:5])

    exc = []
    for t in sel:
        ir = index_return(con, t["entry_date"], _exit_date(con, t))
        if ir is not None:
            exc.append((t["ret"] - COST_ROUND_TRIP) - ir)

    return {
        "n_trades": len(sel),
        "win_rate": len(wins) / len(sel),
        "expectancy_net": sum(net) / len(net),
        "expectancy_gross": sum(t["ret"] for t in sel) / len(sel),
        "excess_vs_topix": (sum(exc) / len(exc)) if exc else None,
        "n_excess_measurable": len(exc),
        "max_drawdown": max_drawdown(net),
        "top5_profit_share": (top5 / total_profit) if total_profit > 0 else None,
        "total_profit": total_profit,
    }


def _exit_date(con, t):
    """別枠のトレードは exit 日を持たないので、価格系列から k 営業日後を引く。"""
    rows = con.execute(
        "SELECT date FROM daily_prices WHERE ticker=? AND date>=? ORDER BY date LIMIT ?",
        (t["ticker"], t["entry_date"], t["entry_n"] + t["exit_k"] + 1)).fetchall()
    dates = [r[0] for r in rows]
    want = t["entry_n"] + t["exit_k"]
    return dates[want] if len(dates) > want else dates[-1]


BAND_MIN_TRADES = 100                  # §4: 1層100件未満は合否材料にしない


def band_report(trades, con, entry_n=MAIN_ENTRY, exit_k=MAIN_EXIT):
    """P3-4 時価総額帯ごとの5項目。**記述統計のみ**（合否は全体の主セル）。

    trades は tag_mktcap_bands 済みであること。MDD は帯内のトレードだけで
    組んだ 10% 固定の資産曲線（evaluate と同じ定義）。
    """
    C.log("  帯          trades   win%   期待値(純)  TOPIX超過   MDD      上位5集中")
    for band in band_names():
        sub = [t for t in trades if t.get("mktcap_band") == band]
        m = evaluate(sub, con, entry_n, exit_k)
        if not m:
            C.log("  %-10s %7s" % (band, "0"))
            continue
        C.log("  %-10s %7s  %5.1f  %+10.4f  %s  %+.4f  %s%s"
              % (band, format(m["n_trades"], ","), m["win_rate"] * 100,
                 m["expectancy_net"],
                 "%+9.4f" % m["excess_vs_topix"] if m["excess_vs_topix"] is not None
                 else "      n/a",
                 m["max_drawdown"],
                 "%.3f" % m["top5_profit_share"] if m["top5_profit_share"] is not None
                 else "n/a",
                 "" if m["n_trades"] >= BAND_MIN_TRADES
                 else "  (<%d件: 合否材料にしない)" % BAND_MIN_TRADES))


def verdict(m):
    """事前登録 §2 の5項目。1つでも欠ければ不合格。"""
    checks = [
        ("トレード数 >= %d" % MIN_TRADES, m["n_trades"] >= MIN_TRADES, m["n_trades"]),
        ("期待値 > 0（コスト控除後）", m["expectancy_net"] > 0, m["expectancy_net"]),
        ("TOPIX超過 > 0",
         m["excess_vs_topix"] is not None and m["excess_vs_topix"] > 0,
         m["excess_vs_topix"]),
        ("最大DD >= %.0f%%" % (MAX_DD * 100), m["max_drawdown"] >= MAX_DD,
         m["max_drawdown"]),
        ("上位5トレード利益集中 < %.0f%%" % (MAX_TOP5_PROFIT_SHARE * 100),
         m["top5_profit_share"] is not None
         and m["top5_profit_share"] < MAX_TOP5_PROFIT_SHARE,
         m["top5_profit_share"]),
    ]
    return checks, all(ok for _, ok, _ in checks)


def main(argv=None):
    sys.stdout.reconfigure(encoding="utf-8")
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--quick", action="store_true", help="主セルのみ・シフト無し")
    a = p.parse_args(argv)

    import sqlite3
    shifts = (0,) if a.quick else SHIFTS
    results = {}
    for s in shifts:
        C.log("=== バックテスト shift=%+d ===" % s)
        res, trades = run_one(s)
        con = sqlite3.connect("file:%s?mode=ro" % _db_path(s).replace("\\", "/"), uri=True)
        m = evaluate(trades, con)
        results[s] = {"res": res, "trades": trades, "metrics": m, "con": con}
        C.log("  イベント %s / スコア試行 %s / 閾値通過 %s"
              % (format(res["n_events"], ","), format(res["n_scored"], ","),
                 format(res["n_passed"], ",")))

    base = results[0]
    m = base["metrics"]
    C.log("")
    C.log("=== 主セル T-%d / T+%d（事前登録 §1）===" % (MAIN_ENTRY, MAIN_EXIT))
    checks, passed = verdict(m)
    for label, ok, val in checks:
        shown = "%.4f" % val if isinstance(val, float) else str(val)
        C.log("  [%s] %-34s 実測 %s" % ("合" if ok else "否", label, shown))
    C.log("  → 総合判定: %s" % ("合格" if passed else "不合格"))
    C.log("  （参考）勝率 %.1f%% / コスト控除前の期待値 %.4f"
          % (m["win_rate"] * 100, m["expectancy_gross"]))

    C.log("")
    C.log("=== グリッド9通り（頑健性レポート・合否には使わない）===")
    C.log("  entry exit  trades  win%   期待値(粗)  期待値(純)")
    for n in (15, 10, 5):
        for k in (1, 2, 5):
            g = evaluate(base["trades"], base["con"], n, k)
            if g:
                C.log("  T-%-3d T+%-2d %7s  %5.1f  %10.4f  %10.4f"
                      % (n, k, format(g["n_trades"], ","), g["win_rate"] * 100,
                         g["expectancy_gross"], g["expectancy_net"]))

    C.log("")
    C.log("=== 時価総額帯別 T-%d / T+%d（P3-4・記述統計・合否には使わない）==="
          % (MAIN_ENTRY, MAIN_EXIT))
    band_report(base["trades"], base["con"])

    if not a.quick:
        C.log("")
        C.log("=== 発表日 ±3営業日シフト（事前登録 §3）===")
        C.log("  shift  trades   win%    期待値(純)")
        wr = []
        for s in SHIFTS:
            mm = results[s]["metrics"]
            wr.append(mm["win_rate"] * 100)
            C.log("  %+d   %7s  %5.1f   %10.4f"
                  % (s, format(mm["n_trades"], ","), mm["win_rate"] * 100,
                     mm["expectancy_net"]))
        all_pos = all(results[s]["metrics"]["expectancy_net"] > 0 for s in SHIFTS)
        spread = max(wr) - min(wr)
        C.log("  3本すべて期待値プラス: %s" % ("はい" if all_pos else "いいえ"))
        C.log("  勝率の変動幅: %.1f pt（許容 %.1f pt以内）: %s"
              % (spread, WIN_RATE_TOLERANCE_PT,
                 "満たす" if spread <= WIN_RATE_TOLERANCE_PT else "満たさない"))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
