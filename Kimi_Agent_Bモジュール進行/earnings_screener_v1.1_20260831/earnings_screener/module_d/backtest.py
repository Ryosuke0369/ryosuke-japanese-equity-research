"""
backtest.py — Module D: バックテストフレームワーク
過去2年分（2024-09〜2026-08）の全決算イベントに対し：
- エントリー T-15/T-10/T-5（営業日）× 出口 T+1/T+2/T+5 のグリッド比較
- エントリー条件：PIT証拠スコア >= 0.10（その時点の公開情報のみで再計算）
- 層別分析：セクター別・市場環境別（指数60営業日リターンの正負）
- 出力：勝率・平均利益率・平均損失率・期待値・最大ドローダウン・シャープレシオ

注：シャープは per-trade リターンの mean/std × sqrt(年間トレード数) の簡易版。
"""
import csv
import statistics
import sys
from collections import defaultdict
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from common.jp_calendar import add_business_days
from module_b.data_access import connect, DB
from module_b.run_scorers import score_ticker, SCORERS_ALL

WINDOW_START, WINDOW_END = date(2024, 9, 1), date(2026, 8, 31)
ENTRY_DAYS = [15, 10, 5]
EXIT_DAYS = [1, 2, 5]
SCORE_THRESHOLD = 0.10


def load_prices(conn):
    prices = {}
    for t, d, c in conn.execute("SELECT ticker, date, close FROM daily_prices"):
        prices.setdefault(t, {})[d] = c
    index = {d: c for d, c in conn.execute("SELECT date, close FROM market_index")}
    return prices, index


def market_regime(index, d: date) -> str:
    past = add_business_days(d, -60).isoformat()
    ds = d.isoformat()
    if ds in index and past in index:
        return "上昇相場" if index[ds] / index[past] - 1 > 0 else "下降相場"
    return "不明"


POS_SIZE = 0.10  # 1トレードあたりNAVの10%（最大同時10ポジション想定）

def max_drawdown(returns_chronological):
    """固定比率ポジションサイジングでの簡易エクイティカーブMDD"""
    equity, peak, mdd = 1.0, 1.0, 0.0
    for r in returns_chronological:
        equity *= 1 + r * POS_SIZE
        peak = max(peak, equity)
        mdd = min(mdd, equity / peak - 1)
    return mdd


def metrics(trades, years=2.0):
    """trades: list of (exit_date, ret)"""
    if not trades:
        return None
    rets = [r for _, r in trades]
    wins = [r for r in rets if r > 0]
    losses = [r for r in rets if r <= 0]
    mean = statistics.mean(rets)
    std = statistics.stdev(rets) if len(rets) > 1 else 0.0
    tpy = len(rets) / years  # 年間トレード数
    sharpe = (mean / std * (tpy ** 0.5)) if std > 0 else 0.0
    return {
        "n_trades": len(rets),
        "win_rate": round(len(wins) / len(rets), 4),
        "avg_win": round(statistics.mean(wins), 4) if wins else 0.0,
        "avg_loss": round(statistics.mean(losses), 4) if losses else 0.0,
        "expectancy": round(mean, 4),
        "max_drawdown": round(max_drawdown([r for _, r in sorted(trades)]), 4),
        "sharpe": round(sharpe, 3),
    }


def run_backtest(db_path=None, out_dir=None):
    conn = connect(db_path)
    prices, index = load_prices(conn)
    sectors = {r["ticker"]: r["sector"] for r in conn.execute(
        "SELECT ticker, sector FROM universe").fetchall()}

    events = conn.execute(
        "SELECT ticker, filing_date, quarter_type FROM filings "
        "WHERE filing_date BETWEEN ? AND ? AND generation=1 ORDER BY filing_date",
        (WINDOW_START.isoformat(), WINDOW_END.isoformat())).fetchall()

    trades = []  # (ticker, event_date, entry_n, exit_k, ret, score, sector, regime)
    n_scored = n_passed = 0
    for ev in events:
        t, d_str, qt = ev["ticker"], ev["filing_date"], ev["quarter_type"]
        d = date.fromisoformat(d_str)
        if t not in prices:
            continue
        px = prices[t]
        exits = {}
        for k in EXIT_DAYS:
            ex = add_business_days(d, k).isoformat()
            if ex in px:
                exits[k] = px[ex]
        if not exits:
            continue
        regime = market_regime(index, d)
        sector = sectors.get(t, "不明")

        for n in ENTRY_DAYS:
            entry_d = add_business_days(d, -n)
            if entry_d.isoformat() not in px or entry_d < WINDOW_START:
                continue
            res = score_ticker(conn, t, SCORERS_ALL, as_of=entry_d)
            score = res["evidence_score"]
            n_scored += 1
            if score is None or score < SCORE_THRESHOLD:
                continue
            n_passed += 1
            for k, ex_close in exits.items():
                trades.append({
                    "ticker": t, "event_date": d_str, "quarter_type": qt,
                    "entry_n": n, "exit_k": k,
                    "entry_date": entry_d.isoformat(),
                    "ret": round(ex_close / px[entry_d.isoformat()] - 1, 5),
                    "score": score, "sector": sector, "regime": regime,
                })

    # ---- グリッド集計 ----
    grid = {}
    for n in ENTRY_DAYS:
        for k in EXIT_DAYS:
            sel = [(t["event_date"], t["ret"]) for t in trades
                   if t["entry_n"] == n and t["exit_k"] == k]
            m = metrics(sel)
            if m:
                grid[(n, k)] = m

    # ---- 最適コンボの層別 ----
    best = max(grid, key=lambda c: grid[c]["expectancy"]) if grid else None
    strata_sector, strata_regime = {}, {}
    if best:
        n0, k0 = best
        sel = [t for t in trades if t["entry_n"] == n0 and t["exit_k"] == k0]
        by_s = defaultdict(list)
        for t in sel:
            by_s[t["sector"]].append((t["event_date"], t["ret"]))
        strata_sector = {s: metrics(v) for s, v in by_s.items()}
        by_r = defaultdict(list)
        for t in sel:
            by_r[t["regime"]].append((t["event_date"], t["ret"]))
        strata_regime = {r: metrics(v) for r, v in by_r.items()}

    # ---- CSV出力 ----
    out_dir = Path(out_dir or Path(__file__).resolve().parents[1] / "data")
    out_dir.mkdir(parents=True, exist_ok=True)
    p_grid = out_dir / "backtest_grid.csv"
    with open(p_grid, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["entry_n", "exit_k", *list(next(iter(grid.values())).keys())] if grid else [])
        for (n, k), m in grid.items():
            w.writerow([n, k, *m.values()])
    p_tr = out_dir / "backtest_trades.csv"
    with open(p_tr, "w", newline="", encoding="utf-8-sig") as f:
        if trades:
            w = csv.DictWriter(f, fieldnames=list(trades[0].keys()))
            w.writeheader()
            w.writerows(trades)
    p_st = out_dir / "backtest_strata.csv"
    with open(p_st, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        if best:
            w.writerow(["axis", "key", *list(next(iter(strata_sector.values())).keys())])
            for s, m in strata_sector.items():
                w.writerow(["sector", s, *m.values()])
            for r, m in strata_regime.items():
                w.writerow(["regime", r, *m.values()])
    conn.close()
    return {"grid": grid, "best": best, "strata_sector": strata_sector,
            "strata_regime": strata_regime, "n_events": len(events),
            "n_scored": n_scored, "n_passed": n_passed,
            "files": (p_grid, p_tr, p_st)}


if __name__ == "__main__":
    res = run_backtest()
    print(f"イベント数: {res['n_events']} / PITスコアリング: {res['n_scored']} / エントリー通過: {res['n_passed']}")
    print("\n=== エントリー×出口グリッド ===")
    for (n, k), m in sorted(res["grid"].items()):
        print(f"T-{n:>2}→T+{k}: trades={m['n_trades']:>4} 勝率={m['win_rate']:.1%} "
              f"期待値={m['expectancy']:+.2%} 平均勝ち={m['avg_win']:+.2%} 平均負け={m['avg_loss']:+.2%} "
              f"MDD={m['max_drawdown']:.1%} Sharpe={m['sharpe']:.2f}")
    if res["best"]:
        n0, k0 = res["best"]
        print(f"\n最適: エントリー T-{n0} / 出口 T+{k0}")
        print("=== セクター別 ===")
        for s, m in sorted(res["strata_sector"].items(), key=lambda x: -x[1]["expectancy"]):
            print(f"  {s}: n={m['n_trades']} 勝率={m['win_rate']:.1%} 期待値={m['expectancy']:+.2%}")
        print("=== 市場環境別 ===")
        for r, m in res["strata_regime"].items():
            print(f"  {r}: n={m['n_trades']} 勝率={m['win_rate']:.1%} 期待値={m['expectancy']:+.2%}")
    print("\n出力:", *[str(p) for p in res["files"]], sep="\n  ")
