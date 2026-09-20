"""screener/technical/acceptance_zone.py — 受容帯（簡易ボリュームプロファイル）。

定義の正本は docs/backtest_acceptance_criteria.md シャドウE B-2。
**このモジュールは定義を実装するだけで、パラメータを決めない。**
N=60 / B=20 / X=70% / M=3 は事前登録の値で、結果を見て動かさない。

テクニカル層の制約（E-0）: これは**撤退判定にしか使わない**。候補の追加も
スコアへの加点もしない。DB には書かない（この module は DB を開かない）。

    zone = value_area(bars)                      # bars は N 日ぶんの OHLCV
    broken_at = first_break(zone, closes, M)     # 終値が VAL を M 日連続で下回った日
"""
from __future__ import annotations

# ---- 事前登録の値（B-2）。動かさない -------------------------------------
N_DAYS = 60          # 受容帯を作る営業日数（エントリー日の前営業日まで）
N_BINS = 20          # 価格帯の本数（期間中の最安値〜最高値を等幅に）
VALUE_AREA = 0.70    # POC から隣接拡張して累積がこの割合に達するまで
BREAK_DAYS = 3       # 終値が VAL を下回る連続日数


def bucket_turnover(bars, n_bins=N_BINS):
    """各日の売買代金を [安値, 高値] に一様配分して価格帯に積む。

    bars: [{"high", "low", "close", "turnover"}]（すべて分割調整済み）。
    戻り値: (edges, weights)。edges は n_bins+1 本の境界。
    """
    lo = min(b["low"] for b in bars)
    hi = max(b["high"] for b in bars)
    if not (hi > lo):
        return None, None
    width = (hi - lo) / n_bins
    edges = [lo + width * i for i in range(n_bins + 1)]
    w = [0.0] * n_bins

    def idx(p):
        return min(int((p - lo) / width), n_bins - 1)

    for b in bars:
        t = b.get("turnover") or 0.0
        if t <= 0:
            continue
        a, z = idx(b["low"]), idx(b["high"])
        if a == z or b["high"] <= b["low"]:
            w[idx(b["close"])] += t
            continue
        # 値幅に一様配分。端の帯は重なった長さぶんだけ受け取る。
        span = b["high"] - b["low"]
        for i in range(a, z + 1):
            over = min(b["high"], edges[i + 1]) - max(b["low"], edges[i])
            if over > 0:
                w[i] += t * over / span
    return edges, w


def value_area(bars, n_bins=N_BINS, share=VALUE_AREA):
    """POC から隣接する帯へ広げ、累積売買代金が share に達した範囲を返す。

    戻り値: {"val": 下限, "vah": 上限, "poc": POC帯の中央値, "share": 実際の割合}。
    作れなければ None（推測で埋めない）。
    """
    edges, w = bucket_turnover(bars, n_bins)
    if not w:
        return None
    total = sum(w)
    if total <= 0:
        return None
    poc = max(range(len(w)), key=lambda i: w[i])
    lo = hi = poc
    acc = w[poc]
    while acc < total * share and (lo > 0 or hi < len(w) - 1):
        below = w[lo - 1] if lo > 0 else -1.0
        above = w[hi + 1] if hi < len(w) - 1 else -1.0
        # 同値なら下側を先に取る（決定的にする）
        if below >= above:
            lo -= 1
            acc += below
        else:
            hi += 1
            acc += above
    return {"val": edges[lo], "vah": edges[hi + 1],
            "poc": (edges[poc] + edges[poc + 1]) / 2,
            "share": acc / total, "n_bins": n_bins, "n_bars": len(bars)}


def first_break(val, closes, m=BREAK_DAYS):
    """終値が val を m 日連続で下回った、その m 日目の位置。無ければ None。

    closes: [(日付, 終値)] の昇順。戻り値は (日付, 終値, 位置)。
    """
    run = 0
    for i, (d, c) in enumerate(closes):
        if c is None:
            run = 0
            continue
        if c < val:
            run += 1
            if run >= m:
                return d, c, i
        else:
            run = 0
    return None
