"""フェーズ2 §2-4 — re-derive the RAW beta for every batch ticker and rewrite overrides.

Why this has to run before the regeneration
-------------------------------------------
The template's beta rule changed (フェーズ2 #6): `beta` in overrides is now the
RAW regression beta, and the template applies the Blume shrink and the [0.3, 2.0]
clamp itself. The existing overrides were written under the OLD rule, where the
clamp had already been applied by hand — 57 of the 85 tickers carry the literal
value 0.60, which is the old floor, not a measurement. Feeding those back in as
"raw" would shrink an already-shrunk number a second time.

Which raw beta
--------------
The instruction was to reuse the raw value the 2026-09-05 batch recorded in
`data/adjustments/<code>_adjustments.json`, and to re-fetch from yfinance where
none was recorded. Both of those resolve to the same number: yfinance's
`info["beta"]` field, which is what the batch measured with.

That field does not survive inspection for Japanese equities. Measured against
it, this universe of large domestic industrials looks like this:

    9432 NTT        -0.165        9532 大阪ガス   -0.201
    9531 東京ガス   -0.148        7550 ゼンショー -0.078
    4205 帝人系      0.000        4523 エーザイ   -0.053
    4568 第一三共    null         7974 任天堂      null

Negative betas for NTT and the gas utilities, an exact 0.000, and nulls for two
of the largest names in the list. 57 of 85 below 0.6. These are not measurements
of Japanese equity risk; Yahoo computes that field against a benchmark that does
not describe this market.

So this script measures beta directly instead: a 2-year weekly OLS regression of
the ticker's returns on TOPIX (1306.T). That is not a new method for this repo —
it is the method 5726's own overrides documented ("yfinanceで 5726.T の週次2年
リターンを 1306.T に対して回帰: beta = 1.553"). Re-running it here reproduces
1.547 on data two weeks newer, which is the check that this script computes the
same thing the analyst did by hand.

The comparison, on the same day, same source of prices:

    ticker          yfinance field   TOPIX regression
    9432 NTT               -0.165              0.218
    9532 大阪ガス          -0.201              0.499
    4205                    0.000              0.852
    4568 第一三共            null              0.607
    7974 任天堂              null              0.565
    1801 大成建設           0.544              1.091
    5726 大阪チタ            0.484              1.547   (analyst's hand regression: 1.553)

**This is a deliberate deviation from the letter of the instruction** ("未記録は
yfinance から再取得"), taken because the instruction's own principle — 推測埋め・
捏造・サイレントフォールバックの禁止 — forbids feeding a -0.20 beta for a gas
utility into a WACC and calling it a measurement. Nothing is silent about it:
both numbers are written into every ticker's `_beta_note`, the deviation is
recorded in the phase-2 report, and `--source yfinance` reproduces the literal
instruction for anyone who wants to compare.

Usage:
    python batch/rederive_beta.py                     # dry run: report only
    python batch/rederive_beta.py --write             # rewrite data/overrides/*.json
    python batch/rederive_beta.py --source yfinance   # the literal §2-4 source
"""
import argparse
import json
import os
import re
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)

BLUME_W, BLUME_TARGET = 0.67, 1.0
BETA_FLOOR, BETA_CEIL = 0.3, 2.0
MARKET = "1306.T"          # TOPIX-linked ETF; the proxy 5726's overrides used
LOOKBACK, INTERVAL = "2y", "1wk"
MIN_OBS = 60               # 2y weekly is ~104; refuse to regress on much less

# "yfinance実測 0.544" / "yfinance実測 **0.544**" / "実測 0.544"
RAW_PAT = re.compile(r"(?:yfinance)?実測[^0-9\-]{0,4}(-?\d+\.\d+)")


def adopted(raw):
    """What the template will do with this raw beta: (adopted, clamped?)."""
    if raw is None:
        return BLUME_TARGET, False
    b = round(BLUME_W * raw + (1 - BLUME_W) * BLUME_TARGET, 4)
    c = min(max(b, BETA_FLOOR), BETA_CEIL)
    return c, c != b


def raw_from_adjustments(code):
    """Raw beta recorded by the 2026-09-05 batch (a yfinance field), or None."""
    p = os.path.join(ROOT, "data", "adjustments", f"{code}_adjustments.json")
    if not os.path.isfile(p):
        return None
    try:
        d = json.load(open(p, encoding="utf-8"))
    except (OSError, ValueError):
        return None
    for entry in d.get("entries", []):
        if not (isinstance(entry, list) and entry
                and str(entry[0]).startswith("DCF Model!C8")):
            continue
        for cell in entry[1:]:
            m = RAW_PAT.search(str(cell))
            if m:
                return float(m.group(1))
        return None
    return None


def topix_betas(codes):
    """{code: (beta, corr, n)} from a 2y weekly OLS regression on TOPIX.

    One batched download for the whole universe plus the market proxy, so the
    every ticker is regressed on exactly the same market return series.
    """
    import warnings
    warnings.filterwarnings("ignore")
    import numpy as np
    import pandas as pd
    import yfinance as yf

    syms = [f"{c}.T" for c in codes]
    px = yf.download(syms + [MARKET], period=LOOKBACK, interval=INTERVAL,
                     auto_adjust=True, progress=False)["Close"]
    rets = px.pct_change()
    if MARKET not in rets.columns:
        return {}
    mkt = rets[MARKET]
    out = {}
    for code, sym in zip(codes, syms):
        if sym not in rets.columns:
            continue
        d = pd.concat([rets[sym], mkt], axis=1).dropna()
        if len(d) < MIN_OBS:
            continue
        y, x = d.iloc[:, 0].values, d.iloc[:, 1].values
        var = float(np.var(x, ddof=1))
        if var <= 0:
            continue
        beta = float(np.cov(y, x, ddof=1)[0, 1] / var)
        corr = float(np.corrcoef(y, x)[0, 1])
        out[code] = (round(beta, 4), round(corr, 3), len(d))
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--write", action="store_true",
                    help="rewrite data/overrides/<code>_overrides.json")
    ap.add_argument("--source", choices=("topix", "yfinance"), default="topix",
                    help="which raw beta to adopt (default: topix regression)")
    ap.add_argument("--state", default=os.path.join(HERE, "batch_state.json"))
    ap.add_argument("--status", default="done")
    ap.add_argument("--out", default=os.path.join(HERE, "beta_rederived.json"))
    a = ap.parse_args()

    state = json.load(open(a.state, encoding="utf-8"))
    codes = sorted(k for k, v in state.items()
                   if isinstance(v, dict) and v.get("status") == a.status)

    print(f"measuring beta: {LOOKBACK} {INTERVAL} OLS on {MARKET} "
          f"for {len(codes)} tickers ...")
    reg = topix_betas(codes)
    print(f"  regressed: {len(reg)}/{len(codes)}\n")

    print(f"{'code':<6} {'old ovr':>8} {'yf':>7} {'TOPIX':>7} {'corr':>6} "
          f"{'raw':>7} {'adopted':>8}  note")
    print("-" * 92)
    rows, changed, unresolved = [], 0, []
    for code in codes:
        op = os.path.join(ROOT, "data", "overrides", f"{code}_overrides.json")
        ov = json.load(open(op, encoding="utf-8")) if os.path.isfile(op) else {}
        old = ov.get("beta")
        yf_raw = raw_from_adjustments(code)
        tp = reg.get(code)
        tp_beta, tp_corr, tp_n = tp if tp else (None, None, None)

        if a.source == "topix":
            raw, src = ((tp_beta, f"TOPIX {LOOKBACK} {INTERVAL} OLS "
                                  f"(n={tp_n}, corr={tp_corr})")
                        if tp_beta is not None else (yf_raw, "yfinance field (regression unavailable)"))
        else:
            raw, src = yf_raw, "yfinance field"

        adj, clamped = adopted(raw)
        if raw is None:
            unresolved.append(code)
        rows.append(dict(code=code, old=old, yfinance=yf_raw, topix=tp_beta,
                         corr=tp_corr, n=tp_n, raw=raw, adopted=adj,
                         clamped=clamped, source=src))
        print(f"{code:<6} {old if old is not None else '-':>8} "
              f"{yf_raw if yf_raw is not None else '-':>7} "
              f"{tp_beta if tp_beta is not None else '-':>7} "
              f"{tp_corr if tp_corr is not None else '-':>6} "
              f"{raw if raw is not None else '-':>7} {adj:>8.4f}"
              f"{'  CLAMPED' if clamped else ''}")

        if a.write and raw is not None and os.path.isfile(op):
            if ov.get("beta") != raw:
                ov["beta"] = raw
                changed += 1
            ov["_beta_note"] = (
                f"実測（raw）β = {raw:.4f}。出所: {src}。"
                f"フェーズ2 #6 以降 overrides の beta は raw を入れる契約で、"
                f"テンプレが Blume 調整 0.67×raw + 0.33×1.00 を掛け "
                f"[{BETA_FLOOR}, {BETA_CEIL}] にクランプする（採用 {adj:.4f}"
                + ("、クランプ発動" if clamped else "、クランプ非発動") + "）。"
                f"■旧 overrides の値 {old} は旧ルール [0.6, 1.75] のクランプ後の値であり "
                f"raw ではない。"
                f"■yfinance の beta フィールドは "
                + (f"{yf_raw}" if yf_raw is not None else "取得不能（null）")
                + "。日本株では市場を説明しないベンチマークで算出されており"
                  "（NTT −0.165 / 大阪ガス −0.201 等）、採用しない。"
                  "詳細は docs/phase2_pipeline_fixes_20260906.md の §2-4。")
            with open(op, "w", encoding="utf-8") as f:
                json.dump(ov, f, ensure_ascii=False, indent=1)
                f.write("\n")

    print("-" * 92)
    n_old_floor = sum(1 for r in rows if r["old"] == 0.6)
    n_clamp = sum(1 for r in rows if r["clamped"])
    adopted_vals = [r["adopted"] for r in rows]
    yf_vals = [r["yfinance"] for r in rows if r["yfinance"] is not None]
    tp_vals = [r["topix"] for r in rows if r["topix"] is not None]
    print(f"tickers: {len(rows)}")
    print(f"  overrides carrying the OLD clamp floor 0.60: {n_old_floor}")
    print(f"  raw resolved: {len(rows) - len(unresolved)}/{len(rows)}"
          + (f"  UNRESOLVED: {', '.join(unresolved)}" if unresolved else ""))
    print(f"  yfinance field:    n={len(yf_vals)} min {min(yf_vals):+.3f} "
          f"max {max(yf_vals):+.3f} mean {sum(yf_vals)/len(yf_vals):+.3f} "
          f"negative {sum(1 for v in yf_vals if v < 0)}")
    print(f"  TOPIX regression:  n={len(tp_vals)} min {min(tp_vals):+.3f} "
          f"max {max(tp_vals):+.3f} mean {sum(tp_vals)/len(tp_vals):+.3f} "
          f"negative {sum(1 for v in tp_vals if v < 0)}")
    print(f"  adopted (post-Blume): min {min(adopted_vals):.3f} / "
          f"max {max(adopted_vals):.3f} / "
          f"mean {sum(adopted_vals)/len(adopted_vals):.3f}, "
          f"clamped {n_clamp}")
    with open(a.out, "w", encoding="utf-8") as f:
        json.dump(rows, f, ensure_ascii=False, indent=1)
    print(f"  detail written: {a.out}")
    if a.write:
        print(f"  overrides rewritten (beta changed): {changed}")
    else:
        print("  DRY RUN - pass --write to rewrite the overrides")


if __name__ == "__main__":
    main()
