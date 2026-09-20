"""screener/report/s5_revision.py — シャドウG: S5 は上方修正を先行して当てるか。

定義の正本は docs/backtest_acceptance_criteria.md「シャドウG 事前登録」（2026-09-20・測定前に確定）。
**このスクリプトは事前登録を実装するだけで、窓・群・ベースレートを決めない。**

測るもの
--------
判定日ごとに S5 の状態（発火 / 非発火 / 評価不能 / guidance_dead）で銘柄を分け、
**その後 90 暦日以内に出た業績予想の修正開示の方向**（up / down / flat / unknown / none）を数える。
比較対象は**同じ判定日・同じ窓のユニバース全体**（ベースレート）。

これは記述統計であり、シグナルではない。スコア・採否・出口は変えない。

    python -m screener.report.s5_revision
"""
from __future__ import annotations

import argparse
import csv
import math
import os
import sqlite3
import sys
from collections import Counter, defaultdict
from datetime import date, timedelta

try:
    from screener import common as C
except ImportError:                                     # pragma: no cover
    sys.path.insert(0, os.path.dirname(os.path.dirname(
        os.path.dirname(os.path.abspath(__file__)))))
    from screener import common as C

from screener.report import backtest_eval as V1
from screener.signals.span_runner import score_ticker

# ---- G-1 / G-2 の値。測定後に動かさない ---------------------------------
WINDOW_DAYS = 90                 # 追跡窓（暦日）
FIRST_DATE = date(2026, 7, 23)   # TDnet アーカイブの開始
LAST_DATE = date(2026, 9, 18)    # 収集済みの最終日
POLICY = "evidence_strict"
OUTCOMES = ("up", "down", "flat", "unknown", "none")
GROUPS = ("S5発火(進捗超過)", "S5非発火", "S5評価不能", "ユニバース全体")


def scorers_all():
    root = str(V1._external_root())
    if root not in sys.path:
        sys.path.insert(0, root)
    from module_b.run_scorers import SCORERS_ALL
    return SCORERS_ALL


def judgment_dates(cal: list[str]) -> list[str]:
    """2026-07-23 以降の各週の月曜。休場なら翌営業日（G-1）。"""
    out, d = [], FIRST_DATE
    while d.weekday() != 0:                 # 最初の月曜まで進める
        d += timedelta(days=1)
    while d <= LAST_DATE:
        s = d.isoformat()
        nxt = [c for c in cal if c >= s]
        if nxt and nxt[0] <= LAST_DATE.isoformat():
            out.append(nxt[0])
        d += timedelta(days=7)
    return out


def revisions_by_code(mcon) -> dict:
    """(code) -> [(開示日, 方向)]。方向は通期営業利益の guidance から。"""
    out = defaultdict(list)
    rows = mcon.execute(
        "SELECT f.code, f.date, g.revision_direction "
        "FROM filings f LEFT JOIN guidance g "
        "  ON g.code=f.code AND g.date=f.date AND g.item='operating_income' "
        "     AND (g.q_no IS NULL OR g.q_no=4) "
        "WHERE f.source='tdnet' AND f.subtype='業績予想修正' AND f.code IS NOT NULL "
        "ORDER BY f.date").fetchall()
    for code, d, direction in rows:
        out[code].append((d, direction if direction in ("up", "down", "flat") else "unknown"))
    return out


def outcome_for(revs, code, as_of, horizon_end):
    """窓内の**最初の**修正開示の方向。無ければ none（G-1）。"""
    for d, direction in revs.get(code, ()):
        if as_of < d <= horizon_end:
            return direction
    return "none"


def s5_state(res):
    """S5 の状態と、その連続量（スコア）。"""
    s5 = res.get("S5") if isinstance(res, dict) else None
    if not isinstance(s5, dict) or not s5.get("available"):
        return "S5評価不能", None, False
    dead = bool((s5.get("details") or {}).get("guidance_dead") or s5.get("guidance_dead"))
    if (s5.get("score") or 0) > 0:
        return "S5発火(進捗超過)", s5.get("score"), dead
    return "S5非発火", s5.get("score"), dead


def collect(mcon, pcon, dates, codes, revs):
    """(判定日 × 銘柄) の観測を作る。"""
    scorers = scorers_all()
    last_day = LAST_DATE.isoformat()
    rows, n_err = [], 0
    for as_of in dates:
        horizon = (date.fromisoformat(as_of) + timedelta(days=WINDOW_DAYS)).isoformat()
        truncated = horizon > last_day
        for code in codes:
            try:
                res = score_ticker(pcon, code, scorers, as_of=as_of, policy=POLICY)
            except Exception:
                n_err += 1
                res = {}
            group, score, dead = s5_state(res)
            rows.append({
                "as_of": as_of, "code": code, "group": group, "s5_score": score,
                "guidance_dead": int(dead), "truncated": int(truncated),
                "window_days": min(WINDOW_DAYS,
                                   (date.fromisoformat(last_day)
                                    - date.fromisoformat(as_of)).days),
                "outcome": outcome_for(revs, code, as_of, min(horizon, last_day)),
            })
    return rows, n_err


# ------------------------------------------------------------------ 集計
def share(rows, outcome="up"):
    n = len(rows)
    return (sum(r["outcome"] == outcome for r in rows) / n) if n else None


def diff_ci(p1, n1, p0, n0):
    """2つの割合の差と95%信頼区間（二項の正規近似・G-4）。"""
    if not n1 or not n0 or p1 is None or p0 is None:
        return None, None, None
    se = math.sqrt(p1 * (1 - p1) / n1 + p0 * (1 - p0) / n0)
    d = p1 - p0
    return d, d - 1.96 * se, d + 1.96 * se


def counts_table(rows):
    c = Counter(r["outcome"] for r in rows)
    return [c[o] for o in OUTCOMES]


def render(rows, dates, codes, n_err):
    base = rows
    L = ["# シャドウG: S5 → 上方修正の先行力（記述統計・シグナルではない）", ""]
    L.append("- 判定日 %d 日（%s 〜 %s）/ ユニバース %d 銘柄 / のべ %d 観測（スコア例外 %d）"
             % (len(dates), dates[0], dates[-1], len(codes), len(rows), n_err))
    L.append("- 追跡窓 %d 暦日。収集の最終日 %s を越える窓は**途中打ち切り**として別に出す"
             % (WINDOW_DAYS, LAST_DATE.isoformat()))
    L.append("")
    L.append("## 群ごとの結果（のべ 銘柄×判定日）")
    L.append("")
    L.append("| 群 | n | " + " | ".join(OUTCOMES) + " | up率 | ベースレート差 [95%CI] |")
    L.append("|---|---|" + "---|" * (len(OUTCOMES) + 2))
    p0, n0 = share(base), len(base)
    for g in GROUPS:
        sub = base if g == "ユニバース全体" else [r for r in rows if r["group"] == g]
        if not sub:
            continue
        p1 = share(sub)
        d, lo, hi = diff_ci(p1, len(sub), p0, n0)
        ci = "–" if d is None or g == "ユニバース全体" else "%+.1fpt [%+.1f, %+.1f]" % (
            100 * d, 100 * lo, 100 * hi)
        L.append("| %s | %d | %s | %.1f%% | %s |"
                 % (g, len(sub), " | ".join(str(x) for x in counts_table(sub)),
                    100 * p1, ci))
    L.append("")
    L.append("## 窓が満了した観測だけ（打ち切りを除く）")
    L.append("")
    full = [r for r in rows if not r["truncated"]]
    if not full:
        L.append("**満了した観測は0件。** 収集開始が 2026-07-23 なので、90日窓が閉じる判定日が無い。")
        L.append("下の表は窓が %d〜%d 日で切れた観測での割合であり、**90日ぶんの確率ではない**。"
                 % (min(r["window_days"] for r in rows), max(r["window_days"] for r in rows)))
    else:
        p0f, n0f = share(full), len(full)
        L.append("| 群 | n | " + " | ".join(OUTCOMES) + " | up率 | ベースレート差 [95%CI] |")
        L.append("|---|---|" + "---|" * (len(OUTCOMES) + 2))
        for g in GROUPS:
            sub = full if g == "ユニバース全体" else [r for r in full if r["group"] == g]
            if not sub:
                continue
            p1 = share(sub)
            d, lo, hi = diff_ci(p1, len(sub), p0f, n0f)
            ci = "–" if d is None or g == "ユニバース全体" else "%+.1fpt [%+.1f, %+.1f]" % (
                100 * d, 100 * lo, 100 * hi)
            L.append("| %s | %d | %s | %.1f%% | %s |"
                     % (g, len(sub), " | ".join(str(x) for x in counts_table(sub)),
                        100 * p1, ci))
    L.append("")
    L.append("## 判定日ごとの up 率（単純平均とのべ割合の両方・G-2）")
    L.append("")
    per_day = []
    for d0 in dates:
        day = [r for r in rows if r["as_of"] == d0]
        fired = [r for r in day if r["group"] == "S5発火(進捗超過)"]
        per_day.append((d0, len(day), share(day), len(fired), share(fired),
                        day[0]["window_days"] if day else 0))
    L.append("| 判定日 | 窓(日) | ユニバース n | up率 | S5発火 n | up率 |")
    L.append("|---|---|---|---|---|---|")
    for d0, n, p, nf, pf, wd in per_day:
        L.append("| %s | %d | %d | %.1f%% | %d | %s |"
                 % (d0, wd, n, 100 * (p or 0), nf,
                    "–" if pf is None else "%.1f%%" % (100 * pf)))
    days_u = [p for _d, _n, p, _nf, _pf, _w in per_day if p is not None]
    days_f = [pf for _d, _n, _p, nf, pf, _w in per_day if pf is not None and nf]
    L.append("")
    L.append("- 判定日の単純平均: ユニバース %.1f%% / S5発火 %.1f%%"
             % (100 * sum(days_u) / len(days_u), 100 * sum(days_f) / len(days_f)))
    L.append("- のべ割合: ユニバース %.1f%% / S5発火 %.1f%%"
             % (100 * share(rows), 100 * (share([r for r in rows if r["group"] == "S5発火(進捗超過)"]) or 0)))
    L.append("")
    L.append("## S5 スコアの四分位（連続量として効いているか・G-4）")
    L.append("")
    fired = [r for r in rows if r["group"] == "S5発火(進捗超過)" and r["s5_score"] is not None]
    if fired:
        vals = sorted(r["s5_score"] for r in fired)
        qs = [vals[int(len(vals) * k / 4)] for k in (1, 2, 3)]
        L.append("| 四分位 | 範囲 | n | up率 |")
        L.append("|---|---|---|---|")
        edges = [(-1e9, qs[0]), (qs[0], qs[1]), (qs[1], qs[2]), (qs[2], 1e9)]
        for i, (lo, hi) in enumerate(edges, 1):
            sub = [r for r in fired if lo <= r["s5_score"] < hi or (i == 4 and r["s5_score"] >= lo)]
            if sub:
                L.append("| Q%d | %.2f〜%.2f | %d | %.1f%% |"
                         % (i, max(lo, min(vals)), min(hi, max(vals)), len(sub),
                            100 * share(sub)))
    L.append("")
    L.append("## guidance_dead（記録のみ・100件未満は判定不能）")
    L.append("")
    dead = [r for r in rows if r["guidance_dead"]]
    L.append("のべ %d 観測 / 銘柄ユニーク %d。" % (len(dead), len({r["code"] for r in dead})))
    if dead:
        L.append("内訳: " + " / ".join("%s %d" % (o, sum(r["outcome"] == o for r in dead))
                                      for o in OUTCOMES))
    L.append("**100件未満なら判定不能として記録する（§4 と同じ）。**")
    L.append("")
    L.append("## 銘柄ユニークで数え直す（G-4-5）")
    L.append("")
    L.append("| 群 | 銘柄ユニーク | うち窓内に up が1回でもあった銘柄 | 割合 |")
    L.append("|---|---|---|---|")
    for g in GROUPS:
        sub = rows if g == "ユニバース全体" else [r for r in rows if r["group"] == g]
        codes_g = {r["code"] for r in sub}
        up_codes = {r["code"] for r in sub if r["outcome"] == "up"}
        if codes_g:
            L.append("| %s | %d | %d | %.1f%% |"
                     % (g, len(codes_g), len(up_codes), 100 * len(up_codes) / len(codes_g)))
    return "\n".join(L) + "\n"


def main(argv=None) -> int:
    sys.stdout.reconfigure(encoding="utf-8")
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--db", default=C.DB_PATH)
    p.add_argument("--projection", default=os.path.join(C.DATA_DIR, "projection.db"))
    p.add_argument("--out", default=os.path.join(C.DATA_DIR, "s5_revision_report.md"))
    p.add_argument("--csv-out", default=os.path.join(C.DATA_DIR, "s5_revision_observations.csv"))
    a = p.parse_args(argv)

    mcon = sqlite3.connect("file:%s?mode=ro" % a.db.replace("\\", "/"), uri=True)
    pcon = sqlite3.connect("file:%s?mode=ro" % a.projection.replace("\\", "/"), uri=True)
    pcon.row_factory = sqlite3.Row          # 採点器は行を名前で引く（§38 で踏んだ）
    cal = [r[0] for r in mcon.execute("SELECT date FROM market_index ORDER BY date")]
    dates = judgment_dates(cal)
    codes = [r[0] for r in mcon.execute(
        "SELECT code FROM companies WHERE universe_flag=1 ORDER BY code")]
    C.log("シャドウG: 判定日 %d × 銘柄 %d を採点する" % (len(dates), len(codes)))
    rows, n_err = collect(mcon, pcon, dates, codes, revisions_by_code(mcon))
    with open(a.csv_out, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.DictWriter(fh, fieldnames=list(rows[0].keys()))
        w.writeheader()
        w.writerows(rows)
    text = render(rows, dates, codes, n_err)
    with open(a.out, "w", encoding="utf-8") as fh:
        fh.write(text)
    print(text)
    C.log("出力: %s / %s" % (a.out, a.csv_out))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
