"""screener/report/revision_gap_5y.py — シャドウH-5y: 初動の分解を5年サンプルで。

定義の正本は docs/backtest_acceptance_criteria.md「シャドウH-5y」（2026-09-21 承認・登録）。
**H の定義は1文字も変えていない。** 変えたのは**イベント集合だけ**:

    H     : TDnet アーカイブ（2か月）の業績予想修正 61件
    H-5y  : `fins_summary` の修正イベント（`f_op` の変化）5年ぶん

量（ギャップ / 寄付後 / 寄付→T+1 / 初動）、T+0 の決め方、ベンチマーク（ユニバース中央値）、
値幅なしの扱いは H のまま。追加で **規模別・方向別・場中・年別**（H5-3）を出す。

    python -m screener.report.revision_gap_5y
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

from screener.report import revision_gap as H
from screener.technical import event_response as E

# H5-3 の規模区分。結果を見て動かさない。
SIZE_BINS = ((None, 0.10, "10%未満"), (0.10, 0.20, "10-20%"), (0.20, 0.30, "20-30%"),
             (0.30, 0.50, "30-50%"), (0.50, None, "50%+"))


def revision_events(con):
    """`f_op` が前の開示から変わった行（§52 J-1 と同じ）。開示時刻も持ち出す。"""
    out, last = [], {}
    for code, d, tm, fy, f_op in con.execute(
            "SELECT code, disc_date, disc_time, fy_end, f_op FROM fins_summary "
            "WHERE f_op IS NOT NULL AND fy_end IS NOT NULL "
            "ORDER BY disc_date, disc_time, code"):
        key = (code, fy)
        prev = last.get(key)
        last[key] = f_op
        if prev is None or prev == 0 or f_op == prev:
            continue
        out.append({"code": code, "date": d, "time": (tm or "")[:5],
                    "dir": "上方" if f_op > prev else "下方",
                    "size": (f_op - prev) / abs(prev), "fy_end": fy})
    return out


def build(con, cal, universe):
    """イベントに T+0 と時刻区分を付け、H と同じ分解を計算する。"""
    px = H.load_prices(con, universe, "2021-08-01")
    med = H.market_median(px, cal)
    evs = revision_events(con)
    rows, skip = [], Counter()
    for e in evs:
        if e["code"] not in universe:
            skip["ユニバース外"] += 1
            continue
        timing, t0 = E.timing_and_t0(e["date"], e["time"], cal)
        if not t0:
            skip["T+0 が窓の外"] += 1
            continue
        d = H.decompose(px, med, cal, e["code"], t0)
        if d is None:
            skip["株価が足りない"] += 1
            continue
        d.update({"code": e["code"], "date": e["date"], "time": e["time"],
                  "t0": t0, "timing": timing, "rev_dir": e["dir"],
                  "size": e["size"], "fy_end": e["fy_end"],
                  "year": t0[:4], "name": ""})
        rows.append(d)
    return rows, skip, len(evs)


def _table(title, groups, keys=("gap", "intraday", "open_to_t1", "total")):
    labels = {"gap": "ギャップ", "intraday": "寄付後", "open_to_t1": "寄付→T+1", "total": "初動"}
    L = ["### %s\n" % title, E.HDR]
    for name, sub in groups:
        for k in keys:
            L.append(E._stat_row("%s %s" % (name, labels[k]),
                                 E.describe([(e.get(k), e["t0"]) for e in sub])))
    L.append("")
    return L


def render(rows, skip, n_events, cover):
    ups = [r for r in rows if r["rev_dir"] == "上方"]
    downs = [r for r in rows if r["rev_dir"] == "下方"]
    L = ["# シャドウH-5y: 上方修正の初動の分解（5年・記録専用）", ""]
    L.append("- 修正イベント %s 件 → 分解できた **%s 件**（上方 %s / 下方 %s）"
             % (f"{n_events:,}", f"{len(rows):,}", f"{len(ups):,}", f"{len(downs):,}"))
    L.append("- 落とした内訳: " + " / ".join("%s %s" % (k, f"{v:,}") for k, v in skip.most_common()))
    L.append("- **株価のカバー率 %.1f%%**（イベント数ベース）" % (100 * cover))
    L.append("- 定義は H のまま（T+0・ベンチマーク=ユニバース中央値・値幅なしの扱い）")
    L.append("")
    L.append("## 全体（上方修正）")
    L.append("")
    L += _table("上方修正 %s 件" % f"{len(ups):,}", [("", ups)])
    tradable = [r for r in ups if not r["no_range"]]
    L.append("値幅なし（高値=安値）を除くと %s 件" % f"{len(tradable):,}")
    L.append("")
    L += _table("値幅なしを除く", [("", tradable)], keys=("intraday", "open_to_t1"))
    L.append("## 開示時刻の区分別（上方修正）")
    L.append("")
    groups = [(t, [r for r in ups if r["timing"] == t])
              for t in ("引け後", "場中", "場外", "不明")]
    L += _table("区分別", [(t, s) for t, s in groups if s])
    L.append("**場中開示では始値が開示より前の価格**なので、寄付後の値は「反応して取れる分」ではない。")
    L.append("日中足が無いので開示時刻より後だけを取り出すことは**できない**（§48）。")
    L.append("")
    L.append("## 修正の規模別（上方修正・H5-3 の1）")
    L.append("")
    L += _table("規模別", [(lab, [r for r in ups
                                 if (lo is None or r["size"] >= lo)
                                 and (hi is None or r["size"] < hi)])
                        for lo, hi, lab in SIZE_BINS])
    L.append("## 方向別（H5-3 の2）")
    L.append("")
    L += _table("方向別", [("上方", ups), ("下方", downs)], keys=("gap", "intraday"))
    L.append("## 年別（H5-3 の4・上方修正の中央値）")
    L.append("")
    L.append("| 年 | 件数 | ギャップ | 寄付後 | 初動 |")
    L.append("|---|---|---|---|---|")
    for y in sorted({r["year"] for r in ups}):
        sub = [r for r in ups if r["year"] == y]
        def med(k):
            st = E.describe([(e.get(k), e["t0"]) for e in sub])
            return "–" if not st.get("n") else "%+.2f%%" % (100 * st["median"])
        L.append("| %s | %d | %s | %s | %s |" % (y, len(sub), med("gap"),
                                                 med("intraday"), med("total")))
    L.append("")
    L.append("## 寄り付きの内訳（上方修正）")
    L.append("")
    L.append("| 区分 | 件数 |")
    L.append("|---|---|")
    L.append("| 値幅なし（寄らず・比例配分を含む） | %d |" % sum(r["no_range"] for r in ups))
    L.append("| 始値=高値 | %d |" % sum(r["open_is_high"] for r in ups))
    bad = [r for r in ups
           if abs((1 + r["gap"]) * (1 + r["open_to_t1"]) - (1 + r["total"])) > 1e-9]
    L.append("")
    L.append("検算（ギャップ × 寄付→T+1 = 初動）不一致 %d 件" % len(bad))
    return "\n".join(L) + "\n"


def main(argv=None) -> int:
    sys.stdout.reconfigure(encoding="utf-8")
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--db", default=C.DB_PATH)
    p.add_argument("--out", default=os.path.join(C.DATA_DIR, "revision_gap_5y_report.md"))
    p.add_argument("--csv-out", default=os.path.join(C.DATA_DIR, "revision_gap_5y.csv"))
    a = p.parse_args(argv)

    con = sqlite3.connect("file:%s?mode=ro" % a.db.replace("\\", "/"), uri=True)
    cal = [r[0] for r in con.execute(
        "SELECT date FROM market_index WHERE date >= '2021-08-01' ORDER BY date")]
    universe = {r[0] for r in con.execute(
        "SELECT code FROM companies WHERE universe_flag=1")}
    names = dict(con.execute("SELECT code, name FROM companies"))
    C.log("シャドウH-5y: ユニバース %d / カレンダー %d 日" % (len(universe), len(cal)))
    rows, skip, n_events = build(con, cal, universe)
    for r in rows:
        r["name"] = names.get(r["code"], "")
    C.log("分解できたイベント %d / 全イベント %d" % (len(rows), n_events))
    with open(a.csv_out, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.DictWriter(fh, fieldnames=list(rows[0].keys()))
        w.writeheader()
        w.writerows(rows)
    text = render(rows, skip, n_events, len(rows) / n_events if n_events else 0)
    with open(a.out, "w", encoding="utf-8") as fh:
        fh.write(text)
    print(text)
    C.log("出力: %s / %s" % (a.out, a.csv_out))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
