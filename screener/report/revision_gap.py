"""screener/report/revision_gap.py — シャドウH: 上方修正の初動をギャップと寄付後に分ける。

定義の正本は docs/backtest_acceptance_criteria.md「シャドウH」（2026-09-20・測定前に確定）。
**このスクリプトは定義を実装するだけで、分け方やベンチマークを決めない。**

    ギャップ   = 始値(T+0) / 終値(T−1) − 1
    寄付後     = 終値(T+0) / 始値(T+0) − 1
    寄付→翌引け = 終値(T+1) / 始値(T+0) − 1
    初動(既報) = 終値(T+1) / 終値(T−1) − 1

ベンチマークは**ユニバース中央値**（TOPIX は終値しか無く、ギャップと寄付後に分けられない）。

これは記述統計であり、シグナルではない。

    python -m screener.report.revision_gap
"""
from __future__ import annotations

import argparse
import csv
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

from screener.technical import event_response as E

MEASURES = (("gap", "ギャップ 前日終値→始値"),
            ("intraday", "寄付後 始値→終値"),
            ("open_to_t1", "寄付→T+1終値"),
            ("total", "初動 前日終値→T+1終値（②と同じ）"))


def load_prices(mcon, codes, date_from):
    """code -> {date: (始値, 高値, 安値, 終値)}（すべて調整後）。"""
    px = defaultdict(dict)
    for code, d, o, h, l, c, ac in mcon.execute(
            "SELECT code, date, open, high, low, close, adj_close FROM prices "
            "WHERE date >= ? AND adj_close IS NOT NULL AND close IS NOT NULL "
            "AND open IS NOT NULL", (date_from,)):
        if code not in codes or not c:
            continue
        k = ac / c
        px[code][d] = (o * k, h * k, l * k, ac)
    return px


def market_median(px, cal):
    """各営業日の (ギャップ中央値, 寄付後中央値)。前日と当日に値がある銘柄だけ。"""
    out = {}
    for i in range(1, len(cal)):
        d, prev = cal[i], cal[i - 1]
        gaps, intra = [], []
        for series in px.values():
            a, b = series.get(prev), series.get(d)
            if not a or not b or not a[3] or not b[0]:
                continue
            gaps.append(b[0] / a[3] - 1)
            intra.append(b[3] / b[0] - 1)
        out[d] = (E.percentile(sorted(gaps), 0.5) if gaps else None,
                  E.percentile(sorted(intra), 0.5) if intra else None)
    return out


def decompose(px, med, cal, code, t0):
    """1イベントの分解。取れない値は None。"""
    i = cal.index(t0) if t0 in cal else None
    if i is None or i == 0 or i + 1 >= len(cal):
        return None
    prev, cur, nxt = cal[i - 1], cal[i], cal[i + 1]
    s = px.get(code) or {}
    a, b, c = s.get(prev), s.get(cur), s.get(nxt)
    if not a or not b or not c or not a[3] or not b[0]:
        return None
    mg, mi = med.get(cur, (None, None))
    gap = b[0] / a[3] - 1
    intra = b[3] / b[0] - 1
    row = {
        "prev_close": a[3], "open": b[0], "high": b[1], "low": b[2],
        "close": b[3], "next_close": c[3],
        "gap": gap, "intraday": intra,
        "open_to_t1": c[3] / b[0] - 1,
        "total": c[3] / a[3] - 1,
        "gap_rel": None if mg is None else gap - mg,
        "intraday_rel": None if mi is None else intra - mi,
        "no_range": int(b[1] == b[2]),                 # 1本値（寄らず・比例配分を含む）
        "open_is_high": int(b[0] == b[1]),
    }
    row["open_to_t1_rel"] = (None if mg is None or mi is None
                             else row["open_to_t1"] - mi)     # 参考（T+1 は補正しない）
    return row


def stat_rows(events, key, label):
    st = E.describe([(e.get(key), e["t0"]) for e in events])
    return E._stat_row(label, st)


def render(events, meta):
    L = ["# シャドウH: 上方修正の初動の分解（記述統計・シグナルではない）", ""]
    for k, v in meta.items():
        L.append("- %s: %s" % (k, v))
    ups = [e for e in events if e["rev_dir"] == "上方"]
    tradable = [e for e in ups if not e["no_range"]]
    L.append("")
    L.append("## 上方修正 %d 件（値幅なし %d 件を除くと %d 件）"
             % (len(ups), len(ups) - len(tradable), len(tradable)))
    L.append("")
    L.append(E.HDR)
    for key, label in MEASURES:
        L.append(stat_rows(ups, key, label))
    L.append(stat_rows(ups, "gap_rel", "ギャップ（対ユニバース中央値）"))
    L.append(stat_rows(ups, "intraday_rel", "寄付後（対ユニバース中央値）"))
    L.append("")
    L.append("### 値幅なし（1本値）を除いた寄付起点")
    L.append("")
    L.append(E.HDR)
    for key, label in (("intraday", "寄付後 始値→終値"),
                       ("open_to_t1", "寄付→T+1終値"),
                       ("intraday_rel", "寄付後（対ユニバース中央値）")):
        L.append(stat_rows(tradable, key, label))
    L.append("")
    L.append("## 開示時刻の区分別（上方修正）")
    L.append("")
    for timing in ("引け後", "場中", "場外", "不明"):
        sub = [e for e in ups if e["timing"] == timing]
        if not sub:
            continue
        L.append("### %s（%d件・うち値幅なし %d）"
                 % (timing, len(sub), sum(e["no_range"] for e in sub)))
        L.append("")
        L.append(E.HDR)
        for key, label in MEASURES:
            L.append(stat_rows(sub, key, label))
        L.append("")
    L.append("## 参考: 下方・不明")
    L.append("")
    L.append(E.HDR)
    for d in ("下方", "不明"):
        sub = [e for e in events if e["rev_dir"] == d]
        for key, label in (("gap", "ギャップ"), ("intraday", "寄付後")):
            L.append(stat_rows(sub, "%s" % key, "%s %s" % (d, label)))
    L.append("")
    L.append("## 値幅・寄り付きの内訳（上方修正）")
    L.append("")
    L.append("| 区分 | 件数 |")
    L.append("|---|---|")
    L.append("| 値幅なし（高値=安値。寄らず・比例配分を含む） | %d |"
             % sum(e["no_range"] for e in ups))
    L.append("| 始値=高値（寄り天の形） | %d |" % sum(e["open_is_high"] for e in ups))
    L.append("| 始値=高値 かつ 寄付後がマイナス | %d |"
             % sum(1 for e in ups if e["open_is_high"] and (e["intraday"] or 0) < 0))
    L.append("| 出来高0の T+0 | %d |" % sum(1 for e in ups if e.get("zero_volume")))
    L.append("")
    L.append("## 検算: ギャップ × 寄付後 = 当日、×翌日 = 初動")
    L.append("")
    bad = [e for e in ups
           if abs((1 + e["gap"]) * (1 + e["open_to_t1"]) - (1 + e["total"])) > 1e-9]
    L.append("不一致 %d 件（許容 1e-9）" % len(bad))
    return "\n".join(L) + "\n"


def main(argv=None) -> int:
    sys.stdout.reconfigure(encoding="utf-8")
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--db", default=C.DB_PATH)
    p.add_argument("--scored", default=os.path.join(
        C.DATA_DIR, "event_price_response_scored.csv"))
    p.add_argument("--events", default=os.path.join(
        C.DATA_DIR, "event_price_response_events.csv"))
    p.add_argument("--out", default=os.path.join(C.DATA_DIR, "revision_gap_report.md"))
    p.add_argument("--csv-out", default=os.path.join(C.DATA_DIR, "revision_gap.csv"))
    p.add_argument("--from", dest="date_from", default="2026-06-01")
    a = p.parse_args(argv)

    dirs = {}
    with open(a.scored, encoding="utf-8-sig") as fh:
        for r in csv.DictReader(fh):
            if r["type"] == E.TYPE_REV:
                dirs[r["event_id"]] = r["rev_dir"]
    events = []
    with open(a.events, encoding="utf-8-sig") as fh:
        for r in csv.DictReader(fh):
            if r["type"] != E.TYPE_REV or r["window_ok"] != "1":
                continue
            r["rev_dir"] = dirs.get(r["event_id"], "不明")
            events.append(r)

    mcon = sqlite3.connect("file:%s?mode=ro" % a.db.replace("\\", "/"), uri=True)
    cal = [r[0] for r in mcon.execute(
        "SELECT date FROM market_index WHERE date >= ? ORDER BY date", (a.date_from,))]
    codes = {e["code"] for e in events}
    px_all = load_prices(mcon, None if False else
                         {r[0] for r in mcon.execute(
                             "SELECT code FROM companies WHERE universe_flag=1")},
                         a.date_from)
    med = market_median(px_all, cal)
    vol = {}
    for code, d, v in mcon.execute(
            "SELECT code, date, volume FROM prices WHERE date >= ?", (a.date_from,)):
        if code in codes:
            vol[(code, d)] = v

    out = []
    for e in events:
        d = decompose(px_all, med, cal, e["code"], e["t0"])
        if d is None:
            continue
        d.update({k: e[k] for k in ("event_id", "code", "name", "timing", "t0",
                                    "date", "time", "rev_dir")})
        d["zero_volume"] = int(not vol.get((e["code"], e["t0"])))
        out.append(d)
    with open(a.csv_out, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.DictWriter(fh, fieldnames=list(out[0].keys()))
        w.writeheader()
        w.writerows(out)
    meta = {"イベント（業績予想の修正・窓あり）": len(events),
            "分解できた": len(out),
            "方向": "上方 %d / 下方 %d / 不明 %d" % (
                sum(e["rev_dir"] == "上方" for e in out),
                sum(e["rev_dir"] == "下方" for e in out),
                sum(e["rev_dir"] == "不明" for e in out)),
            "ベンチマーク": "ユニバース中央値（始値/前日終値・終値/始値）"}
    text = render(out, meta)
    with open(a.out, "w", encoding="utf-8") as fh:
        fh.write(text)
    print(text)
    C.log("出力: %s / %s" % (a.out, a.csv_out))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
