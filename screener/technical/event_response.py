"""screener/technical/event_response.py — 開示イベントに対する価格・出来高の反応の記録（シャドウE フェーズA）。

定義の正本は docs/backtest_acceptance_criteria.md「シャドウE 事前登録」E-A。
**このモジュールは定義を実装するだけで、定義を決めない。** 結果を見て種別・窓・
基準期間を動かすことは禁止されている。記録は calibration_backlog §37。

これは事後の記述統計であり、シグナルではない。スコア・採否・エグジットには
接続しない（パッケージ docstring の E-0 と test_technical_isolation.py）。

Point-in-Time
-------------
イベント単位の出力のうち `f_*`（特徴量）は T−1 の終値までの配列だけを受け取る
関数で計算する（`features()` には T+0 以降の要素を渡さない）。T+0 以降は
`y_*`（結果変数）としてのみ使う。`d_*` は標本全体を見て決まる記述用の列で、
特徴量ではない（例: 種別内の上位20%判定は後のイベントも含めた分位点を使う）。

株価の連鎖
----------
リターンは生の終値と `adj_factor`（分割日の係数。2:1 分割なら 0.5）から自前で
連鎖させる。`adj_close` は取得時点で遡及調整されるため、取得日の違う行が混ざると
分割をまたいで不連続になりうる（1447 の事例、tasks/todo.md 2026-09-19）。
出来高も同じ係数で分割前の株数単位に揃える。

    python -m screener.technical.event_response
    python -m screener.technical.event_response --report-out C:\\screener_data\\event_price_response_report.md
"""
from __future__ import annotations

import argparse
import bisect
import csv
import math
import os
import re
import sqlite3
import sys
from collections import Counter, defaultdict
from pathlib import Path

# ------------------------------------------------------------------ 定義（E-A。動かさない）
TYPE_Q = "決算短信_四半期"
TYPE_FY = "決算短信_本決算"
TYPE_REV = "業績予想修正"
TYPE_DIV = "配当予想修正"
TYPE_ORDER = "受注関連"
TYPE_OTHER = "その他"
PRIMARY_TYPES = (TYPE_Q, TYPE_FY, TYPE_REV, TYPE_DIV, TYPE_ORDER)

# 新しい情報の開示ではない行（訂正・期中レビュー完了に伴う再掲・補足資料・差替・再掲載・XBRLデータ追加）。
# 2026-09-19 追記（事前登録 A-1 の補正。経緯は backtest_acceptance_criteria.md E-A）
RESTATE_WORDS = ("訂正", "期中レビュー", "補足", "差替", "再掲載", "データ追加")
QUARTER_RE = re.compile(r"第\s*[1-3１-３一二三]\s*四半期|中間")
ORDER_EXCLUDE = ("訂正", "借入", "金銭消費貸借", "当座", "融資", "コミットメント", "社債",
                 "株式譲渡", "株式取得", "子会社", "合併", "分割", "株式交換", "公開買付",
                 "資本業務提携", "第三者割当", "新株予約権", "終了", "解除", "解約")
ORDER_STRONG = ("受注", "大口", "落札", "受託")
ORDER_CONTRACT_OBJECT = ("供給", "販売", "納入", "ライセンス", "業務委託", "売買", "取引基本")

MARKET_OPEN = "09:00"
MARKET_CLOSE = "15:30"   # 2024-11-05 以降の東証の大引け。15:30 ちょうどの開示は引け後

TIMING_INTRADAY = "場中"
TIMING_AFTER = "引け後"
TIMING_OFF = "場外"
TIMING_UNKNOWN = "不明"

PRE_DAYS = 15            # ① T−15〜T−1
BASE_OFFSET = -16        # 窓の基準日（①の起点の前日）
POST_DAYS = 20           # T+20
VOL_BASE = (-36, -17)    # 出来高倍率（主）の基準期間。窓の外に固定
TRAIL = 20               # 出来高倍率（副）= 直近20日平均
TOP_PRE_PCT = 0.80       # ⑤ 種別内で①が80パーセンタイル以上
MIN_N = 30               # 1群30件未満は「件数不足」と注記

INTERVALS = {            # 名前: (a, b) → P(b)/P(a−1) − B(b)/B(a−1)
    "pre": (-15, -1),
    "init": (0, 1),
    "drift": (2, 20),
    "drift5": (2, 5),
    "drift10": (2, 10),
}


# ------------------------------------------------------------------ A-1 分類
def is_order_related(title: str | None) -> bool:
    """受注関連（暫定・タイトルのみ）。E-A の判定ルールそのまま。"""
    t = title or ""
    if any(w in t for w in ORDER_EXCLUDE):
        return False
    if any(w in t for w in ORDER_STRONG):
        return True
    return ("契約" in t and "締結" in t
            and any(w in t for w in ORDER_CONTRACT_OBJECT))


def classify(subtype: str | None, title: str | None) -> str:
    t = title or ""
    if any(w in t for w in RESTATE_WORDS):
        return TYPE_OTHER
    if subtype == "決算短信":
        return TYPE_Q if QUARTER_RE.search(t) else TYPE_FY
    if subtype == "業績予想修正":
        return TYPE_REV
    if subtype == "配当予想修正":
        return TYPE_DIV
    if subtype is None and is_order_related(t):
        return TYPE_ORDER
    return TYPE_OTHER


def other_bucket(subtype: str | None, title: str | None) -> str:
    """その他の内訳（件数のみ）。"""
    if any(w in (title or "") for w in RESTATE_WORDS):
        return "訂正・再掲"
    if subtype:
        return subtype
    return "非決算の適時開示"


def _hhmm(s: str | None) -> str | None:
    if not s:
        return None
    m = re.match(r"\s*(\d{1,2}):(\d{2})", s)
    return "%02d:%s" % (int(m.group(1)), m.group(2)) if m else None


def timing_and_t0(d: str, time_str: str | None, cal: list[str]) -> tuple[str, str | None]:
    """開示日・時刻 → (区分, T+0)。T+0 は終値が開示を初めて織り込みうる営業日。

    cal は昇順の営業日。T+0 がカレンダーの末尾を越えるなら None。
    """
    hm = _hhmm(time_str)
    i = bisect.bisect_left(cal, d)
    is_td = i < len(cal) and cal[i] == d
    if is_td:
        nxt = cal[i + 1] if i + 1 < len(cal) else None
    else:
        nxt = cal[i] if i < len(cal) else None
    if hm is None:
        return TIMING_UNKNOWN, nxt
    if not is_td:
        return TIMING_OFF, nxt
    if hm < MARKET_OPEN:
        return TIMING_OFF, d
    if hm < MARKET_CLOSE:
        return TIMING_INTRADAY, d
    return TIMING_AFTER, nxt


def build_events(raw: list[dict], cal: list[str]) -> tuple[list[dict], Counter]:
    """raw: {code, date, time, subtype, title, source} の列 → (主種別イベント, その他の内訳件数)。

    同一 (銘柄, 種別, T+0) は最も早い開示を1件に。同一 (銘柄, T+0) の別種別は concurrent。
    """
    best: dict[tuple, dict] = {}
    other = Counter()
    for r in raw:
        typ = classify(r["subtype"], r["title"])
        if typ == TYPE_OTHER:
            other[other_bucket(r["subtype"], r["title"])] += 1
            continue
        timing, t0 = timing_and_t0(r["date"], r["time"], cal)
        key = (r["code"], typ, t0)
        ev = dict(r, type=typ, timing=timing, t0=t0)
        cur = best.get(key)
        if cur is None or (ev["date"], _hhmm(ev["time"]) or "99:99") < (cur["date"], _hhmm(cur["time"]) or "99:99"):
            best[key] = ev
    events = sorted(best.values(), key=lambda e: (e["t0"] or "9999", e["code"], e["type"]))
    by_ct = defaultdict(set)
    for e in events:
        by_ct[(e["code"], e["t0"])].add(e["type"])
    for e in events:
        e["concurrent"] = "|".join(sorted(by_ct[(e["code"], e["t0"])] - {e["type"]}))
        e["event_id"] = "%s_%s_%s" % (e["code"], e["type"], e["t0"])
    return events, other


# ------------------------------------------------------------------ 価格パネル
def build_series(rows: list[tuple], cal_idx: dict[str, int], n_cal: int) -> dict:
    """1銘柄の行 (date, close, volume, turnover, adj_factor) → カレンダーに揃えた系列。

    px: 連鎖価格指数（初日=1）。売買の無い日は直前値、上場前・最終行より後は None。
    vol: 分割前の株数単位に揃えた出来高（売買の無い日は 0）。close は生終値。
    """
    px = [None] * n_cal
    vol = [None] * n_cal
    tv = [None] * n_cal
    close = [None] * n_cal
    got = {}
    for d, c, v, t, f in rows:
        i = cal_idx.get(d)
        if i is not None and c is not None:
            got[i] = (c, v, t, f)
    if not got:
        return {"px": px, "vol": vol, "tv": tv, "close": close}
    lo, hi = min(got), max(got)
    p = cum = prev_c = None
    for i in range(lo, hi + 1):
        if i in got:
            c, v, t, f = got[i]
            f = f if f else 1.0
            if prev_c is None:
                p, cum = 1.0, 1.0
            else:
                p = p * c / (prev_c * f)
                cum *= f
            prev_c = c
            px[i], close[i] = p, c
            vol[i] = (v or 0.0) * cum
            tv[i] = t or 0.0
        else:
            px[i], close[i], vol[i], tv[i] = p, prev_c, 0.0, 0.0
    return {"px": px, "vol": vol, "tv": tv, "close": close}


def universe_median_index(panel: dict, n_cal: int, traded: dict) -> list[float]:
    """ユニバースの日次リターン中央値を累積した指数（副ベンチマーク）。

    その日と前日の両方に約定がある銘柄だけを使う（売買の無い日の 0 リターンで薄めない）。
    """
    idx = [1.0] * n_cal
    for i in range(1, n_cal):
        rs = []
        for code, s in panel.items():
            t = traded[code]
            if i in t and (i - 1) in t:
                rs.append(s["px"][i] / s["px"][i - 1] - 1.0)
        med = percentile(sorted(rs), 0.5) if rs else 0.0
        idx[i] = idx[i - 1] * (1.0 + med)
    return idx


# ------------------------------------------------------------------ A-2 窓
def _rel(px, bm, a_i, b_i):
    """P(b)/P(a−1) − B(b)/B(a−1)。インデックスは配列上の位置。"""
    if a_i - 1 < 0 or b_i >= len(px):
        return None
    p0, p1, b0, b1 = px[a_i - 1], px[b_i], bm[a_i - 1], bm[b_i]
    if None in (p0, p1, b0, b1) or p0 == 0 or b0 == 0:
        return None
    return p1 / p0 - b1 / b0


def _mean(xs):
    return sum(xs) / len(xs) if xs else None


def vol_base(vol, idx0):
    lo, hi = idx0 + VOL_BASE[0], idx0 + VOL_BASE[1]
    if lo < 0:
        return None
    w = vol[lo:hi + 1]
    if len(w) != hi - lo + 1 or any(v is None for v in w):
        return None
    m = _mean(w)
    return m if m and m > 0 else None


def features(px_pre, vol_pre, topix_pre, univ_pre, idx0):
    """特徴量。**引数は T−1 までで切った配列**（長さ idx0）。T+0 以降は構造的に見えない。"""
    assert len(px_pre) == idx0
    a, b = idx0 + INTERVALS["pre"][0], idx0 + INTERVALS["pre"][1]
    base = vol_base(vol_pre, idx0)
    ratios = []
    if base:
        for k in range(-PRE_DAYS, 0):
            v = vol_pre[idx0 + k] if idx0 + k >= 0 else None
            ratios.append(v / base if v is not None else None)
    rr = [x for x in ratios if x is not None]
    raw = None
    if a - 1 >= 0 and px_pre[a - 1] and px_pre[b] is not None:
        raw = px_pre[b] / px_pre[a - 1] - 1.0
    return {
        "f_pre_rel_topix": _rel(px_pre, topix_pre, a, b),
        "f_pre_rel_univ": _rel(px_pre, univ_pre, a, b),
        "f_pre_raw": raw,
        "f_vol_base": base,
        "f_vol_ratio_pre_max": max(rr) if rr else None,
        "f_vol_ratio_pre_mean": _mean(rr),
        "f_vol_ratio_tm1": ratios[-1] if ratios else None,
    }


def outcomes(px, vol, topix, univ, idx0, base):
    out = {}
    for name in ("init", "drift", "drift5", "drift10"):
        a, b = INTERVALS[name]
        out["y_%s_rel_topix" % name] = _rel(px, topix, idx0 + a, idx0 + b)
        out["y_%s_rel_univ" % name] = _rel(px, univ, idx0 + a, idx0 + b)
    post = []
    for k in range(0, POST_DAYS + 1):
        i = idx0 + k
        v = vol[i] if i < len(vol) else None
        post.append(v / base if (base and v is not None) else None)
    pp = [x for x in post if x is not None]
    out["y_vol_ratio_t0"] = post[0]
    out["y_vol_ratio_t1"] = post[1] if len(post) > 1 else None
    out["y_vol_ratio_post_max"] = max(pp) if pp else None
    return out


def window_rows(ev, s, topix, univ, idx0, cal, base):
    """long 形式の行（T−16〜T+20）。累積相対は基準日 T−16 から。"""
    rows = []
    bi = idx0 + BASE_OFFSET
    for k in range(BASE_OFFSET, POST_DAYS + 1):
        i = idx0 + k
        if i < 0 or i >= len(cal) or s["px"][i] is None:
            continue
        prev = i - 1
        ret = (s["px"][i] / s["px"][prev] - 1.0) if prev >= 0 and s["px"][prev] else None
        t_ret = topix[i] / topix[prev] - 1.0 if prev >= 0 and topix[prev] and topix[i] else None
        u_ret = univ[i] / univ[prev] - 1.0 if prev >= 0 else None
        trail = None
        if i - TRAIL >= 0:
            w = s["vol"][i - TRAIL:i]
            if all(v is not None for v in w):
                m = _mean(w)
                trail = s["vol"][i] / m if m else None
        rows.append({
            "event_id": ev["event_id"], "code": ev["code"], "type": ev["type"],
            "timing": ev["timing"], "t0": ev["t0"], "offset": k, "date": cal[i],
            "role": "feature" if k < 0 else "outcome",
            "close": s["close"][i], "px_index": s["px"][i],
            "turnover_value": s["tv"][i], "volume_adj": s["vol"][i],
            "ret": ret, "ret_topix": t_ret, "ret_univ_median": u_ret,
            "cum_rel_topix": _rel(s["px"], topix, bi + 1, i) if i > bi else None,
            "cum_rel_univ": _rel(s["px"], univ, bi + 1, i) if i > bi else None,
            "vol_ratio_base": (s["vol"][i] / base) if base else None,
            "vol_ratio_trail20": trail,
        })
    return rows


def measure(events, panel, topix, univ, cal):
    """各イベントの f_* / y_* と long 行。窓が組めないイベントは f/y が None のまま残す。"""
    cal_idx = {d: i for i, d in enumerate(cal)}
    ev_rows, long_rows = [], []
    for ev in events:
        rec = dict(ev)
        idx0 = cal_idx.get(ev["t0"]) if ev["t0"] else None
        s = panel.get(ev["code"])
        if idx0 is None or s is None or idx0 + BASE_OFFSET < 0 or s["px"][idx0 + BASE_OFFSET] is None:
            rec["window_ok"] = 0
            ev_rows.append(rec)
            continue
        rec["window_ok"] = 1
        f = features(s["px"][:idx0], s["vol"][:idx0], topix[:idx0], univ[:idx0], idx0)
        rec.update(f)
        rec.update(outcomes(s["px"], s["vol"], topix, univ, idx0, f["f_vol_base"]))
        ev_rows.append(rec)
        long_rows.extend(window_rows(ev, s, topix, univ, idx0, cal, f["f_vol_base"]))
    return ev_rows, long_rows


# ------------------------------------------------------------------ A-3 集計
def percentile(sorted_vals, q):
    """線形補間の分位点（numpy の既定と同じ）。"""
    n = len(sorted_vals)
    if n == 0:
        return None
    pos = (n - 1) * q
    lo = int(math.floor(pos))
    hi = min(lo + 1, n - 1)
    return sorted_vals[lo] + (sorted_vals[hi] - sorted_vals[lo]) * (pos - lo)


def _t(xs):
    n = len(xs)
    if n < 2:
        return None
    m = sum(xs) / n
    var = sum((x - m) ** 2 for x in xs) / (n - 1)
    return m / math.sqrt(var / n) if var > 0 else None


def describe(pairs):
    """pairs: [(値, T+0)] → n / 中央値 / 四分位 / 平均 / 正の割合 / t（素朴・日付クラスタ）。"""
    vals = sorted(v for v, _ in pairs if v is not None)
    n = len(vals)
    if n == 0:
        return {"n": 0}
    by_d = defaultdict(list)
    for v, d in pairs:
        if v is not None:
            by_d[d].append(v)
    cl = [sum(x) / len(x) for x in by_d.values()]
    return {
        "n": n, "median": percentile(vals, 0.5), "q1": percentile(vals, 0.25),
        "q3": percentile(vals, 0.75), "mean": sum(vals) / n,
        "pos": sum(v > 0 for v in vals) / n, "t": _t(vals),
        "n_dates": len(cl), "t_cluster": _t(cl),
    }


def aggregate(ev_rows, long_rows):
    ok = [e for e in ev_rows if e.get("window_ok")]
    res = {"stats": {}, "sign": {}, "volume": {}, "pre_top": {}, "repeats": []}
    for scope in ("all", "solo"):
        for typ in PRIMARY_TYPES:
            es = [e for e in ok if e["type"] == typ and (scope == "all" or not e["concurrent"])]
            st = {}
            for key in ("f_pre_rel_topix", "y_init_rel_topix", "y_drift_rel_topix",
                        "y_drift5_rel_topix", "y_drift10_rel_topix",
                        "f_pre_rel_univ", "y_init_rel_univ", "y_drift_rel_univ"):
                st[key] = describe([(e.get(key), e["t0"]) for e in es])
            res["stats"][(scope, typ)] = st
            sg = {}
            for lab, cond in (("init>0", lambda x: x > 0), ("init<=0", lambda x: x <= 0)):
                sub = [e for e in es if e.get("y_init_rel_topix") is not None and cond(e["y_init_rel_topix"])]
                sg[lab] = describe([(e.get("y_drift_rel_topix"), e["t0"]) for e in sub])
            res["sign"][(scope, typ)] = sg
    # ④ 出来高倍率（主）の相対日ごとの中央値
    prof = defaultdict(lambda: defaultdict(list))
    for r in long_rows:
        if r["vol_ratio_base"] is not None and r["offset"] >= -PRE_DAYS:
            prof[r["type"]][r["offset"]].append(r["vol_ratio_base"])
    for typ in PRIMARY_TYPES:
        med = {k: percentile(sorted(v), 0.5) for k, v in prof[typ].items()}
        pre = {k: v for k, v in med.items() if k < 0}
        post = {k: v for k, v in med.items() if k >= 0}
        res["volume"][typ] = {
            "profile": med, "n_by_offset": {k: len(v) for k, v in prof[typ].items()},
            "peak_pre": max(pre, key=pre.get) if pre else None,
            "peak_post": max(post, key=post.get) if post else None,
        }
    # ⑤ 事前に動いた（種別内で①が上位20%）
    flagged = []
    for typ in PRIMARY_TYPES:
        es = [e for e in ok if e["type"] == typ and e.get("f_pre_rel_topix") is not None]
        if not es:
            continue
        thr = percentile(sorted(e["f_pre_rel_topix"] for e in es), TOP_PRE_PCT)
        hit = [e for e in es if e["f_pre_rel_topix"] >= thr]
        for e in hit:
            e["d_pre_top20"] = 1
        for e in es:
            e.setdefault("d_pre_top20", 0)
        res["pre_top"][typ] = {"threshold": thr, "events": sorted(hit, key=lambda e: -e["f_pre_rel_topix"])}
        flagged.extend(hit)
    by_code = defaultdict(list)
    for e in flagged:
        by_code[e["code"]].append(e)
    for code, es in sorted(by_code.items()):
        if len({e["t0"] for e in es}) >= 2:
            res["repeats"].append((code, sorted(es, key=lambda e: e["t0"])))
    return res


# ------------------------------------------------------------------ 報告
def _pct(x):
    return "–" if x is None else "%+.2f%%" % (100 * x)


def _num(x, fmt="%.2f"):
    return "–" if x is None else fmt % x


def _stat_row(label, st):
    if not st or st.get("n", 0) == 0:
        return "| %s | 0 | – | – | – | – | – | – | – |" % label
    note = " ※件数不足" if st["n"] < MIN_N else ""
    return "| %s | %d%s | %s | %s | %s | %s | %.0f%% | %s | %s (%d日) |" % (
        label, st["n"], note, _pct(st["median"]), _pct(st["q1"]), _pct(st["q3"]),
        _pct(st["mean"]), 100 * st["pos"], _num(st["t"]), _num(st["t_cluster"]), st["n_dates"])


HDR = ("| 種別 | n | 中央値 | Q1 | Q3 | 平均 | 正の割合 | t(素朴) | t(日付クラスタ) |\n"
       "|---|---|---|---|---|---|---|---|---|")


def render(meta, events, other, ev_rows, res, names):
    L = []
    L.append("# シャドウE フェーズA —— イベント→価格の反応（記述統計・シグナルではない）\n")
    for k, v in meta.items():
        L.append("- %s: %s" % (k, v))
    L.append("\n## A-1. イベント件数（ユニバース内・重複除去後）\n")
    L.append("| 種別 | 件数 | 場中 | 引け後 | 場外 | 不明 | 同時開示あり | 窓あり(①) | ②あり | ③あり(T+20) |")
    L.append("|---|---|---|---|---|---|---|---|---|---|")
    for typ in PRIMARY_TYPES:
        es = [e for e in ev_rows if e["type"] == typ]
        tm = Counter(e["timing"] for e in es)
        L.append("| %s | %d | %d | %d | %d | %d | %d | %d | %d | %d |" % (
            typ, len(es), tm[TIMING_INTRADAY], tm[TIMING_AFTER], tm[TIMING_OFF], tm[TIMING_UNKNOWN],
            sum(bool(e["concurrent"]) for e in es),
            sum(e.get("f_pre_rel_topix") is not None for e in es),
            sum(e.get("y_init_rel_topix") is not None for e in es),
            sum(e.get("y_drift_rel_topix") is not None for e in es)))
    L.append("\nその他（件数のみ・行数）: " + " / ".join("%s %d" % kv for kv in other.most_common())
             + " / 計 %d" % sum(other.values()))
    blocks = (("① 事前 T−15〜T−1（対TOPIX）", "f_pre_rel_topix"),
              ("② 初動 T+0〜T+1（対TOPIX）", "y_init_rel_topix"),
              ("③ ドリフト T+2〜T+20（対TOPIX）", "y_drift_rel_topix"),
              ("③' 部分ドリフト T+2〜T+5（対TOPIX・補助）", "y_drift5_rel_topix"),
              ("③' 部分ドリフト T+2〜T+10（対TOPIX・補助）", "y_drift10_rel_topix"),
              ("① 事前（対ユニバース中央値・副）", "f_pre_rel_univ"),
              ("② 初動（対ユニバース中央値・副）", "y_init_rel_univ"),
              ("③ ドリフト（対ユニバース中央値・副）", "y_drift_rel_univ"))
    for scope, lab in (("all", "全件"), ("solo", "単独（同時開示なし）")):
        L.append("\n## A-3. 累積相対リターン —— %s\n" % lab)
        for title, key in blocks:
            L.append("### %s\n" % title)
            L.append(HDR)
            for typ in PRIMARY_TYPES:
                L.append(_stat_row(typ, res["stats"][(scope, typ)][key]))
            L.append("")
        L.append("### ③'' 初動の符号別ドリフト T+2〜T+20（対TOPIX・%s）\n" % lab)
        L.append(HDR)
        for typ in PRIMARY_TYPES:
            for sg in ("init>0", "init<=0"):
                L.append(_stat_row("%s ②%s" % (typ, sg[4:]), res["sign"][(scope, typ)][sg]))
        L.append("")
    L.append("\n## A-3 ④. 出来高倍率（主: 基準 T−36〜T−17 平均）の相対日別中央値\n")
    offs = [-15, -10, -5, -3, -2, -1, 0, 1, 2, 3, 5, 10, 20]
    L.append("| 種別 | 事前ピーク | 事後ピーク | " + " | ".join("T%+d" % k for k in offs) + " |")
    L.append("|---|---|---|" + "---|" * len(offs))
    for typ in PRIMARY_TYPES:
        v = res["volume"][typ]
        pr = v["profile"]
        pk = lambda k: "–" if k is None else "T%+d (%.2f倍)" % (k, pr[k])
        L.append("| %s | %s | %s | %s |" % (typ, pk(v["peak_pre"]), pk(v["peak_post"]),
                                           " | ".join(_num(pr.get(k)) for k in offs)))
    L.append("\n（各相対日の n は種別ごとに異なる。T+20 側は T+0 が新しいイベントほど欠ける）")
    L.append("\n## A-3 ⑤. 事前に動いていたイベント（種別内で①が上位20%）\n")
    L.append("上位20%の閾値は**標本全体**の分位点（事後の記述。特徴量ではない）。\n")
    for typ in PRIMARY_TYPES:
        pt = res["pre_top"].get(typ)
        if not pt:
            continue
        L.append("### %s（閾値 ①≥%s、%d件）\n" % (typ, _pct(pt["threshold"]), len(pt["events"])))
        L.append("| コード | 銘柄 | T+0 | 区分 | ① | ② | ③ | 出来高 T−1 倍率 | 同時開示 |")
        L.append("|---|---|---|---|---|---|---|---|---|")
        for e in pt["events"]:
            L.append("| %s | %s | %s | %s | %s | %s | %s | %s | %s |" % (
                e["code"], names.get(e["code"], ""), e["t0"], e["timing"], _pct(e["f_pre_rel_topix"]),
                _pct(e.get("y_init_rel_topix")), _pct(e.get("y_drift_rel_topix")),
                _num(e.get("f_vol_ratio_tm1")), e["concurrent"] or ""))
        L.append("")
    L.append("### 異なる T+0 で2回以上該当した銘柄\n")
    if not res["repeats"]:
        L.append("なし")
    for code, es in res["repeats"]:
        L.append("- %s %s: " % (code, names.get(code, "")) + " / ".join(
            "%s %s ①%s" % (e["t0"], e["type"], _pct(e["f_pre_rel_topix"])) for e in es))
    return "\n".join(L) + "\n"


# ------------------------------------------------------------------ 入出力
LONG_COLS = ["event_id", "code", "type", "timing", "t0", "offset", "date", "role", "close",
             "px_index", "turnover_value", "volume_adj", "ret", "ret_topix", "ret_univ_median",
             "cum_rel_topix", "cum_rel_univ", "vol_ratio_base", "vol_ratio_trail20"]
EVENT_COLS = ["event_id", "code", "name", "type", "timing", "date", "time", "t0", "concurrent",
              "source", "title", "window_ok",
              "f_pre_rel_topix", "f_pre_rel_univ", "f_pre_raw", "f_vol_base",
              "f_vol_ratio_pre_max", "f_vol_ratio_pre_mean", "f_vol_ratio_tm1",
              "y_init_rel_topix", "y_init_rel_univ", "y_drift_rel_topix", "y_drift_rel_univ",
              "y_drift5_rel_topix", "y_drift5_rel_univ", "y_drift10_rel_topix", "y_drift10_rel_univ",
              "y_vol_ratio_t0", "y_vol_ratio_t1", "y_vol_ratio_post_max", "d_pre_top20"]


def _write_csv(path, cols, rows):
    with open(path, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.DictWriter(fh, fieldnames=cols, extrasaction="ignore")
        w.writeheader()
        for r in rows:
            w.writerow({k: ("" if r.get(k) is None else r.get(k)) for k in cols})


def connect_ro(path):
    """読み取り専用で開く（E-0: このパッケージは DB に書かない）。"""
    return sqlite3.connect(Path(path).resolve().as_uri() + "?mode=ro", uri=True)


def load(con, price_from):
    cal = [r[0] for r in con.execute("SELECT date FROM market_index ORDER BY date")]
    topix_by = dict(con.execute("SELECT date, close FROM market_index"))
    uni = {r[0]: r[1] for r in con.execute("SELECT code, name FROM companies WHERE universe_flag=1")}
    raw = []
    for code, d, dat, st, title in con.execute(
            "SELECT code, date, disclosed_at, subtype, title FROM filings WHERE source='tdnet'"):
        if code in uni:
            day, _, tm = (dat or d or "").partition(" ")
            raw.append({"code": code, "date": day or d, "time": tm or None, "subtype": st,
                        "title": title, "source": "filings"})
    for code, d, tm, title in con.execute(
            "SELECT code, date, time, title FROM disclosure_titles WHERE subtype IS NULL"):
        if code in uni:
            raw.append({"code": code, "date": d, "time": tm, "subtype": None,
                        "title": title, "source": "disclosure_titles"})
    rows = defaultdict(list)
    for code, d, c, v, t, f in con.execute(
            "SELECT code, date, close, volume, turnover_value, adj_factor FROM prices "
            "WHERE date >= ? ORDER BY code, date", (price_from,)):
        if code in uni:
            rows[code].append((d, c, v, t, f))
    return cal, topix_by, uni, raw, rows


def run(db_path, long_out, events_out, report_out, price_from="2026-04-01"):
    con = connect_ro(db_path)
    cal_all, topix_by, names, raw, rows = load(con, price_from)
    con.close()
    cal = [d for d in cal_all if d >= price_from]
    cal_idx = {d: i for i, d in enumerate(cal)}
    topix = [topix_by[d] for d in cal]
    panel, traded = {}, {}
    for code, rs in rows.items():
        panel[code] = build_series(rs, cal_idx, len(cal))
        traded[code] = {cal_idx[d] for d, c, *_ in rs if c is not None and d in cal_idx}
    univ = universe_median_index(panel, len(cal), traded)
    events, other = build_events(raw, cal)
    ev_rows, long_rows = measure(events, panel, topix, univ, cal)
    for e in ev_rows:
        e["name"] = names.get(e["code"], "")
    res = aggregate(ev_rows, long_rows)
    fil = [r for r in raw if r["source"] == "filings"]
    tit = [r for r in raw if r["source"] == "disclosure_titles"]
    meta = {
        "営業日カレンダー（TOPIX）": "%s 〜 %s（%d日）" % (cal[0], cal[-1], len(cal)),
        "ユニバース": "%d社（実行時点の universe_flag=1）" % len(names),
        "filings(TDnet)": "%d行 %s 〜 %s" % (len(fil), min(r["date"] for r in fil), max(r["date"] for r in fil)),
        "disclosure_titles(非決算)": "%d行 %s 〜 %s" % (len(tit), min(r["date"] for r in tit), max(r["date"] for r in tit)),
        "株価パネル": "%d銘柄（%s 以降）" % (len(panel), price_from),
        "イベント（主種別）": "%d件" % len(events),
    }
    _write_csv(long_out, LONG_COLS, long_rows)
    _write_csv(events_out, EVENT_COLS, ev_rows)
    text = render(meta, events, other, ev_rows, res, names)
    if report_out:
        with open(report_out, "w", encoding="utf-8") as fh:
            fh.write(text)
    return text, ev_rows, long_rows, res


def main(argv=None) -> int:
    sys.stdout.reconfigure(encoding="utf-8")
    from screener import common as C
    p = argparse.ArgumentParser(description="シャドウE フェーズA: イベント→価格の反応記録（測定のみ）")
    p.add_argument("--db", default=C.DB_PATH)
    p.add_argument("--out", default=os.path.join(C.DATA_DIR, "event_price_response.csv"))
    p.add_argument("--events-out", default=os.path.join(C.DATA_DIR, "event_price_response_events.csv"))
    p.add_argument("--report-out", default=os.path.join(C.DATA_DIR, "event_price_response_report.md"))
    a = p.parse_args(argv)
    text, ev_rows, long_rows, _ = run(a.db, a.out, a.events_out, a.report_out)
    print(text)
    print("出力: %s（%d行） / %s（%d行） / %s" % (a.out, len(long_rows), a.events_out, len(ev_rows), a.report_out))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
