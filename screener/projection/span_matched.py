"""screener/projection/span_matched.py — span-matched YoY モードの中核。

設計の正本: `docs/span_matched_design.md`（2026-09-02 承認・補強2点込み）。

何を解くのか
------------
四半期報告書の廃止で Q2・Q4 が半期粒度(span_q=2)でしか作れなくなり、
投影層がそれを落としていた結果、2025年以降は
**「四半期として渡さない」が「何も渡さない」になっていた**
（available シグナルが最大1本、B3 は評価可能 2.1%）。

中核の仮定
----------
前期比は「3ヶ月 vs 3ヶ月」でなくてよい。**「6ヶ月 vs 前年の同じ6ヶ月」**
でも意味は保たれる。壊れるのは**異なる長さを比べたとき**だけ。

比較の規則（規則B・2026-09-02 承認。設計書 §9）
------------------------------------------------
> **期ごとに最も細かい粒度の行を採り、常に同じ span の前年と比較する。**

当初は「四半期チェーンが5期連続で切れた時点で銘柄ごと切り替える」
（規則A）としていたが、実データは「Q1・Q2 は span=1 のまま、Q3 が欠け、
Q4 が span=2」という**年内混在**の形をしていて、欠け方が連続しないため
規則Aは永久に不発だった（設計書 §8-2 / §9-2、FY2025以降の比較可能ペアで
2,954 → 3,851）。モード判定は比較の**前提**をやめ、**記録**にした。

`mode` は「その比較が真の四半期(span=1)だったか否か」を意味する。
時系列でスコアを並べるときに粒度が混ざっていないかを人間が確認できる
ようにするための印であって、比較の可否を決めるものではない。

is_valid=0 への依存（設計書 §4-2）
----------------------------------
同じ span の前年値と比べるのは、**その2点のあいだに会社の形が変わって
いない**という前提に立つ。決算期変更・連結範囲変更・遡及修正はいずれも
既存の無効化機構が捕まえているので、**無効な期は比較相手にしない**。
比較間隔が長い（6ヶ月・12ヶ月）ぶん、quarter モードより依存度は強い。
"""
from __future__ import annotations

import os
import re
import sys

try:
    from screener import common as C            # noqa: F401  (呼び出し側の互換)
except ImportError:                             # pragma: no cover
    sys.path.insert(0, os.path.dirname(os.path.dirname(
        os.path.dirname(os.path.abspath(__file__)))))
    from screener import common as C            # noqa: F401

CHAIN_BREAK_LEN = 5          # 5期連続で真の四半期が無ければ切り替える
_RE_PE = re.compile(r"FY(\d{4})-Q([1-4])")


def parse_pe(period_end):
    """'FY2026-Q2' -> (2026, 2)。読めなければ None。"""
    m = _RE_PE.match(period_end or "")
    return (int(m.group(1)), int(m.group(2))) if m else None


def _seq(fy, q):
    """(年, 四半期) を通し番号に。連続性の判定に使う。"""
    return fy * 4 + (q - 1)


def switch_point(con, ticker):
    """真の四半期が途切れた通し番号。**規則B では比較に使わない**（診断用）。

    「この銘柄はいつ四半期開示をやめたか」は運用上は知りたい情報なので
    残す。規則Aではこれがモード切替の条件だったが、年内混在の形では
    永久に発火しないことが実測で分かったため、判定からは外した（§9-2）。

    真の四半期(span_q=1・有効)が **CHAIN_BREAK_LEN 期連続で欠けた**
    最初の時点を返す。以降その銘柄は span-matched モード。
    """
    rows = con.execute(
        "SELECT period_end, span_q, is_valid FROM quarterly_standalone_all "
        "WHERE ticker=? AND sales IS NOT NULL", (ticker,)).fetchall()
    have = set()
    allq = set()
    for r in rows:
        pe = parse_pe(r["period_end"])
        if not pe:
            continue
        n = _seq(*pe)
        allq.add(n)
        if (r["span_q"] or 1) == 1 and r["is_valid"]:
            have.add(n)
    if not allq:
        return None
    if not have:
        # **真の四半期が1本も無い銘柄。** 半期・通期しか開示していない
        # （2026-09-02 時点で 3,379 銘柄中 774 = 23%）。None を返すと
        # mode=quarter に落ちて span>=2 の行が全部捨てられ、この 774 銘柄は
        # S1/S2/S4/S5 から永久に見えなくなる。切れ目が「無い」のではなく
        # **最初から span-matched でしか見られない**が正しい。
        return min(allq)
    # **走査は最初の真の四半期以降から**。それ以前の空白は「チェーンが
    # 切れた」のではなく「まだ履歴が無い」だけ。ここを区別しないと、
    # 古い有報しか無い初期を切れ目と誤判定して全銘柄が FY2018 で
    # span-matched に落ちる（2026-09-02 に実際に踏んだ）。
    lo, hi = min(have), max(allq)
    run = 0
    for n in range(lo, hi + 1):
        if n in have:
            run = 0
        else:
            run += 1
            if run >= CHAIN_BREAK_LEN:
                return n - CHAIN_BREAK_LEN + 1     # 欠け始めた時点
    return None


def granularity_mode(span_q):
    """その比較の粒度。span=1 なら quarter、それ以外は span_matched。

    **規則B では比較の可否を決めない。**「何と何を比べたか」を後から
    層別するための記録。スコアを時系列で並べるとき、途中で粒度が
    変わっていないかはこれを見れば分かる。
    """
    return "quarter" if (span_q or 1) == 1 else "span_matched"


def mode_for(con, ticker, period_end, _cache={}):
    """互換用の別名。**規則B ではその期に実際に使う行の span で決まる**ので、
    行を持っている呼び出し側は `granularity_mode(span_q)` を直接使うこと。
    ここでは「その期で最も細かい粒度」を引いて判定する。
    """
    r = con.execute(
        "SELECT MIN(span_q) FROM quarterly_standalone_all WHERE ticker=? "
        "AND period_end=? AND is_valid=1 AND span_q<4", (ticker, period_end)).fetchone()
    return granularity_mode(r[0] if r and r[0] else 1)


def clear_cache():
    """規則Aのモードキャッシュの名残。規則Bでは持つ状態が無い。"""
    return None


def find_yoy_peer(con, ticker, period_end, span_q, allow_synth=True,
                  item=None):
    """同じ span・同じ起点の前年値を探す。**無ければ None（推測しない）。**

    1. 同じ (四半期, span) の1年前
    2. 見つからなければ「span は同じだが期末点が違う」候補のうち、
       **起点の四半期が一致するもの**だけ許す（決算期変更で q_no がずれる）
    3. それも無ければ **span=1 の行を足して合成**する（allow_synth）。
       合成した行は `synthesized=1` を持つ。開示された値そのものでは
       ないので、使う側は証拠にそう書く義務がある。
    4. どれも無ければ None

    無効化された期(is_valid=0)は比較相手にしない。

    `item` を渡すと **その項目が NULL の候補は相手にしない**。
    前年に「span は合っているが当該項目が NULL」の行が1本あるだけで
    合成に到達せず、比較そのものが落ちていた（3565 は FY2025-Q2 に
    operating_cf だけの span=2 行があり、sales が NULL だったために
    1年遡っていた。2026-09-02）。**行の存在と値の存在は別物。**
    """
    def usable(r):
        return r is not None and (item is None or r[item] is not None)
    pe = parse_pe(period_end)
    if not pe:
        return None
    fy, q = pe
    want = "FY%d-Q%d" % (fy - 1, q)
    r = con.execute(
        "SELECT * FROM quarterly_standalone_all WHERE ticker=? AND period_end=? "
        "AND span_q=? AND is_valid=1", (ticker, want, span_q)).fetchone()
    if usable(r):
        return r
    # 起点の四半期が一致するものだけ許す。期末点が違っても、同じ長さで
    # 同じところから始まっているなら季節性は揃う。
    start_q = max(1, q - span_q + 1)
    for cand in con.execute(
            "SELECT * FROM quarterly_standalone_all WHERE ticker=? AND span_q=? "
            "AND is_valid=1 AND period_end LIKE ?",
            (ticker, span_q, "FY%d-%%" % (fy - 1))):
        cpe = parse_pe(cand["period_end"])
        if not cpe:
            continue
        c_start = max(1, cpe[1] - span_q + 1)
        if c_start == start_q and usable(cand):
            return cand
    if allow_synth and span_q > 1:
        syn = synth_peer(con, ticker, fy - 1, start_q, q, span_q)
        if usable(syn):
            return syn
    return None


def synth_peer(con, ticker, fy, start_q, end_q, span_q):
    """前年の同じ区間を **span=1 の行を足して作る**。作れなければ None。

    なぜ要るか（2026-09-02 の実測）
    ------------------------------
    stale の発火が「**前年 H1 が span=1 でしか存在しない**」銘柄に全集中して
    いた。当期が半期(span=2)でも、前年 H1 が Q1・Q2 の四半期2本でしか無いと
    同じ span の相手が見つからず、組める最後の期まで1年遡っていた。
    Q1+Q2 は定義上 H1 そのものなので、足して作れば当期と組める。

    守ること
    --------
    - **区間を隙間なく覆う span=1 の行だけを使う。** 1本でも欠けたら作らない
      （欠けた四半期をゼロとして足すのは S5 で既に踏んだ誤り）。
    - 構成要素はすべて is_valid=1。無効な期を混ぜたら合成値も無効。
    - 項目ごとに独立に足す。1項目が欠けても他は作れる。**欠けた項目は
      None**（0 で埋めない）。
    - **合成であることを必ず持たせる**（synthesized / synth_parts）。
      開示された値そのものではないので、証拠にそう書けなければ出せない。
    """
    want = ["FY%d-Q%d" % (fy, i) for i in range(start_q, end_q + 1)]
    if len(want) != span_q:
        return None
    rows = {r["period_end"]: r for r in con.execute(
        "SELECT * FROM quarterly_standalone_all WHERE ticker=? AND span_q=1 "
        "AND is_valid=1 AND period_end IN (%s)" % ",".join("?" * len(want)),
        (ticker, *want))}
    if len(rows) != len(want):
        return None                     # 隙間があるなら作らない
    out = {"ticker": ticker, "fiscal_year": fy, "quarter_type": "Q%d" % end_q,
           "period_end": "FY%d-Q%d" % (fy, end_q), "span_q": span_q,
           "period_start": "FY%d-Q%d" % (fy, start_q), "is_valid": 1,
           "invalid_reason": None, "generation": 1,
           "synthesized": 1, "synth_parts": "+".join(want)}
    for item in ("sales", "operating_profit", "gross_profit", "operating_cf"):
        vals = [rows[k][item] for k in want]
        out[item] = sum(vals) if all(v is not None for v in vals) else None
    return out


def yoy_series(con, ticker, item, as_of_period=None):
    """(当期行, 前年行, mode) の並び。span-matched の入力そのもの。

    item は 'sales' / 'operating_profit' / 'gross_profit' / 'operating_cf'。
    比較できないものは**返さない**（0 を作らない）。
    """
    rows = con.execute(
        "SELECT * FROM quarterly_standalone_all WHERE ticker=? AND is_valid=1 "
        "AND %s IS NOT NULL AND span_q<4 ORDER BY period_end" % item,
        (ticker,)).fetchall()
    # **期ごとに最も細かい粒度を1本だけ採る**（規則B）。同じ期に span=1 と
    # span=2 の両方があるなら span=1 を使う —— 細かいほうが情報が多く、
    # 前年も同じ粒度で取れる可能性が高い。
    # span=4（通期）は四半期系列に入れない。年次の比較は別の話であり、
    # 混ぜると「前四半期比」の意味が壊れる。
    finest = {}
    for r in rows:
        pe, sp = r["period_end"], (r["span_q"] or 1)
        if pe not in finest or sp < (finest[pe]["span_q"] or 1):
            finest[pe] = r
    out = []
    for pe in sorted(finest):
        r = finest[pe]
        m = granularity_mode(r["span_q"])
        peer = find_yoy_peer(con, ticker, r["period_end"], r["span_q"] or 1,
                             item=item)
        if not peer or peer[item] is None:
            continue
        if as_of_period and r["period_end"] > as_of_period:
            continue
        synth = peer["synthesized"] if isinstance(peer, dict) else 0
        out.append({"period_end": r["period_end"], "span_q": r["span_q"] or 1,
                    "period_start": r["period_start"], "mode": m,
                    "value": r[item], "peer_period": peer["period_end"],
                    "peer_value": peer[item],
                    "peer_synthesized": 1 if synth else 0,
                    "peer_synth_parts": (peer.get("synth_parts")
                                         if isinstance(peer, dict) else None),
                    "yoy": (r[item] / peer[item] - 1) if peer[item] else None})
    return out


def coverage(con, limit=None):
    """モード別・span別にどれだけ比較可能になるかを測る。"""
    tickers = [r[0] for r in con.execute(
        "SELECT DISTINCT ticker FROM quarterly_standalone_all")]
    if limit:
        tickers = tickers[:limit]
    stat = {"tickers": len(tickers), "quarter": 0, "span_matched": 0,
            "by_span": {}, "switched": 0}
    for t in tickers:
        if switch_point(con, t) is not None:
            stat["switched"] += 1
        for row in yoy_series(con, t, "sales"):
            stat[row["mode"]] += 1
            stat["by_span"][row["span_q"]] = stat["by_span"].get(row["span_q"], 0) + 1
    return stat
