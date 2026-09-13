"""screener/signals/freshness.py — シグナル鮮度（stale）の判定。**表示のみ。**

なぜ要るのか
------------
2026-09-02 に人間が原文精読して見つけた、スクリーンの最大の欠陥。

    3565 アセンテック: 契約負債(S2)の根拠期は FY2025-Q2（開示 2024-09-11）。
    その後 FY2026-Q2 まで開示があり、実際の最新期では契約負債は横ばい。
    **約2年前の証拠で発火し続け、週次スコア 0.822 の第2位に載っていた。**

原因は `find_yoy_peer` が「同じ span の前年」を要求することにある。
前年が別の粒度でしか無いと当期は組めず、**組める最後の期まで遡る**。
遡ること自体は正しい（間違った相手と比べるよりよい）が、**遡ったことを
黙っているのが誤り**だった。`MAX_STALE_Q` で4期を超える組は落としているが、
4期以内なら「1年前の証拠」でも今の証拠として出てしまう。

スコアは変えない
----------------
stale の減点・除外は**値の変更**（§8 の区別）にあたるので、ここでは
一切行わない。まず「どれだけ古い証拠に依存していたか」を測る。
将来やるなら別途事前登録する。
"""
from __future__ import annotations

import os
import sys

try:
    from screener import common as C            # noqa: F401
except ImportError:                             # pragma: no cover
    sys.path.insert(0, os.path.dirname(os.path.dirname(
        os.path.dirname(os.path.abspath(__file__)))))
    from screener import common as C            # noqa: F401

from screener.projection import span_matched as SM
from screener.signals import span_scorers as SP


def latest_period(con, ticker, as_of=None):
    """as_of 時点で**パース済みの最新期**。無ければ None。

    「開示された最新期」ではなく「こちらが数字として持っている最新期」。
    比較の相手はこちらなので、持っていない期を基準にすると
    「取り込めていない」を「証拠が新しい」と誤って読むことになる。
    """
    vis = SP.visible_periods(con, ticker, as_of)
    best = None
    for (pe,) in con.execute(
            "SELECT DISTINCT period_end FROM quarterly_standalone_all "
            "WHERE ticker=? AND is_valid=1 AND sales IS NOT NULL", (ticker,)):
        if vis is not None and pe not in vis:
            continue
        if not SM.parse_pe(pe):
            continue
        if best is None or pe > best:
            best = pe
    return best


def lag_quarters(period, latest):
    """period が latest より何四半期古いか。判定できなければ None。

    **定義（2026-09-02 明文化）**: lag = 期インデックスの差
    （latest の通し番号 − basis の通し番号）。単位は四半期。
    通し番号は `span_matched._seq(fy, q) = fy*4 + (q-1)`。

        basis=FY2025-Q2, latest=FY2026-Q2 → 8105 − 8101 = 4
        basis=FY2025-Q2, latest=FY2027-Q1 → 8108 − 8101 = 7

    span の長さは見ない。**「どれだけ前の期の話か」であって
    「何ヶ月ぶんの期か」ではない。** 半期(span=2)の行でも、その期末が
    1四半期ぶん前なら lag=1 になる。
    """
    a, b = SM.parse_pe(latest or ""), SM.parse_pe(period or "")
    if not a or not b:
        return None
    return SM._seq(*a) - SM._seq(*b)


def annotate(scores, con, ticker, as_of=None):
    """各シグナルに `stale` / `stale_lag` / `latest_period` を付ける。

    scores は別枠形の {signal_id: result_dict}。**スコアには触らない。**
    根拠期を持たないシグナル（別枠の位置ベース比較）は判定不能なので
    `stale=None` にする —— 0（新しい）と混ぜない。
    """
    latest = latest_period(con, ticker, as_of)
    for v in scores.values():
        if not isinstance(v, dict) or not v.get("available"):
            continue
        per = v.get("period")
        if not per or not latest:
            v["stale"], v["stale_lag"] = None, None
        else:
            lag = lag_quarters(per, latest)
            v["stale"] = None if lag is None else int(lag > 0)
            v["stale_lag"] = lag
        v["latest_period"] = latest
    return latest


def summarize(scores, fired=None):
    """行レベルの要約。(stale_flag, stale_lag, 明細文字列) を返す。

    stale_flag は「**発火したシグナルのうち1本でも古いものがあるか**」。
    発火していないシグナルの鮮度は推薦の根拠ではないので数えない。
    """
    detail, flag, worst = [], 0, 0
    for k, v in sorted(scores.items()):
        if not (isinstance(v, dict) and v.get("available")):
            continue
        if fired is not None and k not in fired:
            continue
        st, lag = v.get("stale"), v.get("stale_lag")
        if st is None:
            detail.append("%s:判定不能(根拠期なし)" % k)
        elif st:
            flag = 1
            worst = max(worst, lag or 0)
            detail.append("%s:%s(最新%s より%d期古い)"
                          % (k, v.get("period"), v.get("latest_period"), lag))
        else:
            detail.append("%s:%s(最新)" % (k, v.get("period")))
    return flag, worst, " ".join(detail)
