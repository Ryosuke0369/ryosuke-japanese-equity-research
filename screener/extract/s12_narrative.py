"""screener/extract/s12_narrative.py — S12: 定性文言diffシグナル（ルールベース v1）。

設計は事前登録済み。**重みは初期値であり、評価結果を見て同一データで
再調整することは禁止**（変更するなら新規事前登録）。辞書は
`config/s12_dictionary.yaml` に外出ししてある。

入力
----
EDINET の同一スパンの前期・当期ペア（140→前年同 q_no の140、160→前年160、
120→前年120）。決算期変更・is_valid=0 をまたぐペアは**構築しない**。
ペアが無い期間は `available=0`（欠損）を返す。**0点と混同しない。**

前処理
------
1. `0102010_honbun` から「経営成績の分析」本文のみを切り出す。
   経済環境の一般論・中計方針の引用・「重要な変更はありません」・
   財政状態/CF分析は除外する。
2. 行分割された数字断片（「652\\n億\\n82\\n百万円」型）を結合してから
   正規表現を当てる。結合しないと金額のパターンが全部素通りする。
3. 「報告セグメントの変更」を検出したら `segment_changed=1`。その期は
   セグメント文言（Tier B）の評価を無効化し、会社レベルの総括文だけ使う。

スコアリング
------------
Tier A 硬い事実(±0.30) / B セグメント方向(±0.15) / C 語彙diff(±0.15)、
合成は ±0.30 で clamp。Tier D はフラグのみで**加点しない**。
各ヒットは `s12_evidence` に (filing_id, section, マッチ文, tier, 点数) を保存。
**根拠文の無い点数は存在してはならない。**

    python -m screener.extract.s12_narrative --all
    python -m screener.extract.s12_narrative --code 1301
    python -m screener.extract.s12_narrative --report
"""
from __future__ import annotations

import argparse
import os
import re
import sqlite3
import sys
import unicodedata
import zipfile

try:
    from screener import common as C
except ImportError:                                     # pragma: no cover
    sys.path.insert(0, os.path.dirname(os.path.dirname(
        os.path.dirname(os.path.abspath(__file__)))))
    from screener import common as C

_CFG = None

# 本文中の節見出し。「経営成績」の節だけを採り、次の見出しで切る。
#
# 書類種別で構造が違う（2026-09-02 に実測して判明）:
#   四半期報告書(140): 「(1) 経営成績の分析」
#   半期報告書(160)  : 「２【経営者による財政状態、経営成績及びキャッシュ・
#                       フローの状況の分析】→ (1) 財政状態及び経営成績の状況
#                       → (経営成績の状況)」
# 160 を拾えず available が 55% に留まっていた。内側の見出しほど正確なので、
# 見つかった中で**最も内側（最後に出る）**開始位置を採る。
_SEC_STARTS = (
    re.compile(r"[(（]\s*経営成績(?:等)?(?:の状況|の分析|の概況)?\s*[)）]"),
    re.compile(r"[(（]?\s*[1１一]\s*[)）]?\s*(?:【)?\s*"
               r"(?:経営成績(?:等)?(?:に関する説明|の分析|の状況|の概況)"
               r"|業績(?:の概況|の状況))"),
    re.compile(r"[(（]?\s*[1１一]\s*[)）]?\s*(?:【)?\s*"
               r"財政状態(?:及び|、)経営成績(?:及び|、)?(?:の)?状況"),
    re.compile(r"経営者による財政状態[、,]\s*経営成績[^】]{0,20}分析\s*[】]?"),
)
# 経営成績パートの終わり。財政状態・CF・次の番号付き見出しのどれか最初。
_SEC_ENDS = (
    re.compile(r"[(（]\s*財政状態(?:の状況)?\s*[)）]"),
    re.compile(r"[(（]\s*キャッシュ・フロー"),
    re.compile(r"[(（]?\s*[2２二]\s*[)）]\s*(?:【)?\s*"
               r"(?:財政状態|キャッシュ・フロー|資産|経営方針|優先的に)"),
    re.compile(r"[(（]?\s*[2２二]\s*[)）]?\s*(?:【)?\s*"
               r"(財政状態|キャッシュ・フロー|資産|当[中四半期]{1,4}の財政状態)"),
)
# 行に分割された数字断片。「652 / 億 / 82 / 百万円」を1行に畳む。
_RE_FRAGMENT = re.compile(r"^[\d,.　 ]{0,12}(?:億|百万|千|万|円|%|％|倍|ポイント)?$")


def cfg():
    global _CFG
    if _CFG is None:
        _CFG = C.load_yaml("s12_dictionary.yaml")
    return _CFG


def nfkc(s):
    return unicodedata.normalize("NFKC", s or "")


def join_number_fragments(text):
    """行分割された数字断片を結合する。

    PDF/HTML 由来の本文は「652\\n億\\n82\\n百万円」のように数量が行で
    バラける。結合しないと金額を含むパターンが全部素通りするので、
    正規表現を当てる前に必ず通す。
    """
    out = []
    for ln in text.split("\n"):
        t = ln.strip()
        if not t:
            out.append("")
            continue
        if out and out[-1] and _RE_FRAGMENT.match(t) and len(t) <= 12:
            out[-1] = out[-1] + t          # 直前の行に畳む
        else:
            out.append(t)
    return "\n".join(out)


def extract_section(text):
    """「経営成績の分析」本文のみを返す。見つからなければ None。"""
    t = join_number_fragments(text)
    # 候補となる開始位置を集める。**行頭にあるものだけ**を見出しとみなす。
    # 行の途中に出るのは「『経営者による財政状態、経営成績…の分析』中の
    # 会計上の見積り」のような**参照文**で、これを見出しと取ると本文が
    # 30字しか取れない（9252 / 7279 で実際に起きた）。
    cands = []
    for rx in _SEC_STARTS:
        for m in rx.finditer(t):
            at_line_head = (m.start() == 0) or (t[m.start() - 1] == chr(10))
            if at_line_head:
                cands.append(m.end())
    if not cands:
        return None

    def _cut(pos):
        b = t[pos:]
        end = None
        for rx in _SEC_ENDS:
            e = rx.search(b)
            if e and (end is None or e.start() < end):
                end = e.start()
        if end is not None:
            b = b[:end]
        for pat in cfg().get("exclude_sections") or []:
            x = re.search(pat, b)
            if x:
                b = b[:x.start()]
        return b.strip()

    # 内側の見出しほど正確だが、短すぎるものは誤爆。**本文が最も長く取れる
    # 候補**を採る。外側から始めても _SEC_ENDS で切れるので巻き込みは限定的。
    best = max((_cut(p) for p in sorted(set(cands))), key=len, default="")
    return best or None



def sentences(body):
    """文に割る。根拠として1文を保存するため。"""
    return [s.strip() for s in re.split(r"(?<=。)\s*", body or "") if s.strip()]


# ------------------------------------------------------------ 段落の種類
# **マクロ経済の定型文を採点根拠にしてはいけない。** 会社固有の情報が
# 何も入っていないうえ、ほぼ全社が同じ文言を書くので、差分を取っても
# 「その年に流行した言い回し」しか出てこない。2026-09-02 に人間が
# 原文精読して確認した実例（ベステラ短信 P.2、四半期報告書にも転載）:
#
#   「当第１四半期連結累計期間におけるわが国経済は、雇用・所得環境の改善や
#     各種政策の効果により緩やかな回復基調を維持しております」
#
# ここから「改善」「回復」が Tier C で加点されていた。同様に 3415 の
# 「堅調」、9692 の cost_pressure(-0.06) もすべて経済環境段落由来だった。
#
# **業界環境の段落は残す。** 「当社グループの属する解体・メンテナンス業界では
# 〜需要が」は会社の置かれた状況を語っており、会社固有の情報である。
# マクロと業界を区別することがこの修正の核心で、両方落とすなら
# シグナルの中身が痩せるだけになる。
_MACRO_SUBJ = re.compile(
    r"(わが国|我が国|国内|世界|海外|米国|中国|欧州|日本)(の)?(経済|景気)"
    r"|世界経済|国内経済|海外経済|日本経済"
    r"|金融市場|為替(相場|市場)|消費者物価|物価上昇"
    r"|雇用・所得環境|各種政策の効果|政府の経済対策"
    r"|通商政策|地政学的(な)?リスク|インフレ(率|懸念)")
_INDUSTRY_SUBJ = re.compile(
    r"(当社|当社グループ|当グループ)(の属する|が属する)"
    r"|(業界|市場)(に|で|では|におきまして|については)"
    r"|当業界|同業界|需要(は|が)(堅調|旺盛|低調)")
# 会社固有の主語。マクロ語と同居しても、こちらがあれば会社の話。
_COMPANY_SUBJ = re.compile(
    r"当社|当社グループ|当グループ|当第|セグメント|事業(は|が|に|の)"
    r"|売上高|営業利益|受注|製品|サービス|顧客|子会社")


def paragraphs(body):
    """本文を段落に割る。改行が段落境界。"""
    return [p.strip() for p in (body or "").split(chr(10)) if p.strip()]


def classify_paragraph(text):
    """macro / industry / company を返す。

    会社固有の主語があれば company を優先する —— マクロ語を1つ含むだけで
    落とすと、「当社の主力製品の受注が回復した。為替は円安に推移した」の
    ような混在段落まで捨ててしまう。
    """
    if not text:
        return "company"
    if _INDUSTRY_SUBJ.search(text) and not _MACRO_SUBJ.search(text):
        # 「当社グループの属する〜業界では」は会社固有の主語を含むが、
        # 語っているのは業界環境。監査のために industry と記録する
        # （採点対象である点は company と同じ）。
        return "industry"
    if _COMPANY_SUBJ.search(text) and not _MACRO_SUBJ.search(text):
        return "company"
    if _MACRO_SUBJ.search(text):
        # マクロ語がある。会社固有の主語も**強く**あるなら会社の話とみなす。
        if _INDUSTRY_SUBJ.search(text):
            return "industry"
        if len(_COMPANY_SUBJ.findall(text)) >= 2:
            return "company"
        return "macro"
    if _INDUSTRY_SUBJ.search(text):
        return "industry"
    return "company"


def classified_sentences(body):
    """[(文, 段落の種類)]。段落の種類を文に引き継ぐ。

    文単位で分類しないのは、マクロ段落の2文目以降が主語を省くから。
    「一方、米国の保護主義的な通商政策の再強化は〜」は単独でもマクロだが、
    「これにより消費者マインドが下振れする可能性がある」は単独では
    判定できない。段落で決めて引き継ぐのが正しい。
    """
    out = []
    for para in paragraphs(body):
        kind = classify_paragraph(para)
        for s in sentences(para):
            out.append((s, kind))
    return out


def _no_credit(sentence):
    for pat in cfg().get("no_credit") or []:
        if re.search(pat, sentence):
            return True
    return False


SCORED_CLASSES = ("company", "industry")


def _hits(sents, patterns):
    """パターンに当たる文を返す。加点禁止・マクロ段落の文は除く。

    sents は [(文, 段落種別)] でも [文] でも受ける（後者は company 扱い）。
    戻り値は (文, パターン, 段落種別)。
    """
    out = []
    for item in sents:
        s, kind = item if isinstance(item, tuple) else (item, "company")
        if kind not in SCORED_CLASSES:
            continue                      # マクロ定型文は採点根拠にしない
        if _no_credit(s):
            continue
        for pat in patterns:
            if re.search(pat, s):
                out.append((s, pat, kind))
                break
    return out


def detect_flags(cur_sents):
    """Tier D。**加点しない。** 層別のための印だけを返す。"""
    flags = {}
    for rule in cfg().get("tier_d") or []:
        hit = _hits(cur_sents, rule["patterns"])
        # segment_changed と forecast_revision は方針文の中にも出るので
        # 加点禁止フィルタを通さずに素で見る（フラグなので害がない）
        if not hit:
            raw = [s for s in cur_sents
                   if any(re.search(p, s) for p in rule["patterns"])]
            hit = [(s, None) for s in raw]
        flags[rule["key"]] = hit
    return flags


def score_pair(cur_body, prior_body, *, segment_changed=False):
    """当期と前期の本文から Tier A/B/C を計算する。

    戻り値: (tier_a, tier_b, tier_c, evidence[list])
    evidence は (tier, rule_key, matched_text, score)。
    """
    conf = cfg()
    caps = conf["caps"]
    # **段落の種類を持ったまま文に割る。** マクロ経済の定型文を
    # 採点根拠から外すため（設計判断は classify_paragraph の注記を参照）。
    cur = classified_sentences(cur_body)
    prior = classified_sentences(prior_body)
    ev = []

    # --- Tier A: 硬い事実。当期にあり前期に無いものを効かせる
    a = 0.0
    for rule in conf.get("tier_a") or []:
        w = float(rule.get("weight") or 0.0)
        cur_hits = _hits(cur, rule["patterns"])
        prior_hit = bool(_hits(prior, rule["patterns"]))
        if not cur_hits or prior_hit:
            continue                      # 前期にもあるなら「変化」ではない
        s, _pat, kind = cur_hits[0]
        a += w
        ev.append(("A", rule["key"], s, w, kind))
    a = max(-caps["tier_a"], min(caps["tier_a"], a))

    # --- Tier B: セグメント方向。区分が変わった期は評価しない
    b = 0.0
    if not segment_changed:
        tb = conf.get("tier_b") or {}
        for side in ("positive", "negative"):
            spec = tb.get(side) or {}
            w = float(spec.get("weight") or 0.0)
            n_cur = len(_hits(cur, spec.get("patterns") or []))
            n_pri = len(_hits(prior, spec.get("patterns") or []))
            if n_cur > n_pri:
                d = min(n_cur - n_pri, 3) * w
                b += d
                hit = _hits(cur, spec.get("patterns") or [])
                ev.append(("B", "segment_%s" % side, hit[0][0], d, hit[0][2]))
        b = max(-caps["tier_b"], min(caps["tier_b"], b))

    # --- Tier C: 語彙diff。前期に無く当期に現れた語
    c = 0.0
    tc = conf.get("tier_c") or {}
    for side in ("positive", "negative"):
        spec = tc.get(side) or {}
        w = float(spec.get("weight") or 0.0)
        for word in spec.get("words") or []:
            # **マクロ段落の語は数えない。** ここが今回の修正の核心。
            in_cur = [(s, k) for s, k in cur
                      if word in s and k in SCORED_CLASSES and not _no_credit(s)]
            in_pri = any(word in s for s, _k in prior)
            if in_cur and not in_pri:
                c += w
                ev.append(("C", "appeared:%s" % word, in_cur[0][0], w, in_cur[0][1]))
    c = max(-caps["tier_c"], min(caps["tier_c"], c))
    return a, b, c, ev


def apply_dedup(con, code, period_label, ev):
    """辞書が dedup を指定したルールを 0 点にする。

    修正言及: TDnet の業績予想修正が既に採用されているなら 0
    減損:     pl_adjustments に登録済みなら 0
    二重に数えないための処理で、**根拠は消さずに点だけ落とす**。
    """
    dedup_of = {r["key"]: r.get("dedup") for r in (cfg().get("tier_a") or [])}
    out = []
    for item in ev:
        # 段落種別を足す前の4要素タプルも受ける（呼び出し側の互換）。
        tier, key, text, sc = item[:4]
        kind = item[4] if len(item) > 4 else "company"
        reason = None
        d = dedup_of.get(key)
        if d == "tdnet_revision":
            n = con.execute(
                "SELECT COUNT(*) FROM filings WHERE code=? AND source='tdnet' "
                "AND subtype='業績予想修正'", (code,)).fetchone()[0]
            if n:
                reason = "tdnet_revision(%d件)" % n
        elif d == "pl_adjustments":
            n = con.execute(
                "SELECT COUNT(*) FROM pl_adjustments WHERE code=?", (code,)).fetchone()[0]
            if n:
                reason = "pl_adjustments(%d件)" % n
        out.append((tier, key, text, 0.0 if reason else sc, reason, kind))
    return out


def revision_direction(con, code, as_of):
    """方向は本文から判定せず guidance から補完する。"""
    r = con.execute(
        "SELECT revision_direction FROM guidance WHERE code=? AND date<=? "
        "AND revision_direction IS NOT NULL ORDER BY date DESC LIMIT 1",
        (code, as_of)).fetchone()
    return r[0] if r else "unknown"


# ----------------------------------------------------------------- ペア構築
_RE_Q = re.compile(r"第\s*(\d)\s*四半期")


def build_pairs(con, code=None):
    """同一スパンの前期・当期ペア。決算期変更・無効期をまたぐものは作らない。"""
    sql = ("SELECT id, code, date, subtype, xbrl_path, title FROM filings "
           "WHERE source='edinet' AND subtype IN ('140','160','120') "
           "AND xbrl_ok=1 AND xbrl_path IS NOT NULL")
    args = ()
    if code:
        sql += " AND code=?"
        args = (code,)
    rows = con.execute(sql + " ORDER BY code, date", args).fetchall()

    buckets = {}
    for r in rows:
        t = nfkc(r["title"])
        q = _RE_Q.search(t)
        key = (r["code"], r["subtype"], q.group(1) if q else "FY")
        buckets.setdefault(key, []).append(r)

    pairs = []
    for key, v in buckets.items():
        for a, b in zip(v[:-1], v[1:]):
            gap = int(b["date"][:4]) - int(a["date"][:4])
            if gap != 1:
                continue                    # 1年離れていないものは組まない
            if _pair_is_invalid(con, key[0], b["id"], a["id"]):
                continue                    # 対象期が無効（決算期変更・遡及修正等）
            pairs.append((b, a))            # (当期, 前期)
    return pairs


def _periods_of(con, filing_id):
    """その書類が語っている (period, q_no) の集合。"""
    return {(r[0], r[1]) for r in con.execute(
        "SELECT DISTINCT period, q_no FROM financials_cum "
        "WHERE filing_id=? AND period IS NOT NULL AND q_no IS NOT NULL",
        (filing_id,))}


def _pair_is_invalid(con, code, cur_id, prior_id):
    """**そのペアの対象期**が無効化されているか。

    当初「2つの開示のあいだに無効期があるか」で見ていたが、それだと
    その銘柄にどこか1つでも連結範囲変更があると全ペアが落ちる
    （1301 で候補8組が全滅した）。判定すべきは
    **比べようとしている当期と前期そのものが有効か**。
    """
    for fid in (cur_id, prior_id):
        periods = _periods_of(con, fid)
        if not periods:
            continue
        # 1つの書類は当期に加えて前年の比較数値も載せる。**その書類自身の
        # 対象期**（最も新しい期）だけを見る。全期を要求すると、比較用に
        # 載っている無効期のせいで全ペアが落ちる（1301 で8組が全滅した）。
        period, q_no = max(periods)
        r = con.execute(
            "SELECT COUNT(*) FROM financials_q WHERE code=? AND period=? "
            "AND q_no=? AND valid_flag=0", (code, period, q_no)).fetchone()
        if r[0]:
            return True
    return False


def body_of(con, filing_row):
    """EDINET zip から 0102010_honbun の本文を取る。"""
    from bs4 import BeautifulSoup
    p = C.full_path(filing_row["xbrl_path"])
    if not os.path.exists(p):
        return None
    try:
        with zipfile.ZipFile(p) as z:
            names = [n for n in z.namelist()
                     if "PublicDoc" in n and "0102010_honbun" in n and n.endswith(".htm")]
            if not names:
                return None
            soup = BeautifulSoup(z.read(names[0]).decode("utf-8", "replace"), "lxml")
            for t in soup(["style", "script"]):
                t.decompose()
            return re.sub(r"\n{3,}", "\n\n", soup.get_text("\n", strip=True))
    except Exception:
        return None


def evaluate(con, cur, prior, *, store=True):
    """1ペアを評価して s12_scores / s12_evidence に書く。"""
    label = "%s-%s" % (cur["date"][:4], cur["subtype"])
    raw_cur, raw_pri = body_of(con, cur), body_of(con, prior)
    if not raw_cur or not raw_pri:
        return _unavailable(con, cur, prior, label, "本文が取れない", store)
    cur_body = extract_section(raw_cur)
    pri_body = extract_section(raw_pri)
    if not cur_body or not pri_body:
        return _unavailable(con, cur, prior, label, "経営成績の節が見つからない", store)

    cur_sents = sentences(cur_body)
    flags = detect_flags(cur_sents)
    seg_changed = bool(flags.get("segment_changed"))
    a, b, c, ev = score_pair(cur_body, pri_body, segment_changed=seg_changed)
    ev = apply_dedup(con, cur["code"], label, ev)
    total = sum(x[3] for x in ev)
    cap = cfg()["caps"]["composite"]
    total = max(-cap, min(cap, total))

    res = {
        "filing_id": cur["id"], "prior_filing_id": prior["id"],
        "code": cur["code"], "period_label": label, "available": 1,
        "score": round(total, 4), "tier_a": round(a, 4),
        "tier_b": round(b, 4), "tier_c": round(c, 4),
        "segment_changed": int(seg_changed),
        "new_product_mention": int(bool(flags.get("new_product_mention"))),
        "forecast_revision_mentioned": int(bool(flags.get("forecast_revision_mentioned"))),
        "forecast_revision_direction": revision_direction(con, cur["code"], cur["date"]),
        "evidence": ev, "flags": flags,
    }
    if store:
        _store(con, res)
    return res


def _unavailable(con, cur, prior, label, reason, store):
    res = {"filing_id": cur["id"], "prior_filing_id": prior["id"] if prior else None,
           "code": cur["code"], "period_label": label, "available": 0,
           "unavailable_reason": reason, "score": None, "evidence": [], "flags": {}}
    if store:
        con.execute(
            "INSERT OR REPLACE INTO s12_scores (filing_id, prior_filing_id, code, "
            " period_label, available, unavailable_reason, computed_at) "
            "VALUES (?,?,?,?,0,?,?)",
            (res["filing_id"], res["prior_filing_id"], res["code"], label,
             reason, C.utcnow()))
        # **評価不能になったら古い根拠も消す。** スコアだけ available=0 に
        # 更新して根拠を残すと、どのスコアも支えていない根拠が s12_evidence に
        # 溜まる。2026-09-02 のマクロ除外の再計算後、787行(314書類分)の
        # 孤児が残っていて集計を狂わせた —— 「もう成り立っていない根拠」が
        # 現役の根拠と同じ表に同じ顔で並ぶのは、このプロジェクトが一貫して
        # 避けてきた失敗そのもの。
        con.execute("DELETE FROM s12_evidence WHERE filing_id=?",
                    (res["filing_id"],))
    return res


def _store(con, r):
    con.execute(
        "INSERT OR REPLACE INTO s12_scores (filing_id, prior_filing_id, code, "
        " period_label, available, score, tier_a, tier_b, tier_c, segment_changed, "
        " new_product_mention, forecast_revision_mentioned, "
        " forecast_revision_direction, computed_at) "
        "VALUES (?,?,?,?,1,?,?,?,?,?,?,?,?,?)",
        (r["filing_id"], r["prior_filing_id"], r["code"], r["period_label"],
         r["score"], r["tier_a"], r["tier_b"], r["tier_c"], r["segment_changed"],
         r["new_product_mention"], r["forecast_revision_mentioned"],
         r["forecast_revision_direction"], C.utcnow()))
    con.execute("DELETE FROM s12_evidence WHERE filing_id=?", (r["filing_id"],))
    for tier, key, text, sc, reason, kind in r["evidence"]:
        args = (r["filing_id"], r["prior_filing_id"], r["code"],
                r["period_label"], "経営成績の分析", tier, key, text[:500],
                sc, reason)
        try:
            con.execute(
                "INSERT INTO s12_evidence (filing_id, prior_filing_id, code, "
                " period_label, section, tier, rule_key, matched_text, score, "
                " dedup_applied, paragraph_class, created_at)"
                " VALUES (?,?,?,?,?,?,?,?,?,?,?,?)", args + (kind, C.utcnow()))
        except sqlite3.OperationalError:
            # **移行前でも走れるようにする。** 段落種別の列が無いだけで
            # 再計算そのものが落ちるのは筋が悪い（列の追加と再計算は
            # 別のジョブで、順番が前後しうる）。監査情報が欠けるだけで
            # スコアは同じものが入る。
            con.execute(
                "INSERT INTO s12_evidence (filing_id, prior_filing_id, code, "
                " period_label, section, tier, rule_key, matched_text, score, "
                " dedup_applied, created_at) VALUES (?,?,?,?,?,?,?,?,?,?,?)",
                args + (C.utcnow(),))


def main(argv=None):
    sys.stdout.reconfigure(encoding="utf-8")
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--all", action="store_true")
    p.add_argument("--code")
    p.add_argument("--limit", type=int)
    p.add_argument("--report", action="store_true")
    p.add_argument("--lock-wait", type=float, default=3600.0,
                   help="他の書き込みジョブを待つ秒数（既定1時間）")
    a = p.parse_args(argv)
    con = C.init_db()

    if a.report:
        n = con.execute("SELECT COUNT(*) c FROM s12_scores").fetchone()["c"]
        av = con.execute("SELECT COUNT(*) c FROM s12_scores WHERE available=1").fetchone()["c"]
        C.log("=== S12 ===")
        C.log("  評価したペア: %d / available %d (%.0f%%)" % (n, av, av / max(n, 1) * 100))
        for r in con.execute("SELECT unavailable_reason x, COUNT(*) n FROM s12_scores "
                             "WHERE available=0 GROUP BY x ORDER BY n DESC"):
            C.log("    欠損理由 %s: %d" % (r["x"], r["n"]))
        for r in con.execute("SELECT tier, COUNT(*) n, ROUND(SUM(score),2) s "
                             "FROM s12_evidence GROUP BY tier ORDER BY tier"):
            C.log("  Tier %s: %d ヒット / 合計 %s" % (r["tier"], r["n"], r["s"]))
        return 0

    if not (a.all or a.code):
        p.error("--all か --code か --report を指定する")
    pairs = build_pairs(con, a.code)
    if a.limit:
        pairs = pairs[:a.limit]
    C.log("ペア %d 組を評価" % len(pairs))
    # **書き込みは writer_lock で直列化する。** 日次パースと同時に走ると
    # database is locked で落ちる（2026-09-02 に踏んだ）。
    try:
        with C.writer_lock("s12", wait_seconds=a.lock_wait):
            n_av = 0
            for i, (cur, pri) in enumerate(pairs, 1):
                r = evaluate(con, cur, pri)
                n_av += r["available"]
                if i % 200 == 0:
                    con.commit()
                    C.log("  [%d/%d] available %d" % (i, len(pairs), n_av))
            con.commit()
            C.log("完了: %d 組 / available %d (%.0f%%)"
                  % (len(pairs), n_av, n_av / max(len(pairs), 1) * 100))
            return 0

    except C.WriterBusy as e:
        C.log("  ! 他の書き込みジョブが実行中: %s" % e)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
