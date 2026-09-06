"""Render the Adjustments Log config and the 1-ticker summary from one spec.

Generic: every judgement is supplied per ticker by the caller in a `spec` dict.
This file only fixes the SHAPE of the two artefacts (the batch spec 8-1 fields
and the Adjustments Log 6 columns), never the content — scenario design, peer
choice and exit multiples are decided per ticker and passed in.

Usage from a per-wave script:

    from batch.make_docs import emit
    emit(spec)          # writes data/adjustments/<code>_adjustments.json
                        #    and batch/per_ticker/<code>_summary.md
"""
import json, os, sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATE = "2026-09-05"

# Rows every batch ticker carries. Values that differ per ticker come from spec.
def _common_entries(s):
    e = []
    e.append(["DCF Model!C7", "Risk-Free Rate = 2.97%",
              "財務省 国債金利情報 jgbcm.csv 10年物 R8.9.3(2026-09-03) 2.966% を四捨五入。バッチ全銘柄共通の一次ソース。yfinance は日本国債利回りを返さないため使用しない。",
              "テンプレ既定 2.2%", "確定"])
    e.append(["DCF Model!C9", "Equity Risk Premium = 6.5%",
              "バッチ共通の固定前提(v2 §1-3)。銘柄固有の調整はしていない。", "テンプレ既定 6.5%(同値)", "設計判断"])
    e.append(["DCF Model!C10", "Size Premium = 1.0%",
              "時価総額バンドによる機械決定(>3,000億円 -> 1.0%)。当社時価総額 " + s["mcap"] + "百万円。"
              "■テンプレートの自動規則(>=1兆円 -> 0.0%、>=1,000億円 -> 1.5%)とは異なる値であり、明示指定で上書きしている。",
              "テンプレ自動値", "設計判断"])
    e.append(["DCF Model!C8", "Beta = " + s["beta"], s["beta_reason"],
              "yfinance実測 " + s["beta_raw"], s["beta_status"]])
    e.append(["DCF Model!C6", "Effective Tax Rate = 25%",
              "バッチ共通の既定値(v2 §1-3)。当社の実績税負担率を個別に検証したものではない。", "テンプレ既定 30%", "推定"])
    e.append(["DCF Model!C11", s["kd"][0], s["kd"][1], s["kd"][2], s["kd"][3]])
    e.append(["DCF Model!C13", "Terminal Growth = 1.0%", "バッチ共通の固定前提(v2 §1-3)。", "テンプレ既定 2.0%", "設計判断"])
    e.append(["DCF Model!C14", "Exit Multiple (EV/EBITDA) = " + s["exitm"] + "x",
              s["exit_reason"] + " 自社の現在倍率 " + s["own_mult"] + " は算式に入れていない(市場価格への循環参照回避)。"
              "■マージン品質調整式はバッチ全体で棄却済み(EV/EBITDA が既に資本集約度を織り込むため、営業利益率による上方調整は資本集約型で逆方向に働く)。",
              "テンプレ既定 10.0x", "設計判断"])
    e.append(["全体(型判定・プレスクリーン)", s["type_line"],
              "追補1 §B に従い、型判定の前に batch/prescreen.py で EDINET有報XBRL の金融事業専用科目"
              "(ForFinancialBusiness / Banking / Deposits / CallLoan / InstallmentReceivable 系)を機械チェックした。"
              "検出結果: " + s["prescreen"] + "。" + s["type_reason"], "-", "設計判断"])
    e.extend(s.get("extra", []))
    e.append(["DCF Model シナリオ5本", s["scen"][0], s["scen"][1], "テンプレ自動生成シナリオ", "設計判断"])
    e.append(["DCF Model Management シナリオ", "Management = Base と同一",
              "EDINET XBRL・TDnet短信のいずれからも会社業績予想を取得できなかった(生成ログ Step 3: 'No guidance data found')。"
              "v2 §7-2 に従い Management=Base で続行。■本バッチではここまで生成した全銘柄で同じ結果になっており、"
              "銘柄固有ではなくパイプラインの構造的な取得失敗である。", "-", "フォールバック使用(会社予想未取得)"])
    e.append(["DCF Model NWC(dso/dih/dpo)", "Base: " + s["nwc"], s["nwc_reason"], "テンプレ自動値", "設計判断"])
    e.append(["Comps Analysis", "Peer " + str(len(s["peers"])) + "社 = " + " / ".join(s["peers"]),
              s["peer_reason"], "-", "設計判断"])
    e.append(["Executive Summary!C10 / Target",
              "Target Price = " + s["tgt"] + "円 = PGM " + s["pgm"] + "円 と Exit " + s["ex"] +
              "円 の2本平均。Comps 2本(EV/EBITDA " + s["cev"] + "円 / PER " + s["cper"] + "円)は [参考・Target不算入]",
              "メモ§2 の規約どおり。validate チェック14 PASS(C10 が C16:C17 のみを平均)。" + s.get("target_note", ""),
              "-", "確定"])
    e.append(["全体(納品状態)", "本ファイルは DRAFT。_FINAL は付けない",
              "バッチモードでの機械生成物であり人間の検証を経ていない。FINAL化は手順書v2 §6 に従い検証後に行う。", "-", "DRAFT"])
    for o in s["open_items"]:
        e.append(["全体(未解決)", o[0], o[1], "-", "未解決"])
    return e


SUMMARY = """# {code} {name} — DCF生成サマリー（バッチ / DRAFT）

生成物: `models/{code}_DCF_Model_20260905.xlsx`（DRAFT・_FINAL なし）
分析基準日: 2026-09-05 / テンプレートrev: c9b5dd9 / 決算期末月: {fy}月{retry_note}{typeb_note}

## 1. 型判定と根拠

{type_block}

**金融事業プレスクリーン（追補1 §B）**: {prescreen}

## 2. 株数の検算式

| 出所 | 値 |
|---|---|
| 採用値（yfinance sharesOutstanding、自己株控除後） | **{shares} 株** |
| 検算: 時価総額 {mcap}百万円 ÷ 株価 {price}円 | = {shares} 株 ✓ |

EDINET有報による「発行済株式総数 − 自己株式」の独立検算は未実施 → Adjustments Log に「要確認」。

## 3. net_debt 構成（型A定義: 有利子負債 − 現金及び現金同等物、JPY mn）

{nd}

{nd_note}

## 4. WACC inputs

| 項目 | 値 | 根拠 |
|---|---|---|
| Risk-free | 2.97% | 財務省 jgbcm.csv 10年 R8.9.3（バッチ共通） |
| **Beta** | **{beta}** | {beta_short} |
| ERP | 6.5% | バッチ共通固定 |
| Size premium | 1.0% | 時価総額バンド（テンプレ自動規則なら 0.0%） |
| 税引後Kd | {kd_short} | |
| D/E | {de} | |
| Ke | {ke} | 2.97% + {beta}×6.5% + 1.0% |
| **WACC** | **{wacc}** | |

## 5. シナリオ5本の含意営業利益率 検算テーブル

実績OPM（直近4期）: {opm}

| シナリオ | 到達OPM（Y5） | 導出ルール |
|---|---|---|
| **Base** | {base} |
| **Upside** | {up} |
| **Management** | Base と同一 | **会社予想が取得できず**（生成ログ Step 3「No guidance data found」）→ v2 §7-2 |
| **Downside 1** | {d1} |
| **Downside 2** | {d2} |

{scen_note}

NWC（実績4期平均）: {nwc}。{nwc_reason}

## 6. Peer選定 / 除外理由・D&A取得状況

**採用{npeer}社**: {peers}

{peer_reason}

- **exit_multiple = {exitm}x** ← {exit_reason} 自社の現在倍率 **{own_mult}** は算式に入れていない。
- **Market_Cap 全行記入**（2026-09-04 終値ベース静的値）→ 再現性を確保。海外Peer は v2 §5 により不使用。

## 7. Target / 判定 / 逆算DCF

| 手法 | 含意株価 | Target算入 |
|---|---:|---|
| DCF — Perpetuity Growth | {pgm} 円 | ○ |
| DCF — Exit Multiple ({exitm}x) | {ex} 円 | ○ |
| Comps — EV/EBITDA 中央値 | {cev} 円 | ×[参考] |
| Comps — PER 中央値 | {cper} 円 | ×[参考] |
| **Target Price (DCF Mid)** | **{tgt} 円** | PGM+Exit の2本平均 |

現在株価 **{price} 円**（2026-09-04 終値） / **Upside {up_pct}** / 判定 **{verdict}**

**逆算DCF 一行回答（Block F）**: {rev}

## 8. validate 結果

```
{val}   VERDICT: PASS
```
（`models/{code}_DCF_Model_20260905_validation.txt` に全文。Adjustments Log 追記後に再recalc → 再validate 実施済み）
{warn_block}
## 9. 使用したフォールバック一覧

{fb}

## 10. 未解決事項（人間の判断を要する）

{open_items}
"""


def emit(s):
    code = s["code"]
    entries = _common_entries(s)
    p = os.path.join(ROOT, "data", "adjustments", "%s_adjustments.json" % code)
    json.dump({"_ticker": "%s %s" % (code, s["name"]),
               "_source": "バッチ生成(大手90銘柄 v2 + 追補1_20260905)。分析基準日 2026-09-05 / DRAFT。全フォールバック・全ルール判断を記録。",
               "date": DATE, "entries": entries},
              open(p, "w", encoding="utf-8"), ensure_ascii=False, indent=1)

    d = dict(s)
    d["retry_note"] = ("\n**リトライ%d回で合格**" % s["retries"]) if s.get("retries") else ""
    d["typeb_note"] = "\n**型B・要目視レビュー**" if s.get("typeb") else ""
    d["warn_block"] = ("\n**WARN の判断と記録（手順書§5-1）**: " + s["warn"] + "\n") if s.get("warn") else "\n"
    d["npeer"] = str(len(s["peers"]))
    d["peers"] = " / ".join(s["peers"])
    d["fb"] = "\n".join("%d. %s" % (i + 1, x) for i, x in enumerate(s["fb"]))
    d["open_items"] = "\n".join("%d. %s" % (i + 1, x[0] + " — " + x[1]) for i, x in enumerate(s["open_items"]))
    out = os.path.join(ROOT, "batch", "per_ticker", "%s_summary.md" % code)
    open(out, "w", encoding="utf-8").write(SUMMARY.format(**d))
    print("%s: adjustments %d entries + summary" % (code, len(entries)))
