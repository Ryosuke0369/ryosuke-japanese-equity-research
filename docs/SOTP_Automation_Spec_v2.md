# SOTP自動化 設計仕様書 v2.0 (Final)
# 2026/04/09 作成

---

## 0. 設計方針

### 3つの原則
1. **既存パイプラインに一切触れない** — generate_dcf.py, dcf_comps_template.py, edinet_parser.py は変更なし
2. **完全に独立した新規ファイルのみ作成** — generate_sotp.py + sotp_template.py の2ファイル
3. **overrides JSONの`sotp`セクションで全てを制御** — sotpセクションがなければスキップ

### 自動/手動の切り分け（確定）

| 工程 | Phase 1（今回） | Phase 2（将来） | 理由 |
|---|---|---|---|
| セグメント別OP | **手動** → overrides JSON | **自動** → EDINET XBRL パーサー拡張 | 現パーサーのセグメント対応が不完全 |
| D&A配分比率 | **手動** → overrides JSON | 手動のまま | IHI自体が非開示。100%人間の判断 |
| Peer企業リスト | **手動** → overrides JSON | 手動のまま | どの企業をPeerにするかは判断 |
| Peerマルチプル | **手動** → overrides JSON | **自動** → yfinance | Phase 1は手動で正確性優先 |
| Selected Multiple | **手動** → overrides JSON | 手動のまま | Median vs 調整判断は人間 |
| ネットデット | **自動** → EDINET BS | 自動 | 既存パーサーで取得済み |
| 少数株主持分 | **自動** → EDINET BS | 自動 | 同上 |
| 株式数 | **自動** → yfinance | 自動 | 同上 |
| 連結D&A | **自動** → EDINET IS | 自動 | 既存パーサーで取得済み |
| EBITDA計算 | **自動** | 自動 | OP + D&A |
| セグメントEV計算 | **自動** | 自動 | EBITDA × Multiple |
| エクイティブリッジ | **自動** | 自動 | 固定ロジック |
| 感応度テーブル | **自動** | 自動 | パラメトリック計算 |
| Excel 6シート生成 | **自動** | 自動 | openpyxl |

**人間が入力するもの（Phase 1で3つ）：**
1. セグメント別OP + D&A配分比率
2. Peer企業リスト + EV/EBITDA + selected multiple
3. 投資テーシス（テキスト）

**全て overrides JSON の `sotp` セクションに記入する。**

---

## 1. ファイル構成

### 新規作成（2ファイルのみ）
```
scripts/generate_sotp.py        ← エントリーポイント
templates/sotp_template.py      ← Excel生成エンジン
```

### 変更するファイル（1ファイルのみ）
```
data/overrides/7013_overrides.json   ← `sotp`セクション追加（既存キーは変更不可）
```

### 変更しないファイル
```
scripts/generate_dcf.py         ← 触れない
scripts/edinet_fetcher.py       ← 触れない
scripts/edinet_parser.py        ← 触れない（Phase 2で拡張予定）
templates/dcf_comps_template.py ← 触れない
data/overrides/2359_overrides.json  ← 触れない
data/overrides/6246_overrides.json  ← 触れない
```

---

## 2. overrides JSON — `sotp`セクション仕様

### 7013_overrides.json に追加する構造

```json
{
  "sotp": {
    "thesis": "IHI's consolidated valuation obscures value concentration in Aero segment. 76% of EV from Aerospace & Defense.",
    "valuation_date": "April 2026",
    "stock_split_ratio": 7,

    "consolidated": {
      "da_total": 70000,
      "net_debt": 360000,
      "minority_interest": 27000,
      "equity_method_op": 6200,
      "equity_method_multiple": 10.0,
      "shares_outstanding": 151364
    },

    "segments": [
      {
        "key": "aero",
        "label": "Aero, Space & Defense",
        "label_jp": "航空・宇宙・防衛",
        "revenue": [270400, 555700, 620000],
        "op": [-102800, 122700, 130000],
        "fiscal_years": ["FY23", "FY24", "FY25E"],
        "da_allocation_pct": 0.31,
        "selected_multiple": 18.0,
        "multiple_source": "Median of Safran/MTU/Rolls-Royce (excl. TransDigm)",
        "peers": [
          {"name": "Safran", "ticker": "SAF.PA", "ev_ebitda": 20.5, "opm": 0.166, "market_cap_usd_b": 130, "note": "CFM LEAP, MRO, Defense", "excluded": false},
          {"name": "MTU Aero Engines", "ticker": "MTX.DE", "ev_ebitda": 11.9, "opm": 0.199, "market_cap_usd_b": 24, "note": "V2500/PW1100G partner", "excluded": false},
          {"name": "Rolls-Royce", "ticker": "RR.L", "ev_ebitda": 18.0, "opm": 0.15, "market_cap_usd_b": 65, "note": "Large engine, V-shape recovery", "excluded": false},
          {"name": "TransDigm", "ticker": "TDG", "ev_ebitda": 25.0, "opm": 0.54, "market_cap_usd_b": 75, "note": "Aftermarket monopoly — EXCLUDED", "excluded": true}
        ]
      },
      {
        "key": "energy",
        "label": "Resources, Energy & Environment",
        "label_jp": "資源・エネルギー・環境",
        "revenue": [404900, 411400, 370000],
        "op": [17700, 16100, 24000],
        "fiscal_years": ["FY23", "FY24", "FY25E"],
        "da_allocation_pct": 0.30,
        "selected_multiple": 8.0,
        "multiple_source": "Energy equipment peers",
        "peers": [
          {"name": "Siemens Energy", "ticker": "ENR.DE", "ev_ebitda": 12.0, "opm": 0.06, "market_cap_usd_b": 50, "note": "Gas turbine, grid", "excluded": false},
          {"name": "MHI (Energy seg.)", "ticker": "7011.T", "ev_ebitda": 8.0, "opm": 0.05, "market_cap_usd_b": 0, "note": "Comparable segment estimate", "excluded": false}
        ]
      },
      {
        "key": "industrial",
        "label": "Industrial Systems & General Machinery",
        "label_jp": "産業システム・汎用機械",
        "revenue": [466100, 484800, 440000],
        "op": [12700, 10800, 25000],
        "fiscal_years": ["FY23", "FY24", "FY25E"],
        "da_allocation_pct": 0.27,
        "selected_multiple": 10.0,
        "multiple_source": "Turbocharger/machinery blend",
        "peers": [
          {"name": "BorgWarner", "ticker": "BWA", "ev_ebitda": 6.5, "opm": 0.10, "market_cap_usd_b": 7, "note": "Turbocharger OEM", "excluded": false},
          {"name": "Garrett Motion", "ticker": "GTX", "ev_ebitda": 7.0, "opm": 0.15, "market_cap_usd_b": 3, "note": "Turbocharger pure-play", "excluded": false}
        ]
      },
      {
        "key": "social_infra",
        "label": "Social Infrastructure",
        "label_jp": "社会基盤",
        "revenue": [154300, 146000, 140000],
        "op": [4900, -4200, 0],
        "fiscal_years": ["FY23", "FY24", "FY25E"],
        "da_allocation_pct": 0.12,
        "selected_multiple": 6.0,
        "multiple_source": "Infrastructure floor",
        "peers": [
          {"name": "Yokogawa Bridge", "ticker": "5911.T", "ev_ebitda": 5.5, "opm": 0.04, "market_cap_usd_b": 0, "note": "Bridge construction — low margin", "excluded": false}
        ]
      }
    ],

    "conglomerate_discount": {
      "base": 0.0,
      "sensitivity_range": [0.0, 0.05, 0.10, 0.15, 0.20]
    },

    "sensitivity": {
      "primary_segment_key": "aero",
      "primary_multiples": [14, 16, 18, 20, 22, 25],
      "table2": {
        "row_segment_key": "industrial",
        "row_multiples": [7, 8, 10, 12, 14],
        "col_segment_key": "energy",
        "col_multiples": [6, 8, 10, 12]
      }
    },

    "dcf_crosscheck": {
      "pgm_fair_value": 1982,
      "exit_fair_value": 3463,
      "comps_ev_ebitda": 3794,
      "comps_per": 4973
    }
  }
}
```

### 単位規則
- `revenue`, `op`: **百万円**（EDINET準拠）
- `da_total`, `net_debt`, `minority_interest`: **百万円**
- `shares_outstanding`: **千株**
- `ev_ebitda`, `selected_multiple`: 倍率（x）
- `opm`, `da_allocation_pct`: 小数（0.31 = 31%）
- `market_cap_usd_b`: $B（表示用のみ、計算には使わない）
- `dcf_crosscheck`の値: 分割後株価（¥）

---

## 3. generate_sotp.py — 処理フロー

```python
#!/usr/bin/env python3
"""
SOTP Valuation Model Generator
Usage: python scripts/generate_sotp.py <ticker>
Example: python scripts/generate_sotp.py 7013
"""
import sys, json, os

def main():
    ticker = sys.argv[1]
    overrides_path = f"data/overrides/{ticker}_overrides.json"
    with open(overrides_path) as f:
        overrides = json.load(f)

    if "sotp" not in overrides:
        print(f"No SOTP config in {overrides_path}. Exiting.")
        sys.exit(0)

    sotp = overrides["sotp"]
    output_path = f"models/{ticker}_SOTP_Model.xlsx"

    from templates.sotp_template import generate_sotp_excel
    generate_sotp_excel(sotp, output_path)

    os.system(f"python scripts/recalc.py {output_path}")
    print(f"Done: {output_path}")

if __name__ == "__main__":
    main()
```

---

## 4. sotp_template.py — Excel生成仕様

### 6シート構成

| シート | 内容 | 数式/入力 |
|---|---|---|
| Cover & Thesis | テーシス + 比較表 | テキスト + dcf_crosscheck値 |
| Segment Data | 3年×セグメント売上/OP/OPM | 青=入力, 黒=OPM数式, 緑=合計数式 |
| Peer Comps | セグメント別Peer + Median + Selected | 青=Peer値, MEDIAN数式, 黄=Selected |
| SOTP Valuation | EBITDA→EV→株価 | 緑=他シートリンク, 黒=全計算数式 |
| Sensitivity | 2つの感応度テーブル | 全セル数式（固定前提をセル参照） |
| D&A Allocation | 配分方法論 | 青=比率, 黒=計算数式 |

### 色規則（既存DCFと統一）
- **青文字** (0,0,255): ハードコード入力
- **黒文字** (0,0,0): Excel数式
- **緑文字** (0,128,0): 他シートリンク
- **黄色背景**: 変更可能な主要前提
- **水色背景**: ベースケース / 結論セル

### 重要: 全計算はExcel数式で行う（Pythonでハードコードしない）
- EBITDA = Excel数式 `=B5+C5`
- Segment EV = Excel数式 `=D5*E5`
- Fair Value = Excel数式 `=F16*100000/F18/7`
- 感応度テーブル = Excel数式（固定前提セルを参照）

---

## 5. テスト計画

### 5-A. IHI (7013) 照合

| チェック項目 | 期待値 |
|---|---|
| 航空 EBITDA | ~152,000百万円 |
| Total Segment EV | ~3,584,000百万円 |
| Fair Value (post-split) | ~¥3,077 |
| 感応度(18x, 0%) | ~¥3,122 |
| recalc.pyエラー | 0 |

### 5-B. 既存非破壊テスト
```bash
python scripts/generate_dcf.py 2359  # 変化なし
python scripts/generate_dcf.py 6246  # 変化なし
python scripts/generate_dcf.py 7013  # 変化なし
```

---

## 6. Phase 2 予定（今回スコープ外）

- edinet_parser.py セグメントXBRL拡張
- yfinance Peer自動取得
- LBO自動化（別仕様書）
- PDF自動生成

---

## 7. ClaudeCodeへの実装指示

### 渡すもの
1. この仕様書
2. リポジトリアクセス

### 指示文
> この仕様書（SOTP_Automation_Spec_v2.md）を読んで実装してください。
> 既存ファイルは7013_overrides.jsonのsotp追加以外は一切変更しないこと。
> generate_sotp.pyとsotp_template.pyを新規作成し、IHI (7013)で動作確認してください。
> テスト計画セクション5-A/5-Bを実行し結果を報告してください。
