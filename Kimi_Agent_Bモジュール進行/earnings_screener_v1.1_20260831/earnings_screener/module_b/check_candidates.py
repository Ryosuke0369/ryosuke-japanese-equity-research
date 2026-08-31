"""check_candidates.py — T-15候補全社へのS1-S4一括適用（分布確認用）"""
import csv
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from module_b.data_access import connect
from module_b.run_scorers import score_ticker, SCORERS_ALL

conn = connect()
with open(Path(__file__).resolve().parents[1] / "data" / "candidates_2026-08-31.csv",
          encoding="utf-8-sig") as f:
    tickers = [r["ticker"] for r in csv.DictReader(f)]

rows = []
for t in tickers:
    res = score_ticker(conn, t, SCORERS_ALL)
    rows.append({"ticker": t, "evidence_score": res["evidence_score"],
                 **{k: v["score"] for k, v in res.items() if k != "evidence_score"}})

import pandas as pd
df = pd.DataFrame(rows).sort_values("evidence_score", ascending=False)
pd.set_option("display.width", 200)
print("=== 証拠スコア上位10社 ===")
print(df.head(10).to_string(index=False))
print("\n=== 下位5社 ===")
print(df.tail(5).to_string(index=False))
print("\n=== 分布 ===")
print(df.describe().loc[["mean", "std", "min", "max"]].to_string())
