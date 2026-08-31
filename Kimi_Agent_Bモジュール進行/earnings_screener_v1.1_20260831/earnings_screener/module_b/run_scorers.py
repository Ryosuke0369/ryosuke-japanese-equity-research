"""
run_scorers.py — スコアラー検証ランナー（S1-S4 / 後続でS5-S8も統合）
使い方: python -m module_b.run_scorers [ticker ...]（省略時はテスト銘柄＋T-15候補全社）
"""
import csv
import sys
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from module_b.data_access import connect, DB
from module_b.scorers_s1_s4 import SCORERS_S1_S4, SIGNAL_LABELS as L14
from module_b.scorers_s5_s8 import SCORERS_S5_S8, SIGNAL_LABELS as L58

SCORERS_ALL = {**SCORERS_S1_S4, **SCORERS_S5_S8}
SIGNAL_LABELS = {**L14, **L58}

TEST_TICKERS = ["3905", "3441", "278A", "5726", "6501", "7203"]


def score_ticker(conn, ticker, scorers, as_of=None):
    out = {}
    for name, fn in scorers.items():
        try:
            out[name] = fn(conn, ticker, as_of)
        except Exception as e:
            out[name] = {"score": 0.0, "available": False, "evidence": f"ERROR: {e}"}
    avail = [v["score"] for v in out.values() if v["available"]]
    out["evidence_score"] = round(sum(avail) / len(avail), 3) if avail else None
    return out


def main():
    tickers = sys.argv[1:]
    conn = connect()
    if not tickers:
        tickers = TEST_TICKERS
    for t in tickers:
        res = score_ticker(conn, t, SCORERS_ALL)
        print(f"=== {t} ===  証拠スコア平均(利用可能な指標のみ): {res['evidence_score']}")
        for name, v in res.items():
            if name == "evidence_score":
                continue
            mark = "○" if v["available"] else "―"
            label = SIGNAL_LABELS.get(name, name)
            print(f"  [{mark}] {name}({label}): {v['score']:+.3f}  {v['evidence']}")
    conn.close()


if __name__ == "__main__":
    main()
