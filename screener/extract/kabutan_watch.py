"""screener/extract/kabutan_watch.py — 外部リスト（株探「業績上方修正が有望」等）の手動アーカイブ取り込み。

**記録専用。スコアリング・採否・出口には一切使わない。**
用途は将来の PIT 検証（「外部が有望と言った銘柄は、その後どうなったか」を後から測れるようにする）。
スコア側がこのテーブルを読まないことは `screener/tests/test_kabutan_watch.py` が検査する。

**自動取得はしない。** 規約上、取得は人手のコピーに限る。このスクリプトが触るのは
`C:\\screener_data\\kabutan_watch\\` に**手で置かれた CSV** だけで、ネットワークには出ない。

CSV の列（順不同・余分な列は数値として保存）
--------------------------------------------
    date        取得日（YYYY-MM-DD / YYYY/M/D）。列が無ければ --date で渡す
    list_type   リストの種類（例: 上方修正有望）。無ければ --list-type
    code        銘柄コード（4桁 / 5桁 / 全角も可。正規化して保存）
    name        銘柄名（任意）
    その他       進捗率・乖離率など。数値に読めるものだけ metrics_json に入る

    python -m screener.extract.kabutan_watch --import C:\\screener_data\\kabutan_watch
    python -m screener.extract.kabutan_watch --report
"""
from __future__ import annotations

import argparse
import csv
import glob
import json
import os
import re
import sys

try:
    from screener import common as C
except ImportError:                                     # pragma: no cover
    sys.path.insert(0, os.path.dirname(os.path.dirname(
        os.path.dirname(os.path.abspath(__file__)))))
    from screener import common as C

WATCH_DIRNAME = "kabutan_watch"
ENCODINGS = ("utf-8-sig", "cp932")      # 手でコピーした CSV は Excel 経由が多い
# 銘柄コードの形（4桁 or 3桁+英字）。`common.normalise_code` は桁を詰めるだけで
# 妥当性を見ないので、ここで形を確かめる（"----" のような行を黙って入れない）。
CODE_RE = re.compile(r"^(?:\d{4}|\d{3}[A-Z])$")
KEY_COLS = {"date": ("date", "日付", "取得日"),
            "list_type": ("list_type", "種類", "リスト", "list"),
            "code": ("code", "コード", "銘柄コード", "証券コード"),
            "name": ("name", "銘柄名", "名称", "銘柄")}

DDL = """
CREATE TABLE IF NOT EXISTS kabutan_watch (
    date        TEXT NOT NULL,          -- リストを取得した日
    list_type   TEXT NOT NULL,          -- リストの種類
    code        TEXT NOT NULL,          -- 正規化後（4桁 or 3桁+英字）
    code_raw    TEXT,                   -- CSV にあったままの表記
    name        TEXT,
    metrics_json TEXT,                  -- 進捗率・乖離率など（数値のみ）
    in_universe INTEGER,                -- 取り込み時点の universe_flag
    source_file TEXT NOT NULL,
    imported_at TEXT NOT NULL,
    PRIMARY KEY (date, list_type, code)
);
CREATE INDEX IF NOT EXISTS ix_kabutan_code ON kabutan_watch (code, date);
"""


def watch_dir() -> str:
    return os.path.join(C.DATA_DIR, WATCH_DIRNAME)


def _norm_date(s: str | None) -> str | None:
    s = (s or "").strip()
    m = re.match(r"^(\d{4})[-/年](\d{1,2})[-/月](\d{1,2})", s)
    return "%s-%02d-%02d" % (m.group(1), int(m.group(2)), int(m.group(3))) if m else None


def _zen2han(s: str) -> str:
    return s.translate(str.maketrans("０１２３４５６７８９ＡＢＣＤＥＦＧＨＩＪＫＬＭ"
                                     "ＮＯＰＱＲＳＴＵＶＷＸＹＺ",
                                     "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"))


def _num(v):
    """'12.3%' / '+1,234' → float。読めなければ None（推測で0にしない）。"""
    if v is None:
        return None
    t = str(v).strip().replace(",", "").replace("%", "").replace("％", "")
    t = _zen2han(t).replace("＋", "").replace("+", "")
    t = t.replace("△", "-").replace("▲", "-").replace("−", "-")
    try:
        return float(t)
    except ValueError:
        return None


def _pick(row: dict, names) -> str | None:
    for n in names:
        for k in row:
            if k and k.strip().lower() == n.lower():
                return row[k]
    return None


def read_csv(path: str) -> list[dict]:
    """手置きの CSV を読む。文字コードは utf-8(BOM) → cp932 の順に試す。"""
    last = None
    for enc in ENCODINGS:
        try:
            with open(path, encoding=enc, newline="") as fh:
                return list(csv.DictReader(fh))
        except UnicodeDecodeError as e:
            last = e
    raise UnicodeDecodeError(*last.args)                # pragma: no cover


def parse_rows(rows, source_file, default_date=None, default_type=None):
    """CSV の行 → 取り込み用の dict。コードが読めない行は理由つきで捨てる。"""
    out, skipped = [], []
    for r in rows:
        raw_code = (_pick(r, KEY_COLS["code"]) or "").strip()
        code = C.normalise_code(_zen2han(raw_code))
        if code and not CODE_RE.match(code):
            code = None
        date = _norm_date(_pick(r, KEY_COLS["date"])) or default_date
        list_type = (_pick(r, KEY_COLS["list_type"]) or default_type or "").strip()
        if not code or not date or not list_type:
            skipped.append({"row": r, "reason": (
                "コードが読めない" if not code else
                "日付が無い（--date で渡す）" if not date else
                "list_type が無い（--list-type で渡す）")})
            continue
        used = {v.lower() for names in KEY_COLS.values() for v in names}
        metrics = {}
        for k, v in r.items():
            if not k or k.strip().lower() in used:
                continue
            n = _num(v)
            if n is not None:
                metrics[k.strip()] = n
        out.append({"date": date, "list_type": list_type, "code": code,
                    "code_raw": raw_code, "name": (_pick(r, KEY_COLS["name"]) or "").strip(),
                    "metrics_json": json.dumps(metrics, ensure_ascii=False) if metrics else None,
                    "source_file": os.path.basename(source_file)})
    return out, skipped


def import_rows(con, rows) -> dict:
    """取り込み（冪等）。取り込み時点の universe_flag を一緒に残す。"""
    con.executescript(DDL)
    uni = {r[0] for r in con.execute(
        "SELECT code FROM companies WHERE universe_flag=1")}
    known = {r[0] for r in con.execute("SELECT code FROM companies")}
    n_new = n_upd = 0
    for r in rows:
        exists = con.execute(
            "SELECT 1 FROM kabutan_watch WHERE date=? AND list_type=? AND code=?",
            (r["date"], r["list_type"], r["code"])).fetchone()
        con.execute(
            "INSERT INTO kabutan_watch (date, list_type, code, code_raw, name,"
            " metrics_json, in_universe, source_file, imported_at)"
            " VALUES (?,?,?,?,?,?,?,?,?)"
            " ON CONFLICT(date, list_type, code) DO UPDATE SET"
            " name=excluded.name, metrics_json=excluded.metrics_json,"
            " in_universe=excluded.in_universe, source_file=excluded.source_file,"
            " imported_at=excluded.imported_at",
            (r["date"], r["list_type"], r["code"], r["code_raw"], r["name"],
             r["metrics_json"], int(r["code"] in uni), r["source_file"], C.utcnow()))
        n_new += 0 if exists else 1
        n_upd += 1 if exists else 0
    con.commit()
    return {"new": n_new, "updated": n_upd,
            "in_universe": sum(r["code"] in uni for r in rows),
            "unknown_code": sorted({r["code"] for r in rows if r["code"] not in known})}


def report(con) -> str:
    con.executescript(DDL)
    L = ["# 株探ウォッチの取り込み状況（記録専用・スコアには使わない）", ""]
    tot = con.execute("SELECT COUNT(*) FROM kabutan_watch").fetchone()[0]
    L.append("行 %d" % tot)
    if not tot:
        L.append("")
        L.append("まだ1件も入っていない。%s に CSV を置いてから --import する。" % watch_dir())
        return "\n".join(L)
    L.append("")
    L.append("| 取得日 | 種類 | 銘柄 | うちユニバース内 | 元ファイル |")
    L.append("|---|---|---|---|---|")
    for r in con.execute(
            "SELECT date, list_type, COUNT(*), SUM(in_universe),"
            " GROUP_CONCAT(DISTINCT source_file) FROM kabutan_watch "
            "GROUP BY date, list_type ORDER BY date DESC, list_type"):
        L.append("| %s | %s | %d | %d | %s |" % (r[0], r[1], r[2], r[3] or 0, r[4]))
    L.append("")
    rep = con.execute(
        "SELECT code, COUNT(DISTINCT date) n FROM kabutan_watch "
        "GROUP BY code HAVING n > 1 ORDER BY n DESC LIMIT 10").fetchall()
    L.append("複数回載った銘柄（上位10）: "
             + (", ".join("%s(%d回)" % (c, n) for c, n in rep) if rep else "なし"))
    return "\n".join(L)


TEMPLATE = """date,list_type,code,name,進捗率,乖離率
2026-09-20,上方修正有望,6758,ソニーグループ,72.5,18.3
"""

README = """kabutan_watch — 外部リストの手動アーカイブ（記録専用）

・ここには **手でコピーした CSV** だけを置く。自動取得はしない（規約上の判断）。
・列: date / list_type / code / name / （あれば 進捗率・乖離率 などの数値列）
  - date と list_type は列が無ければ取り込み時に --date / --list-type で渡せる。
  - code は4桁・5桁・全角でもよい（取り込み時に正規化する）。
・取り込み:
    python -m screener.extract.kabutan_watch --import C:\\screener_data\\kabutan_watch
・**このデータはスコアリング・採否・出口に一切使わない。** 将来の PIT 検証用の記録。
  スコア側が参照していないことはテスト（test_kabutan_watch.py）が検査する。
"""


def ensure_dir() -> str:
    d = watch_dir()
    os.makedirs(d, exist_ok=True)
    for name, body in (("README.txt", README), ("_template.csv", TEMPLATE)):
        p = os.path.join(d, name)
        if not os.path.exists(p):
            with open(p, "w", encoding="utf-8") as fh:
                fh.write(body)
    return d


def main(argv=None) -> int:
    sys.stdout.reconfigure(encoding="utf-8")
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--import", dest="src", nargs="?", const="", metavar="PATH",
                   help="CSV かディレクトリ。省略時は既定の kabutan_watch ディレクトリ")
    p.add_argument("--date", help="CSV に date 列が無いときの取得日")
    p.add_argument("--list-type", help="CSV に list_type 列が無いときの種類")
    p.add_argument("--report", action="store_true")
    a = p.parse_args(argv)

    d = ensure_dir()
    con = C.init_db()
    if a.report or a.src is None:
        print(report(con))
        if a.src is None:
            C.log("置き場: %s（--import で取り込む）" % d)
        return 0

    src = a.src or d
    files = ([src] if src.lower().endswith(".csv")
             else sorted(f for f in glob.glob(os.path.join(src, "*.csv"))
                         if not os.path.basename(f).startswith("_")))
    if not files:
        C.log("CSV が無い: %s" % src)
        return 0
    total = {"new": 0, "updated": 0, "in_universe": 0}
    for f in files:
        rows, skipped = parse_rows(read_csv(f), f, a.date, a.list_type)
        st = import_rows(con, rows)
        for k in total:
            total[k] += st[k]
        C.log("%s: 取り込み %d（新規 %d / 更新 %d）/ ユニバース内 %d / 捨てた行 %d"
              % (os.path.basename(f), len(rows), st["new"], st["updated"],
                 st["in_universe"], len(skipped)))
        for s in skipped[:5]:
            C.log("  ! %s: %s" % (s["reason"], s["row"]))
        if st["unknown_code"]:
            C.log("  ! companies に無いコード %d 件: %s"
                  % (len(st["unknown_code"]), st["unknown_code"][:10]))
    C.log("合計: 新規 %d / 更新 %d / ユニバース内 %d"
          % (total["new"], total["updated"], total["in_universe"]))
    print(report(con))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
