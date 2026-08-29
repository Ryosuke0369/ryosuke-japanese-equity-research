"""screener/extract/xbrl_parser.py - TDnet 短信XBRL → financials_cum / guidance.

仕様書 §3-1。ここでの最重要の設計判断は2つ。

1. 短信の会社予想は「予想専用のタグ」では表現されない。実績と **同じ要素名** に
   `..._ForecastMember` という contextRef が付くだけである。したがって値の意味は
   (要素名, contextRef) の組でしか決まらず、パーサは contextRef を
   (年の相対位置 / 四半期 / 連結・単体 / ロール) に必ず分解する。
   これが S5(進捗率・死んだガイダンス検出) の前提になる。

2. マッピングできなかったタグは捨てない。unknown_tags に頻度を積み、
   上位を見て account_mapping.yaml を育てる(仕様書 §3-1)。

Usage
    python -m screener.extract.xbrl_parser --all           # 未処理の短信を全部
    python -m screener.extract.xbrl_parser --date 20260828
    python -m screener.extract.xbrl_parser --unknown-top 40   # 頻度上位のレビュー
"""
from __future__ import annotations

import argparse
import os
import re
import sys
import zipfile
from collections import Counter, defaultdict

from bs4 import BeautifulSoup

try:
    from screener import common as C
except ImportError:                                     # pragma: no cover
    sys.path.insert(0, os.path.dirname(os.path.dirname(
        os.path.dirname(os.path.abspath(__file__)))))
    from screener import common as C

MAPPING_FILE = "account_mapping.yaml"


# --------------------------------------------------------------- mapping
class Mapping:
    def __init__(self, cfg: dict):
        self.cfg = cfg
        self.by_qname: dict[str, str] = {}
        self.by_localname: dict[str, str] = {}
        for item, tags in (cfg.get("items") or {}).items():
            for tag in tags:
                self.by_qname.setdefault(tag, item)
                self.by_localname.setdefault(tag.split(":")[-1], item)
        self.allow_localname = bool(cfg.get("match_localname_when_prefix_unknown"))
        self.noise = tuple(cfg.get("noise_prefixes") or ())
        cx = cfg.get("contexts") or {}
        self.ctx_year = cx.get("year_rel") or {}
        self.ctx_q = cx.get("quarter") or {}
        self.ctx_cons = cx.get("consolidation") or {}
        self.ctx_role = cx.get("role") or {}
        self.prefer_cons = cx.get("prefer_consolidation", "consolidated")

    def item_for(self, qname: str) -> str | None:
        if qname in self.by_qname:
            return self.by_qname[qname]
        if self.allow_localname:
            return self.by_localname.get(qname.split(":")[-1])
        return None

    def is_noise(self, qname: str) -> bool:
        return any(qname.startswith(p) for p in self.noise)

    def parse_context(self, ctx: str) -> dict:
        """contextRef → dims. Unrecognised fragments are returned in `rest` so a
        new taxonomy shows up in review instead of being silently ignored."""
        parts = (ctx or "").split("_")
        out = {"year_rel": None, "q_no": None, "consolidation": None,
               "role": None, "rest": []}
        for p in parts:
            if p in self.ctx_year:
                out["year_rel"] = self.ctx_year[p]
            elif p in self.ctx_q:
                out["q_no"] = self.ctx_q[p]
            elif p in self.ctx_cons:
                out["consolidation"] = self.ctx_cons[p]
            elif p in self.ctx_role:
                out["role"] = self.ctx_role[p]
            else:
                out["rest"].append(p)
        return out


def load_mapping() -> Mapping:
    return Mapping(C.load_yaml(MAPPING_FILE))


# --------------------------------------------------------------- extraction
_NUM = re.compile(r"^-?[\d,]+(\.\d+)?$")


def _to_float(text: str, sign: str | None, scale: str | None):
    t = (text or "").strip().replace(",", "").replace("△", "-").replace("▲", "-")
    if not t or not _NUM.match(t.replace("-", "", 1) if t.startswith("-") else t):
        try:
            v = float(t)
        except (TypeError, ValueError):
            return None
    else:
        v = float(t)
    if scale:
        try:
            v *= 10 ** int(scale)
        except ValueError:
            pass
    if sign == "-":
        v = -v
    return v


def facts_from_zip(zip_path: str) -> list[dict]:
    """Every numeric iXBRL fact in a TDnet zip, tagged with which part it came
    from (Summary = 短信サマリー, Attachment = 財務諸表本体)."""
    out: list[dict] = []
    with zipfile.ZipFile(zip_path) as z:
        for name in z.namelist():
            if not name.endswith("-ixbrl.htm"):
                continue
            part = "summary" if "/Summary/" in name else "attachment"
            try:
                soup = BeautifulSoup(z.read(name).decode("utf-8", "replace"), "lxml-xml")
            except Exception:
                continue
            for t in soup.find_all("nonFraction"):
                qname = t.get("name")
                if not qname:
                    continue
                out.append({
                    "part": part,
                    "tag": qname,
                    "context": t.get("contextRef") or "",
                    "unit": t.get("unitRef") or "",
                    "value": _to_float(t.get_text(strip=True), t.get("sign"),
                                       t.get("scale")),
                    "raw": t.get_text(strip=True)[:32],
                })
    return out


def period_label(filing_row, dims: dict) -> str:
    """A stable period key. The 短信 does not carry the FY label as a fact, so it
    is derived from the disclosure date and the year-relative dimension. This is
    approximate by construction and is why financials_q carries valid_flag."""
    year = int((filing_row["date"] or "1900-01-01")[:4])
    rel = dims.get("year_rel")
    if rel == "prior":
        year -= 1
    elif rel == "prior2":
        year -= 2
    elif rel == "next":
        year += 1
    return f"FY{year}"


# --------------------------------------------------------------- persistence
def store_filing(con, mapping: Mapping, filing_row, facts: list[dict],
                 unknown: Counter, unknown_files: defaultdict) -> dict:
    n_cum = n_guid = n_unknown = n_noise = 0
    seen_unknown_in_file = set()

    for f in facts:
        if f["value"] is None:
            continue
        item = mapping.item_for(f["tag"])
        if item is None:
            src = f"tdnet_{f['part']}"
            if mapping.is_noise(f["tag"]):
                n_noise += 1
            n_unknown += 1
            unknown[(src, f["tag"])] += 1
            key = (src, f["tag"])
            if key not in seen_unknown_in_file:
                unknown_files[key] += 1
                seen_unknown_in_file.add(key)
            continue

        dims = mapping.parse_context(f["context"])
        # 単体しか無い会社もあるので、連結が無いときに単体を落とすことはしない。
        if dims["role"] in ("forecast", "forecast_upper", "forecast_lower"):
            con.execute(
                "INSERT INTO guidance (code, date, fy, item, value, "
                " revision_direction, filing_id) VALUES (?,?,?,?,?,?,?) "
                "ON CONFLICT(code, date, fy, item) DO UPDATE SET "
                " value=excluded.value, filing_id=excluded.filing_id",
                (filing_row["code"], filing_row["date"],
                 period_label(filing_row, dims), item, f["value"],
                 "initial", filing_row["id"]),
            )
            n_guid += 1
        else:
            con.execute(
                "INSERT OR REPLACE INTO financials_cum "
                "(filing_id, code, period, q_no, item, value, unit, context_ref, "
                " source_tag) VALUES (?,?,?,?,?,?,?,?,?)",
                (filing_row["id"], filing_row["code"],
                 period_label(filing_row, dims), dims["q_no"], item, f["value"],
                 f["unit"], f["context"], f["tag"]),
            )
            n_cum += 1
    return {"cum": n_cum, "guidance": n_guid, "unknown": n_unknown, "noise": n_noise}


def flush_unknown(con, unknown: Counter, unknown_files: defaultdict,
                  samples: dict) -> None:
    now = C.utcnow()
    for (src, tag), n in unknown.items():
        con.execute(
            "INSERT INTO unknown_tags (source, tag, n, n_filings, sample_value, "
            " first_seen, last_seen) VALUES (?,?,?,?,?,?,?) "
            "ON CONFLICT(source, tag) DO UPDATE SET "
            " n = unknown_tags.n + excluded.n, "
            " n_filings = unknown_tags.n_filings + excluded.n_filings, "
            " last_seen = excluded.last_seen",
            (src, tag, n, unknown_files[(src, tag)], samples.get((src, tag)),
             now, now),
        )
    con.commit()


# --------------------------------------------------------------------- run
def parse_archive(con, mapping: Mapping, where_sql: str, params: tuple) -> dict:
    rows = con.execute(
        "SELECT id, code, date, subtype, xbrl_path FROM filings "
        "WHERE source='tdnet' AND xbrl_ok=1 AND xbrl_path IS NOT NULL "
        + where_sql + " ORDER BY date, code", params).fetchall()
    C.log(f"parsing {len(rows)} filing(s) with XBRL")

    unknown, unknown_files, samples = Counter(), defaultdict(int), {}
    tot = {"cum": 0, "guidance": 0, "unknown": 0, "noise": 0, "files": 0,
           "failed": 0}
    for r in rows:
        path = os.path.join(C.ROOT, r["xbrl_path"])
        if not os.path.exists(path):
            tot["failed"] += 1
            C.log(f"  ! missing file for filing {r['id']}: {r['xbrl_path']}")
            continue
        try:
            facts = facts_from_zip(path)
        except Exception as e:
            tot["failed"] += 1
            C.log(f"  ! {r['code']} {os.path.basename(path)}: {type(e).__name__}: {e}")
            continue
        for f in facts:
            if f["value"] is not None:
                samples.setdefault((f"tdnet_{f['part']}", f["tag"]), f["raw"])
        got = store_filing(con, mapping, r, facts, unknown, unknown_files)
        for k in ("cum", "guidance", "unknown", "noise"):
            tot[k] += got[k]
        tot["files"] += 1
        con.commit()

    flush_unknown(con, unknown, unknown_files, samples)
    return tot


def unknown_report(con, top: int = 30, source: str | None = None) -> list[dict]:
    sql = ("SELECT source, tag, n, n_filings, sample_value FROM unknown_tags "
           + ("WHERE source=? " if source else "")
           + "ORDER BY n DESC LIMIT ?")
    args = ((source, top) if source else (top,))
    return [dict(r) for r in con.execute(sql, args)]


# 仕様書 §3-1 が「最低限のセット」として名指しした内部項目。
# カバレッジ率はこの集合に対して測る（DB に何行入ったかではなく、
# 必要な科目がどれだけ取れたか、が意味のある指標）。
REQUIRED_ITEMS = (
    "revenue", "cogs", "gross_profit", "sga", "operating_income",
    "ordinary_income", "net_income_parent", "cash_and_deposits",
    "trade_receivables", "merchandise_and_finished_goods", "work_in_process",
    "raw_materials_and_supplies", "inventories_total", "construction_in_progress",
    "machinery_and_equipment", "contract_liabilities", "advances_received",
    "short_term_loans", "long_term_loans", "treasury_stock", "net_assets",
)


def coverage_report(con) -> dict:
    """Per-required-item coverage over the 短信 that carry a full financial
    statement attachment. A 四半期短信 legitimately omits some BS lines, so this
    is a "how often is it there when we look" figure, not a defect count."""
    n_filings = con.execute(
        "SELECT COUNT(DISTINCT filing_id) AS c FROM financials_cum").fetchone()["c"]
    rows = []
    for item in REQUIRED_ITEMS:
        c = con.execute(
            "SELECT COUNT(DISTINCT filing_id) AS c FROM financials_cum WHERE item=?",
            (item,)).fetchone()["c"]
        rows.append({"item": item, "filings": c,
                     "pct": (c / n_filings * 100 if n_filings else 0.0)})
    guid = con.execute(
        "SELECT COUNT(DISTINCT filing_id) AS c FROM guidance").fetchone()["c"]
    return {"n_filings": n_filings, "items": rows, "filings_with_guidance": guid}


def print_coverage(con) -> None:
    rep = coverage_report(con)
    C.log(f"§3-1 required-item coverage over {rep['n_filings']} parsed filing(s) "
          f"({rep['filings_with_guidance']} of them carry 会社予想)")
    for r in rep["items"]:
        bar = "#" * int(r["pct"] / 5)
        C.log(f"  {r['item']:<32} {r['filings']:>3}/{rep['n_filings']:<3} "
              f"{r['pct']:5.1f}%  {bar}")


def print_unknown(con, top: int) -> None:
    mapping = load_mapping()
    rows = unknown_report(con, top)
    total = con.execute("SELECT COALESCE(SUM(n),0) AS s FROM unknown_tags").fetchone()["s"]
    distinct = con.execute("SELECT COUNT(*) AS c FROM unknown_tags").fetchone()["c"]
    mapped = con.execute("SELECT COUNT(*) AS c FROM financials_cum").fetchone()["c"]
    C.log(f"unknown tags: {distinct} distinct / {total} occurrences "
          f"(mapped facts stored: {mapped})")
    for r in con.execute("SELECT source, COUNT(*) AS d, SUM(n) AS n "
                         "FROM unknown_tags GROUP BY source ORDER BY n DESC"):
        C.log(f"  by source: {r['source']:<20} {r['d']:>4} distinct / {r['n']:>5} occ.")
    C.log(f"top {len(rows)} by frequency  [noise = account_mapping.yaml の "
          f"noise_prefixes 該当 = 意図的に項目化していない]")
    C.log(f"  {'n':>5} {'files':>5} {'noise':>5}  {'source':<18} tag")
    for r in rows:
        C.log(f"  {r['n']:>5} {r['n_filings']:>5} "
              f"{'yes' if mapping.is_noise(r['tag']) else '   ':>5}  "
              f"{r['source']:<18} {r['tag']}")


def main(argv=None) -> int:
    sys.stdout.reconfigure(encoding="utf-8")
    p = argparse.ArgumentParser(description="TDnet 短信XBRL parser (仕様書 §3-1)")
    p.add_argument("--all", action="store_true", help="parse every archived filing")
    p.add_argument("--date", help="parse one disclosure date (YYYYMMDD)")
    p.add_argument("--code", help="parse one securities code")
    p.add_argument("--unknown-top", type=int, metavar="N",
                   help="print the top-N unmapped tags and exit")
    p.add_argument("--coverage", action="store_true",
                   help="print §3-1 required-item coverage and exit")
    p.add_argument("--reset", action="store_true",
                   help="clear financials_cum / guidance / unknown_tags first")
    a = p.parse_args(argv)

    con = C.init_db()
    if a.unknown_top:
        print_unknown(con, a.unknown_top)
        return 0
    if a.coverage:
        print_coverage(con)
        return 0
    if a.reset:
        for t in ("financials_cum", "guidance", "unknown_tags"):
            con.execute(f"DELETE FROM {t}")
        con.commit()
        C.log("cleared financials_cum / guidance / unknown_tags")

    where, params = "", ()
    if a.date:
        where, params = " AND date=?", (C.parse_date_arg(a.date).isoformat(),)
    elif a.code:
        where, params = " AND code=?", (a.code,)
    elif not a.all:
        p.error("choose one of --all / --date / --code / --unknown-top")

    tot = parse_archive(con, load_mapping(), where, params)
    C.log(f"parsed {tot['files']} file(s): {tot['cum']} cum facts, "
          f"{tot['guidance']} guidance facts, {tot['unknown']} unmapped "
          f"({tot['noise']} of them配当明細等のノイズ), {tot['failed']} failed")
    print_coverage(con)
    print_unknown(con, 25)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
