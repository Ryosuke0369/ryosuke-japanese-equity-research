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
        self.ctx_q_by_ctx = cx.get("quarter_by_context") or {}
        self.default_cons = cx.get("default_consolidation_by_source") or {}

    def item_for(self, qname: str) -> str | None:
        if qname in self.by_qname:
            return self.by_qname[qname]
        if self.allow_localname:
            return self.by_localname.get(qname.split(":")[-1])
        return None

    def is_noise(self, qname: str) -> bool:
        return any(qname.startswith(p) for p in self.noise)

    def parse_context(self, ctx: str, source: str = "tdnet") -> dict:
        """contextRef → dims. Unrecognised fragments are returned in `rest` so a
        new taxonomy shows up in review instead of being silently ignored.

        EDINET は四半期を member ではなくコンテキスト名そのもの
        (InterimDuration / CurrentYearDuration) で表し、連結には member を
        付けない。短信の語彙だけで読むと q_no も consolidation も None になり、
        「読めているのに何期のものか分からない」行が量産される。
        """
        parts = (ctx or "").split("_")
        out = {"year_rel": None, "q_no": None, "consolidation": None,
               "role": None, "rest": []}
        for p in parts:
            if p in self.ctx_year:
                out["year_rel"] = self.ctx_year[p]
                if out["q_no"] is None and p in self.ctx_q_by_ctx:
                    out["q_no"] = self.ctx_q_by_ctx[p]
            elif p in self.ctx_q:
                out["q_no"] = self.ctx_q[p]
            elif p in self.ctx_cons:
                out["consolidation"] = self.ctx_cons[p]
            elif p in self.ctx_role:
                out["role"] = self.ctx_role[p]
            else:
                out["rest"].append(p)
        if out["consolidation"] is None and source in self.default_cons:
            out["consolidation"] = self.default_cons[source]
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


def _ixbrl_members(names: list[str], source: str) -> list[tuple[str, str]]:
    """zip 内の iXBRL ファイルと、その「部位」ラベル。

    TDnet: `-ixbrl.htm`。/Summary/ が短信サマリー、他が財務諸表本体。
    EDINET: `_ixbrl.htm`(区切りがハイフンではなくアンダースコア)。
            AuditDoc は監査報告書で財務数値を持たないので読まない ——
            読むと監査文言のタグが unknown を無意味に膨らませる。
    """
    out = []
    for n in names:
        if source == "edinet":
            if not n.endswith("_ixbrl.htm") or "/AuditDoc/" in n:
                continue
            out.append((n, "publicdoc"))
        else:
            if not n.endswith("-ixbrl.htm"):
                continue
            out.append((n, "summary" if "/Summary/" in n else "attachment"))
    return out


def edinet_fy_end(zip_path: str) -> str | None:
    """EDINET の DEI から当期の決算期末日 (YYYY-MM-DD) を取る。

    短信は会計期間をファクトとして持たないので提出日から推定するしかないが、
    EDINET は jpdei_cor:CurrentFiscalYearEndDateDEI を持っている。12月期・
    11月期の会社は提出年と会計年度がずれるので、推定ではなく実値を使う。
    """
    try:
        with zipfile.ZipFile(zip_path) as z:
            for name, _ in _ixbrl_members(z.namelist(), "edinet"):
                soup = BeautifulSoup(z.read(name).decode("utf-8", "replace"),
                                     "lxml-xml")
                for t in soup.find_all("nonNumeric"):
                    if (t.get("name") or "").endswith(
                            ":CurrentFiscalYearEndDateDEI"):
                        v = t.get_text(strip=True)
                        return v[:10] if v else None
    except Exception:
        return None
    return None


def facts_from_zip(zip_path: str, source: str = "tdnet") -> list[dict]:
    """Every numeric iXBRL fact in a zip, tagged with which part it came from."""
    out: list[dict] = []
    with zipfile.ZipFile(zip_path) as z:
        for name, part in _ixbrl_members(z.namelist(), source):
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


_YEAR_OFFSET = {"prior": -1, "prior2": -2, "prior3": -3, "prior4": -4,
                "next": 1, "current": 0, None: 0}


def period_label(filing_row, dims: dict, fy_end: str | None = None) -> str:
    """A stable period key (FY<year>).

    短信は会計期間をファクトとして持たないので開示日から推定するしかなく、
    それが financials_q に valid_flag がある理由。EDINET は DEI に当期の
    決算期末日を持っているので、渡されたらそちらを基準にする —— 12月期・
    11月期の会社は提出年と会計年度がずれ、開示日推定だと1年ずれる。
    """
    base = fy_end or filing_row["date"] or "1900-01-01"
    year = int(str(base)[:4])
    return f"FY{year + _YEAR_OFFSET.get(dims.get('year_rel'), 0)}"


# --------------------------------------------------------------- persistence
def store_filing(con, mapping: Mapping, filing_row, facts: list[dict],
                 unknown: Counter, unknown_files: defaultdict,
                 source: str = "tdnet", fy_end: str | None = None) -> dict:
    n_cum = n_guid = n_unknown = n_noise = n_dim = 0
    seen_unknown_in_file = set()

    for f in facts:
        if f["value"] is None:
            continue
        item = mapping.item_for(f["tag"])
        if item is None:
            src = f"{source}_{f['part']}"
            if mapping.is_noise(f["tag"]):
                n_noise += 1
            n_unknown += 1
            unknown[(src, f["tag"])] += 1
            key = (src, f["tag"])
            if key not in seen_unknown_in_file:
                unknown_files[key] += 1
                seen_unknown_in_file.add(key)
            continue

        dims = mapping.parse_context(f["context"], source)

        # 未知の ...Member が付いた文脈は「見出し数値」ではなく内訳
        # （セグメント別、株主資本等変動計算書の資本金/利益剰余金、大株主 等）。
        # 同じ item 名で見出しと内訳が混ざると、下流が net_assets を引いたとき
        # に資本金や自己株式まで一緒に返る。落とすのではなく数えて報告する。
        breakdown = [p for p in dims["rest"] if p.endswith("Member")]
        if breakdown:
            n_dim += 1
            if source == "edinet":
                # EDINET は有報の変動計算書だけで数十の内訳を持つ。見出しを
                # 汚すので入れない。内訳自体は将来セグメント分析で使うので、
                # 件数を残して「取り込んでいない」ことを可視にしておく。
                continue
            # TDnet は既存パイプライン(quarterly_builder)が現在の挙動を前提に
            # しているため、ここでは落とさず件数だけ数える。同じ問題はあるので
            # 別途対応する。
        # 単体しか無い会社もあるので、連結が無いときに単体を落とすことはしない。
        if dims["role"] in ("forecast", "forecast_upper", "forecast_lower"):
            con.execute(
                "INSERT INTO guidance (code, date, fy, item, value, "
                " revision_direction, filing_id) VALUES (?,?,?,?,?,?,?) "
                "ON CONFLICT(code, date, fy, item) DO UPDATE SET "
                " value=excluded.value, filing_id=excluded.filing_id",
                (filing_row["code"], filing_row["date"],
                 period_label(filing_row, dims, fy_end), item, f["value"],
                 "initial", filing_row["id"]),
            )
            n_guid += 1
        else:
            con.execute(
                "INSERT OR REPLACE INTO financials_cum "
                "(filing_id, code, period, q_no, item, value, unit, context_ref, "
                " source_tag) VALUES (?,?,?,?,?,?,?,?,?)",
                (filing_row["id"], filing_row["code"],
                 period_label(filing_row, dims, fy_end), dims["q_no"], item, f["value"],
                 f["unit"], f["context"], f["tag"]),
            )
            n_cum += 1
    return {"cum": n_cum, "guidance": n_guid, "unknown": n_unknown,
            "noise": n_noise, "dimensional": n_dim}


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
def parse_archive(con, mapping: Mapping, where_sql: str, params: tuple,
                  source: str = "tdnet") -> dict:
    """source='tdnet' は短信、'edinet' は有報/半期。

    かつてここは source='tdnet' に固定されていて、EDINET を落としても
    永久に解析されなかった —— unknown_tags が空でも「未マップが無い」ので
    はなく「一度も見ていない」だけ、という状態になっていた(2026-08-31 修正)。
    """
    rows = con.execute(
        "SELECT id, code, date, subtype, xbrl_path FROM filings "
        "WHERE source=? AND xbrl_ok=1 AND xbrl_path IS NOT NULL "
        + where_sql + " ORDER BY date, code", (source,) + params).fetchall()
    C.log(f"parsing {len(rows)} {source} filing(s) with XBRL")

    unknown, unknown_files, samples = Counter(), defaultdict(int), {}
    tot = {"cum": 0, "guidance": 0, "unknown": 0, "noise": 0, "files": 0,
           "failed": 0, "dimensional": 0}
    for r in rows:
        path = C.full_path(r["xbrl_path"])
        if not os.path.exists(path):
            tot["failed"] += 1
            C.log(f"  ! missing file for filing {r['id']}: {r['xbrl_path']}")
            continue
        try:
            facts = facts_from_zip(path, source)
        except Exception as e:
            tot["failed"] += 1
            C.log(f"  ! {r['code']} {os.path.basename(path)}: {type(e).__name__}: {e}")
            continue
        for f in facts:
            if f["value"] is not None:
                samples.setdefault((f"{source}_{f['part']}", f["tag"]), f["raw"])
        # EDINET は会計期間を DEI に持っている。開示日からの推定だと 12月期・
        # 11月期の会社で1年ずれるので、1書類につき一度だけ読んで渡す。
        fy_end = edinet_fy_end(path) if source == "edinet" else None
        got = store_filing(con, mapping, r, facts, unknown, unknown_files,
                           source, fy_end)
        for k in ("cum", "guidance", "unknown", "noise", "dimensional"):
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
    p.add_argument("--source", choices=("tdnet", "edinet", "all"),
                   default="tdnet",
                   help="解析対象。tdnet=短信 / edinet=有報・半期 / all=両方")
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

    sources = ("tdnet", "edinet") if a.source == "all" else (a.source,)
    mapping = load_mapping()
    tot = {"files": 0, "cum": 0, "guidance": 0, "unknown": 0,
           "noise": 0, "failed": 0, "dimensional": 0}
    for src in sources:
        got = parse_archive(con, mapping, where, params, src)
        for k in tot:
            tot[k] += got.get(k, 0)
    C.log(f"parsed {tot['files']} file(s): {tot['cum']} cum facts, "
          f"{tot['guidance']} guidance facts, {tot['unknown']} unmapped "
          f"({tot['noise']} of them配当明細等のノイズ), "
          f"{tot['dimensional']} dimensional breakdown facts, "
          f"{tot['failed']} failed")
    print_coverage(con)
    print_unknown(con, 25)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
