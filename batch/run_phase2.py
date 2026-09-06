"""フェーズ2 §4 — regenerate every ticker of the batch under one set of rules.

Every ticker is reprocessed, including the ones already marked done: the point
of the exercise is that all 85 models come out of the SAME pipeline revision and
the SAME beta rule, so "already generated" is not a reason to skip. `--force` is
always passed, and `--date` pins the analysis-basis date so a run that crosses
midnight still writes one file per ticker (フェーズ2 #4).

State lives in batch/phase2_state.json and is written after every ticker, so an
interrupted run resumes where it stopped (`--resume`). Per-ticker stdout goes to
batch/logs_phase2/<code>.log — the console only carries the verdict line, but
nothing is thrown away.

Parallelism is safe for the recalc step: scripts/recalc_excel_com.py uses
DispatchEx, which starts its own Excel instance per process rather than sharing
one. Keep it modest anyway - EDINET and yfinance are shared, rate-limited.

Usage:
    python batch/run_phase2.py --workers 3
    python batch/run_phase2.py --resume --workers 3
    python batch/run_phase2.py --only 2802 2502          # a specific set
"""
import argparse
import concurrent.futures as cf
import json
import os
import re
import subprocess
import sys
import threading
import time

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
LOGDIR = os.path.join(HERE, "logs_phase2")
STATE = os.path.join(HERE, "phase2_state.json")

VERDICT_RE = re.compile(r"^FAIL (\d+) / WARN (\d+) / SKIP (\d+) / PASS (\d+)", re.M)
_lock = threading.Lock()


def load_state():
    return json.load(open(STATE, encoding="utf-8")) if os.path.isfile(STATE) else {}


def save_state(state):
    with _lock:
        tmp = STATE + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(state, f, ensure_ascii=False, indent=1)
        os.replace(tmp, STATE)


def run_one(code, date):
    os.makedirs(LOGDIR, exist_ok=True)
    log = os.path.join(LOGDIR, f"{code}.log")
    t0 = time.time()
    p = subprocess.run(
        [sys.executable, "-u", os.path.join(ROOT, "scripts", "generate_dcf.py"),
         code, "--force", "--date", date],
        cwd=ROOT, capture_output=True, text=True, encoding="utf-8",
        errors="replace", timeout=2400,
    )
    out = (p.stdout or "") + (p.stderr or "")
    with open(log, "w", encoding="utf-8") as f:
        f.write(out)

    m = VERDICT_RE.search(out)
    rec = {
        "exit": p.returncode,
        "seconds": round(time.time() - t0, 1),
        "fail": int(m.group(1)) if m else None,
        "warn": int(m.group(2)) if m else None,
        "skip": int(m.group(3)) if m else None,
        "pass": int(m.group(4)) if m else None,
        "log": os.path.relpath(log, ROOT),
    }
    for key, pat in (
        ("guidance", r"guidance source: (.+)"),
        ("beta_line", r"Beta: raw (.+)"),
        ("wacc", r"=> WACC:\s+([0-9.]+)%"),
        ("no_guidance", r"(NO GUIDANCE - Management scenario falls back)"),
    ):
        mm = re.search(pat, out)
        if mm:
            rec[key] = mm.group(1).strip()
    # Every WARNING line, so nothing the generator said is lost to the summary.
    rec["warnings"] = sorted({w.strip()[:200] for w in
                              re.findall(r"^\s*(WARNING:.*)$", out, re.M)})
    if p.returncode != 0:
        tail = [l for l in out.splitlines() if l.strip()][-6:]
        rec["error_tail"] = tail
    return rec


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--date", default="20260906")
    ap.add_argument("--workers", type=int, default=3)
    ap.add_argument("--resume", action="store_true",
                    help="skip tickers already recorded with exit 0")
    ap.add_argument("--only", nargs="*", default=None)
    ap.add_argument("--limit", type=int, default=None)
    ap.add_argument("--state-in", default=os.path.join(HERE, "batch_state.json"))
    ap.add_argument("--status", default="done")
    a = ap.parse_args()

    src = json.load(open(a.state_in, encoding="utf-8"))
    codes = sorted(k for k, v in src.items()
                   if isinstance(v, dict) and v.get("status") == a.status)
    if a.only:
        codes = [c for c in a.only]
    state = load_state()
    if a.resume:
        codes = [c for c in codes if state.get(c, {}).get("exit") != 0]
    if a.limit:
        codes = codes[:a.limit]

    print(f"phase2 regeneration: {len(codes)} ticker(s), date={a.date}, "
          f"workers={a.workers}")
    t0 = time.time()
    done = 0
    with cf.ThreadPoolExecutor(max_workers=a.workers) as ex:
        futs = {ex.submit(run_one, c, a.date): c for c in codes}
        for fut in cf.as_completed(futs):
            code = futs[fut]
            try:
                rec = fut.result()
            except Exception as e:
                rec = {"exit": -1, "error_tail": [f"{type(e).__name__}: {e}"]}
            state[code] = rec
            save_state(state)
            done += 1
            el = time.time() - t0
            eta = (el / done) * (len(codes) - done)
            flag = ("OK  " if rec.get("exit") == 0 else f"EXIT{rec.get('exit')}")
            print(f"[{done:>3}/{len(codes)}] {flag} {code}  "
                  f"FAIL {rec.get('fail')} WARN {rec.get('warn')} "
                  f"SKIP {rec.get('skip')} PASS {rec.get('pass')}  "
                  f"{rec.get('seconds')}s   ETA {eta/60:.0f}m", flush=True)

    ok = [c for c in codes if state.get(c, {}).get("exit") == 0]
    bad = [c for c in codes if state.get(c, {}).get("exit") != 0]
    print(f"\nfinished in {(time.time() - t0)/60:.1f} min: "
          f"{len(ok)} ok / {len(bad)} failed")
    if bad:
        print("  failed: " + ", ".join(bad))
    print(f"state: {STATE}")


if __name__ == "__main__":
    main()
