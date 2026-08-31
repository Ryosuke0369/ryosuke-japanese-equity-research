#!/usr/bin/env python
"""tasks/ を追記専用にする PreToolUse フック。

なぜ必要か
----------
tasks/todo.md と tasks/lessons.md は「積み上げた作業ログ」で、消えると
git に入っていない分は復元できない。2026-08-31 に2回目の全文上書き事故が
起きたため、注意ではなく仕組みで止める。

事故の実際の経路は Write ツールではなく **Bash のリダイレクト** だった。
Write ツールには「読まずに既存ファイルを上書きすると失敗する」ガードが
あるが、`cat > tasks/todo.md` はそのガードを丸ごと迂回する。
したがって Write と Bash の両方を見る。

判定
----
- 既存の tasks/ 配下ファイルへの Write        -> deny
- tasks/ 配下への切り詰めリダイレクト `>`     -> deny (`>>` は許可)
- tasks/ 配下への tee (-a 無し)               -> deny
- tasks/ 配下への sed -i / cp / mv            -> deny
- 新規ファイルの作成、追記、読み取り          -> allow

deny されたら: まず全文を読み、追記(`>>` / Edit)にするか、どうしても
置換が要るなら Read してから Edit で差分を当てる。
"""
from __future__ import annotations

import json
import os
import re
import sys

TASKS_RE = re.compile(r"(?:^|[\s'\"=/\\])tasks[/\\][^\s;|&'\"()]+")
# `>` だが `>>` ではないもの。`2>` `&>` のような fd 指定も切り詰めなので含む。
TRUNC_REDIR_RE = re.compile(r"(?<![>0-9&])>(?!>)\s*['\"]?([^\s;|&'\"()]+)")


def _decision(allow: bool, reason: str = "") -> None:
    # Windows の既定コードページは cp932。理由文に日本語が入るので、明示的に
    # UTF-8 にしないと UnicodeEncodeError でフック自体が落ちる。
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except (AttributeError, OSError):
        pass
    if allow:
        print(json.dumps({"hookSpecificOutput": {
            "hookEventName": "PreToolUse", "permissionDecision": "allow",
            "permissionDecisionReason": reason}}))
    else:
        print(json.dumps({"hookSpecificOutput": {
            "hookEventName": "PreToolUse", "permissionDecision": "deny",
            "permissionDecisionReason": reason}}))
    sys.exit(0)


def _under_tasks(path: str, cwd: str) -> bool:
    if not path:
        return False
    p = path.replace("\\", "/")
    if not os.path.isabs(p):
        p = os.path.join(cwd, p).replace("\\", "/")
    p = os.path.normpath(p).replace("\\", "/")
    return "/tasks/" in p + "/"


def _exists(path: str, cwd: str) -> bool:
    p = path if os.path.isabs(path) else os.path.join(cwd, path)
    return os.path.isfile(p)


GUIDANCE = (
    "tasks/ は追記専用です（2026-08-31 に全文上書き事故が2回起きたため仕組みで"
    "禁止しています）。既存の作業ログを丸ごと置き換えると、git に入っていない"
    "差分は復元できません。\n"
    "  追記したい      -> `>>` かReadしてからEditで末尾に足す\n"
    "  一部を直したい  -> 先に全文を Read し、Edit で該当箇所だけ置換する\n"
    "  本当に作り直す  -> 先に `git diff -- tasks/` で未コミット差分を確認し、"
    "ユーザーの明示の了解を取ってから行う"
)


def main() -> int:
    try:
        payload = json.load(sys.stdin)
    except (ValueError, OSError):
        _decision(True, "hook input を解釈できなかったので素通しする")
    tool = payload.get("tool_name") or ""
    ti = payload.get("tool_input") or {}
    cwd = payload.get("cwd") or os.getcwd()

    if tool == "Write":
        path = ti.get("file_path") or ""
        if _under_tasks(path, cwd) and _exists(path, cwd):
            _decision(False, f"Write が tasks/ 配下の既存ファイル "
                             f"({os.path.basename(path)}) を上書きしようとしています。\n"
                             + GUIDANCE)
        _decision(True)

    if tool in ("Bash", "PowerShell"):
        cmd = ti.get("command") or ""
        if "tasks" not in cmd:
            _decision(True)
        hits = []
        for m in TRUNC_REDIR_RE.finditer(cmd):
            tgt = m.group(1)
            if _under_tasks(tgt, cwd) and _exists(tgt, cwd):
                hits.append(f"切り詰めリダイレクト `> {tgt}`")
        low = cmd.lower()
        for tok, label in (("tee", "tee"), ("sed -i", "sed -i"),
                           ("set-content", "Set-Content"),
                           ("out-file", "Out-File")):
            if tok in low and not re.search(r"(-a\b|--append|-Append)", cmd, re.I):
                for m in TASKS_RE.finditer(cmd):
                    tgt = m.group(0).strip(" '\"=")
                    if _under_tasks(tgt, cwd) and _exists(tgt, cwd):
                        hits.append(f"{label} による上書き ({tgt})")
                        break
        for m in re.finditer(r"\b(cp|mv|copy|move|Copy-Item|Move-Item)\s+[^;|&]*",
                             cmd):
            seg = m.group(0)
            parts = [p for p in seg.split() if not p.startswith("-")][1:]
            if parts and _under_tasks(parts[-1], cwd) and _exists(parts[-1], cwd):
                hits.append(f"{m.group(1)} による上書き ({parts[-1]})")
        if hits:
            _decision(False, "tasks/ 配下の既存ファイルを上書きしようとしています: "
                             + " / ".join(sorted(set(hits))) + "\n" + GUIDANCE)
        _decision(True)

    _decision(True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
