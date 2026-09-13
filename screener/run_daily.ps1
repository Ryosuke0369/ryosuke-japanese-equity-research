<#
    screener/run_daily.ps1 - the daily job Task Scheduler runs (仕様書 §2-1
    「cron/タスクスケジューラで平日夕方1回+リトライ」).

    Why it is a --backfill and not a --today:
      TDnet keeps only ~1 month, so a day missed for any reason (laptop asleep,
      network down, TDnet 5xx) is unrecoverable after that window. The job
      therefore does not archive "today", it archives *every weekday in the last
      N days that is not already recorded as ok/empty/partial in fetch_runs*.
      A machine that was off for a week catches up on its own, and the same run
      is safe to fire twice (already-downloaded files are skipped).

    Exit code is propagated so a failed day shows as a failed task in the
    Task Scheduler history instead of disappearing.
#>
[CmdletBinding()]
param(
    [int]$BackfillDays = 14,
    [string]$PythonExe = "",
    [switch]$SkipParse
)

$ErrorActionPreference = "Stop"
$repo = Split-Path -Parent $PSScriptRoot
Set-Location $repo

if (-not $PythonExe) {
    $PythonExe = (Get-Command python -ErrorAction SilentlyContinue).Source
}
if (-not $PythonExe) { throw "python not found on PATH; pass -PythonExe" }

# The log directory is NOT hard-coded: screener/common.py resolves it from
# DATA_ROOT in .env (the data tree lives on D:, C: has no room). Asking Python
# keeps one source of truth -- a second copy of the rule here is how the job
# ends up writing its log to a drive the archive no longer uses.
# No 2>$null here: in PS 5.1, redirecting a native exe's stderr under
# $ErrorActionPreference='Stop' raises NativeCommandError before the check
# below can run. Let a traceback print - it is the diagnosis.
$logDir = & $PythonExe -c "import screener.common as C; print(C.LOG_DIR)"
if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($logDir)) {
    # Fail loudly rather than quietly logging somewhere else: if common.py
    # cannot even be imported, the run itself is not going to work.
    throw "cannot resolve screener LOG_DIR via $PythonExe (is the repo on sys.path?)"
}
$logDir = $logDir.Trim()
New-Item -ItemType Directory -Force -Path $logDir | Out-Null
$log = Join-Path $logDir ("run_daily_{0}.log" -f (Get-Date -Format "yyyyMM"))

function Write-Log([string]$m) {
    $line = "[{0}] {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $m
    Write-Output $line
    Add-Content -Path $log -Value $line -Encoding utf8
}

Write-Log "=== run_daily start (backfill $BackfillDays days) ==="
Write-Log "data root: $(& $PythonExe -c ""import screener.common as C; print(C.DATA_DIR)"")"

& $PythonExe -m screener.fetch.tdnet_archiver --backfill $BackfillDays 2>&1 |
    ForEach-Object { Add-Content -Path $log -Value $_ -Encoding utf8; Write-Output $_ }
$archiveExit = $LASTEXITCODE
Write-Log "archiver exit=$archiveExit"

if (-not $SkipParse) {
    & $PythonExe -m screener.extract.xbrl_parser --all 2>&1 |
        ForEach-Object { Add-Content -Path $log -Value $_ -Encoding utf8; Write-Output $_ }
    Write-Log "parser exit=$LASTEXITCODE"
}

# 非決算の適時開示タイトル（2026-09-13 追加。9/02 以降止まっていた）。
# 一覧は約1ヶ月しか遡れないので、日次で直近5日を取り直す（冪等）。
& $PythonExe -m screener.fetch.tdnet_titles --days 5 2>&1 |
    ForEach-Object { Add-Content -Path $log -Value $_ -Encoding utf8; Write-Output $_ }
$titlesExit = $LASTEXITCODE
Write-Log "titles exit=$titlesExit"

# EDINET 日次（2026-09-13 追加。8/31 で停止し日次ジョブが無かった）。
# 直近7日を再索引 + 30日内の欠損日を補完 → 現ユニバースの未取得を取得 → その日付だけ解析。
& $PythonExe -m screener.fetch.edinet_bulk --recent 7 2>&1 |
    ForEach-Object { Add-Content -Path $log -Value $_ -Encoding utf8; Write-Output $_ }
$edinetExit = $LASTEXITCODE
Write-Log "edinet exit=$edinetExit"
if (-not $SkipParse) {
    # 1回で済ませる。xbrl_parser は解析後に全DBを舐める coverage/unknown 集計を毎回走らせ、
    # それだけで約10分かかる（0件でも同じ。2026-09-13 実測）。日付ごとに呼ぶと8回ぶん払う。
    # --resume = financials_cum に行が無い（未解析の）EDINET 書類だけを解析する。
    & $PythonExe -m screener.extract.xbrl_parser --source edinet --all --resume --lock-wait 3600 2>&1 |
        ForEach-Object { Add-Content -Path $log -Value $_ -Encoding utf8; Write-Output $_ }
    Write-Log "edinet parser exit=$LASTEXITCODE"
}

# Coverage is printed on every run: a silent job that has stopped collecting is
# the failure mode this whole design is trying to avoid.
& $PythonExe -m screener.fetch.tdnet_archiver --report 14 2>&1 |
    ForEach-Object { Add-Content -Path $log -Value $_ -Encoding utf8; Write-Output $_ }
$reportExit = $LASTEXITCODE

Write-Log "=== run_daily end (archiver=$archiveExit, coverage=$reportExit, titles=$titlesExit, edinet=$edinetExit) ==="
if ($archiveExit -ne 0 -or $reportExit -ne 0 -or $titlesExit -ne 0 -or $edinetExit -ne 0) { exit 1 }
exit 0
