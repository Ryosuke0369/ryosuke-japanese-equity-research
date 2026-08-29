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

$logDir = Join-Path $repo "screener\data\logs"
New-Item -ItemType Directory -Force -Path $logDir | Out-Null
$log = Join-Path $logDir ("run_daily_{0}.log" -f (Get-Date -Format "yyyyMM"))

function Write-Log([string]$m) {
    $line = "[{0}] {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $m
    Write-Output $line
    Add-Content -Path $log -Value $line -Encoding utf8
}

Write-Log "=== run_daily start (backfill $BackfillDays days) ==="

& $PythonExe -m screener.fetch.tdnet_archiver --backfill $BackfillDays 2>&1 |
    ForEach-Object { Add-Content -Path $log -Value $_ -Encoding utf8; Write-Output $_ }
$archiveExit = $LASTEXITCODE
Write-Log "archiver exit=$archiveExit"

if (-not $SkipParse) {
    & $PythonExe -m screener.extract.xbrl_parser --all 2>&1 |
        ForEach-Object { Add-Content -Path $log -Value $_ -Encoding utf8; Write-Output $_ }
    Write-Log "parser exit=$LASTEXITCODE"
}

# Coverage is printed on every run: a silent job that has stopped collecting is
# the failure mode this whole design is trying to avoid.
& $PythonExe -m screener.fetch.tdnet_archiver --report 14 2>&1 |
    ForEach-Object { Add-Content -Path $log -Value $_ -Encoding utf8; Write-Output $_ }
$reportExit = $LASTEXITCODE

Write-Log "=== run_daily end (archiver=$archiveExit, coverage=$reportExit) ==="
if ($archiveExit -ne 0 -or $reportExit -ne 0) { exit 1 }
exit 0
