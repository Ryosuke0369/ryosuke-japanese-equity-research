<#
    screener/run_paper_weekly.ps1 - ペーパートレード週次バッチのランナー。

    設計の正本: docs/paper_trading_design.md。
    毎週月曜 06:00 にタスクスケジューラから起動される想定。

    リトライは1回まで:
      無限にリトライすると、壊れたまま毎週回り続けて「動いているつもり」に
      なる。2回落ちたら止めて、火曜のチェックに ERROR を拾わせる。

    writer.lock で日次ジョブ(19:00/23:15)と直列化される。同時に走ると
    database is locked で両方壊れる(2026-09-01 に実際に踏んだ)。

        powershell -ExecutionPolicy Bypass -File screener/run_paper_weekly.ps1
        powershell -ExecutionPolicy Bypass -File screener/run_paper_weekly.ps1 -DryRun
#>
[CmdletBinding()]
param(
    [string]$AsOf = "",
    [string]$PythonExe = "",
    [switch]$DryRun,
    [switch]$NoSleepGuard
)

$ErrorActionPreference = "Stop"
$repo = Split-Path -Parent $PSScriptRoot
Set-Location $repo
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

if (-not $PythonExe) { $PythonExe = (Get-Command python -ErrorAction SilentlyContinue).Source }
if (-not $PythonExe) { throw "python not found on PATH; pass -PythonExe" }

$logDir = & $PythonExe -c "import screener.common as C; print(C.LOG_DIR)"
if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($logDir)) {
    throw "could not resolve LOG_DIR from screener.common"
}
$logDir = $logDir.Trim()
$log = Join-Path $logDir ("paper_weekly_{0}.log" -f (Get-Date -Format "yyyyMM"))

function Write-Log([string]$text) {
    $line = "[{0}] {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $text
    Write-Output $line
    Add-Content -Path $log -Value $line -Encoding utf8
}

if (-not $NoSleepGuard) {
    $sig = @"
[DllImport("kernel32.dll", SetLastError = true)]
public static extern uint SetThreadExecutionState(uint esFlags);
"@
    try {
        $p = Add-Type -MemberDefinition $sig -Name Power -Namespace Win32Paper -PassThru
        [void]$p::SetThreadExecutionState([uint32]"0x80000001")
    } catch { Write-Log "WARNING: sleep 抑止に失敗 ($_)" }
}

$args = @("-m", "screener.report.paper_weekly", "--lock-wait", "1800")
if ($AsOf)  { $args += @("--as-of", $AsOf) }
if ($DryRun) { $args += "--dry-run" }

$exit = 1
for ($try = 1; $try -le 2; $try++) {
    Write-Log "=== paper_weekly 実行 (試行 $try/2) ==="
    $prev = $ErrorActionPreference
    $ErrorActionPreference = "Continue"
    try {
        & $PythonExe @args 2>&1 | ForEach-Object { Write-Log ([string]$_) }
    } finally { $ErrorActionPreference = $prev }
    if ($LASTEXITCODE -eq 0) { $exit = 0; break }
    Write-Log "FAILED (exit $LASTEXITCODE)"
    if ($try -eq 1) { Write-Log "60秒待って1回だけ再試行する"; Start-Sleep -Seconds 60 }
}
if ($exit -ne 0) { Write-Log "ERROR: paper_weekly が2回とも失敗した" }
Write-Log "=== paper_weekly 終了 (exit $exit) ==="
exit $exit
