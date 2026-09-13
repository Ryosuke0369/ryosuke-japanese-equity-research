<#
    screener/check_paper_weekly.ps1 - 週次バッチが走ったかの健全性チェック。

    毎週火曜 06:00 に起動され、「今週の月曜の週次バッチが成功しているか」を
    fetch_runs(source='paper_weekly') で確認する。無ければログに ERROR を
    残す。次にユーザーがセッションを開いたとき、ここを見れば気づける。

    チェック自体は何も直さない。直すのは人間の仕事で、
    このスクリプトの仕事は**黙って壊れている状態を作らないこと**。
#>
[CmdletBinding()]
param([string]$PythonExe = "")

$ErrorActionPreference = "Stop"
$repo = Split-Path -Parent $PSScriptRoot
Set-Location $repo
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

if (-not $PythonExe) { $PythonExe = (Get-Command python -ErrorAction SilentlyContinue).Source }
if (-not $PythonExe) { throw "python not found on PATH" }

$logDir = (& $PythonExe -c "import screener.common as C; print(C.LOG_DIR)").Trim()
$log = Join-Path $logDir ("paper_weekly_{0}.log" -f (Get-Date -Format "yyyyMM"))

# チェック本体は Python モジュールに置く。python -c に複数行を渡すと
# ネイティブ実行への引数組み立てで改行が壊れる（2026-09-02 に実際に踏んだ）。
$prev = $ErrorActionPreference
$ErrorActionPreference = "Continue"
$out = & $PythonExe -m screener.report.check_paper_weekly 2>&1
$rc = $LASTEXITCODE
$ErrorActionPreference = $prev
foreach ($line in $out) {
    $l = [string]$line
    Write-Output $l
    Add-Content -Path $log -Value $l -Encoding utf8
}
if ($rc -ne 0) { Write-Output "健全性チェック: 異常 (exit $rc)" }
exit $rc
