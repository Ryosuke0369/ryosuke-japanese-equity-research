<#
    screener/run_edinet_full.ps1 - 確定ユニバース x N年 の EDINET 一括取得
    (仕様書 §2-2)。夜間に一度流す想定で、TDnet の run_daily.ps1 とは別物。

    なぜ索引と取得を分けて呼ぶか:
      index_range は営業日を1日ずつ EDINET の書類一覧APIに当てる(3年で約730回)。
      download_pending は索引済みで未取得の書類だけを落とす。どちらも冪等で、
      途中で落ちても次回は残りだけを再開する。分けて呼ぶのは、索引が終わった
      時点で「何件取れるはずか」が確定し、取得側の進捗が分母付きで読めるため。

    スリープ抑止:
      3年分の取得は数時間かかる。ノートPCが寝ると途中で止まるので、実行中だけ
      SetThreadExecutionState でシステムスリープを抑止する(画面は消えてよい)。
      抑止はプロセス終了で自動的に解除される。

    完了時に $logDir\edinet_full_<日付>.log と同 .summary.txt を残す。
    summary は「カバレッジ率 / 失敗一覧 / unknownタグ頻度」の3点。

        powershell -ExecutionPolicy Bypass -File screener\run_edinet_full.ps1
        powershell -ExecutionPolicy Bypass -File screener\run_edinet_full.ps1 -Years 1 -IndexOnly
#>
[CmdletBinding()]
param(
    [double]$Years = 3.0,
    [string]$PythonExe = "",
    [switch]$IndexOnly,
    [switch]$DownloadOnly,
    [switch]$NoSleepGuard
)

$ErrorActionPreference = "Stop"
$repo = Split-Path -Parent $PSScriptRoot
Set-Location $repo

# Python 側は stdout を UTF-8 に固定している。PowerShell 5.1 の既定は
# コンソールのコードページ(日本語環境では cp932)なので、明示的に合わせないと
# ログとサマリの日本語が全部化ける —— 読めないレポートは無いのと同じ。
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

if (-not $PythonExe) {
    $PythonExe = (Get-Command python -ErrorAction SilentlyContinue).Source
}
if (-not $PythonExe) { throw "python not found on PATH; pass -PythonExe" }

# ログの置き場はハードコードしない。run_daily.ps1 と同じく common.py に訊く
# (DATA_ROOT は .env にあり、データツリーは D: にある)。
$logDir = & $PythonExe -c "import screener.common as C; print(C.LOG_DIR)"
if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($logDir)) {
    throw "could not resolve LOG_DIR from screener.common"
}
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Force $logDir | Out-Null }

$stamp = Get-Date -Format "yyyyMMdd_HHmmss"
$log = Join-Path $logDir "edinet_full_$stamp.log"
$summary = Join-Path $logDir "edinet_full_$stamp.summary.txt"

if (-not $NoSleepGuard) {
    # ES_CONTINUOUS(0x80000000) | ES_SYSTEM_REQUIRED(0x00000001)
    # 画面は消えてよいので ES_DISPLAY_REQUIRED は立てない。
    $sig = @'
[DllImport("kernel32.dll", SetLastError = true)]
public static extern uint SetThreadExecutionState(uint esFlags);
'@
    try {
        $p = Add-Type -MemberDefinition $sig -Name Power -Namespace Win32 -PassThru
        [void]$p::SetThreadExecutionState([uint32]"0x80000001")
        "sleep suppressed for the duration of this run" |
            Out-File -FilePath $log -Append -Encoding utf8
    } catch {
        "WARNING: could not suppress sleep ($_). PC が寝ると取得が中断します。" |
            Out-File -FilePath $log -Append -Encoding utf8
    }
}

function Write-Log([string]$text) {
    Write-Host $text
    $text | Out-File -FilePath $log -Append -Encoding utf8
}

function Invoke-Step([string]$label, [string[]]$stepArgs) {
    Write-Log "=== $label  $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ==="
    & $PythonExe -m screener.fetch.edinet_bulk @stepArgs 2>&1 |
        ForEach-Object { Write-Log ([string]$_) }
    if ($LASTEXITCODE -ne 0) {
        Write-Log "FAILED: $label (exit $LASTEXITCODE)"
        throw "$label failed with exit code $LASTEXITCODE"
    }
}

$exit = 0
try {
    if (-not $DownloadOnly) {
        Invoke-Step "index (universe x $Years years)" @("--full", "--index", "--years", "$Years")
    }
    if (-not $IndexOnly) {
        Invoke-Step "download" @("--download")
    }
} catch {
    Write-Log "RUN ABORTED: $_"
    $exit = 1
}

# サマリは中断時も書く。途中まで何が取れたかが分からないほうが困る。
"=== summary  $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ===" |
    Out-File -FilePath $summary -Encoding utf8
& $PythonExe -m screener.report.edinet_coverage 2>&1 |
    ForEach-Object { [string]$_ | Out-File -FilePath $summary -Append -Encoding utf8 }
Get-Content $summary -Encoding UTF8 | ForEach-Object { Write-Log $_ }

"log:     $log"
"summary: $summary"
exit $exit
