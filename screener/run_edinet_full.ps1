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
        powershell -ExecutionPolicy Bypass -File screener\run_edinet_full.ps1 -From 2024-01-19 -To 2024-12-31 -IndexOnly
        powershell -ExecutionPolicy Bypass -File screener\run_edinet_full.ps1 -DownloadOnly -Subtypes 140,150
#>
[CmdletBinding()]
param(
    [double]$Years = 3.0,
    [string]$From = "",
    [string]$To = "",
    [string]$Subtypes = "",
    [double]$MinInterval = 0,
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
    # PS 5.1 では $ErrorActionPreference='Stop' のまま native exe の stderr を
    # 2>&1 で拾うと、最初の1行で NativeCommandError が飛んでパイプが止まる。
    # 結果、ログには「RUN ABORTED: Traceback (most recent call last):」だけが
    # 残り、肝心の例外本文が消える —— 2026-09-01 のスイープ中断で実際に
    # 原因が分からなくなった。この区間だけ Continue に落として全行を拾う。
    $prev = $ErrorActionPreference
    $ErrorActionPreference = "Continue"
    try {
        & $PythonExe -m screener.fetch.edinet_bulk @stepArgs 2>&1 |
            ForEach-Object { Write-Log ([string]$_) }
    } finally {
        $ErrorActionPreference = $prev
    }
    if ($LASTEXITCODE -ne 0) {
        Write-Log "FAILED: $label (exit $LASTEXITCODE)"
        throw "$label failed with exit code $LASTEXITCODE"
    }
}

$exit = 0
try {
    if (-not $DownloadOnly) {
        # --index 単独で呼ぶ。--full は index と download の両方を起動するので、
        # -IndexOnly を付けても取得まで走ってしまい、冒頭コメントが言う
        # 「索引が終わった時点で分母が確定する」が成り立っていなかった。
        $indexArgs = @("--index")
        if ($From -or $To) {
            if ($From) { $indexArgs += @("--from", $From) }
            if ($To)   { $indexArgs += @("--to",   $To) }
            $label = "index ($From .. $To)"
        } else {
            $indexArgs += @("--years", "$Years")
            $label = "index (universe x $Years years)"
        }
        # 他のジョブと並走させるときは間隔を広げて相手を殺さない。
        if ($MinInterval -gt 0) { $indexArgs += @("--min-interval", "$MinInterval") }
        Invoke-Step $label $indexArgs
    }
    if (-not $IndexOnly) {
        # 種別で絞れると「四半期報告書だけ先に埋める」ができる。数千件・数GBの
        # 取得なので、何から埋めるかを呼び出し側が決められることに意味がある。
        $downloadArgs = @("--download")
        $dlLabel = "download"
        if ($Subtypes) {
            $downloadArgs += @("--subtypes", $Subtypes)
            $dlLabel = "download (subtypes $Subtypes)"
        }
        if ($MinInterval -gt 0) { $downloadArgs += @("--min-interval", "$MinInterval") }
        Invoke-Step $dlLabel $downloadArgs
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
