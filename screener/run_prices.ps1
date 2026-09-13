<#
    screener/run_prices.ps1 - 日次株価(全項目)と TOPIX のバックフィル用ランナー。
    仕様書 §5 の prices / market_index を埋める。

    なぜ python -m を直接叩かないのか:
      2026-08-31 夜、対話セッションから python を直接起動したところ、
      セッションが落ちた瞬間に子プロセスごと道連れになった (索引ジョブは
      fetch_runs に running を残したまま消滅)。数時間級のジョブには
        (1) スリープ抑止  (2) 親プロセスからの切り離し
      の両方が要る。(2) は start_detached.ps1 が担い、このファイルは
      (1) とログの一元化を担う。単体で叩いても動く。

    5年ローリング窓の注意:
      Light の契約範囲は「今日から5年前」で、日付が変わると窓の古い端が
      1営業日ぶん落ちる。2026-08-31 23:48 に取れた 2021-08-31 が
      2026-09-01 00:00 には 400 になった —— 起動が日付を跨ぐだけで初日が
      窓の外に出る。jquants_prices 側で「400 の message が明示する契約範囲
      まで開始日を繰り上げて継続する」ようにしたので停止はしないが、
      意図した窓はログのために明示して渡すこと。

        powershell -ExecutionPolicy Bypass -File screener\run_prices.ps1 -From 2021-09-01
        powershell -ExecutionPolicy Bypass -File screener\run_prices.ps1 -From 2021-09-01 -TopixOnly
#>
[CmdletBinding()]
param(
    [string]$From = "",
    [string]$To = "",
    [int]$Rpm = 60,
    [switch]$PricesOnly,
    [switch]$TopixOnly,
    [string]$PythonExe = "",
    [switch]$NoSleepGuard
)

$ErrorActionPreference = "Stop"
$repo = Split-Path -Parent $PSScriptRoot
Set-Location $repo

# Python 側は stdout を UTF-8 に固定している。PS 5.1 の既定は cp932 なので
# 明示的に合わせないとログの日本語が全部化ける (run_edinet_full.ps1 と同じ)。
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

if (-not $PythonExe) {
    $PythonExe = (Get-Command python -ErrorAction SilentlyContinue).Source
}
if (-not $PythonExe) { throw "python not found on PATH; pass -PythonExe" }

# ログの置き場はハードコードしない。DATA_ROOT は .env にあり D: にある。
$logDir = & $PythonExe -c "import screener.common as C; print(C.LOG_DIR)"
if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($logDir)) {
    throw "could not resolve LOG_DIR from screener.common"
}
$logDir = $logDir.Trim()
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Force $logDir | Out-Null }

$stamp = Get-Date -Format "yyyyMMdd_HHmmss"
$log = Join-Path $logDir "prices_$stamp.log"

function Write-Log([string]$text) {
    Write-Output $text
    $text | Out-File -FilePath $log -Append -Encoding utf8
}

if (-not $NoSleepGuard) {
    # ES_CONTINUOUS(0x80000000) | ES_SYSTEM_REQUIRED(0x00000001)
    # 画面は消えてよいので ES_DISPLAY_REQUIRED は立てない。
    $sig = @'
[DllImport("kernel32.dll", SetLastError = true)]
public static extern uint SetThreadExecutionState(uint esFlags);
'@
    try {
        $p = Add-Type -MemberDefinition $sig -Name Power -Namespace Win32Prices -PassThru
        [void]$p::SetThreadExecutionState([uint32]"0x80000001")
        Write-Log "sleep suppressed for the duration of this run"
    } catch {
        Write-Log "WARNING: could not suppress sleep ($_). PC が寝ると取得が中断します。"
    }
}

$common = @()
if ($From) { $common += @("--from", $From) }
if ($To)   { $common += @("--to",   $To) }
$common += @("--rpm", "$Rpm")

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
        & $PythonExe -m screener.fetch.jquants_prices @stepArgs 2>&1 |
            ForEach-Object { Write-Log ([string]$_) }
    } finally {
        $ErrorActionPreference = $prev
    }
    if ($LASTEXITCODE -ne 0) {
        Write-Log "FAILED: $label (exit $LASTEXITCODE)"
        throw "$label failed with exit code $LASTEXITCODE"
    }
}

Write-Log "=== run_prices start  pid=$PID  repo=$repo ==="
$exit = 0
try {
    if (-not $TopixOnly) { Invoke-Step "prices backfill" (@("--prices") + $common) }
    if (-not $PricesOnly) { Invoke-Step "topix backfill"  (@("--topix")  + $common) }
} catch {
    Write-Log "RUN ABORTED: $_"
    $exit = 1
}

# 到達点は中断時こそ要る。--report は prices/market_index の実測を出す。
& $PythonExe -m screener.fetch.jquants_prices --report 2>&1 |
    ForEach-Object { Write-Log ([string]$_) }

Write-Log "=== run_prices end (exit $exit)  $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ==="
Write-Output "log: $log"
exit $exit
