<#
    screener/run_rebuild.ps1 - 全件再解析 → 四半期単独値ビルド を直列で流す。

    なぜ1本にまとめるか:
      パーサ(--reset --all)とビルダーは順序が意味を持つ。financials_cum を
      作り直してから financials_q を作らないと、古い会計年度のまま単独値が
      積み上がる（2026-08-31 に会計年度を提出日から推定していたバグを直した
      ので、既存の financials_cum は全部作り直しが要る）。
      2つを別々に起動すると、間にセッションが落ちたとき「累計は新・単独値は旧」
      という最悪の中間状態が残る。1プロセスなら中断はどちらか片方で済む。

    --reset は financials_cum / guidance / unknown_tags を消してから作り直す。
    消えるのは全部 raw/ の zip から再生成できるものだけで、取得物には触らない。

        powershell -ExecutionPolicy Bypass -File screener\run_rebuild.ps1
        powershell -ExecutionPolicy Bypass -File screener\run_rebuild.ps1 -NoReset
#>
[CmdletBinding()]
param(
    [string]$PythonExe = "",
    [switch]$NoReset,
    [switch]$Resume,
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
$stamp = Get-Date -Format "yyyyMMdd_HHmmss"
$log = Join-Path $logDir "rebuild_$stamp.log"

function Write-Log([string]$text) {
    Write-Output $text
    $text | Out-File -FilePath $log -Append -Encoding utf8
}

if (-not $NoSleepGuard) {
    $sig = @'
[DllImport("kernel32.dll", SetLastError = true)]
public static extern uint SetThreadExecutionState(uint esFlags);
'@
    try {
        $p = Add-Type -MemberDefinition $sig -Name Power -Namespace Win32Rebuild -PassThru
        [void]$p::SetThreadExecutionState([uint32]"0x80000001")
        Write-Log "sleep suppressed for the duration of this run"
    } catch {
        Write-Log "WARNING: could not suppress sleep ($_)"
    }
}

function Invoke-Step([string]$label, [string]$module, [string[]]$stepArgs) {
    Write-Log "=== $label  $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ==="
    # PS 5.1 では Stop のまま native の stderr を 2>&1 で拾うと最初の1行で
    # NativeCommandError が飛び、肝心のトレースバックが消える。
    $prev = $ErrorActionPreference
    $ErrorActionPreference = "Continue"
    try {
        & $PythonExe -m $module @stepArgs 2>&1 |
            ForEach-Object { Write-Log ([string]$_) }
    } finally {
        $ErrorActionPreference = $prev
    }
    if ($LASTEXITCODE -ne 0) {
        Write-Log "FAILED: $label (exit $LASTEXITCODE)"
        throw "$label failed with exit code $LASTEXITCODE"
    }
}

Write-Log "=== run_rebuild start  pid=$PID ==="
$exit = 0
try {
    $parseArgs = @("--all", "--source", "all")
    # -Resume は「途中から」なので -NoReset を含意する。両方渡すと消してから
    # 再開することになり、意味が反転する。
    # 手動の長時間ジョブは日次タスク(19:00/23:15)の後ろに並ぶ。0 で諦めると
    # 「夜に流したら翌朝まで何もしていなかった」になる。
    $parseArgs += @("--lock-wait", "3600")
    if ($Resume) { $parseArgs += "--resume" }
    elseif (-not $NoReset) { $parseArgs = @("--reset") + $parseArgs }
    # ラベルは実際に渡した引数から作る。-Resume のとき "reset=True" と
    # 書いてしまい、ログだけ見ると全消ししたように読めた（2026-09-01）。
    $mode = if ($Resume) { "resume" } elseif ($NoReset) { "no-reset" } else { "reset" }
    Invoke-Step "parse ($mode)" "screener.extract.xbrl_parser" $parseArgs
    Invoke-Step "quarterly build" "screener.extract.quarterly_builder" @("--all", "--lock-wait", "3600")
    Invoke-Step "quarterly report" "screener.extract.quarterly_builder" @("--report")
} catch {
    Write-Log "RUN ABORTED: $_"
    $exit = 1
}
Write-Log "=== run_rebuild end (exit $exit)  $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ==="
Write-Output "log: $log"
exit $exit
