<#
    screener/start_detached.ps1 - 数時間級のジョブを「親から切り離して」起動する。

    なぜ Start-Process では足りないのか:
      2026-08-31 夜の事故の実因は PC のスリープでも再起動でもなく、
      対話セッションの終了だった。00:00 前後に株価ジョブと EDINET 索引ジョブが
      揃って消え、索引側は fetch_runs に running を残したまま死んだ。
      Windows では対話シェルの子孫プロセスはジョブオブジェクトごと畳まれうるので、
      「子として起動する」限りセッションの寿命に縛られる。

      タスクスケジューラ経由なら、実プロセスの親は Task Scheduler サービスに
      なり、こちらのセッションが何をしようと生き残る。トリガー無しでタスクを
      登録して即 Start-ScheduledTask するのが、常駐タスクを増やさずに
      「一回だけ切り離して走らせる」いちばん短い道。

    管理者権限は不要 (現在のユーザーで登録し、-RunLevel Limited で走る)。
    ログオフすると止まる点だけは Start-Process と同じなので、ログオンは保つこと。

        powershell -ExecutionPolicy Bypass -File screener\start_detached.ps1 `
            -Script screener\run_prices.ps1 -Arguments "-From 2021-09-01"
        powershell -ExecutionPolicy Bypass -File screener\start_detached.ps1 -Status
        powershell -ExecutionPolicy Bypass -File screener\start_detached.ps1 -Remove -TaskName ScreenerJob_run_prices
#>
[CmdletBinding()]
param(
    [string]$Script = "",
    [string]$Arguments = "",
    [string]$TaskName = "",
    [int]$TimeLimitHours = 8,
    [switch]$Status,
    [switch]$Remove,
    [switch]$Force
)

$ErrorActionPreference = "Stop"
$repo = Split-Path -Parent $PSScriptRoot
Set-Location $repo
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$prefix = "ScreenerJob_"

function Get-JobProcesses {
    # タスク側のプロセスは powershell.exe として立つ。CommandLine で自分の
    # ランナーだけを拾う (他人の powershell を巻き込まないため)。
    # -match は正規表現。パス区切りのバックスラッシュがエスケープと
    # 解釈されて一致しないので、ワイルドカードの -like を使う。
    Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" -ErrorAction SilentlyContinue |
        Where-Object { $_.CommandLine -and $_.CommandLine -like "*screener*run_*" }
}

if ($Status) {
    Write-Output "=== registered one-shot tasks ==="
    $tasks = Get-ScheduledTask -ErrorAction SilentlyContinue |
             Where-Object { $_.TaskName -like "$prefix*" }
    if (-not $tasks) { Write-Output "  (none)" }
    foreach ($t in $tasks) {
        $info = $t | Get-ScheduledTaskInfo
        Write-Output ("  {0}  state={1}  last={2}  result={3}" -f `
            $t.TaskName, $t.State, $info.LastRunTime, $info.LastTaskResult)
    }
    Write-Output "=== live runner processes ==="
    $procs = Get-JobProcesses
    if (-not $procs) { Write-Output "  (none)" }
    foreach ($p in $procs) {
        Write-Output ("  pid={0}  started={1}" -f $p.ProcessId, $p.CreationDate)
        Write-Output ("    {0}" -f $p.CommandLine)
    }
    exit 0
}

if ($Remove) {
    if (-not $TaskName) { throw "-Remove には -TaskName が要る" }
    Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue |
        Unregister-ScheduledTask -Confirm:$false
    Write-Output "removed: $TaskName"
    exit 0
}

if (-not $Script) { throw "-Script が要る (例: screener\run_prices.ps1)" }
$runner = if ([System.IO.Path]::IsPathRooted($Script)) { $Script }
          else { Join-Path $repo $Script }
if (-not (Test-Path $runner)) { throw "not found: $runner" }

if (-not $TaskName) {
    $TaskName = $prefix + [System.IO.Path]::GetFileNameWithoutExtension($runner)
}

$pythonExe = (Get-Command python -ErrorAction SilentlyContinue).Source
if (-not $pythonExe) { throw "python not found on PATH" }

$argline = '-NoProfile -ExecutionPolicy Bypass -File "{0}"' -f $runner
if ($Arguments) { $argline += " $Arguments" }
$argline += ' -PythonExe "{0}"' -f $pythonExe

$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument $argline `
                                  -WorkingDirectory $repo

# ExecutionTimeLimit: 既定の3日ではなく明示。走りっぱなしの検出を早める。
# DontStopIfGoingOnBatteries / AllowStartIfOnBatteries: ノートPCで電源が
# 抜けただけで数時間の取得が消えるのを防ぐ。スリープ抑止はランナー側の責務。
$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -MultipleInstances IgnoreNew `
    -ExecutionTimeLimit (New-TimeSpan -Hours $TimeLimitHours) `
    -StartWhenAvailable

# 同名タスクが走っている最中に登録し直すと、Unregister がその実行中の
# ジョブごと消す。既定のタスク名はスクリプト名から作るので、同じランナーを
# 別窓で二重に起動しようとした瞬間に、先に走っている数時間のジョブが死ぬ。
# 名前が衝突したら止める —— 並走させたいなら -TaskName を明示させる。
$existing = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($existing -and $existing.State -eq "Running" -and -not $Force) {
    throw ("タスク $TaskName は実行中。登録し直すとそのジョブを殺す。" +
           "別ジョブなら -TaskName で名前を分ける。承知の上なら -Force。")
}
$existing | Unregister-ScheduledTask -Confirm:$false

# トリガー無しで登録し、直後に手で起動する = 一回きりの切り離し実行。
Register-ScheduledTask -TaskName $TaskName -Action $action -Settings $settings `
    -RunLevel Limited `
    -Description "screener one-shot detached job ($Script $Arguments)" | Out-Null

Start-ScheduledTask -TaskName $TaskName

# 起動確認。Ready のまま/結果コードが即入るのは、登録できても走れていない印。
Start-Sleep -Seconds 3
$t = Get-ScheduledTask -TaskName $TaskName
$info = $t | Get-ScheduledTaskInfo
Write-Output ("task   : {0} [{1}]" -f $t.TaskName, $t.State)
Write-Output ("action : powershell.exe {0}" -f $argline)
Write-Output ("lastrun: {0}  result={1}" -f $info.LastRunTime, $info.LastTaskResult)

$procs = @(Get-JobProcesses)
foreach ($p in $procs) {
    Write-Output ("pid    : {0}  (parent={1})" -f $p.ProcessId, $p.ParentProcessId)
}
if (-not $procs) {
    Write-Output "WARNING: ランナーのプロセスが見つからない。State と result を確認すること。"
}

# 次にセッションが落ちたとき「何が走っていたか」を実測で辿れるようにする。
$logDir = (& $pythonExe -c "import screener.common as C; print(C.LOG_DIR)").Trim()
if ($logDir -and (Test-Path $logDir)) {
    $line = "[{0}] started {1} :: powershell.exe {2} :: pids={3}" -f `
        (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $TaskName, $argline,
        (($procs | ForEach-Object { $_.ProcessId }) -join ",")
    Add-Content -Path (Join-Path $logDir "detached_jobs.log") -Value $line -Encoding utf8
}
exit 0
