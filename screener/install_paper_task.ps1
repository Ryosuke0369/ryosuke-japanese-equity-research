<#
    screener/install_paper_task.ps1 - ペーパートレード週次バッチのタスク登録。

    2本登録する:
      ScreenerPaperWeekly       毎週月曜 06:00  週次バッチ本体
      ScreenerPaperWeeklyCheck  毎週火曜 06:00  前日の成功を確認し、無ければ ERROR

    なぜ月曜06:00か: 日次ジョブ(19:00 / 23:15)と時間帯を離す。writer.lock でも
    直列化されるが、待ち行列に入るより「そもそもぶつからない」ほうが良い。

    なぜチェックを別タスクにするか: 本体が起動すらしなかった場合、本体自身は
    何も書けない。**起動しなかったことを検知できるのは外側だけ**。

    管理者権限は不要（現在のユーザーで登録、-RunLevel Limited）。

        powershell -ExecutionPolicy Bypass -File screener/install_paper_task.ps1
        powershell -ExecutionPolicy Bypass -File screener/install_paper_task.ps1 -Remove
        powershell -ExecutionPolicy Bypass -File screener/install_paper_task.ps1 -RunNow
#>
[CmdletBinding()]
param(
    [string]$WeeklyTime = "06:00",
    [string]$CheckTime = "06:00",
    [switch]$Remove,
    [switch]$RunNow
)

$ErrorActionPreference = "Stop"
$repo = Split-Path -Parent $PSScriptRoot
$TASK_RUN = "ScreenerPaperWeekly"
$TASK_CHK = "ScreenerPaperWeeklyCheck"

if ($Remove) {
    foreach ($n in @($TASK_RUN, $TASK_CHK)) {
        Get-ScheduledTask -TaskName $n -ErrorAction SilentlyContinue |
            Unregister-ScheduledTask -Confirm:$false
        Write-Output "removed: $n"
    }
    exit 0
}
if ($RunNow) {
    Start-ScheduledTask -TaskName $TASK_RUN
    Write-Output "started: $TASK_RUN"
    exit 0
}

$pythonExe = (Get-Command python -ErrorAction SilentlyContinue).Source
if (-not $pythonExe) { throw "python not found on PATH" }

$settings = New-ScheduledTaskSettingsSet `
    -StartWhenAvailable `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -MultipleInstances IgnoreNew `
    -ExecutionTimeLimit (New-TimeSpan -Hours 2)

function Register-One([string]$name, [string]$script, [string]$day, [string]$at, [string]$desc) {
    $runner = Join-Path $PSScriptRoot $script
    if (-not (Test-Path $runner)) { throw "not found: $runner" }
    $argline = ('-NoProfile -ExecutionPolicy Bypass -File "{0}" -PythonExe "{1}"' `
                -f $runner, $pythonExe)
    $action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument $argline `
                                      -WorkingDirectory $repo
    $trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek $day -At $at
    Get-ScheduledTask -TaskName $name -ErrorAction SilentlyContinue |
        Unregister-ScheduledTask -Confirm:$false
    Register-ScheduledTask -TaskName $name -Action $action -Trigger $trigger `
        -Settings $settings -RunLevel Limited -Description $desc | Out-Null
    $t = Get-ScheduledTask -TaskName $name
    Write-Output ("registered: {0} [{1}] {2} {3}" -f $t.TaskName, $t.State, $day, $at)
}

Register-One $TASK_RUN "run_paper_weekly.ps1" "Monday" $WeeklyTime `
    "ペーパートレード週次バッチ (docs/paper_trading_design.md)"
Register-One $TASK_CHK "check_paper_weekly.ps1" "Tuesday" $CheckTime `
    "前日の週次バッチが成功したかの健全性チェック。無ければログに ERROR"
