<#
    screener/install_task.ps1 - register (or remove) the Windows Task Scheduler
    entries that keep the TDnet archive alive. 仕様書 §2-1
    「cron/タスクスケジューラで平日夕方1回+リトライ」.

    Two weekday triggers, not one:
      19:00  the "夕方1回" of the spec.
      23:15  the retry. TDnet disclosures keep arriving until about 22:30
             (observed on 2026-08-28), so a 19:00-only job would systematically
             miss the late filings — and those are unrecoverable a month later.
             run_daily.ps1 is idempotent, so the second pass only adds what is
             new and re-attempts whatever failed.

    Both triggers run the same script, which backfills any weekday still missing
    from fetch_runs. That is what makes a week of downtime self-healing.

        powershell -ExecutionPolicy Bypass -File screener\install_task.ps1
        powershell -ExecutionPolicy Bypass -File screener\install_task.ps1 -Remove
        powershell -ExecutionPolicy Bypass -File screener\install_task.ps1 -RunNow

    No admin rights required: the tasks are registered for the current user and
    run only while that user is logged on (-RunLevel Limited, no stored password).
#>
[CmdletBinding()]
param(
    [string]$TaskName = "ScreenerTdnetArchiver",
    [string]$EveningTime = "19:00",
    [string]$RetryTime = "23:15",
    [int]$BackfillDays = 14,
    [switch]$Remove,
    [switch]$RunNow
)

$ErrorActionPreference = "Stop"
$repo = Split-Path -Parent $PSScriptRoot
$runner = Join-Path $PSScriptRoot "run_daily.ps1"
if (-not (Test-Path $runner)) { throw "not found: $runner" }

if ($Remove) {
    Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue |
        Unregister-ScheduledTask -Confirm:$false
    Write-Output "removed scheduled task: $TaskName"
    exit 0
}

if ($RunNow) {
    Start-ScheduledTask -TaskName $TaskName
    $py = (Get-Command python -ErrorAction SilentlyContinue).Source
    $logDir = if ($py) { (& $py -c "import screener.common as C; print(C.LOG_DIR)") } else { $null }
    if ([string]::IsNullOrWhiteSpace($logDir)) { $logDir = "<screener DATA_ROOT>\logs" }
    Write-Output "started $TaskName; check $($logDir.Trim()) for output"
    exit 0
}

$pythonExe = (Get-Command python -ErrorAction SilentlyContinue).Source
if (-not $pythonExe) { throw "python not found on PATH" }

$argline = ('-NoProfile -ExecutionPolicy Bypass -File "{0}" -BackfillDays {1} -PythonExe "{2}"' `
            -f $runner, $BackfillDays, $pythonExe)
$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument $argline `
                                  -WorkingDirectory $repo

$weekdays = @("Monday","Tuesday","Wednesday","Thursday","Friday")
$triggers = @(
    (New-ScheduledTaskTrigger -Weekly -DaysOfWeek $weekdays -At $EveningTime),
    (New-ScheduledTaskTrigger -Weekly -DaysOfWeek $weekdays -At $RetryTime)
)

# StartWhenAvailable: a laptop that was asleep at 19:00 runs the job on wake
# rather than skipping the day. That is the single most important setting here.
$settings = New-ScheduledTaskSettingsSet `
    -StartWhenAvailable `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -MultipleInstances IgnoreNew `
    -ExecutionTimeLimit (New-TimeSpan -Hours 3) `
    -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 20)

Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue |
    Unregister-ScheduledTask -Confirm:$false

Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $triggers `
    -Settings $settings -RunLevel Limited `
    -Description ("TDnet daily archiver for the 傾き検出スクリーナー " +
                  "(docs/傾き検出スクリーナー仕様書_v1_20260829.md §2-1). " +
                  "Backfills any weekday missing from fetch_runs; TDnet keeps " +
                  "only ~1 month so a missed day is unrecoverable.") | Out-Null

$t = Get-ScheduledTask -TaskName $TaskName
Write-Output ("registered: {0} [{1}]" -f $t.TaskName, $t.State)
foreach ($tr in $t.Triggers) { Write-Output ("  trigger: {0}" -f $tr.StartBoundary) }
Write-Output ("  action : powershell.exe {0}" -f $argline)
