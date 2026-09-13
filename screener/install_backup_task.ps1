<#
    screener/install_backup_task.ps1 — バックアップをタスクスケジューラに登録する。

    実行時刻は既定 03:30。理由: 既存バッチ(run_daily.ps1 の 19:00 / 23:15)と
    夜間の取得ジョブに重ならない時間帯を選ぶ。バックアップ中に DB へ書き込みが
    走っても SQLite の backup API は一貫したコピーを作るが、robocopy とバッチが
    同じディスクを叩けば両方遅くなる。

    管理者権限は不要（現在のユーザーで、ログオン中のみ実行）。

        powershell -ExecutionPolicy Bypass -File screener\install_backup_task.ps1 -Destination E:\screener_backup
        powershell -ExecutionPolicy Bypass -File screener\install_backup_task.ps1 -Remove
        powershell -ExecutionPolicy Bypass -File screener\install_backup_task.ps1 -Destination E:\screener_backup -RunNow

    -DbAndLogsOnly / -AllowSameDisk はそのまま backup_data.ps1 へ渡す。
    ディスクが1台しかない機体で OneDrive 配下を退避先にする構成では両方要る
    (同一ディスク判定を承知で通す + raw/ の 8GB 超を同期に流さない)。
#>
[CmdletBinding()]
param(
    [string]$TaskName = "ScreenerDataBackup",
    [string]$Destination,
    [string]$Time = "03:30",
    [switch]$DbAndLogsOnly,
    [switch]$AllowSameDisk,
    [switch]$Remove,
    [switch]$RunNow
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$script = Join-Path $PSScriptRoot "backup_data.ps1"
if (-not (Test-Path $script)) { throw "not found: $script" }

if ($Remove) {
    Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue |
        Unregister-ScheduledTask -Confirm:$false
    Write-Host "removed: $TaskName"
    exit 0
}

if (-not $Destination) {
    throw "-Destination を指定すること（バックアップ先のパス）"
}

# 登録前に一度ドライランして、同一物理ディスク・容量不足で止まる構成を弾く。
# 「登録はできたが毎晩失敗する」状態を作らない。
Write-Host "事前確認（ドライラン）..."
# ドライランと本番で同じスイッチを使う。片方だけに付けると「ドライランは
# 通ったのに毎晩落ちる」構成をそのまま登録してしまう。
# 登録時点の python を解決して action に焼き込む。タスクは対話シェルとは
# 別の PATH で走るので、「今この shell で動くから大丈夫」は通用しない。
$pythonExe = (Get-Command python -ErrorAction SilentlyContinue).Source
if (-not $pythonExe) { throw "python not found on PATH" }

$passThru = @("-PythonExe", ('"{0}"' -f $pythonExe))
if ($DbAndLogsOnly) { $passThru += "-DbAndLogsOnly" }
if ($AllowSameDisk) { $passThru += "-AllowSameDisk" }

& powershell -NoProfile -ExecutionPolicy Bypass -File $script `
    -Destination $Destination -PythonExe $pythonExe `
    $(if ($DbAndLogsOnly) { "-DbAndLogsOnly" }) `
    $(if ($AllowSameDisk) { "-AllowSameDisk" }) -WhatIf
if ($LASTEXITCODE -ne 0) {
    throw "ドライランが失敗した。構成を直してから登録すること"
}

$action = New-ScheduledTaskAction -Execute "powershell.exe" `
    -Argument ("-NoProfile -ExecutionPolicy Bypass -File `"$script`" " +
               "-Destination `"$Destination`" " + ($passThru -join " "))
$trigger = New-ScheduledTaskTrigger -Daily -At $Time
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable `
    -DontStopIfGoingOnBatteries -AllowStartIfOnBatteries `
    -ExecutionTimeLimit (New-TimeSpan -Hours 4)
$principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -RunLevel Limited

Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger `
    -Settings $settings -Principal $principal -Force | Out-Null
Write-Host ("登録した: $TaskName（毎日 $Time / 先 $Destination" +
            $(if ($passThru) { " " + ($passThru -join " ") } else { "" }) + "）")

if ($RunNow) {
    Start-ScheduledTask -TaskName $TaskName
    Write-Host "今すぐ実行を開始した"
}
Get-ScheduledTask -TaskName $TaskName |
    Select-Object TaskName, State, @{n='NextRun';e={(Get-ScheduledTaskInfo $_).NextRunTime}} |
    Format-List
