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
#>
[CmdletBinding()]
param(
    [string]$TaskName = "ScreenerDataBackup",
    [string]$Destination,
    [string]$Time = "03:30",
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
& powershell -NoProfile -ExecutionPolicy Bypass -File $script -Destination $Destination -WhatIf
if ($LASTEXITCODE -ne 0) {
    throw "ドライランが失敗した。構成を直してから登録すること"
}

$action = New-ScheduledTaskAction -Execute "powershell.exe" `
    -Argument ("-NoProfile -ExecutionPolicy Bypass -File `"$script`" " +
               "-Destination `"$Destination`"")
$trigger = New-ScheduledTaskTrigger -Daily -At $Time
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable `
    -DontStopIfGoingOnBatteries -AllowStartIfOnBatteries `
    -ExecutionTimeLimit (New-TimeSpan -Hours 4)
$principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -RunLevel Limited

Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger `
    -Settings $settings -Principal $principal -Force | Out-Null
Write-Host "登録した: $TaskName（毎日 $Time / 先 $Destination）"

if ($RunNow) {
    Start-ScheduledTask -TaskName $TaskName
    Write-Host "今すぐ実行を開始した"
}
Get-ScheduledTask -TaskName $TaskName |
    Select-Object TaskName, State, @{n='NextRun';e={(Get-ScheduledTaskInfo $_).NextRunTime}} |
    Format-List
