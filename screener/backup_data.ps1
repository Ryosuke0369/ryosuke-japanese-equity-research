<#
    screener/backup_data.ps1 — DATA_ROOT (C:\screener_data) の差分ミラー + 世代バックアップ。

    なぜ必須か
    ----------
    J-Quants Light のカバレッジは **5年ローリング**。窓から落ちた日付は二度と
    取得できない。EDINET の生 zip も、TDnet の約1ヶ月保持を過ぎた短信も同じ。
    つまり DATA_ROOT は「再生成できない一次資料」であり、**ローカルDBが唯一の正本**。
    バックアップはこのシステムの必須構成要素であって、あれば良いものではない。

    2つの層
    -------
      current/    robocopy /MIR の差分ミラー。速いが、元を消すとミラーも消える
      snapshots/  日付フォルダへの週次フルコピー。/MIR の削除伝播に対する保険

    SQLite の扱い
    -------------
    稼働中の DB を robocopy でそのままコピーすると、書き込み途中の不整合な
    ファイルが取れる（WAL 有効なら .db-wal との整合も崩れる）。**DB だけは
    SQLite の backup API で一貫したコピーを作り**、robocopy からは除外する。

    使い方
        powershell -ExecutionPolicy Bypass -File screener\backup_data.ps1 -Destination E:\screener_backup
        powershell -ExecutionPolicy Bypass -File screener\backup_data.ps1 -Destination E:\screener_backup -WhatIf

    -DbAndLogsOnly
        DB (backup API の一貫コピー) と logs/ だけを退避する。raw/ の 8GB 超を
        送らないので、OneDrive のようなクラウド同期先へ週次で回すのに向く。
        raw/ は別途フルバックアップを取ること (これ単独では復旧できない)。
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$Destination,
    [string]$PythonExe = "",
    [switch]$DbAndLogsOnly,
    [int]$KeepSnapshots = 2,
    [int]$SnapshotEveryDays = 7,
    [switch]$AllowSameDisk,
    [switch]$WhatIf
)

$ErrorActionPreference = "Stop"
$repo = Split-Path -Parent $PSScriptRoot
Set-Location $repo
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

# PATH の python に頼らない。タスクスケジューラが渡してくる PATH は
# 対話シェルのそれとは別物で、venv の外のシステム python を拾う。それでは
# screener.common の import が requests で落ち、DATA_DIR を解決できずに
# ログを1行も残さず死ぬ (2026-09-10 の初回登録で実際に踏んだ)。
# run_daily.ps1 と同じく、呼び出し側が -PythonExe を渡せるようにする。
$python = if ($PythonExe) { $PythonExe }
          else { (Get-Command python -ErrorAction SilentlyContinue).Source }
if (-not $python) { throw "python not found on PATH; pass -PythonExe" }
if (-not (Test-Path $python)) { throw "PythonExe が存在しない: $python" }

# バックアップ元はハードコードしない。common.py が .env の DATA_ROOT から解決する
$src = & $python -c "import screener.common as C; print(C.DATA_DIR)"
if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($src)) {
    throw "DATA_DIR を screener.common から解決できなかった"
}
$src = $src.Trim()
if (-not (Test-Path $src)) { throw "バックアップ元が無い: $src" }

Write-Host "元: $src"
Write-Host "先: $Destination"

# ---- 事前確認1: 同一物理ディスクへのバックアップを拒む -------------------
# 同じディスクに置いたコピーはディスク故障で一緒に消える。それはバックアップ
# ではない。誤削除対策にしかならないので、意図的な場合だけ -AllowSameDisk。
function Get-DiskNumber([string]$path) {
    $letter = (Split-Path -Qualifier $path).TrimEnd(':')
    $p = Get-Partition -ErrorAction SilentlyContinue |
         Where-Object { $_.DriveLetter -eq $letter }
    if ($p) { return $p.DiskNumber }
    return $null
}
$srcDisk = Get-DiskNumber $src
$dstDisk = Get-DiskNumber $Destination
if ($null -ne $srcDisk -and $srcDisk -eq $dstDisk -and -not $AllowSameDisk) {
    throw ("バックアップ先が元と同じ物理ディスク(Disk $srcDisk)にある。" +
           "ディスク故障で両方失うのでバックアップにならない。" +
           "別ディスクを指定するか、誤削除対策だけと割り切るなら -AllowSameDisk。")
}

# ---- 事前確認2: 空き容量 -------------------------------------------------
# -DbAndLogsOnly のときは raw/ (再取得不能だが巨大) を対象から外すので、
# 必要容量も DB + logs だけで見積もる。
$sizeRoots = if ($DbAndLogsOnly) {
    @((Join-Path $src "logs")) +
    (Get-ChildItem $src -File -Filter "*.db" -EA SilentlyContinue |
     ForEach-Object { $_.FullName })
} else { @($src) }
$srcSize = ($sizeRoots | Where-Object { Test-Path $_ } | ForEach-Object {
                Get-ChildItem $_ -Recurse -File -ErrorAction SilentlyContinue
            } | Measure-Object -Sum Length).Sum
$need = $srcSize * 2.2      # ミラー + スナップショット1世代ぶん
if (-not (Test-Path $Destination)) {
    if ($WhatIf) { Write-Host "(WhatIf) 作成: $Destination" }
    else { New-Item -ItemType Directory -Force $Destination | Out-Null }
}
$free = (Get-PSDrive -Name (Split-Path -Qualifier $Destination).TrimEnd(':')).Free
Write-Host ("元のサイズ {0:N1} GB / 先の空き {1:N1} GB / 必要 {2:N1} GB" -f `
            ($srcSize/1GB), ($free/1GB), ($need/1GB))
if ($free -lt $need) {
    throw ("空き容量が足りない。ミラーに加えてスナップショット1世代を置くには " +
           "{0:N1} GB 必要。" -f ($need/1GB))
}

$current = Join-Path $Destination "current"
$snapRoot = Join-Path $Destination "snapshots"
$logDir = Join-Path $Destination "logs"
foreach ($d in @($current, $snapRoot)) {
    if (-not (Test-Path $d)) {
        if ($WhatIf) { Write-Host "(WhatIf) 作成: $d" } else { New-Item -ItemType Directory -Force $d | Out-Null }
    }
}
# logDir だけは -WhatIf でも実際に作る。下の robocopy には WhatIf でも
# /LOG+:$log を渡すので、置き場が無いと robocopy が exit 16 で落ち、
# 「ドライランが失敗した」として install_backup_task.ps1 の登録まで止まる。
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Force $logDir | Out-Null }
$stamp = Get-Date -Format "yyyyMMdd_HHmmss"
$log = Join-Path $logDir "backup_$stamp.log"

# ---- SQLite は backup API で一貫したコピーを作る -------------------------
# 稼働中の DB を robocopy でコピーすると不整合なファイルが取れる。
if ($WhatIf) {
    Write-Host "(WhatIf) SQLite の一貫コピーを作成"
} else {
    Write-Host "SQLite の一貫コピーを作成中..."
    $dbOut = Join-Path $current "screener.db"
    & $python -c @"
import sqlite3, sys
import screener.common as C
src = sqlite3.connect(C.DB_PATH)
dst = sqlite3.connect(r'''$dbOut''')
with dst:
    src.backup(dst)          # 書き込み中でも一貫したスナップショットになる
dst.close(); src.close()
print('ok')
"@
    if ($LASTEXITCODE -ne 0) { throw "SQLite backup に失敗した" }
}

# ---- 差分ミラー（DB本体と WAL/SHM は上で処理済みなので除外） -------------
# -DbAndLogsOnly: raw/ と cache/ を外し logs/ だけをミラーする。DB 本体は上の
# backup API で既に取れている。週次でクラウド(OneDrive)へ逃がす用途では、
# 8GB 超の raw/ まで送ると同期が終わらず、肝心の DB の退避が遅れる。
if ($DbAndLogsOnly) {
    $rcArgs = @((Join-Path $src "logs"), (Join-Path $current "logs"),
                "/MIR", "/R:2", "/W:5", "/NFL", "/NDL", "/NP", "/LOG+:$log")
} else {
    $rcArgs = @($src, $current, "/MIR", "/R:2", "/W:5", "/NFL", "/NDL", "/NP",
                "/XF", "screener.db", "screener.db-wal", "screener.db-shm",
                "/LOG+:$log")
}
if ($WhatIf) { $rcArgs += "/L" }
Write-Host $(if ($DbAndLogsOnly) { "差分ミラー中 (logs のみ)..." }
             else { "差分ミラー中..." })
& robocopy @rcArgs | Out-Null
# robocopy の終了コードは 0-7 が成功。8以上が本当の失敗。
if ($LASTEXITCODE -ge 8) { throw "robocopy が失敗した (exit $LASTEXITCODE)。ログ: $log" }
Write-Host "ミラー完了 (robocopy exit $LASTEXITCODE)"

# ---- 世代スナップショット（週次） ----------------------------------------
# /MIR は元の削除をミラーにも伝播する。誤削除・誤上書きから戻せるように
# 日付フォルダへのフルコピーを1世代以上残す。
$latest = Get-ChildItem $snapRoot -Directory -ErrorAction SilentlyContinue |
          Sort-Object Name -Descending | Select-Object -First 1
$needSnap = $true
if ($latest) {
    $age = (Get-Date) - [datetime]::ParseExact($latest.Name.Substring(0,8), "yyyyMMdd", $null)
    if ($age.TotalDays -lt $SnapshotEveryDays) { $needSnap = $false }
}
if ($needSnap) {
    $snap = Join-Path $snapRoot (Get-Date -Format "yyyyMMdd")
    if ($WhatIf) {
        Write-Host "(WhatIf) 世代スナップショット: $snap"
    } else {
        Write-Host "世代スナップショット作成中: $snap"
        & robocopy $current $snap /E /R:2 /W:5 /NFL /NDL /NP "/LOG+:$log" | Out-Null
        if ($LASTEXITCODE -ge 8) { throw "スナップショットに失敗 (exit $LASTEXITCODE)" }
    }
    # 古い世代を落とす。**最低1世代は必ず残す**（KeepSnapshots は 1 未満にしない）
    $keep = [Math]::Max(1, $KeepSnapshots)
    # -WhatIf では snapshots/ を作っていないので、列挙は失敗しうる。
    # 保持世代の計算はドライランの検証対象ではない。
    $old = Get-ChildItem $snapRoot -Directory -ErrorAction SilentlyContinue |
           Sort-Object Name -Descending | Select-Object -Skip $keep
    foreach ($o in $old) {
        if ($WhatIf) { Write-Host "(WhatIf) 削除: $($o.FullName)" }
        else { Remove-Item $o.FullName -Recurse -Force; Write-Host "古い世代を削除: $($o.Name)" }
    }
} else {
    Write-Host "世代スナップショットは $SnapshotEveryDays 日以内に作成済み ($($latest.Name))"
}

Write-Host "完了。ログ: $log"
