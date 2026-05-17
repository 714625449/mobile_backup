# =================== CONFIG ======================
$targetDeviceName = "S23"
$destBaseDir      = "C:\Users\11\OneDrive\Desktop\PhoneBackup"

$sourceFolders = @(
    "DCIM\Camera",
    "DCIM\Screenshots",
    "Pictures\Weixin",
    "Pictures\Screenshots",
    "Recordings\Voice Recorder",
    "迅雷下载"
)

$appearTimeoutSec    = 120
$stabilizeTimeoutSec = 300
$maxRetry            = 2
# =================================================

$shell    = New-Object -ComObject Shell.Application
$computer = $shell.Namespace(17)

# ====================================================
function Get-MTPFolder {
    param([object]$FolderItem)
    if ($null -eq $FolderItem) { return $null }
    $folder = $null
    try { $folder = $FolderItem.GetFolder } catch {}
    if ($null -eq $folder -and (-not [string]::IsNullOrEmpty($FolderItem.Path))) {
        try { $folder = $script:shell.Namespace($FolderItem.Path) } catch {}
    }
    return $folder
}

function Get-PhoneFolder {
    param([object]$RootFolder, [string]$FolderPath)
    $current = $RootFolder
    foreach ($part in ($FolderPath -split "\\")) {
        $next = $null
        foreach ($item in $current.Items()) {
            if ($item.Name -eq $part -and $item.IsFolder) {
                $next = Get-MTPFolder -FolderItem $item
                break
            }
        }
        if ($null -eq $next) { return $null }
        $current = $next
    }
    return $current
}

# 两阶段等待（修复只复制6张的问题）
function Wait-FileCopied {
    param(
        [string]$FilePath,
        [int]$AppearTimeoutSec    = 120,
        [int]$StabilizeTimeoutSec = 300
    )
    # Phase 1: 等文件出现
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    while (-not (Test-Path $FilePath)) {
        if ($sw.Elapsed.TotalSeconds -gt $AppearTimeoutSec) { return $false }
        Start-Sleep -Milliseconds 500
    }
    # Phase 2: 等大小稳定不变
    $sw.Restart()
    $prev = -1
    while ($sw.Elapsed.TotalSeconds -lt $StabilizeTimeoutSec) {
        try {
            $sz = (Get-Item $FilePath -ErrorAction Stop).Length
            if ($sz -gt 0 -and $sz -eq $prev) { return $true }
            $prev = $sz
        } catch {}
        Start-Sleep -Milliseconds 1000
    }
    try { return (Get-Item $FilePath -ErrorAction Stop).Length -gt 0 } catch { return $false }
}

# ====================================================
Write-Host "`n============================================" -ForegroundColor White
Write-Host "   Samsung S23 Backup" -ForegroundColor Cyan
Write-Host "   Dest : $destBaseDir" -ForegroundColor Cyan
Write-Host "============================================`n" -ForegroundColor White

# ─── 找手机设备（原版逻辑，不加 usb# 过滤）────────
$phoneDevice = $null
foreach ($item in $computer.Items()) {
    if ($item.IsFolder -and $item.Name -eq $targetDeviceName) {
        $phoneDevice = $item
        break
    }
}

if ($null -eq $phoneDevice) {
    Write-Host "未找到设备 '$targetDeviceName'，当前可见设备：" -ForegroundColor Red
    foreach ($item in $computer.Items()) {
        if ($item.IsFolder) {
            Write-Host ("  → '{0}'" -f $item.Name) -ForegroundColor Yellow
        }
    }
    Read-Host "`n按 Enter 退出"
    Exit
}
Write-Host "Connected : $($phoneDevice.Name)" -ForegroundColor Green

# ─── 获取设备根命名空间 ───────────────────────────
$phoneRoot = $null
try { $phoneRoot = $shell.Namespace($phoneDevice.Path) } catch {}
if ($null -eq $phoneRoot) {
    try { $phoneRoot = $phoneDevice.GetFolder } catch {}
}
if ($null -eq $phoneRoot) {
    Write-Host "无法打开设备根目录，请解锁手机并选择「文件传输」模式" -ForegroundColor Red
    Read-Host "按 Enter 退出"
    Exit
}

# ─── 找内部存储 ───────────────────────────────────
$internalStorage = $null
foreach ($item in $phoneRoot.Items()) {
    if ($item.Name -match "内部|Internal|Phone|共享") {
        $candidate = Get-MTPFolder -FolderItem $item
        if ($null -ne $candidate) {
            $internalStorage = $candidate
            Write-Host "Storage   : $($item.Name)" -ForegroundColor Green
            break
        }
    }
}
if ($null -eq $internalStorage) {
    foreach ($item in $phoneRoot.Items()) {
        if ($item.IsFolder) {
            $candidate = Get-MTPFolder -FolderItem $item
            if ($null -ne $candidate) {
                $internalStorage = $candidate
                Write-Warning "使用第一个可用分区: '$($item.Name)'"
                break
            }
        }
    }
}
if ($null -eq $internalStorage) {
    Write-Host "无法访问内部存储" -ForegroundColor Red
    Read-Host "按 Enter 退出"
    Exit
}

if (-not (Test-Path $destBaseDir)) {
    New-Item -ItemType Directory -Path $destBaseDir -Force | Out-Null
}

# ─── 主拷贝循环 ───────────────────────────────────
Write-Host "`n=== STARTING BACKUP ===`n" -ForegroundColor White

$totalCopied  = 0
$totalSkipped = 0
$totalFailed  = 0
$failedList   = @()

foreach ($folderPath in $sourceFolders) {
    Write-Host "[Folder] $folderPath" -ForegroundColor Cyan

    $srcFolder = Get-PhoneFolder -RootFolder $internalStorage -FolderPath $folderPath
    if ($null -eq $srcFolder) {
        Write-Warning "  手机上未找到此路径，跳过。`n"
        continue
    }

    $localDest = Join-Path $destBaseDir $folderPath
    if (-not (Test-Path $localDest)) {
        New-Item -ItemType Directory -Path $localDest -Force | Out-Null
    }
    $destShell = $shell.Namespace($localDest)

    $allItems  = @($srcFolder.Items() | Where-Object { -not $_.IsFolder })
    $newItems  = @($allItems | Where-Object { -not (Test-Path (Join-Path $localDest $_.Name)) })
    $skipCount = $allItems.Count - $newItems.Count
    $totalSkipped += $skipCount

    Write-Host ("  Total: {0}  New: {1}  Skip: {2}" -f `
        $allItems.Count, $newItems.Count, $skipCount) -ForegroundColor Gray

    if ($newItems.Count -eq 0) {
        Write-Host "  All backed up.`n" -ForegroundColor Green
        continue
    }

    $folderCopied = 0
    $idx = 0

    foreach ($fileItem in $newItems) {
        $idx++
        $targetFile = Join-Path $localDest $fileItem.Name

        $ok  = $false
        $try = 0

        while (-not $ok -and $try -lt $maxRetry) {
            $try++
            $tag = if ($try -gt 1) { " (retry $($try-1))" } else { "" }
            Write-Host ("  [Copy{0}] [{1}/{2}] {3}" -f `
                $tag, $idx, $newItems.Count, $fileItem.Name) -ForegroundColor White

            if ((Test-Path $targetFile) -and
                (Get-Item $targetFile -EA SilentlyContinue).Length -eq 0) {
                Remove-Item $targetFile -Force -EA SilentlyContinue
            }

            $timer = [System.Diagnostics.Stopwatch]::StartNew()
            try {
                $destShell.CopyHere($fileItem, 20)
                $ok = Wait-FileCopied -FilePath             $targetFile `
                                      -AppearTimeoutSec     $appearTimeoutSec `
                                      -StabilizeTimeoutSec  $stabilizeTimeoutSec
            } catch {
                $ok = $false
            }
            $timer.Stop()

            if ($ok) {
                $kb = [math]::Round((Get-Item $targetFile).Length / 1KB, 0)
                Write-Host ("    OK  {0:N1}s  {1} KB" -f $timer.Elapsed.TotalSeconds, $kb) `
                    -ForegroundColor Green
            } elseif ($try -lt $maxRetry) {
                Write-Warning "    Failed, retrying in 3s..."
                Start-Sleep -Seconds 3
                Remove-Item $targetFile -Force -EA SilentlyContinue
            }
        }

        if ($ok) {
            $folderCopied++; $totalCopied++
        } else {
            Write-Warning ("    FAILED: {0}" -f $fileItem.Name)
            $totalFailed++
            $failedList += "$folderPath\$($fileItem.Name)"
            if ((Test-Path $targetFile) -and
                (Get-Item $targetFile -EA SilentlyContinue).Length -eq 0) {
                Remove-Item $targetFile -Force -EA SilentlyContinue
            }
        }
    }

    Write-Host ("`n  Done — Copied:{0}  Skip:{1}  Failed:{2}`n" -f `
        $folderCopied, $skipCount, ($newItems.Count - $folderCopied)) -ForegroundColor Green
}

# ─── 汇总 ─────────────────────────────────────────
$fc = if ($totalFailed -gt 0) { "Red" } else { "Green" }
Write-Host "============================================" -ForegroundColor White
Write-Host "  Copied  : $totalCopied" -ForegroundColor Green
Write-Host "  Skipped : $totalSkipped" -ForegroundColor Cyan
Write-Host "  Failed  : $totalFailed" -ForegroundColor $fc
if ($failedList.Count -gt 0) {
    $failedList | ForEach-Object { Write-Host "    - $_" -ForegroundColor Red }
}
Write-Host "============================================`n" -ForegroundColor White

# ─── 可选删除 ─────────────────────────────────────
if ($totalCopied -gt 0) {
    Write-Host "输入 DELETE 删除手机上已备份的文件，直接回车跳过" -ForegroundColor Yellow
    $confirmation = Read-Host

    if ($confirmation -ceq 'DELETE') {
        foreach ($folderPath in $sourceFolders) {
            $srcFolder = Get-PhoneFolder -RootFolder $internalStorage -FolderPath $folderPath
            if ($null -eq $srcFolder) { continue }
            Write-Host "[Del] $folderPath" -ForegroundColor Yellow
            foreach ($file in @($srcFolder.Items() | Where-Object { -not $_.IsFolder })) {
                $pcCopy = Join-Path (Join-Path $destBaseDir $folderPath) $file.Name
                if (Test-Path $pcCopy) {
                    try { $file.InvokeVerb("delete"); Write-Host "  Deleted: $($file.Name)" } 
                    catch { Write-Warning "  Failed: $($file.Name)" }
                }
            }
        }
    }
}

Read-Host "`n完成，按 Enter 退出"