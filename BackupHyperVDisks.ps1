<#
.SYNOPSIS
  Backup for stopped Hyper-V VMs using timestamp-based change detection with Export-VM or Robocopy.
.DESCRIPTION
  Supports three backup methods:
  - Export-VM: Copies VM configuration and disks into BackupRoot\<VMName>\<timestamp>\Export
  - Robocopy: Copies only VHD/VHDX files into BackupRoot\<VMName>\<timestamp>\Disks
  - 7zip: Compresses VHD/VHDX files into BackupRoot\<VMName>\<timestamp>\<VMName>.7z
  For each VM, backup runs only when any attached VHD/VHDX has a LastWriteTime later than the most recent backup timestamp.
  Keeps the newest N backups per VM and prunes older ones.
.NOTES
  - Run as Administrator on the Hyper-V host.
  - Requires Hyper-V PowerShell module.
  - Script uses folder timestamps named with format yyyy-MM-dd_HH-mm-ss to decide changes.
  - Configure $BackupMethod: "Export" (full VM), "Robocopy" (disks only), or "7zip" (compressed disks).
  - For 7zip method, ensure 7z.exe is in PATH or set $SevenZipPath to full executable path.
#>

#region Configuration
$BackupRoot = "T:\Arhiv\VMBackups"          # where per-VM backups will be stored
$MaxBackupsPerVM = 2                        # keep the newest N backups per VM
$TimestampFormat = "yyyy-MM-dd_HH-mm-ss"    # folder name timestamp format
$TempSuffix = ".backupTemp"                 # suffix for temporary backup folders
$BackupMethod = "7zip"                      # "Export" for Export-VM, "Robocopy" for file copy, "7zip" for compressed archive
$RobocopyFlags = "/E /Z /B /MT:8 /R:3 /W:10"  # robocopy parameters (multithread, retry, etc.)
$SevenZipPath = "7z"                        # path to 7z.exe (assumes it's in PATH)
$SevenZipFlags = "a -t7z -mx=1 -mmt=4 -ms=off"     # 7z parameters (fastest compression, reasonable multithreading, non-solid)
#endregion

#region Helpers
function Get-Timestamp { (Get-Date).ToString($TimestampFormat) }

function Use-Folder {
    param([string]$Path)
    if (-not (Test-Path $Path)) { New-Item -Path $Path -ItemType Directory -Force | Out-Null }
}

function Get-TimestampFromFolderName {
    param([string]$FolderName)
    try {
        return [datetime]::ParseExact($FolderName, $TimestampFormat, $null)
    }
    catch {
        return $null
    }
}

function Get-LatestBackupTimestamp {
    param([string]$VMFolder)
    if (-not (Test-Path $VMFolder)) { return $null }
    $folders = Get-ChildItem -Path $VMFolder -Directory -ErrorAction SilentlyContinue
    if (-not $folders -or $folders.Count -eq 0) { return $null }
    $timestamps = @()
    foreach ($f in $folders) {
        $ts = Get-TimestampFromFolderName -FolderName $f.Name
        if ($ts) { $timestamps += $ts }
    }
    if ($timestamps.Count -eq 0) { return $null }
    return ($timestamps | Sort-Object)[-1]
}

function Remove-OldBackups {
    param([string]$VMName, [int]$MaxKeep, [string]$BackupRoot)
    $vmFolder = Join-Path $BackupRoot $VMName
    if (-not (Test-Path $vmFolder)) { return }
    $backups = Get-ChildItem -Path $vmFolder -Directory | Sort-Object Name -Descending
    if ($backups.Count -le $MaxKeep) { return }
    $toRemove = $backups | Select-Object -Skip $MaxKeep
    foreach ($b in $toRemove) {
        Write-Host "Pruning old backup $($b.FullName)"
        try {
            Remove-Item -Path $b.FullName -Recurse -Force -ErrorAction Stop
        }
        catch {
            Write-Warning "Failed to remove $($b.FullName): $_"
        }
    }
}

function Remove-TempBackups {
    param([string]$BackupRoot, [string]$TempSuffix)
    Write-Host "Cleaning up any leftover temporary backup folders..." -ForegroundColor Gray
    $tempFolders = Get-ChildItem -Path $BackupRoot -Directory -Recurse -ErrorAction SilentlyContinue | Where-Object { $_.Name.EndsWith($TempSuffix) }
    foreach ($tempFolder in $tempFolders) {
        Write-Host "Removing incomplete backup: $($tempFolder.FullName)"
        try {
            Remove-Item -Path $tempFolder.FullName -Recurse -Force -ErrorAction Stop
        }
        catch {
            Write-Warning "Failed to remove temporary folder $($tempFolder.FullName): $_"
        }
    }
}
#endregion

#region Main
Write-Host "Starting Hyper-V backup..." -ForegroundColor Cyan

Use-Folder $BackupRoot

# Clean up any leftover temporary backup folders from previous failed runs
Remove-TempBackups -BackupRoot $BackupRoot -TempSuffix $TempSuffix

$allVMs = Get-VM
if (-not $allVMs) {
    Write-Host "No Hyper-V VMs detected." -ForegroundColor Yellow
    return
}

$runningVMs = $allVMs | Where-Object { $_.State -ne 'Off' }
if ($runningVMs.Count -gt 0) {
    Write-Host "Warning: The following VMs are not turned off and will NOT be backed up by this script:" -ForegroundColor Yellow
    $runningVMs | ForEach-Object { Write-Host " - $($_.Name): State: $($_.State)" -ForegroundColor Yellow }
}
else {
    Write-Host "No running or paused VMs detected." -ForegroundColor Green
}

$backupVMs = $allVMs | Where-Object { $_.State -eq 'Off' }
if ($backupVMs.Count -eq 0) {
    Write-Host "No stopped VMs to backup." -ForegroundColor Yellow
    return
}

foreach ($vm in $backupVMs) {
    $vmName = $vm.Name
    Write-Host "`nProcessing VM: $vmName" -ForegroundColor Cyan

    if ($vmName -ne "TumbleDev")
    {
        continue;
    }

    $vmFolder = Join-Path $BackupRoot $vmName
    Use-Folder $vmFolder

    # Determine latest backup timestamp for this VM (from folder names)
    $latestBackupTs = Get-LatestBackupTimestamp -VMFolder $vmFolder
    if ($latestBackupTs) {
        Write-Host "Latest backup for ${vmName}: $latestBackupTs"
    }
    else {
        Write-Host "No previous backups found for $vmName"
    }

    # Gather attached VHD/VHDX paths and find the most recent LastWriteTime among them
    $hdDrives = Get-VMHardDiskDrive -VMName $vmName -ErrorAction SilentlyContinue
    if (-not $hdDrives) {
        Write-Warning "No hard disk drives found for $vmName; skipping backup."
        continue
    }

    $latestDiskTime = $null
    foreach ($hd in $hdDrives) {
        $vhdPath = $hd.Path
        if (-not $vhdPath) { continue }
        if (-not (Test-Path $vhdPath)) {
            Write-Warning "Disk file not found: $vhdPath for $vmName; skipping this disk in timestamp check."
            continue
        }
        $t = (Get-Item $vhdPath).LastWriteTimeUtc
        if (-not $latestDiskTime -or $t -gt $latestDiskTime) { $latestDiskTime = $t }
    }

    if (-not $latestDiskTime) {
        Write-Warning "Could not determine disk timestamp for $vmName; skipping backup."
        continue
    }

    # Compare and decide whether to backup
    $doBackup = $false
    if (-not $latestBackupTs) {
        $doBackup = $true
        Write-Host "No previous backup: will backup $vmName using $BackupMethod"
    }
    else {
        # convert latestBackupTs to UTC for fair comparison
        $latestBackupTsUtc = [datetime]::SpecifyKind($latestBackupTs, [System.DateTimeKind]::Unspecified).ToUniversalTime()
        if ($latestDiskTime -gt $latestBackupTsUtc) {
            $doBackup = $true
            Write-Host "Disk changed since last backup: will backup $vmName using $BackupMethod"
        }
        else {
            Write-Host "No change detected for $vmName since latest backup at $latestBackupTs; skipping backup." -ForegroundColor DarkYellow
        }
    }

    if ($doBackup) {
        $timestamp = Get-Timestamp
        $thisBackupFolder = Join-Path $vmFolder $timestamp
        $tempBackupFolder = $thisBackupFolder + $TempSuffix
        
        # Create temporary backup folder
        Use-Folder $tempBackupFolder

        $backupSuccess = $false
        
        if ($BackupMethod -eq "Export") {
            $exportPath = Join-Path $tempBackupFolder "Export"
            Use-Folder $exportPath
            Write-Host "Exporting $vmName to temporary location $exportPath" -ForegroundColor Gray
            try {
                Export-VM -Name $vmName -Path $exportPath -ErrorAction Stop
                $backupSuccess = $true
            }
            catch {
                Write-Warning "Export-VM failed for ${vmName}: $_"
            }
        }
        elseif ($BackupMethod -eq "Robocopy") {
            $disksPath = Join-Path $tempBackupFolder "Disks"
            Use-Folder $disksPath
            Write-Host "Copying disk files for $vmName using robocopy" -ForegroundColor Gray
            
            $allCopySuccess = $true
            foreach ($hd in $hdDrives) {
                $vhdPath = $hd.Path
                if (-not $vhdPath -or -not (Test-Path $vhdPath)) { continue }
                
                $fileName = Split-Path $vhdPath -Leaf
                $destPath = Join-Path $disksPath $fileName
                
                Write-Host "  Copying $fileName..." -ForegroundColor Gray
                $robocopyCmd = "robocopy `"$(Split-Path $vhdPath -Parent)`" `"$disksPath`" `"$fileName`" $RobocopyFlags"
                $result = Invoke-Expression $robocopyCmd
                
                # Robocopy exit codes 0-7 are success, 8+ are errors
                if ($LASTEXITCODE -ge 8) {
                    Write-Warning "Robocopy failed for $fileName (exit code: $LASTEXITCODE)"
                    $allCopySuccess = $false
                }
            }
            $backupSuccess = $allCopySuccess
        }
        elseif ($BackupMethod -eq "7zip") {
            $archivePath = Join-Path $tempBackupFolder "$vmName.7z"
            Write-Host "Compressing disk files for $vmName using 7-zip" -ForegroundColor Gray
            
            # Build list of disk files to compress
            $diskFiles = @()
            foreach ($hd in $hdDrives) {
                $vhdPath = $hd.Path
                if (-not $vhdPath -or -not (Test-Path $vhdPath)) { continue }
                $diskFiles += "`"$vhdPath`""
            }
            
            if ($diskFiles.Count -eq 0) {
                Write-Warning "No valid disk files found for $vmName"
                $backupSuccess = $false
            }
            else {
                $diskFilesString = $diskFiles -join " "
                $sevenZipCmd = "& `"$SevenZipPath`" $SevenZipFlags `"$archivePath`" $diskFilesString"
                
                Write-Host "  Creating archive $vmName.7z..." -ForegroundColor Gray
                try {
                    Invoke-Expression $sevenZipCmd
                    # 7-zip exit code 0 = success, 1 = warning (non-fatal), 2+ = error
                    if ($LASTEXITCODE -le 1) {
                        $backupSuccess = $true
                        if ($LASTEXITCODE -eq 1) {
                            Write-Warning "7-zip completed with warnings for $vmName"
                        }
                    }
                    else {
                        Write-Warning "7-zip failed for $vmName (exit code: $LASTEXITCODE)"
                        $backupSuccess = $false
                    }
                }
                catch {
                    Write-Warning "7-zip execution failed for ${vmName}: $_"
                    $backupSuccess = $false
                }
            }
        }
        else {
            Write-Warning "Unknown backup method: $BackupMethod. Use 'Export', 'Robocopy', or '7zip'."
            $backupSuccess = $false
        }
        
        if ($backupSuccess) {
            # If backup succeeded, atomically rename temp folder to final name
            Write-Host "Backup completed, finalizing backup for $vmName" -ForegroundColor Gray
            try {
                Rename-Item -Path $tempBackupFolder -NewName $timestamp -ErrorAction Stop
                Write-Host "Backup completed successfully for $vmName" -ForegroundColor Green
                
                # After successful backup, prune old backups
                Remove-OldBackups -VMName $vmName -MaxKeep $MaxBackupsPerVM -BackupRoot $BackupRoot
            }
            catch {
                Write-Warning "Failed to finalize backup for ${vmName}: $_"
                try { Remove-Item -Path $tempBackupFolder -Recurse -Force -ErrorAction Stop } catch { }
            }
        }
        else {
            Write-Warning "Removing incomplete backup folder $tempBackupFolder"
            try { Remove-Item -Path $tempBackupFolder -Recurse -Force -ErrorAction Stop } catch { }
            continue
        }
    }
}

Write-Host "`nBackup run completed." -ForegroundColor Green

#endregion