<#
.SYNOPSIS
  Export-only backup for stopped Hyper-V VMs using timestamp-based change detection.
.DESCRIPTION
  Uses Export-VM to copy VM configuration and disks into BackupRoot\<VMName>\<timestamp>\Export.
  For each VM, export runs only when any attached VHD/VHDX has a LastWriteTime later than the most recent backup timestamp.
  Keeps the newest N backups per VM and prunes older ones.
.NOTES
  - Run as Administrator on the Hyper-V host.
  - Requires Hyper-V PowerShell module.
  - Script uses folder timestamps named with format yyyy-MM-dd_HH-mm-ss to decide changes.
#>

#region Configuration
$BackupRoot = "T:\Arhiv\VM"                 # where per-VM backups will be stored
$MaxBackupsPerVM = 2                        # keep the newest N backups per VM
$TimestampFormat = "yyyy-MM-dd_HH-mm-ss"    # folder name timestamp format
$TempSuffix = ".backupTemp"                 # suffix for temporary backup folders
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

    #if ($vmName -ne "TumbleDev")
    #{
    #    continue;
    #}

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
        Write-Warning "No hard disk drives found for $vmName; skipping export."
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
        Write-Warning "Could not determine disk timestamp for $vmName; skipping export."
        continue
    }

    # Compare and decide whether to export
    $doExport = $false
    if (-not $latestBackupTs) {
        $doExport = $true
        Write-Host "No previous backup: will export $vmName"
    }
    else {
        # convert latestBackupTs to UTC for fair comparison
        $latestBackupTsUtc = [datetime]::SpecifyKind($latestBackupTs, [System.DateTimeKind]::Unspecified).ToUniversalTime()
        if ($latestDiskTime -gt $latestBackupTsUtc) {
            $doExport = $true
            Write-Host "Disk changed since last backup: will export $vmName"
        }
        else {
            Write-Host "No change detected for $vmName since latest backup at $latestBackupTs; skipping export." -ForegroundColor DarkYellow
        }
    }

    if ($doExport) {
        $timestamp = Get-Timestamp
        $thisBackupFolder = Join-Path $vmFolder $timestamp
        $tempBackupFolder = $thisBackupFolder + $TempSuffix
        
        # Create temporary backup folder
        Use-Folder $tempBackupFolder

        $exportPath = Join-Path $tempBackupFolder "Export"
        Use-Folder $exportPath

        Write-Host "Exporting $vmName to temporary location $exportPath" -ForegroundColor Gray
        try {
            Export-VM -Name $vmName -Path $exportPath -ErrorAction Stop
            
            # If export succeeded, atomically rename temp folder to final name
            Write-Host "Export completed, finalizing backup for $vmName" -ForegroundColor Gray
            Rename-Item -Path $tempBackupFolder -NewName $timestamp -ErrorAction Stop
            Write-Host "Backup completed successfully for $vmName" -ForegroundColor Green
            
            # After successful backup, prune old backups
            Remove-OldBackups -VMName $vmName -MaxKeep $MaxBackupsPerVM -BackupRoot $BackupRoot
        }
        catch {
            Write-Warning "Export-VM failed for ${vmName}: $_"
            Write-Warning "Removing incomplete backup folder $tempBackupFolder"
            try { Remove-Item -Path $tempBackupFolder -Recurse -Force -ErrorAction Stop } catch { }
            continue
        }
    }
}

Write-Host "`nBackup run completed." -ForegroundColor Green

#endregion