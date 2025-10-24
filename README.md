# BackupHyperVDisks

A PowerShell script that creates backups of stopped Hyper-V virtual machines using timestamp-based change detection with three different backup methods.

## What it does

The script automatically backs up Hyper-V VMs by:
- Backing up stopped VMs using one of three configurable methods
- Only backing up VMs when their disk files have changed since the last backup
- Keeping a configurable number of recent backups per VM
- Cleaning up old backups automatically
- Skipping VMs that are not turned off

## How it works

The script checks the `LastWriteTime` of each VM's VHD/VHDX files against the timestamp of the most recent backup folder. If any disk file is newer than the last backup, it creates a new backup using the configured method.

### Backup Methods

1. **Export-VM** (`$BackupMethod = "Export"`)
   - Creates complete VM backups including configuration and disks
   - Uses Hyper-V's built-in Export-VM command
   - Stored in `BackupRoot\<VMName>\<timestamp>\Export\`
   - Largest backup size but includes everything needed to restore the VM

2. **!!! NOT TESTED !!!** **Robocopy** (`$BackupMethod = "Robocopy"`)
   - Copies only VHD/VHDX disk files
   - Uses robocopy for fast, reliable file copying
   - Stored in `BackupRoot\<VMName>\<timestamp>\Disks\`
   - Faster than Export-VM, smaller backup size, but no VM configuration

3. **7-zip** (`$BackupMethod = "7zip"`)
   - Compresses VHD/VHDX disk files into a single archive
   - Uses fastest compression setting with multithreading
   - Stored as `BackupRoot\<VMName>\<timestamp>\<VMName>.7z`
   - Smallest backup size, moderate speed, no VM configuration

Backups are stored in folders named with timestamps (yyyy-MM-dd_HH-mm-ss format) under `BackupRoot\<VMName>\`.

## Requirements

- Windows with Hyper-V role installed
- PowerShell with Hyper-V module
- Administrator privileges
- Sufficient disk space at the backup destination
- 7-zip command line utility (only if using 7-zip backup method)

## Configuration

Edit the variables at the top of the script:

```powershell
$BackupRoot = "T:\Arhiv\VM"                     # Backup destination path
$MaxBackupsPerVM = 2                            # Number of backups to keep per VM
$TimestampFormat = "yyyy-MM-dd_HH-mm-ss"        # Folder naming format
$BackupMethod = "Export"                        # "Export", "Robocopy", or "7zip"
$RobocopyFlags = "/E /Z /B /MT:8 /R:3 /W:10"    # Robocopy parameters
$SevenZipPath = "7z"                            # Path to 7z.exe
$SevenZipFlags = "a -t7z -mx=1 -mmt=4 -ms=off"  # 7-zip parameters
```

## Usage

1. Stop the VMs you want to back up
2. Run PowerShell as Administrator
3. Execute the script:
   ```powershell
   .\BackupHyperVDisks.ps1
   ```

You can of course create a desktop shortcut to run the script. In this case, the PowerShell switch  `-NoExit` is recommended, so you can inspect the script output when it's done.

For the Export-VM and Robocopy methods, you might want to enable NTFS compression on the backup folder to save space. The 7-zip method already provides compression.


## Notes

- Only stopped VMs are backed up
- Running or paused VMs are listed but skipped
- The script uses atomic operations (temporary folders) to prevent incomplete backups
- Failed backups are automatically cleaned up
- Change detection is based on disk file timestamps, not VM configuration changes
- Export-VM method creates complete backups that can be imported directly
- Robocopy and 7-zip methods only backup disk files - VM configuration must be recreated manually
- 7-zip method requires 7z.exe to be installed and accessible in PATH or specify full path in `$SevenZipPath`
