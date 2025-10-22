# BackupHyperVDisks

A PowerShell script that creates export-only backups of stopped Hyper-V virtual machines using timestamp-based change detection.

## What it does

The script automatically backs up Hyper-V VMs by:
- Exporting stopped VMs (configuration and disks) to a backup location
- Only backing up VMs when their disk files have changed since the last backup
- Keeping a configurable number of recent backups per VM
- Cleaning up old backups automatically
- Skipping VMs that are not turned off

## How it works

The script checks the `LastWriteTime` of each VM's VHD/VHDX files against the timestamp of the most recent backup folder. If any disk file is newer than the last backup, it creates a new backup using Hyper-V's `Export-VM` command.

Backups are stored in folders named with timestamps (yyyy-MM-dd_HH-mm-ss format) under `BackupRoot\<VMName>\`.

## Requirements

- Windows with Hyper-V role installed
- PowerShell with Hyper-V module
- Administrator privileges
- Sufficient disk space at the backup destination

## Configuration

Edit the variables at the top of the script:

```powershell
$BackupRoot = "T:\Arhiv\VM"                 # Backup destination path
$MaxBackupsPerVM = 2                        # Number of backups to keep per VM
$TimestampFormat = "yyyy-MM-dd_HH-mm-ss"    # Folder naming format
```

## Usage

1. Stop the VMs you want to back up
2. Run PowerShell as Administrator
3. Execute the script:
   ```powershell
   .\BackupHyperVDisks.ps1
   ```

## Output structure

```
BackupRoot\
├── VM1\
│   ├── 2024-10-22_14-30-15\
│   │   └── Export\
│   │       ├── Virtual Machines\
│   │       ├── Virtual Hard Disks\
│   │       └── Snapshots\
│   └── 2024-10-23_09-15-42\
│       └── Export\
└── VM2\
    └── 2024-10-22_14-35-20\
        └── Export\
```

## Notes

- Only stopped VMs are backed up
- Running or paused VMs are listed but skipped
- The script uses atomic operations (temporary folders) to prevent incomplete backups
- Failed backups are automatically cleaned up
- Change detection is based on disk file timestamps, not VM configuration changes
