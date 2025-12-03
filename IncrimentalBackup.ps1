<#
.SYNOPSIS
    Performs versioned, incremental backups of a specified folder using Robocopy and tracks changes in a CSV manifest.

.DESCRIPTION
    This script provides an automated backup solution for project folders, creating timestamped, versioned copies of new or modified files. It is designed to be run as a scheduled task.

    - Initial Backup: If the backup destination is empty, it performs a full copy of the source folder into a new timestamped subfolder. It creates a CSV manifest file listing all items and their modification dates.
    - Incremental Backup: On subsequent runs, it compares the source files against the last modification time of the manifest file. It then uses Robocopy to copy only new or changed files into a new timestamped folder. The manifest is updated with the new file versions.
    - Recovery: If the backup destination contains data but the CSV manifest is missing, it quarantines the existing data into a "Missing_Logs" folder and performs a fresh initial backup.
    - Exclusions: The script is configured to exclude any directory named '.godot' from the backup.

.PARAMETER BackupTarget
    The full path to the source folder you want to back up. This parameter is mandatory.

.PARAMETER BackupDestination
    The full path to the root folder where backup versions will be stored. This folder will contain timestamped subfolders for each backup operation. This parameter is mandatory.

.PARAMETER Logs
    The full path to a folder where the backup manifest CSV and Robocopy logs will be stored. It is recommended to make this a subfolder of BackupDestination. This parameter is mandatory.

.EXAMPLE
    .\IncrementalBackup.ps1 -BackupTarget "C:\Users\You\Documents\MyProject" -BackupDestination "D:\Backups\MyProject" -Logs "D:\Backups\MyProject\Logs"

.NOTES
    Author:  Tydorius
    Version: 2.1
    License: CC BY-NC 4.0
    Last Updated: 2025-06-21
    - Robocopy is used for its efficiency and robust logging.
    - The script is designed to be idempotent and handle common error scenarios.
    - Ensure Robocopy is available in the system's PATH. It is included by default in Windows 11.

    I am really bad about remembering to use git for various things, as well as maintaining my own backups. I set this as a scheduled task to make up for my own failings.
#>
[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [Parameter(Mandatory=$true, HelpMessage="Enter the full path to the source folder to be backed up.")]
    [string]$BackupTarget,

    [Parameter(Mandatory=$true, HelpMessage="Enter the full path to the destination folder for backups.")]
    [string]$BackupDestination,

    [Parameter(Mandatory=$true, HelpMessage="Enter the full path to the folder for storing logs and the manifest.")]
    [string]$Logs
)

# --- Script Setup ---
$ErrorActionPreference = "Stop"
$timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"

Write-Verbose "Script execution started at $timestamp."
Write-Verbose "Source Path (BackupTarget): $BackupTarget"
Write-Verbose "Destination Path (BackupDestination): $BackupDestination"
Write-Verbose "Logs Path: $Logs"

# Ensure core paths exist
if (-not (Test-Path -Path $BackupTarget)) {
    Write-Error "The backup target folder '$BackupTarget' does not exist. Halting script."
    exit 1
}
# Create the root destination and logs folder if they don't exist
New-Item -Path $BackupDestination -ItemType Directory -Force | Out-Null
New-Item -Path $Logs -ItemType Directory -Force | Out-Null

$logCsvPath = Join-Path -Path $Logs -ChildPath "BackupManifest.csv"
Write-Verbose "Manifest File Path: $logCsvPath"

# --- Main Logic ---
try {
    # SCENARIO 1: Initial Backup (Destination is empty)
    # Check if anything exists in the backup destination besides the Logs folder itself.
    $destinationItems = Get-ChildItem -Path $BackupDestination | Where-Object { $_.FullName -ne $Logs }
    if (-not $destinationItems) {
        Write-Host "Scenario: Initial Backup. Destination is empty." -ForegroundColor Green
        $newBackupFolderName = "Backup_$timestamp"
        $newBackupFolderPath = Join-Path -Path $BackupDestination -ChildPath $newBackupFolderName
        $robocopyLogPath = Join-Path -Path $Logs -ChildPath "Robocopy_Initial_$timestamp.log"

        Write-Verbose "Creating new backup folder: $newBackupFolderPath"
        New-Item -Path $newBackupFolderPath -ItemType Directory -Force | Out-Null

        Write-Verbose "Performing full backup using Robocopy..."
        $robocopyArgs = @(
            $BackupTarget,
            $newBackupFolderPath,
            "/E",           # Copy subdirectories, including empty ones.
            "/COPY:DAT",    # Copy Data, Attributes, Timestamps.
            "/DCOPY:DA",    # Copy Directory timestamps and attributes.
            "/XD", ".godot",# Exclude directories named .godot
            "/R:3",         # Retry 3 times on failed copies.
            "/W:5",         # Wait 5 seconds between retries.
            "/NP",          # No Progress - don't show % copied.
            "/TEE",         # Output to console window as well as the log file.
            "/LOG+:$robocopyLogPath"
        )
        robocopy @robocopyArgs
        Write-Verbose "Robocopy operation complete. Log file created at: $robocopyLogPath"

        # Generate the initial CSV manifest, excluding specified folders
        Write-Verbose "Generating new backup manifest..."
        $sourceFiles = Get-ChildItem -Path $BackupTarget -Recurse -File | Where-Object { $_.FullName -notmatch '\\.godot\\' }
        $manifestData = foreach ($file in $sourceFiles) {
            [PSCustomObject]@{
                FullName       = $file.FullName
                RelativePath   = $file.FullName.Substring($BackupTarget.Length)
                LastWriteTime  = $file.LastWriteTime
                BackupDate     = $timestamp
                BackupLocation = $newBackupFolderPath
            }
        }
        $manifestData | Export-Csv -Path $logCsvPath -NoTypeInformation
        Write-Host "Initial backup and manifest creation complete." -ForegroundColor Green
    }
    # SCENARIO 2: Destination has data, but the manifest is missing (Recovery)
    elseif (-not (Test-Path -Path $logCsvPath)) {
        Write-Warning "Scenario: Recovery. Backup data found in '$BackupDestination', but the manifest file '$logCsvPath' is missing."
        Write-Warning "Quarantining existing data and performing a fresh initial backup."

        $quarantineFolder = Join-Path -Path $BackupDestination -ChildPath "Missing_Logs_$timestamp"
        Write-Verbose "Creating quarantine folder: $quarantineFolder"
        New-Item -Path $quarantineFolder -ItemType Directory -Force | Out-Null

        Write-Verbose "Moving all items from '$BackupDestination' to quarantine..."
        Get-ChildItem -Path $BackupDestination -Exclude (Split-Path $quarantineFolder -Leaf) | ForEach-Object {
            Write-Verbose "Moving '$($_.FullName)'..."
            Move-Item -Path $_.FullName -Destination $quarantineFolder
        }
        Write-Verbose "Quarantine complete."

        # Re-create the Logs directory after it was moved.
        Write-Verbose "Re-creating the logs directory at '$Logs'."
        New-Item -Path $Logs -ItemType Directory -Force | Out-Null

        # Now, perform a fresh initial backup.
        Write-Verbose "Proceeding with a fresh initial backup..."
        $newBackupFolderName = "Backup_$timestamp"
        $newBackupFolderPath = Join-Path -Path $BackupDestination -ChildPath $newBackupFolderName
        $robocopyLogPath = Join-Path -Path $Logs -ChildPath "Robocopy_Recovery_$timestamp.log"

        Write-Verbose "Creating new backup folder: $newBackupFolderPath"
        New-Item -Path $newBackupFolderPath -ItemType Directory -Force | Out-Null

        Write-Verbose "Performing full backup using Robocopy..."
        $robocopyArgs = @(
            $BackupTarget,
            $newBackupFolderPath,
            "/E", "/COPY:DAT", "/DCOPY:DA", "/XD", ".godot", "/R:3", "/W:5", "/NP", "/TEE", "/LOG+:$robocopyLogPath"
        )
        robocopy @robocopyArgs
        Write-Verbose "Robocopy operation complete. Log file created at: $robocopyLogPath"

        Write-Verbose "Generating new backup manifest..."
        $sourceFiles = Get-ChildItem -Path $BackupTarget -Recurse -File | Where-Object { $_.FullName -notmatch '\\.godot\\' }
        $manifestData = foreach ($file in $sourceFiles) {
            [PSCustomObject]@{
                FullName       = $file.FullName
                RelativePath   = $file.FullName.Substring($BackupTarget.Length)
                LastWriteTime  = $file.LastWriteTime
                BackupDate     = $timestamp
                BackupLocation = $newBackupFolderPath
            }
        }
        $manifestData | Export-Csv -Path $logCsvPath -NoTypeInformation
        Write-Host "Recovery backup and manifest creation complete." -ForegroundColor Green
    }
    # SCENARIO 3: Incremental Backup (Destination and Manifest both exist)
    else {
        Write-Host "Scenario: Incremental Backup." -ForegroundColor Green
        $newBackupFolderName = "Update_$timestamp"
        $newBackupFolderPath = Join-Path -Path $BackupDestination -ChildPath $newBackupFolderName
        $robocopyLogPath = Join-Path -Path $Logs -ChildPath "Robocopy_Update_$timestamp.log"

        # OPTIMIZED: Get last backup time from the manifest file's timestamp.
        $lastBackupTime = (Get-Item -Path $logCsvPath).LastWriteTime
        Write-Verbose "Scanning for files modified after $lastBackupTime"

        # Find new or modified files since the last backup, excluding specified folders
        $filesToBackup = Get-ChildItem -Path $BackupTarget -Recurse -File | Where-Object {
            ($_.LastWriteTime -gt $lastBackupTime) -and ($_.FullName -notmatch '\\.godot\\')
        }

        if ($filesToBackup.Count -gt 0) {
            Write-Host "Found $($filesToBackup.Count) new or modified files to back up."
            Write-Verbose "Creating new backup folder for updates: $newBackupFolderPath"
            New-Item -Path $newBackupFolderPath -ItemType Directory -Force | Out-Null

            # Create a temporary directory structure for Robocopy
            $tempSourcePath = Join-Path $env:TEMP "BackupSource_$timestamp"
            New-Item -Path $tempSourcePath -ItemType Directory -Force | Out-Null

            Write-Verbose "Staging files for Robocopy in '$tempSourcePath'..."
            foreach ($file in $filesToBackup) {
                $relativeDirPath = Split-Path -Path $file.FullName.Substring($BackupTarget.Length) -Parent
                $destDir = Join-Path -Path $tempSourcePath -ChildPath $relativeDirPath
                New-Item -Path $destDir -ItemType Directory -Force | Out-Null
                Copy-Item -Path $file.FullName -Destination $destDir
            }

            Write-Verbose "Performing incremental backup using Robocopy..."
            # Note: No /XD needed here as files are pre-filtered before staging
            $robocopyArgs = @(
                $tempSourcePath,
                $newBackupFolderPath,
                "/E", "/COPY:DAT", "/DCOPY:DA", "/R:3", "/W:5", "/NP", "/TEE", "/LOG+:$robocopyLogPath"
            )
            robocopy @robocopyArgs
            Write-Verbose "Robocopy operation complete. Log file created at: $robocopyLogPath"

            # Clean up the temporary staging folder
            Remove-Item -Path $tempSourcePath -Recurse -Force
            
            # Append new entries to the manifest
            Write-Verbose "Appending new entries to the manifest..."
            $newManifestEntries = foreach ($file in $filesToBackup) {
                [PSCustomObject]@{
                    FullName       = $file.FullName
                    RelativePath   = $file.FullName.Substring($BackupTarget.Length)
                    LastWriteTime  = $file.LastWriteTime
                    BackupDate     = $timestamp
                    BackupLocation = $newBackupFolderPath
                }
            }
            $newManifestEntries | Export-Csv -Path $logCsvPath -Append -NoTypeInformation
            Write-Host "Incremental backup complete. Manifest has been updated." -ForegroundColor Green
        }
        else {
            Write-Host "No new or modified files found. Backup is up to date." -ForegroundColor Yellow
        }
    }
}
catch {
    Write-Error "An unexpected error occurred: $($_.Exception.Message)"
    Write-Error "At: $($_.InvocationInfo.ScriptLineNumber): $($_.InvocationInfo.Line)"
    exit 1
}

Write-Verbose "Script finished."
