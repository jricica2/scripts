# Interactive RobocopyData_BackToPC.ps1
# Purpose: Restore user profile folders from an external drive after computer re-imaging
<#
.SYNOPSIS
    Interactive script to restore user profile folders from a backup using RoboCopy.

.DESCRIPTION
    This script provides an interactive interface to restore previously backed up user profile 
    folders from an external drive or network location. It allows users to:
    - Select the source backup drive and path
    - Choose which folders to restore from those available in the backup
    - Configure RoboCopy parameters like thread count
    - View detailed progress and summary information
    
    The script uses RoboCopy to perform the restore with optimized settings for reliability and performance:
    - Detects and verifies existing backups using timestamp information
    - Shows size information for available backup folders
    - Excludes problematic files like desktop.ini and system files
    - Uses multithreaded copying for better performance
    - Provides detailed logging and error handling

.NOTES
    - This script requires Administrator privileges to restore some system files
    - RoboCopy is a native Windows tool and does not require additional installation
    - The script automatically excludes temporary files and junction points that can cause issues
    - Only folders that actually exist in the backup will be shown as available to restore
    - Detailed logs are saved to help troubleshoot any issues

.AUTHOR
    Updated by Claude on March 20, 2025
#>

# Function to display menu and get user selection
function Show-FolderSelectionMenu {
    param (
        [array]$FolderOptions,
        [string]$BackupDrive
    )
    
    $selectedFolders = @()
    $menuOptions = @{}
    $availableFolders = @()
    
    # Only show folders that actually exist in the backup
    foreach ($folder in $FolderOptions) {
        $folderPath = Join-Path -Path $BackupDrive -ChildPath $folder
        if (Test-Path -Path $folderPath) {
            $availableFolders += $folder
        }
    }
    
    for ($i=0; $i -lt $availableFolders.Count; $i++) {
        $menuOptions.Add(($i+1), $availableFolders[$i])
    }
    
    do {
        Clear-Host
        Write-Host "=== FOLDER SELECTION MENU ===" -ForegroundColor Cyan
        Write-Host "Select folders to restore (toggle on/off):" -ForegroundColor Yellow
        
        foreach ($key in ($menuOptions.Keys | Sort-Object)) {
            $folder = $menuOptions[$key]
            $status = if ($selectedFolders -contains $folder) { "[X]" } else { "[ ]" }
            
            # Get folder size information
            $folderPath = Join-Path -Path $BackupDrive -ChildPath $folder
            $folderSize = (Get-ChildItem -Path $folderPath -Recurse -File | Measure-Object -Property Length -Sum).Sum
            $folderSizeMB = [math]::Round($folderSize / 1MB, 2)
            
            Write-Host "$key. $status $folder ($folderSizeMB MB)"
        }
        
        Write-Host "`nA. Select All"
        Write-Host "N. Select None"
        Write-Host "C. Continue with selected folders"
        Write-Host "Q. Quit"
        
        $choice = Read-Host "`nEnter your choice"
        
        if ($choice -eq "A" -or $choice -eq "a") {
            $selectedFolders = $menuOptions.Values
        }
        elseif ($choice -eq "N" -or $choice -eq "n") {
            $selectedFolders = @()
        }
        elseif ($choice -eq "C" -or $choice -eq "c") {
            break
        }
        elseif ($choice -eq "Q" -or $choice -eq "q") {
            Write-Host "Operation cancelled by user."
            exit 0
        }
        elseif ($menuOptions.ContainsKey([int]$choice)) {
            $folder = $menuOptions[[int]$choice]
            if ($selectedFolders -contains $folder) {
                $selectedFolders = $selectedFolders | Where-Object { $_ -ne $folder }
            } else {
                $selectedFolders += $folder
            }
        }
    } while ($true)
    
    return $selectedFolders
}

# Default folders that might be backed up
$defaultFolders = @(
    "Desktop",
    "Documents",
    "Downloads",
    "Pictures",
    "Videos",
    "Music",
    "Favorites",
    "OneDrive",
    "AppData\Roaming"
)

# Get the user profile path
$userProfile = [System.Environment]::GetFolderPath("UserProfile")

# Ask for source backup drive/path
Write-Host "=== RESTORE SOURCE ===" -ForegroundColor Cyan
Write-Host "Available drives:" -ForegroundColor Yellow
$drives = Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Free -gt 0 }
foreach ($drive in $drives) {
    $freeGB = [math]::Round($drive.Free / 1GB, 2)
    Write-Host "$($drive.Name): - $($drive.Description) - $freeGB GB free"
}

$backupFound = $false
do {
    $driveLetter = Read-Host "`nEnter drive letter where backup is located (e.g., D)"
    
    if ($driveLetter -eq "") {
        Write-Host "Operation cancelled by user."
        exit 0
    }
    
    $externalDrive = "${driveLetter}:"
    if (!(Test-Path -Path $externalDrive)) {
        Write-Host "Drive $externalDrive not found. Please enter a valid drive letter." -ForegroundColor Red
        continue
    }
    
    # Check if this is potentially a backup drive
    $timestampFile = Join-Path -Path $externalDrive -ChildPath "backup_timestamp.txt"
    $customPath = ""
    
    if (Test-Path -Path $timestampFile) {
        $backupFound = $true
        $backupDate = Get-Content -Path $timestampFile
        Write-Host "Backup found on root of drive $externalDrive!" -ForegroundColor Green
        Write-Host "Backup was created on: $backupDate" -ForegroundColor Green
    } else {
        # If timestamp not found at root, ask for subfolder
        $customPath = Read-Host "Backup not found at drive root. Enter subfolder path where backup is stored"
        if ($customPath -ne "") {
            $externalDrive = Join-Path -Path $externalDrive -ChildPath $customPath
            $timestampFile = Join-Path -Path $externalDrive -ChildPath "backup_timestamp.txt"
            
            if (Test-Path -Path $timestampFile) {
                $backupFound = $true
                $backupDate = Get-Content -Path $timestampFile
                Write-Host "Backup found at $externalDrive!" -ForegroundColor Green
                Write-Host "Backup was created on: $backupDate" -ForegroundColor Green
            } else {
                Write-Host "No backup timestamp found at $externalDrive. This might not be a valid backup location." -ForegroundColor Yellow
                $confirm = Read-Host "Do you want to continue anyway? (Y/N)"
                if ($confirm -eq 'Y' -or $confirm -eq 'y') {
                    $backupFound = $true
                }
            }
        } else {
            Write-Host "No path specified. Please try again." -ForegroundColor Yellow
        }
    }
} while (!$backupFound)

# Show folder selection menu from available folders
Write-Host "`nVerifying available folders in backup..."
$foldersToCopy = Show-FolderSelectionMenu -FolderOptions $defaultFolders -BackupDrive $externalDrive

if ($foldersToCopy.Count -eq 0) {
    Write-Host "No folders selected for restore. Operation cancelled." -ForegroundColor Yellow
    exit 0
}

# Ask for the number of threads to use
$defaultThreads = 16
do {
    $threadsInput = Read-Host "`nEnter number of threads to use for copying (default: $defaultThreads)"
    if ($threadsInput -eq "") {
        $threads = $defaultThreads
        break
    }
    elseif ($threadsInput -match "^\d+$" -and [int]$threadsInput -gt 0 -and [int]$threadsInput -le 128) {
        $threads = [int]$threadsInput
        break
    }
    else {
        Write-Host "Please enter a valid number between 1 and 128." -ForegroundColor Red
    }
} while ($true)

# Set up logging
$logPath = Join-Path -Path $userProfile -ChildPath "Logs"
$dateTime = Get-Date -Format "yyyyMMdd_HHmmss"
$logFile = Join-Path -Path $logPath -ChildPath "robocopy_restore_$dateTime.log"

# Create logs directory if it doesn't exist
if (!(Test-Path -Path $logPath)) {
    try {
        New-Item -Path $logPath -ItemType Directory -Force | Out-Null
        Write-Host "Created logs directory: $logPath" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to create logs directory: $_"
        $logFile = Join-Path -Path $userProfile -ChildPath "robocopy_restore_$dateTime.log"
        Write-Host "Logging to: $logFile" -ForegroundColor Yellow
    }
}

# Show a summary of what will be restored
Write-Host "`n=== RESTORE OPERATION SUMMARY ===" -ForegroundColor Cyan
Write-Host "Source: $externalDrive"
Write-Host "Destination: $userProfile"
Write-Host "Folders to restore: $($foldersToCopy -join ', ')"
Write-Host "Threads: $threads"
Write-Host "Log file: $logFile"
Write-Host "=================================" -ForegroundColor Cyan

# Calculate total size to be restored
$totalSizeToRestore = 0
foreach ($folder in $foldersToCopy) {
    $folderPath = Join-Path -Path $externalDrive -ChildPath $folder
    if (Test-Path -Path $folderPath) {
        $size = (Get-ChildItem -Path $folderPath -Recurse -File | Measure-Object -Property Length -Sum).Sum
        $totalSizeToRestore += $size
    }
}

$sizeToRestoreGB = [math]::Round($totalSizeToRestore / 1GB, 2)
$sizeToRestoreMB = [math]::Round($totalSizeToRestore / 1MB, 2)

if ($sizeToRestoreMB -lt 1000) {
    Write-Host "Total size to restore: $sizeToRestoreMB MB" -ForegroundColor Yellow
} else {
    Write-Host "Total size to restore: $sizeToRestoreGB GB" -ForegroundColor Yellow
}

# Ask for confirmation
$confirmation = Read-Host "Do you want to proceed with the restore? (Y/N)"
if ($confirmation -ne 'Y' -and $confirmation -ne 'y') {
    Write-Host "Restore operation cancelled by user." -ForegroundColor Yellow
    exit 0
}

# Copy each folder back to the user profile
$totalFilesCopied = 0
$totalBytesCopied = 0

foreach ($folder in $foldersToCopy) {
    $sourcePath = Join-Path -Path $externalDrive -ChildPath $folder
    $destinationPath = Join-Path -Path $userProfile -ChildPath $folder
    
    # Skip if source folder doesn't exist in backup
    if (!(Test-Path -Path $sourcePath)) {
        Write-Host "Skipping $folder - folder not found in backup." -ForegroundColor Yellow
        continue
    }
    
    # Create destination directory if it doesn't exist
    if (!(Test-Path -Path $destinationPath)) {
        New-Item -Path $destinationPath -ItemType Directory -Force | Out-Null
    }

    # Run RoboCopy
    Write-Host "`nRestoring $folder back to the user profile..." -ForegroundColor Cyan
    
    # Robocopy parameters
    $result = robocopy $sourcePath $destinationPath /MIR /MT:$threads /R:5 /W:5 `
        /XF "desktop.ini" "ntuser.dat*" "NTUSER.DAT*" "*.tmp" `
        /XJ /ZB /NP /BYTES /DCOPY:T /COPY:DAT /LOG+:$logFile /TEE
    
    # Evaluate the robocopy exit code
    switch ($LASTEXITCODE) {
        0 { Write-Host "No files were copied for $folder. Source and destination are synchronized." -ForegroundColor Green }
        1 { 
            Write-Host "Files were copied successfully for $folder." -ForegroundColor Green 
            # Try to parse robocopy output for statistics
            if ($result -match "Files : \s*(\d+)") { $totalFilesCopied += [int]$Matches[1] }
            if ($result -match "Bytes : \s*(\d+)") { $totalBytesCopied += [int64]$Matches[1] }
        }
        2 { Write-Host "Extra files or directories were detected in $folder." -ForegroundColor Yellow }
        3 { 
            Write-Host "Some files were copied, some failed for $folder." -ForegroundColor Yellow 
            # Try to parse robocopy output for statistics
            if ($result -match "Files : \s*(\d+)") { $totalFilesCopied += [int]$Matches[1] }
            if ($result -match "Bytes : \s*(\d+)") { $totalBytesCopied += [int64]$Matches[1] }
        }
        {$_ -ge 8} { 
            Write-Host "At least one failure occurred during copying of $folder." -ForegroundColor Red
            if ($_ -band 8) { Write-Host "Some files or directories could not be copied (copy errors occurred)." -ForegroundColor Red }
            if ($_ -band 16) { Write-Host "Serious error - robocopy did not copy any files." -ForegroundColor Red }
        }
        default { Write-Host "Completed with return code $LASTEXITCODE" -ForegroundColor Yellow }
    }
}

$copiedSizeGB = [math]::Round($totalBytesCopied / 1GB, 2)
$copiedSizeMB = [math]::Round($totalBytesCopied / 1MB, 2)

Write-Host "`n=== RESTORE SUMMARY ===" -ForegroundColor Cyan
Write-Host "Restore completed at: $(Get-Date)" -ForegroundColor Green
if ($copiedSizeMB -lt 1000) {
    Write-Host "Total size copied: $copiedSizeMB MB" -ForegroundColor Green
} else {
    Write-Host "Total size copied: $copiedSizeGB GB" -ForegroundColor Green
}
Write-Host "Files copied: $totalFilesCopied" -ForegroundColor Green
Write-Host "Restore log saved to: $logFile" -ForegroundColor Green

Write-Host "`nRestore operation completed. Please verify your files."
