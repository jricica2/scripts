# Interactive RobocopyData.ps1
# Purpose: Back up user profile folders to an external drive before computer re-imaging
<#
.SYNOPSIS
    Interactive script to back up user profile folders using RoboCopy.

.DESCRIPTION
    This script provides an interactive interface to back up user profile folders to an external drive 
    or network location. It allows users to:
    - Select the destination drive and path
    - Choose which folders to back up from a predefined list
    - Configure RoboCopy parameters like thread count
    - View detailed progress and summary information
    
    The script uses RoboCopy to perform the backup with optimized settings for reliability and performance:
    - Excludes problematic files like desktop.ini and system files
    - Uses multithreaded copying for better performance
    - Provides detailed logging and error handling
    - Creates timestamped backups for verification

.NOTES
    - This script requires Administrator privileges to copy some system files
    - RoboCopy is a native Windows tool and does not require additional installation
    - The script automatically excludes temporary files and junction points that can cause issues
    - A timestamp file is created to verify when the backup was made
    - Detailed logs are saved to help troubleshoot any issues

.AUTHOR
    Updated by Claude on March 20, 2025
#>

# Function to display menu and get user selection
function Show-FolderSelectionMenu {
    param (
        [array]$FolderOptions
    )
    
    # Initialize arrays to avoid null array issues
    $selectedFolders = @()
    $menuOptions = @{}
    
    # Safety check for empty folder options
    if ($null -eq $FolderOptions -or $FolderOptions.Count -eq 0) {
        Write-Host "No folders available for selection." -ForegroundColor Red
        return @()
    }
    
    for ($i=0; $i -lt $FolderOptions.Count; $i++) {
        $menuOptions.Add(($i+1), $FolderOptions[$i])
    }
    
    do {
        Clear-Host
        Write-Host "=== FOLDER SELECTION MENU ===" -ForegroundColor Cyan
        Write-Host "Select folders to back up (toggle on/off):" -ForegroundColor Yellow
        
        foreach ($key in ($menuOptions.Keys | Sort-Object)) {
            $folder = $menuOptions[$key]
            $status = if ($selectedFolders -contains $folder) { "[X]" } else { "[ ]" }
            Write-Host "$key. $status $folder"
        }
        
        Write-Host "`nA. Select All"
        Write-Host "N. Select None"
        Write-Host "C. Continue with selected folders"
        Write-Host "Q. Quit"
        
        $choice = Read-Host "`nEnter your choice"
        
        if ($choice -eq "A" -or $choice -eq "a") {
            $selectedFolders = @($menuOptions.Values)
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
                $selectedFolders = @($selectedFolders | Where-Object { $_ -ne $folder })
            } else {
                $selectedFolders += $folder
            }
        }
    } while ($true)
    
    return $selectedFolders
}

# Default folders that can be backed up
$defaultFolders = @(
    "Desktop",
    "Documents",
    "Downloads",
    "Pictures",
    "Videos",
    "Music",
    "Favorites",
    "OneDrive"
)

# Get the user profile path
$userProfile = [System.Environment]::GetFolderPath("UserProfile")

# Ask for destination drive/path
Write-Host "=== BACKUP DESTINATION ===" -ForegroundColor Cyan
Write-Host "Available drives:" -ForegroundColor Yellow
$drives = @(Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Free -gt 0 })
if ($drives.Count -eq 0) {
    Write-Host "No available drives found." -ForegroundColor Red
    exit 1
}

foreach ($drive in $drives) {
    # Safely calculate free space
    $freeGB = 0
    if ($null -ne $drive.Free) {
        $freeGB = [math]::Round($drive.Free / 1GB, 2)
    }
    Write-Host "$($drive.Name): - $($drive.Description) - $freeGB GB free"
}

$validPath = $false
$externalDrive = ""
do {
    $driveLetter = Read-Host "`nEnter drive letter for backup destination (e.g., D)"
    
    if ($driveLetter -eq "") {
        Write-Host "Operation cancelled by user."
        exit 0
    }
    
    $externalDrive = "${driveLetter}:"
    if (!(Test-Path -Path $externalDrive)) {
        Write-Host "Drive $externalDrive not found. Please enter a valid drive letter." -ForegroundColor Red
        continue
    }
    
    $customPath = Read-Host "Enter subfolder path for backup (press Enter for root of $externalDrive)"
    if ($customPath -ne "") {
        $externalDrive = Join-Path -Path $externalDrive -ChildPath $customPath
    }
    
    $validPath = $true
    
    # Try to create directory if it doesn't exist
    if (!(Test-Path -Path $externalDrive)) {
        try {
            New-Item -Path $externalDrive -ItemType Directory -Force | Out-Null
            Write-Host "Created directory: $externalDrive" -ForegroundColor Green
        }
        catch {
            Write-Host "Could not create directory $externalDrive. Error: $_" -ForegroundColor Red
            $validPath = $false
        }
    }
} while (!$validPath)

# Show folder selection menu
Write-Host "`nSelect folders to back up:"
$foldersToCopy = @(Show-FolderSelectionMenu -FolderOptions $defaultFolders)

if ($foldersToCopy.Count -eq 0) {
    Write-Host "No folders selected for backup. Operation cancelled." -ForegroundColor Yellow
    exit 0
}

# Ask for the number of threads to use
$defaultThreads = 16
$threads = $defaultThreads
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
$logPath = Join-Path -Path $externalDrive -ChildPath "Logs"
$dateTime = Get-Date -Format "yyyyMMdd_HHmmss"
$logFile = Join-Path -Path $logPath -ChildPath "robocopy_backup_$dateTime.log"

# Create logs directory if it doesn't exist
if (!(Test-Path -Path $logPath)) {
    try {
        New-Item -Path $logPath -ItemType Directory -Force | Out-Null
        Write-Host "Created logs directory: $logPath" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to create logs directory: $_"
        $logFile = Join-Path -Path $externalDrive -ChildPath "robocopy_backup_$dateTime.log"
        Write-Host "Logging to: $logFile" -ForegroundColor Yellow
    }
}

# Show a summary of what will be backed up
Write-Host "`n=== BACKUP OPERATION SUMMARY ===" -ForegroundColor Cyan
Write-Host "Source: $userProfile"
Write-Host "Destination: $externalDrive"
Write-Host "Folders to copy: $($foldersToCopy -join ', ')"
Write-Host "Threads: $threads"
Write-Host "Log file: $logFile"
Write-Host "===============================" -ForegroundColor Cyan

# Ask for confirmation
$confirmation = Read-Host "Do you want to proceed with the backup? (Y/N)"
if ($confirmation -ne 'Y' -and $confirmation -ne 'y') {
    Write-Host "Backup operation cancelled by user." -ForegroundColor Yellow
    exit 0
}

# Create a timestamp file to verify backup date later
$timestampFile = Join-Path -Path $externalDrive -ChildPath "backup_timestamp.txt"
Get-Date | Out-File -FilePath $timestampFile -Force

# Copy each folder
$totalFilesCopied = 0
$totalBytesCopied = 0

foreach ($folder in $foldersToCopy) {
    $sourcePath = Join-Path -Path $userProfile -ChildPath $folder
    $destinationPath = Join-Path -Path $externalDrive -ChildPath $folder
    
    # Skip if source folder doesn't exist
    if (!(Test-Path -Path $sourcePath)) {
        Write-Host "Skipping $folder - source folder doesn't exist." -ForegroundColor Yellow
        continue
    }

    # Run RoboCopy with improved parameters
    Write-Host "`nCopying $folder..." -ForegroundColor Cyan
    
    # Run robocopy and capture output
    # Convert output to array to avoid null array issues
    $robocopyOutput = @(robocopy $sourcePath $destinationPath /MIR /MT:$threads /R:5 /W:5 `
        /XF "desktop.ini" "ntuser.dat*" "NTUSER.DAT*" "*.tmp" `
        /XD "AppData\Local\Temp" "AppData\LocalLow" `
        /XJ /ZB /NP /BYTES /DCOPY:T /COPY:DAT /LOG+:$logFile /TEE)
    
    # Save the last exit code before it gets overwritten
    $robocopyExitCode = $LASTEXITCODE
    
    # Safely extract statistics from output
    try {
        $fileStats = $robocopyOutput | Select-String -Pattern "Files :" -SimpleMatch
        $byteStats = $robocopyOutput | Select-String -Pattern "Bytes :" -SimpleMatch
        
        if ($fileStats) {
            $fileMatch = $fileStats -match "Files :\s*(\d+)"
            if ($fileMatch -and $Matches -and $Matches[1]) {
                $totalFilesCopied += [int]$Matches[1]
            }
        }
        
        if ($byteStats) {
            $byteMatch = $byteStats -match "Bytes :\s*(\d+)"
            if ($byteMatch -and $Matches -and $Matches[1]) {
                $totalBytesCopied += [int64]$Matches[1]
            }
        }
    }
    catch {
        Write-Host "Error parsing robocopy output: $_" -ForegroundColor Yellow
    }
    
    # Evaluate the robocopy exit code
    # See https://learn.microsoft.com/en-us/windows-server/administration/windows-commands/robocopy-exit-codes
    switch ($robocopyExitCode) {
        0 { Write-Host "No files were copied for $folder. Source and destination are synchronized." -ForegroundColor Green }
        1 { Write-Host "Files were copied successfully for $folder." -ForegroundColor Green }
        2 { Write-Host "Extra files or directories were detected in $folder." -ForegroundColor Yellow }
        3 { Write-Host "Some files were copied, some failed for $folder." -ForegroundColor Yellow }
        {$_ -ge 8} { 
            Write-Host "At least one failure occurred during copying of $folder." -ForegroundColor Red
            if ($_ -band 8) { Write-Host "Some files or directories could not be copied (copy errors occurred)." -ForegroundColor Red }
            if ($_ -band 16) { Write-Host "Serious error - robocopy did not copy any files." -ForegroundColor Red }
        }
        default { Write-Host "Completed with return code $robocopyExitCode" -ForegroundColor Yellow }
    }
}

# Calculate and display total backup size - using safer methods
$totalSize = 0
foreach ($folder in $foldersToCopy) {
    $folderPath = Join-Path -Path $externalDrive -ChildPath $folder
    if (Test-Path -Path $folderPath) {
        try {
            $folderFiles = @(Get-ChildItem -Path $folderPath -Recurse -File -ErrorAction SilentlyContinue)
            if ($folderFiles.Count -gt 0) {
                $folderSize = ($folderFiles | Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
                if ($null -ne $folderSize) {
                    $totalSize += $folderSize
                }
            }
        } 
        catch {
            Write-Host "Error calculating size for $folder : $_" -ForegroundColor Yellow
        }
    }
}

$sizeGB = [math]::Round($totalSize / 1GB, 2)
$sizeMB = [math]::Round($totalSize / 1MB, 2)

Write-Host "`n=== BACKUP SUMMARY ===" -ForegroundColor Cyan
Write-Host "Backup completed at: $(Get-Date)" -ForegroundColor Green
if ($sizeMB -lt 1000) {
    Write-Host "Total size: $sizeMB MB" -ForegroundColor Green
} else {
    Write-Host "Total size: $sizeGB GB" -ForegroundColor Green
}
Write-Host "Backup log saved to: $logFile" -ForegroundColor Green
Write-Host "Files copied: $totalFilesCopied" -ForegroundColor Green

Write-Host "`nTo restore this backup, run the RobocopyData_BackToPC.ps1 script."
