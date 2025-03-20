# UnifiedRobocopyTool.ps1
# Purpose: All-in-one tool for backing up and restoring user profile folders using RoboCopy
<#
.SYNOPSIS
    Interactive unified script to back up or restore user profile folders using RoboCopy.

.DESCRIPTION
    This script provides an interactive interface for two primary operations:
    1. BACKUP: Copy user profile folders to an external drive before computer re-imaging
    2. RESTORE: Copy previously backed up user profile folders back to the user profile

    The script allows users to:
    - Select the operation mode (backup or restore)
    - Choose source and destination locations
    - Select which folders to copy
    - Configure RoboCopy parameters
    - Monitor progress with a real-time progress bar
    
    The script uses RoboCopy with optimized settings for reliability and performance:
    - Excludes problematic files like desktop.ini and system files
    - Uses multithreaded copying for better performance
    - Provides detailed logging and error handling
    - Shows real-time progress information with ETA

.NOTES
    - This script requires Administrator privileges to copy some system files
    - RoboCopy is a native Windows tool and does not require additional installation
    - The script automatically excludes temporary files and junction points that can cause issues
    - A timestamp file is created during backup to verify when the backup was made
    - Detailed logs are saved to help troubleshoot any issues

.AUTHOR
    Updated by Claude on March 20, 2025
#>

#region FUNCTIONS

# Function to display menu and get user selection for backup folders
function Show-FolderSelectionMenu {
    param (
        [array]$FolderOptions,
        [string]$BackupDrive = "",
        [bool]$CheckExistence = $false
    )
    
    # Initialize arrays to avoid null array issues
    $selectedFolders = @()
    $menuOptions = @{}
    $availableFolders = @()
    
    # Safety check for empty folder options
    if ($null -eq $FolderOptions -or $FolderOptions.Count -eq 0) {
        Write-Host "No folders available for selection." -ForegroundColor Red
        return @()
    }
    
    # If we need to check for existence (in restore mode)
    if ($CheckExistence -and $BackupDrive -ne "") {
        foreach ($folder in $FolderOptions) {
            $folderPath = Join-Path -Path $BackupDrive -ChildPath $folder
            if (Test-Path -Path $folderPath) {
                $availableFolders += $folder
            }
        }
        
        # Check if we found any available folders
        if ($availableFolders.Count -eq 0) {
            Write-Host "No valid folders found in backup location $BackupDrive." -ForegroundColor Red
            return @()
        }
    } else {
        $availableFolders = $FolderOptions
    }
    
    for ($i=0; $i -lt $availableFolders.Count; $i++) {
        $menuOptions.Add(($i+1), $availableFolders[$i])
    }
    
    do {
        Clear-Host
        Write-Host "=== FOLDER SELECTION MENU ===" -ForegroundColor Cyan
        Write-Host "Select folders to process (toggle on/off):" -ForegroundColor Yellow
        
        foreach ($key in ($menuOptions.Keys | Sort-Object)) {
            $folder = $menuOptions[$key]
            $status = if ($selectedFolders -contains $folder) { "[X]" } else { "[ ]" }
            
            # Show folder size if in restore mode
            if ($CheckExistence -and $BackupDrive -ne "") {
                # Get folder size information - using safer methods
                $folderSizeMB = 0
                $folderPath = Join-Path -Path $BackupDrive -ChildPath $folder
                
                try {
                    $folderFiles = @(Get-ChildItem -Path $folderPath -Recurse -File -ErrorAction SilentlyContinue)
                    if ($folderFiles.Count -gt 0) {
                        $folderSize = ($folderFiles | Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
                        if ($null -ne $folderSize) {
                            $folderSizeMB = [math]::Round($folderSize / 1MB, 2)
                        }
                    }
                }
                catch {
                    # Silently handle errors and just show 0 MB
                }
                
                Write-Host "$key. $status $folder ($folderSizeMB MB)"
            } else {
                Write-Host "$key. $status $folder"
            }
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

# Function to handle the backup operation
function Invoke-BackupOperation {
    # Get the user profile path
    $userProfile = [System.Environment]::GetFolderPath("UserProfile")

    # Ask for destination drive/path
    Write-Host "=== BACKUP DESTINATION ===" -ForegroundColor Cyan
    Write-Host "Available drives:" -ForegroundColor Yellow
    $drives = @(Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Free -gt 0 })
    if ($drives.Count -eq 0) {
        Write-Host "No available drives found." -ForegroundColor Red
        return
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
            return
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
    $foldersToCopy = @(Show-FolderSelectionMenu -FolderOptions $script:defaultFolders)

    if ($foldersToCopy.Count -eq 0) {
        Write-Host "No folders selected for backup. Operation cancelled." -ForegroundColor Yellow
        return
    }

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
    Write-Host "Threads: $script:threads"
    Write-Host "Log file: $logFile"
    Write-Host "===============================" -ForegroundColor Cyan

    # Ask for confirmation
    $confirmation = Read-Host "Do you want to proceed with the backup? (Y/N)"
    if ($confirmation -ne 'Y' -and $confirmation -ne 'y') {
        Write-Host "Backup operation cancelled by user." -ForegroundColor Yellow
        return
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
		
# First, count total files and size to calculate progress
        Write-Host "Analyzing $folder to calculate total files and size..." -ForegroundColor Yellow
        $totalFilesInFolder = 0
        $totalSizeInFolder = 0
        try {
            $folderFiles = @(Get-ChildItem -Path $sourcePath -Recurse -File -ErrorAction SilentlyContinue)
            $totalFilesInFolder = $folderFiles.Count
            $totalSizeInFolder = ($folderFiles | Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
            
            $displaySize = if ($totalSizeInFolder -gt 1GB) {
                "$([math]::Round($totalSizeInFolder / 1GB, 2)) GB"
            } else {
                "$([math]::Round($totalSizeInFolder / 1MB, 2)) MB"
            }
            
            Write-Host "Found $totalFilesInFolder files totaling $displaySize" -ForegroundColor Yellow
        }
        catch {
            Write-Host "Error calculating folder statistics: $_" -ForegroundColor Red
            Write-Host "Will proceed without progress indication for this folder." -ForegroundColor Yellow
        }
        
# This is a replacement for the progress monitoring section in both backup and restore functions

# Run robocopy with progress monitoring - IMPROVED VERSION
if ($totalFilesInFolder -gt 0) {
    Write-Host "Starting copy with progress monitoring..." -ForegroundColor Green
    
    # Use a temporary file to store robocopy output for live monitoring
    $tempOutputFile = [System.IO.Path]::GetTempFileName()
    
    # Start robocopy as a process so we can monitor it in real-time
    $processStartInfo = New-Object System.Diagnostics.ProcessStartInfo
    $processStartInfo.FileName = "robocopy.exe"
    $processStartInfo.Arguments = "`"$sourcePath`" `"$destinationPath`" /MIR /MT:$script:threads /R:5 /W:5 " +
                                  "/XF `"desktop.ini`" `"ntuser.dat*`" `"NTUSER.DAT*`" `"*.tmp`" " +
                                  "/XD `"AppData\Local\Temp`" `"AppData\LocalLow`" " +
                                  "/XJ /Z /BYTES /DCOPY:T /COPY:DAT /LOG+:`"$logFile`" /TEE /NP"
    $processStartInfo.RedirectStandardOutput = $true
    $processStartInfo.RedirectStandardError = $true
    $processStartInfo.UseShellExecute = $false
    $processStartInfo.CreateNoWindow = $true
    
    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $processStartInfo
    $process.EnableRaisingEvents = $true
    
    # Setup output handlers
    $outputBuilder = New-Object System.Text.StringBuilder
    $errorBuilder = New-Object System.Text.StringBuilder
    
    $filesCopied = 0
    $lastReportedFiles = 0
    $bytesTransferred = 0
    $copyComplete = $false
    
    # Register output event handler
    $outputHandler = {
        if ($EventArgs.Data -match '^\s*(\d+)\s+(\S.+)') {
            # This is likely a file being copied
            $filesCopied++
        }
        elseif ($EventArgs.Data -match 'Files\s*:\s*(\d+)') {
            # This is the summary line with total files copied
            $filesCopied = [int]$Matches[1]
        }
        elseif ($EventArgs.Data -match 'Bytes\s*:\s*(\d+)') {
            # This is the summary line with total bytes transferred
            $bytesTransferred = [int64]$Matches[1]
        }
        elseif ($EventArgs.Data -match 'Speed\s*:\s*(.+)') {
            # This contains the transfer speed
            $transferSpeed = $Matches[1]
        }

        # Save to the output builder
        [void]$outputBuilder.AppendLine($EventArgs.Data)
    }
    
    $errorHandler = {
        [void]$errorBuilder.AppendLine($EventArgs.Data)
    }
    
    $processExitedHandler = {
        $copyComplete = $true
    }
    
    # Register the event handlers
    $outputEvent = Register-ObjectEvent -InputObject $process -EventName OutputDataReceived -Action $outputHandler
    $errorEvent = Register-ObjectEvent -InputObject $process -EventName ErrorDataReceived -Action $errorHandler
    $exitedEvent = Register-ObjectEvent -InputObject $process -EventName Exited -Action $processExitedHandler
    
    # Start the process
    [void]$process.Start()
    $process.BeginOutputReadLine()
    $process.BeginErrorReadLine()
    
    # Monitor progress
    $startTime = Get-Date
    $spinChars = '|', '/', '-', '\'
    $spinIndex = 0
    
    while (-not $copyComplete) {
        $elapsedTime = (Get-Date) - $startTime
        $elapsedSeconds = [math]::Max(1, $elapsedTime.TotalSeconds)
        
        # Calculate progress based on file count
        if ($filesCopied -gt $lastReportedFiles) {
            $lastReportedFiles = $filesCopied
        }
        
        $percentComplete = [math]::Min(100, [math]::Max(0, [math]::Round(($filesCopied / $totalFilesInFolder) * 100, 0)))
        
        # Create a spinner for activity indication even when no progress is visible
        $spinChar = $spinChars[$spinIndex % $spinChars.Length]
        $spinIndex++
        
        # Format status line - include spinner to show activity
        $status = "$spinChar $filesCopied of $totalFilesInFolder files"
        
        # Add transfer rate if available
        if ($bytesTransferred -gt 0) {
            $rate = [math]::Round($bytesTransferred / $elapsedSeconds / 1MB, 2)
            $status += " | $rate MB/s"
        }
        
        # Calculate ETA if we have progress
        if ($percentComplete -gt 0) {
            $totalSecondsRemaining = ($elapsedSeconds / $percentComplete) * (100 - $percentComplete)
            $timeRemaining = [timespan]::FromSeconds($totalSecondsRemaining)
            $eta = "{0:hh\:mm\:ss}" -f $timeRemaining
            $status += " | ETA: $eta"
        }
        
        # Create a progress bar
        $progressBarWidth = 50
        $filledWidth = [math]::Round(($percentComplete / 100) * $progressBarWidth)
        $progressBar = "[" + ("#" * $filledWidth) + (" " * ($progressBarWidth - $filledWidth)) + "]"
        
        Write-Host "`r$progressBar $percentComplete% $status     " -NoNewline
        
        Start-Sleep -Milliseconds 200
        
        # Check if process has exited
        if ($process.HasExited) {
            $copyComplete = $true
        }
    }
    
    # Complete the progress line
    Write-Host "`r" + (" " * 120) + "`r" -NoNewline
    
    # Cleanup events
    Unregister-Event -SourceIdentifier $outputEvent.Name
    Unregister-Event -SourceIdentifier $errorEvent.Name
    Unregister-Event -SourceIdentifier $exitedEvent.Name
    
    # Get the final output and cleanup
    $output = $outputBuilder.ToString()
    $errorOutput = $errorBuilder.ToString()
    
    # Wait for process to fully exit
    $process.WaitForExit()
    
    # Get exit code
    $robocopyExitCode = $process.ExitCode
    
    # Cleanup
    $process.Close()
    
    Write-Host "Copy operation completed for $folder" -ForegroundColor Green
    
    # Try to extract statistics from output
    try {
        if ($output -match "Files\s*:\s*(\d+)") {
            $totalFilesCopied += [int]$Matches[1]
        }
        
        if ($output -match "Bytes\s*:\s*(\d+)") {
            $totalBytesCopied += [int64]$Matches[1]
        }
    }
    catch {
        Write-Host "Error parsing robocopy output: $_" -ForegroundColor Yellow
    }
}
else {
    # Fallback to standard robocopy if we couldn't calculate file count
    $robocopyProcess = Start-Process -FilePath "robocopy.exe" -ArgumentList @(
        "`"$sourcePath`"",
        "`"$destinationPath`"",
        "/MIR",
        "/MT:$script:threads",
        "/R:5",
        "/W:5",
        "/XF", "desktop.ini", "ntuser.dat*", "NTUSER.DAT*", "*.tmp",
        "/XD", "AppData\Local\Temp", "AppData\LocalLow",
        "/XJ",
        "/ZB",
        "/NP",
        "/BYTES",
        "/DCOPY:T",
        "/COPY:DAT",
        "/LOG+:`"$logFile`"",
        "/TEE"
    ) -NoNewWindow -Wait -PassThru
    
    $robocopyExitCode = $robocopyProcess.ExitCode
}
        
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

    Write-Host "`nTo restore this backup, run this script again and select 'Restore'."
}

# Function to handle the restore operation
function Invoke-RestoreOperation {
    # Get the user profile path
    $userProfile = [System.Environment]::GetFolderPath("UserProfile")

    # Ask for source backup drive/path
    Write-Host "=== RESTORE SOURCE ===" -ForegroundColor Cyan
    Write-Host "Available drives:" -ForegroundColor Yellow
    $drives = @(Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Free -gt 0 })
    if ($drives.Count -eq 0) {
        Write-Host "No available drives found." -ForegroundColor Red
        return
    }

    foreach ($drive in $drives) {
        # Safely calculate free space
        $freeGB = 0
        if ($null -ne $drive.Free) {
            $freeGB = [math]::Round($drive.Free / 1GB, 2)
        }
        Write-Host "$($drive.Name): - $($drive.Description) - $freeGB GB free"
    }

    $backupFound = $false
    $externalDrive = ""
    do {
        $driveLetter = Read-Host "`nEnter drive letter where backup is located (e.g., D)"
        
        if ($driveLetter -eq "") {
            Write-Host "Operation cancelled by user."
            return
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
    $foldersToCopy = @(Show-FolderSelectionMenu -FolderOptions $script:defaultFolders -BackupDrive $externalDrive -CheckExistence $true)

    if ($null -eq $foldersToCopy -or $foldersToCopy.Count -eq 0) {
        Write-Host "No folders selected for restore. Operation cancelled." -ForegroundColor Yellow
        return
    }

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
    Write-Host "Threads: $script:threads"
    Write-Host "Log file: $logFile"
    Write-Host "=================================" -ForegroundColor Cyan

    # Calculate total size to be restored - using safer methods
    $totalSizeToRestore = 0
    foreach ($folder in $foldersToCopy) {
        $folderPath = Join-Path -Path $externalDrive -ChildPath $folder
        if (Test-Path -Path $folderPath) {
            try {
                $folderFiles = @(Get-ChildItem -Path $folderPath -Recurse -File -ErrorAction SilentlyContinue)
                if ($folderFiles.Count -gt 0) {
                    $folderSize = ($folderFiles | Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
                    if ($null -ne $folderSize) {
                        $totalSizeToRestore += $folderSize
                    }
                }
            }
            catch {
                Write-Host "Error calculating size for $folder : $_" -ForegroundColor Yellow
            }
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
        return
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
            try {
                New-Item -Path $destinationPath -ItemType Directory -Force | Out-Null
            }
            catch {
                Write-Host "Error creating destination directory $destinationPath : $_" -ForegroundColor Red
                continue
            }
        }

        # First, count total files and size to calculate progress
        Write-Host "Analyzing $folder to calculate total files and size..." -ForegroundColor Yellow
        $totalFilesInFolder = 0
        $totalSizeInFolder = 0
        try {
            $folderFiles = @(Get-ChildItem -Path $sourcePath -Recurse -File -ErrorAction SilentlyContinue)
            $totalFilesInFolder = $folderFiles.Count
            $totalSizeInFolder = ($folderFiles | Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
            
            $displaySize = if ($totalSizeInFolder -gt 1GB) {
                "$([math]::Round($totalSizeInFolder / 1GB, 2)) GB"
            } else {
                "$([math]::Round($totalSizeInFolder / 1MB, 2)) MB"
            }
            
            Write-Host "Found $totalFilesInFolder files totaling $displaySize" -ForegroundColor Yellow
        }
        catch {
            Write-Host "Error calculating folder statistics: $_" -ForegroundColor Red
            Write-Host "Will proceed without progress indication for this folder." -ForegroundColor Yellow
        }
        
# This is a replacement for the progress monitoring section in both backup and restore functions

# Run robocopy with progress monitoring - IMPROVED VERSION
if ($totalFilesInFolder -gt 0) {
    Write-Host "Starting copy with progress monitoring..." -ForegroundColor Green
    
    # Use a temporary file to store robocopy output for live monitoring
    $tempOutputFile = [System.IO.Path]::GetTempFileName()
    
    # Start robocopy as a process so we can monitor it in real-time
    $processStartInfo = New-Object System.Diagnostics.ProcessStartInfo
    $processStartInfo.FileName = "robocopy.exe"
    $processStartInfo.Arguments = "`"$sourcePath`" `"$destinationPath`" /MIR /MT:$script:threads /R:5 /W:5 " +
                                  "/XF `"desktop.ini`" `"ntuser.dat*`" `"NTUSER.DAT*`" `"*.tmp`" " +
                                  "/XD `"AppData\Local\Temp`" `"AppData\LocalLow`" " +
                                  "/XJ /Z /BYTES /DCOPY:T /COPY:DAT /LOG+:`"$logFile`" /TEE /NP"
    $processStartInfo.RedirectStandardOutput = $true
    $processStartInfo.RedirectStandardError = $true
    $processStartInfo.UseShellExecute = $false
    $processStartInfo.CreateNoWindow = $true
    
    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $processStartInfo
    $process.EnableRaisingEvents = $true
    
    # Setup output handlers
    $outputBuilder = New-Object System.Text.StringBuilder
    $errorBuilder = New-Object System.Text.StringBuilder
    
    $filesCopied = 0
    $lastReportedFiles = 0
    $bytesTransferred = 0
    $copyComplete = $false
    
    # Register output event handler
    $outputHandler = {
        if ($EventArgs.Data -match '^\s*(\d+)\s+(\S.+)') {
            # This is likely a file being copied
            $filesCopied++
        }
        elseif ($EventArgs.Data -match 'Files\s*:\s*(\d+)') {
            # This is the summary line with total files copied
            $filesCopied = [int]$Matches[1]
        }
        elseif ($EventArgs.Data -match 'Bytes\s*:\s*(\d+)') {
            # This is the summary line with total bytes transferred
            $bytesTransferred = [int64]$Matches[1]
        }
        elseif ($EventArgs.Data -match 'Speed\s*:\s*(.+)') {
            # This contains the transfer speed
            $transferSpeed = $Matches[1]
        }

        # Save to the output builder
        [void]$outputBuilder.AppendLine($EventArgs.Data)
    }
    
    $errorHandler = {
        [void]$errorBuilder.AppendLine($EventArgs.Data)
    }
    
    $processExitedHandler = {
        $copyComplete = $true
    }
    
    # Register the event handlers
    $outputEvent = Register-ObjectEvent -InputObject $process -EventName OutputDataReceived -Action $outputHandler
    $errorEvent = Register-ObjectEvent -InputObject $process -EventName ErrorDataReceived -Action $errorHandler
    $exitedEvent = Register-ObjectEvent -InputObject $process -EventName Exited -Action $processExitedHandler
    
    # Start the process
    [void]$process.Start()
    $process.BeginOutputReadLine()
    $process.BeginErrorReadLine()
    
    # Monitor progress
    $startTime = Get-Date
    $spinChars = '|', '/', '-', '\'
    $spinIndex = 0
    
    while (-not $copyComplete) {
        $elapsedTime = (Get-Date) - $startTime
        $elapsedSeconds = [math]::Max(1, $elapsedTime.TotalSeconds)
        
        # Calculate progress based on file count
        if ($filesCopied -gt $lastReportedFiles) {
            $lastReportedFiles = $filesCopied
        }
        
        $percentComplete = [math]::Min(100, [math]::Max(0, [math]::Round(($filesCopied / $totalFilesInFolder) * 100, 0)))
        
        # Create a spinner for activity indication even when no progress is visible
        $spinChar = $spinChars[$spinIndex % $spinChars.Length]
        $spinIndex++
        
        # Format status line - include spinner to show activity
        $status = "$spinChar $filesCopied of $totalFilesInFolder files"
        
        # Add transfer rate if available
        if ($bytesTransferred -gt 0) {
            $rate = [math]::Round($bytesTransferred / $elapsedSeconds / 1MB, 2)
            $status += " | $rate MB/s"
        }
        
        # Calculate ETA if we have progress
        if ($percentComplete -gt 0) {
            $totalSecondsRemaining = ($elapsedSeconds / $percentComplete) * (100 - $percentComplete)
            $timeRemaining = [timespan]::FromSeconds($totalSecondsRemaining)
            $eta = "{0:hh\:mm\:ss}" -f $timeRemaining
            $status += " | ETA: $eta"
        }
        
        # Create a progress bar
        $progressBarWidth = 50
        $filledWidth = [math]::Round(($percentComplete / 100) * $progressBarWidth)
        $progressBar = "[" + ("#" * $filledWidth) + (" " * ($progressBarWidth - $filledWidth)) + "]"
        
        Write-Host "`r$progressBar $percentComplete% $status     " -NoNewline
        
        Start-Sleep -Milliseconds 200
        
        # Check if process has exited
        if ($process.HasExited) {
            $copyComplete = $true
        }
    }
    
    # Complete the progress line
    Write-Host "`r" + (" " * 120) + "`r" -NoNewline
    
    # Cleanup events
    Unregister-Event -SourceIdentifier $outputEvent.Name
    Unregister-Event -SourceIdentifier $errorEvent.Name
    Unregister-Event -SourceIdentifier $exitedEvent.Name
    
    # Get the final output and cleanup
    $output = $outputBuilder.ToString()
    $errorOutput = $errorBuilder.ToString()
    
    # Wait for process to fully exit
    $process.WaitForExit()
    
    # Get exit code
    $robocopyExitCode = $process.ExitCode
    
    # Cleanup
    $process.Close()
    
    Write-Host "Copy operation completed for $folder" -ForegroundColor Green
    
    # Try to extract statistics from output
    try {
        if ($output -match "Files\s*:\s*(\d+)") {
            $totalFilesCopied += [int]$Matches[1]
        }
        
        if ($output -match "Bytes\s*:\s*(\d+)") {
            $totalBytesCopied += [int64]$Matches[1]
        }
    }
    catch {
        Write-Host "Error parsing robocopy output: $_" -ForegroundColor Yellow
    }
}
else {
    # Fallback to standard robocopy if we couldn't calculate file count
    $robocopyProcess = Start-Process -FilePath "robocopy.exe" -ArgumentList @(
        "`"$sourcePath`"",
        "`"$destinationPath`"",
        "/MIR",
        "/MT:$script:threads",
        "/R:5",
        "/W:5",
        "/XF", "desktop.ini", "ntuser.dat*", "NTUSER.DAT*", "*.tmp",
        "/XD", "AppData\Local\Temp", "AppData\LocalLow",
        "/XJ",
        "/ZB",
        "/NP",
        "/BYTES",
        "/DCOPY:T",
        "/COPY:DAT",
        "/LOG+:`"$logFile`"",
        "/TEE"
    ) -NoNewWindow -Wait -PassThru
    
    $robocopyExitCode = $robocopyProcess.ExitCode
}
        
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
}

#endregion FUNCTIONS

#region MAIN SCRIPT

# Default folders that can be backed up
$script:defaultFolders = @(
    "Desktop",
    "Documents",
    "Downloads",
    "Pictures",
    "Videos",
    "Music",
    "Favorites",
    "OneDrive"
)

# Default number of threads
$script:threads = 16

# Display script header
Clear-Host
Write-Host "======================================================" -ForegroundColor Cyan
Write-Host "               UNIFIED ROBOCOPY TOOL                  " -ForegroundColor Cyan
Write-Host "======================================================" -ForegroundColor Cyan
Write-Host "This tool helps you back up and restore user profile folders" -ForegroundColor Yellow
Write-Host "using RoboCopy with optimized settings, progress monitoring," -ForegroundColor Yellow
Write-Host "and interactive folder selection." -ForegroundColor Yellow
Write-Host "======================================================" -ForegroundColor Cyan

# Ask for operation mode
Write-Host "`nSelect operation mode:" -ForegroundColor Yellow
Write-Host "1. Back up user profile folders to external drive"
Write-Host "2. Restore user profile folders from backup"
Write-Host "3. Exit"

$operationMode = 0
do {
    $operationInput = Read-Host "`nEnter your choice (1-3)"
    if ($operationInput -match "^\d+$" -and [int]$operationInput -ge 1 -and [int]$operationInput -le 3) {
        $operationMode = [int]$operationInput
        break
    } else {
        Write-Host "Invalid choice. Please enter a number between 1 and 3." -ForegroundColor Red
    }
} while ($true)

# Ask for the number of threads to use
$script:threads = 16  # Default value
do {
    $threadsInput = Read-Host "`nEnter number of threads to use for copying (default: $script:threads, press Enter for default)"
    if ($threadsInput -eq "") {
        # Keep default
        break
    }
    elseif ($threadsInput -match "^\d+$" -and [int]$threadsInput -gt 0 -and [int]$threadsInput -le 128) {
        $script:threads = [int]$threadsInput
        break
    }
    else {
        Write-Host "Please enter a valid number between 1 and 128." -ForegroundColor Red
    }
} while ($true)

# Execute the selected operation
switch ($operationMode) {
    1 {
        Write-Host "`n*** BACKUP MODE ***" -ForegroundColor Green
        Invoke-BackupOperation
    }
    2 {
        Write-Host "`n*** RESTORE MODE ***" -ForegroundColor Green
        Invoke-RestoreOperation
    }
    3 {
        Write-Host "`nExiting script." -ForegroundColor Yellow
        exit 0
    }
}

Write-Host "`nThank you for using the Unified RoboCopy Tool." -ForegroundColor Cyan
Write-Host "Script completed at $(Get-Date)" -ForegroundColor Cyan
#endregion MAIN SCRIPT
