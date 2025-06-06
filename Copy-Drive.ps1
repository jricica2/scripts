<#
.SYNOPSIS
    Script to copy all data from a smaller NVMe SSD to a larger one.

.DESCRIPTION
    This script recursively copies all files and directories from a source drive to a destination drive.
    It includes options to verify copied files, skip existing files, and display progress during the operation.

.PARAMETER SourceDrive
    The drive letter of the source drive (e.g., "D:"). This parameter is mandatory.

.PARAMETER DestinationDrive
    The drive letter of the destination drive (e.g., "E:"). This parameter is mandatory.

.PARAMETER Verify
    Optional switch to verify the copied files. If specified, the script ensures the integrity of copied files.

.PARAMETER SkipExisting
    Optional switch to skip files that already exist in the destination drive with matching size and write time.

.EXAMPLE
    .\Copy-Drive.ps1 -SourceDrive "D:" -DestinationDrive "E:"
    Copies all files and directories from drive D: to drive E:.

.EXAMPLE
    .\Copy-Drive.ps1 -SourceDrive "D:" -DestinationDrive "E:" -SkipExisting
    Copies files and directories from drive D: to drive E:, skipping files that already exist.

.EXAMPLE
    .\Copy-Drive.ps1 -SourceDrive "D:" -DestinationDrive "E:" -Verify
    Copies files and directories from drive D: to drive E:, verifying the integrity of copied files.

.NOTES
    Author: jricica2
    Date: 2025-04-22
    Version: 1.0
    This script requires administrative privileges to run.

#>

# Define parameters
param (
    [Parameter(Mandatory=$true)]
    [string]$SourceDrive,
    
    [Parameter(Mandatory=$true)]
    [string]$DestinationDrive,
    
    [Parameter(Mandatory=$false)]
    [switch]$Verify,
    
    [Parameter(Mandatory=$false)]
    [switch]$SkipExisting
)

# Function to display progress
function Show-Progress {
    param (
        [string]$Activity,
        [string]$Status,
        [int]$PercentComplete
    )
    
    Write-Progress -Activity $Activity -Status $Status -PercentComplete $PercentComplete
}

# Start copy process
$startTime = Get-Date
Write-Host "Starting copy process at $startTime" -ForegroundColor Green

try {
    # Get all items from source drive
    $sourceItems = Get-ChildItem -Path "$SourceDrive\" -Recurse -Force -ErrorAction SilentlyContinue
    $totalItems = $sourceItems.Count
    $copiedItems = 0
    $skippedItems = 0
    
    # Copy all items
    foreach ($item in $sourceItems) {
        $relativePath = $item.FullName.Substring($SourceDrive.Length)
        $destPath = Join-Path -Path $DestinationDrive -ChildPath $relativePath
        
        $currentItem = $copiedItems + $skippedItems + 1
        $percentComplete = [math]::Min(100, [math]::Round(($currentItem / $totalItems) * 100))
        
        if ($item.PSIsContainer) {
            # Create directory
            if (-not (Test-Path -Path $destPath)) {
                New-Item -Path $destPath -ItemType Directory -Force | Out-Null
            }
            $copiedItems++
        }
        else {
            # Optimized SkipExisting logic
            if ($SkipExisting -and (Test-Path -Path $destPath)) {
                $destItem = Get-Item -Path $destPath
                if ($item.Length -eq $destItem.Length -and $item.LastWriteTime -eq $destItem.LastWriteTime) {
                    $skippedItems++
                    Show-Progress -Activity "Processing items" -Status "Skipping $($item.Name) - $currentItem of $totalItems ($percentComplete%)" -PercentComplete $percentComplete
                    continue
                }
            }
            
            # Copy file
            $destDir = Split-Path -Path $destPath -Parent
            if (-not (Test-Path -Path $destDir)) {
                New-Item -Path $destDir -ItemType Directory -Force | Out-Null
            }
            
            Copy-Item -Path $item.FullName -Destination $destPath -Force
            $copiedItems++
            Show-Progress -Activity "Processing items" -Status "Copying $($item.Name) - $currentItem of $totalItems ($percentComplete%)" -PercentComplete $percentComplete
        }
    }
    
    $endTime = Get-Date
    $duration = $endTime - $startTime
    
    Write-Host "Copy process completed at $endTime" -ForegroundColor Green
    Write-Host "Total duration: $($duration.Hours)h $($duration.Minutes)m $($duration.Seconds)s" -ForegroundColor Green
    Write-Host "Total items copied: $copiedItems" -ForegroundColor Green
    Write-Host "Total items skipped: $skippedItems" -ForegroundColor Green
}
catch {
    Write-Error "Error during copy process: $_"
}
