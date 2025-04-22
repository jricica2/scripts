# Script to copy all data from a smaller NVMe SSD to a larger one
# Usage Context:
# .\Copy-Drive.ps1 -SourceDrive "D:" -DestinationDrive "E:" -SkipExisting OR -SkipExisting:$true

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
            # Check if file exists and should be skipped
            $shouldCopy = $true
            
            if ($SkipExisting -and (Test-Path -Path $destPath)) {
                $destItem = Get-Item -Path $destPath
                
                # Skip if file exists and has same length and write time
                if ($item.Length -eq $destItem.Length -and $item.LastWriteTime -eq $destItem.LastWriteTime) {
                    $shouldCopy = $false
                    $skippedItems++
                    Show-Progress -Activity "Processing items" -Status "Skipping $($item.Name) - $currentItem of $totalItems ($percentComplete%)" -PercentComplete $percentComplete
                }
            }
            
            if ($shouldCopy) {
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
