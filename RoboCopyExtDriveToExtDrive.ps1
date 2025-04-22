# Script to copy all data from a smaller NVMe SSD to a larger one
# Context to run:
# .\Copy-NVMeDrive.ps1 -SourceDrive "D:" -DestinationDrive "E:" -Verify

# Define parameters
param (
    [Parameter(Mandatory=$true)]
    [string]$SourceDrive,
    
    [Parameter(Mandatory=$true)]
    [string]$DestinationDrive,
    
    [Parameter(Mandatory=$false)]
    [switch]$Verify = $false
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

# Check if running as administrator
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Error "This script requires administrator privileges. Please run PowerShell as Administrator."
    exit 1
}

# Validate drives
try {
    $sourceDriveInfo = Get-Volume -DriveLetter $SourceDrive.TrimEnd(':')
    $destDriveInfo = Get-Volume -DriveLetter $DestinationDrive.TrimEnd(':')
    
    Write-Host "Source Drive: $SourceDrive - $($sourceDriveInfo.SizeRemaining / 1GB) GB free of $($sourceDriveInfo.Size / 1GB) GB" -ForegroundColor Cyan
    Write-Host "Destination Drive: $DestinationDrive - $($destDriveInfo.SizeRemaining / 1GB) GB free of $($destDriveInfo.Size / 1GB) GB" -ForegroundColor Cyan
    
    # Check if destination drive is larger than source drive
    if ($destDriveInfo.Size -lt $sourceDriveInfo.Size) {
        Write-Error "Destination drive is smaller than source drive. Please select a larger destination drive."
        exit 1
    }
    
    # Check if destination has enough free space
    if ($destDriveInfo.SizeRemaining -lt $sourceDriveInfo.Size) {
        Write-Error "Destination drive does not have enough free space to copy all data from the source drive."
        exit 1
    }
}
catch {
    Write-Error "Error validating drives: $_"
    exit 1
}

# Confirm operation
Write-Host "`nThis will copy all data from $SourceDrive to $DestinationDrive." -ForegroundColor Yellow
$confirmation = Read-Host "Are you sure you want to proceed? (y/n)"
if ($confirmation -ne 'y') {
    Write-Host "Operation cancelled by user." -ForegroundColor Red
    exit 0
}

# Start copy process
$startTime = Get-Date
Write-Host "`nStarting copy process at $startTime" -ForegroundColor Green

try {
    # Get all items from source drive
    $sourceItems = Get-ChildItem -Path "$SourceDrive\" -Recurse -Force -ErrorAction SilentlyContinue
    $totalItems = $sourceItems.Count
    $copiedItems = 0
    
    # Create destination root if it doesn't exist
    if (-not (Test-Path -Path "$DestinationDrive\")) {
        New-Item -Path "$DestinationDrive\" -ItemType Directory -Force | Out-Null
    }
    
    # Copy all items
    foreach ($item in $sourceItems) {
        $relativePath = $item.FullName.Substring($SourceDrive.Length)
        $destPath = Join-Path -Path $DestinationDrive -ChildPath $relativePath
        
        $copiedItems++
        $percentComplete = [math]::Min(100, [math]::Round(($copiedItems / $totalItems) * 100))
        Show-Progress -Activity "Copying files" -Status "$copiedItems of $totalItems items ($percentComplete%)" -PercentComplete $percentComplete
        
        try {
            if ($item.PSIsContainer) {
                # Create directory
                if (-not (Test-Path -Path $destPath)) {
                    New-Item -Path $destPath -ItemType Directory -Force | Out-Null
                }
            }
            else {
                # Copy file
                $destDir = Split-Path -Path $destPath -Parent
                if (-not (Test-Path -Path $destDir)) {
                    New-Item -Path $destDir -ItemType Directory -Force | Out-Null
                }
                
                Copy-Item -Path $item.FullName -Destination $destPath -Force
            }
        }
        catch {
            Write-Warning "Error copying $($item.FullName): $_"
        }
    }
    
    # Verify copy if requested
    if ($Verify) {
        Write-Host "`nVerifying copied files..." -ForegroundColor Yellow
        $verifiedItems = 0
        $failedItems = 0
        
        foreach ($item in $sourceItems) {
            if (-not $item.PSIsContainer) {
                $relativePath = $item.FullName.Substring($SourceDrive.Length)
                $destPath = Join-Path -Path $DestinationDrive -ChildPath $relativePath
                
                $verifiedItems++
                $percentComplete = [math]::Min(100, [math]::Round(($verifiedItems / $totalItems) * 100))
                Show-Progress -Activity "Verifying files" -Status "$verifiedItems of $totalItems items ($percentComplete%)" -PercentComplete $percentComplete
                
                if (Test-Path -Path $destPath) {
                    $sourceHash = Get-FileHash -Path $item.FullName -Algorithm SHA256
                    $destHash = Get-FileHash -Path $destPath -Algorithm SHA256
                    
                    if ($sourceHash.Hash -ne $destHash.Hash) {
                        Write-Warning "Verification failed for $($item.FullName)"
                        $failedItems++
                    }
                }
                else {
                    Write-Warning "File not found at destination: $destPath"
                    $failedItems++
                }
            }
        }
        
        if ($failedItems -gt 0) {
            Write-Host "`nVerification completed with $failedItems failures." -ForegroundColor Red
        }
        else {
            Write-Host "`nVerification completed successfully. All files match." -ForegroundColor Green
        }
    }
    
    $endTime = Get-Date
    $duration = $endTime - $startTime
    
    Write-Host "`nCopy process completed at $endTime" -ForegroundColor Green
    Write-Host "Total duration: $($duration.Hours) hours, $($duration.Minutes) minutes, $($duration.Seconds) seconds" -ForegroundColor Green
    Write-Host "Total items copied: $copiedItems" -ForegroundColor Green
}
catch {
    Write-Error "Error during copy process: $_"
    exit 1
}
