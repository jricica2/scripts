# Get the current user's Downloads folder path
$downloadsPath = [Environment]::GetFolderPath("UserProfile") + "\Downloads"

# Create the drivers-backup folder
$backupPath = Join-Path -Path $downloadsPath -ChildPath "drivers-backup"
New-Item -ItemType Directory -Force -Path $backupPath

# Run DISM command to export drivers
$dismCommand = "dism /online /export-driver /destination:$backupPath"

# Execute the DISM command
try {
    Invoke-Expression -Command $dismCommand
    Write-Host "Drivers have been successfully backed up to: $backupPath"
} catch {
    Write-Host "An error occurred while backing up the drivers: $_"
}
