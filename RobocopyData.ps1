# Define variables
$externalDrive = "D:\FrameworkData"  # Change this to the path of your external drive
$userProfile = [System.Environment]::GetFolderPath("UserProfile")
$foldersToCopy = @("Desktop", "Documents", "Downloads", "Pictures", "Videos")
$threads = 16  # Number of threads to use

# Create backup directory if it doesn't exist
if (!(Test-Path -Path $externalDrive)) {
    New-Item -Path $externalDrive -ItemType Directory
}

# Copy each folder
foreach ($folder in $foldersToCopy) {
    $sourcePath = Join-Path -Path $userProfile -ChildPath $folder
    $destinationPath = Join-Path -Path $externalDrive -ChildPath $folder

    # Run RoboCopy
    Write-Output "Copying $folder..."
    robocopy $sourcePath $destinationPath /MIR /MT:$threads /R:3 /W:5 /LOG+:robocopy_log.txt /TEE /NFL /NDL
}

Write-Output "Backup completed."
