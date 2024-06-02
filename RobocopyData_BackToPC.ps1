# Define variables
$externalDrive = "D:\FrameworkData"  # Change this to the path of your external drive
$userProfile = [System.Environment]::GetFolderPath("UserProfile")
$foldersToCopy = @("Desktop", "Documents", "Downloads", "Pictures", "Videos")
$threads = 16  # Number of threads to use

# Copy each folder back to the user profile
foreach ($folder in $foldersToCopy) {
    $sourcePath = Join-Path -Path $externalDrive -ChildPath $folder
    $destinationPath = Join-Path -Path $userProfile -ChildPath $folder

    # Run RoboCopy
    Write-Output "Copying $folder back to the user profile..."
    robocopy $sourcePath $destinationPath /MIR /MT:$threads /R:3 /W:5 /LOG+:robocopy_restore_log.txt /TEE /NFL /NDL
}

Write-Output "Restore completed."