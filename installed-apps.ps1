# Get all installed applications from both 32-bit and 64-bit paths, and for the current user
$installedApps = Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*, HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*, HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName, DisplayVersion, Publisher, InstallDate

# Remove entries without a display name
$installedApps = $installedApps | Where-Object {$_.DisplayName -ne $null}

# Sort the applications in alphabetical order
$sortedApps = $installedApps | Sort-Object DisplayName

# Export the sorted list to a .txt file
$sortedApps | Out-File -FilePath "C:\Users\Jeff\Desktop\installed_apps.txt"
