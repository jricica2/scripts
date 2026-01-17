# Enhanced-DataMigration.ps1
# Purpose: Comprehensive data migration tool for PC re-imaging and migration
<#
.SYNOPSIS
    Interactive unified script for complete PC data migration using RoboCopy.

.DESCRIPTION
    This script provides a full-featured interactive interface for migrating user data
    between computers or during PC re-imaging. Features include:

    - Interactive TUI with arrow-key navigation
    - User profile folder backup/restore (Desktop, Documents, etc.)
    - Application configuration backup (browsers, VS Code, Notepad++)
    - Browser bookmarks and extension list export
    - Desktop wallpaper and taskbar pins backup
    - Installed application inventory with Winget/Chocolatey matching
    - Auto-generated reinstallation script
    - Real-time progress indicators
    - Detailed error/skip logging
    - Dry run mode for testing

.NOTES
    - Designed for Windows 10/11 migrations
    - RoboCopy is built into Windows - no additional installation required
    - Winget is used as primary package manager, Chocolatey as fallback
    - Some features require Administrator privileges

.AUTHOR
    Created by Claude Code - January 2026
#>

#Requires -Version 5.1

#region CONFIGURATION

# Script version
$script:Version = "1.0.0"

# Default folders to backup (pre-selected)
# These use Environment.SpecialFolder enum values where available
$script:DefaultFolders = @(
    @{ Name = "Desktop"; Selected = $true; SpecialFolder = "Desktop" },
    @{ Name = "Documents"; Selected = $true; SpecialFolder = "MyDocuments" },
    @{ Name = "Downloads"; Selected = $true; KnownFolder = "Downloads" },
    @{ Name = "Pictures"; Selected = $true; SpecialFolder = "MyPictures" },
    @{ Name = "Videos"; Selected = $true; SpecialFolder = "MyVideos" },
    @{ Name = "Music"; Selected = $false; SpecialFolder = "MyMusic" },
    @{ Name = "Favorites"; Selected = $false; SpecialFolder = "Favorites" }
)

# Application config whitelist - known-safe configurations that can be restored
# Format: Name, Source path (relative to user profile), files/folders to copy
$script:AppConfigWhitelist = @(
    # Browsers - Bookmarks and Extension lists only
    @{
        Name = "Google Chrome"
        Enabled = $true
        Source = "AppData\Local\Google\Chrome\User Data\Default"
        Items = @("Bookmarks", "Bookmarks.bak", "Extensions")
        ExtensionManifest = $true
    },
    @{
        Name = "Mozilla Firefox"
        Enabled = $true
        Source = "AppData\Roaming\Mozilla\Firefox\Profiles"
        Items = @("places.sqlite", "extensions.json", "addons.json")
        IsProfileFolder = $true
    },
    @{
        Name = "Microsoft Edge"
        Enabled = $true
        Source = "AppData\Local\Microsoft\Edge\User Data\Default"
        Items = @("Bookmarks", "Bookmarks.bak", "Extensions")
        ExtensionManifest = $true
    },
    # Code Editors
    @{
        Name = "VS Code"
        Enabled = $true
        Source = "AppData\Roaming\Code\User"
        Items = @("settings.json", "keybindings.json", "snippets")
        ExtensionList = $true
        ExtensionCmd = "code --list-extensions"
    },
    @{
        Name = "Notepad++"
        Enabled = $true
        Source = "AppData\Roaming\Notepad++"
        Items = @("config.xml", "session.xml", "shortcuts.xml", "stylers.xml", "langs.xml", "userDefineLangs")
    }
)

# RoboCopy settings
$script:Threads = 16
$script:RetryCount = 3
$script:RetryWait = 5

# Files and directories to always exclude
$script:ExcludeFiles = @(
    "desktop.ini",
    "thumbs.db",
    "*.tmp",
    "ntuser.dat*",
    "NTUSER.DAT*"
)

$script:ExcludeDirs = @(
    "AppData\Local\Temp",
    "AppData\LocalLow",
    "AppData\Local\Microsoft\Windows\INetCache",
    "AppData\Local\Microsoft\Windows\Temporary Internet Files"
)

#endregion CONFIGURATION

#region TUI FUNCTIONS

function Show-InteractiveMenu {
    <#
    .SYNOPSIS
        Displays an interactive menu with arrow key navigation
    .PARAMETER Title
        The menu title
    .PARAMETER Items
        Array of menu items (strings or hashtables with Name and Selected properties)
    .PARAMETER MultiSelect
        Allow multiple selections with spacebar
    .PARAMETER ShowSize
        Show size information for items (requires SizePath property)
    .PARAMETER AllowBack
        Allow user to go back to previous step (Backspace key)
    #>
    param(
        [string]$Title,
        [array]$Items,
        [bool]$MultiSelect = $false,
        [bool]$ShowSize = $false,
        [string]$BasePath = "",
        [bool]$AllowBack = $false
    )

    if ($Items.Count -eq 0) {
        Write-Host "No items available." -ForegroundColor Red
        return @()
    }

    # Convert items to menu format - copy ALL properties to avoid losing data
    $menuItems = @()
    foreach ($item in $Items) {
        if ($item -is [string]) {
            $menuItems += @{ Name = $item; Selected = $false; SizeStr = "" }
        } else {
            # Copy all properties from the original hashtable
            $newItem = @{
                SizeStr = ""  # Will be calculated once below if ShowSize is true
            }
            # Copy every property from the source item
            foreach ($key in $item.Keys) {
                $newItem[$key] = $item[$key]
            }
            # Ensure required properties exist with defaults
            if (-not $newItem.Name) { $newItem.Name = "Unknown" }
            if (-not $newItem.ContainsKey('Selected')) { $newItem.Selected = $false }
            $menuItems += $newItem
        }
    }

    # PRE-CALCULATE sizes once (not on every redraw!)
    if ($ShowSize) {
        Write-Host "Calculating folder sizes..." -ForegroundColor DarkGray
        $totalItems = $menuItems.Count
        $currentItem = 0
        foreach ($item in $menuItems) {
            $currentItem++
            if ($item.Path -and $item.Path.Length -gt 0) {
                $fullPath = if ([System.IO.Path]::IsPathRooted($item.Path)) {
                    $item.Path
                } elseif ($BasePath) {
                    Join-Path $BasePath $item.Path
                } else {
                    $null
                }

                if ($fullPath -and (Test-Path -LiteralPath $fullPath -ErrorAction SilentlyContinue)) {
                    Write-Host "  [$currentItem/$totalItems] $($item.Name)..." -NoNewline
                    try {
                        $size = (Get-ChildItem -LiteralPath $fullPath -Recurse -File -ErrorAction SilentlyContinue |
                                 Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
                        if ($size) {
                            $item.SizeStr = if ($size -gt 1GB) {
                                "{0:N2} GB" -f ($size / 1GB)
                            } elseif ($size -gt 1MB) {
                                "{0:N2} MB" -f ($size / 1MB)
                            } else {
                                "{0:N2} KB" -f ($size / 1KB)
                            }
                            Write-Host " $($item.SizeStr)" -ForegroundColor Green
                        } else {
                            Write-Host " (empty)" -ForegroundColor DarkGray
                        }
                    } catch {
                        Write-Host " (error)" -ForegroundColor Yellow
                    }
                }
            }
        }
        Write-Host ""
    }

    $currentIndex = 0
    $done = $false

    while (-not $done) {
        Clear-Host
        Write-Host "============================================" -ForegroundColor Cyan
        Write-Host "  $Title" -ForegroundColor Cyan
        Write-Host "============================================" -ForegroundColor Cyan
        Write-Host ""

        if ($MultiSelect) {
            $helpText = "  [UP/DOWN] Navigate  [SPACE] Toggle  [A] All  [N] None  [ENTER] Confirm"
        } else {
            $helpText = "  [UP/DOWN] Navigate  [ENTER] Select"
        }
        if ($AllowBack) {
            $helpText += "  [BACKSPACE] Back"
        }
        Write-Host $helpText -ForegroundColor DarkGray
        Write-Host ""

        for ($i = 0; $i -lt $menuItems.Count; $i++) {
            $item = $menuItems[$i]
            $prefix = if ($i -eq $currentIndex) { " >> " } else { "    " }

            # Get display name with fallback
            $displayName = if ($item.Name -and $item.Name.Length -gt 0) {
                $item.Name
            } else {
                "Item $($i + 1)"
            }

            if ($MultiSelect) {
                $checkbox = if ($item.Selected) { "[X]" } else { "[ ]" }
                $line = "$prefix$checkbox $displayName"
            } else {
                $line = "$prefix$displayName"
            }

            # Add pre-calculated size if available
            if ($ShowSize -and $item.SizeStr) {
                $line += " ($($item.SizeStr))"
            }

            # Always output the line
            if ($i -eq $currentIndex) {
                Write-Host $line -ForegroundColor Yellow
            } else {
                Write-Host $line
            }
        }

        Write-Host ""

        # Read key input
        $key = [Console]::ReadKey($true)

        switch ($key.Key) {
            "UpArrow" {
                $currentIndex = if ($currentIndex -gt 0) { $currentIndex - 1 } else { $menuItems.Count - 1 }
            }
            "DownArrow" {
                $currentIndex = if ($currentIndex -lt $menuItems.Count - 1) { $currentIndex + 1 } else { 0 }
            }
            "Spacebar" {
                if ($MultiSelect) {
                    $menuItems[$currentIndex].Selected = -not $menuItems[$currentIndex].Selected
                }
            }
            "Enter" {
                $done = $true
            }
            "A" {
                if ($MultiSelect) {
                    foreach ($item in $menuItems) { $item.Selected = $true }
                }
            }
            "N" {
                if ($MultiSelect) {
                    foreach ($item in $menuItems) { $item.Selected = $false }
                }
            }
            "Escape" {
                return $null
            }
            "Backspace" {
                if ($AllowBack) {
                    # Return special marker to indicate "go back"
                    return @{ _GoBack = $true }
                }
            }
            "B" {
                if ($AllowBack) {
                    return @{ _GoBack = $true }
                }
            }
        }
    }

    if ($MultiSelect) {
        return $menuItems | Where-Object { $_.Selected }
    } else {
        return $menuItems[$currentIndex]
    }
}

function Test-GoBack {
    <#
    .SYNOPSIS
        Checks if menu result indicates user wants to go back
    #>
    param($Result)

    if ($null -eq $Result) { return $false }
    if ($Result -is [hashtable] -and $Result._GoBack -eq $true) { return $true }
    if ($Result -is [array] -and $Result.Count -eq 1 -and $Result[0]._GoBack -eq $true) { return $true }
    return $false
}

function Show-Confirmation {
    <#
    .SYNOPSIS
        Shows a yes/no confirmation prompt
    #>
    param(
        [string]$Message,
        [bool]$DefaultYes = $false
    )

    $options = if ($DefaultYes) { "(Y/n)" } else { "(y/N)" }
    Write-Host ""
    Write-Host "$Message $options " -ForegroundColor Yellow -NoNewline

    $response = Read-Host

    if ($DefaultYes) {
        return $response -ne "n" -and $response -ne "N"
    } else {
        return $response -eq "y" -or $response -eq "Y"
    }
}

function Show-ProgressBar {
    <#
    .SYNOPSIS
        Displays a progress bar on the console
    #>
    param(
        [int]$Percent,
        [string]$Status = "",
        [int]$Width = 50
    )

    $filled = [math]::Round(($Percent / 100) * $Width)
    $empty = $Width - $filled
    $bar = "[" + ("#" * $filled) + ("-" * $empty) + "]"

    Write-Host "`r$bar $Percent% $Status     " -NoNewline
}

function Write-Header {
    param([string]$Text)

    Clear-Host
    Write-Host ""
    Write-Host "============================================" -ForegroundColor Cyan
    Write-Host "  $Text" -ForegroundColor Cyan
    Write-Host "============================================" -ForegroundColor Cyan
    Write-Host ""
}

function Write-Section {
    param([string]$Text)

    Write-Host ""
    Write-Host "--- $Text ---" -ForegroundColor Yellow
    Write-Host ""
}

function Write-Success {
    param([string]$Text)
    Write-Host "[OK] $Text" -ForegroundColor Green
}

function Write-Warning {
    param([string]$Text)
    Write-Host "[WARNING] $Text" -ForegroundColor Yellow
}

function Write-Error {
    param([string]$Text)
    Write-Host "[ERROR] $Text" -ForegroundColor Red
}

function Write-Info {
    param([string]$Text)
    Write-Host "[INFO] $Text" -ForegroundColor Cyan
}

#endregion TUI FUNCTIONS

#region UTILITY FUNCTIONS

function Get-SpecialFolderPath {
    <#
    .SYNOPSIS
        Gets the actual path of a Windows special folder, handling OneDrive redirection
    #>
    param(
        [string]$FolderName,
        [string]$SpecialFolder,
        [string]$KnownFolder
    )

    $path = $null

    # Try Environment.SpecialFolder first
    if ($SpecialFolder) {
        try {
            $path = [Environment]::GetFolderPath($SpecialFolder)
            if ($path -and (Test-Path $path)) {
                return $path
            }
        } catch { }
    }

    # Try Known Folder API for folders not in SpecialFolder enum (like Downloads)
    if ($KnownFolder) {
        try {
            # Use Shell COM object to get known folder path
            $shell = New-Object -ComObject Shell.Application

            # Known folder GUIDs
            $knownFolderGuids = @{
                "Downloads" = "{374DE290-123F-4565-9164-39C4925E467B}"
            }

            if ($knownFolderGuids.ContainsKey($KnownFolder)) {
                $folder = $shell.Namespace($knownFolderGuids[$KnownFolder])
                if ($folder) {
                    $path = $folder.Self.Path
                    if ($path -and (Test-Path $path)) {
                        return $path
                    }
                }
            }
        } catch { }

        # Fallback: Try registry for Downloads
        if ($KnownFolder -eq "Downloads") {
            try {
                $regPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
                $path = (Get-ItemProperty -Path $regPath -Name "{374DE290-123F-4565-9164-39C4925E467B}" -ErrorAction SilentlyContinue)."{374DE290-123F-4565-9164-39C4925E467B}"
                if ($path -and (Test-Path $path)) {
                    return $path
                }
            } catch { }

            # Last resort: User profile + Downloads
            $path = Join-Path $env:USERPROFILE "Downloads"
            if (Test-Path $path) {
                return $path
            }
        }
    }

    # Final fallback: try user profile + folder name
    $fallbackPath = Join-Path $env:USERPROFILE $FolderName
    if (Test-Path $fallbackPath) {
        return $fallbackPath
    }

    return $null
}

function Get-UserFoldersForBackup {
    <#
    .SYNOPSIS
        Gets all available user folders with their actual paths
    #>
    $folders = [System.Collections.ArrayList]@()

    foreach ($folderDef in $script:DefaultFolders) {
        $folderName = $folderDef.Name
        $actualPath = $null

        # Try to get the actual path
        $actualPath = Get-SpecialFolderPath -FolderName $folderDef.Name -SpecialFolder $folderDef.SpecialFolder -KnownFolder $folderDef.KnownFolder

        # If shell folder detection failed, try direct path in user profile
        if (-not $actualPath) {
            $directPath = Join-Path $env:USERPROFILE $folderName
            if (Test-Path $directPath) {
                $actualPath = $directPath
            }
        }

        # If still not found, check common OneDrive locations
        if (-not $actualPath) {
            $oneDrivePaths = @(
                (Join-Path $env:OneDrive $folderName -ErrorAction SilentlyContinue),
                (Join-Path $env:OneDriveCommercial $folderName -ErrorAction SilentlyContinue),
                (Join-Path $env:OneDriveConsumer $folderName -ErrorAction SilentlyContinue)
            ) | Where-Object { $_ -and (Test-Path $_ -ErrorAction SilentlyContinue) }

            if ($oneDrivePaths.Count -gt 0) {
                $actualPath = $oneDrivePaths[0]
            }
        }

        if ($actualPath -and (Test-Path $actualPath -ErrorAction SilentlyContinue)) {
            $folderEntry = @{
                Name = [string]$folderName
                Selected = [bool]$folderDef.Selected
                Path = [string]$actualPath
                RelativeName = [string]$folderName
                Exists = $true
            }
            [void]$folders.Add($folderEntry)
        }
    }

    return $folders
}

function Get-RemovableDrives {
    <#
    .SYNOPSIS
        Gets a list of removable/external drives
    #>
    $drives = @()

    # Get all fixed and removable drives with free space
    Get-WmiObject Win32_LogicalDisk | Where-Object {
        $_.DriveType -in @(2, 3) -and  # 2=Removable, 3=Fixed
        $_.FreeSpace -gt 0 -and
        $_.DeviceID -ne $env:SystemDrive  # Exclude system drive
    } | ForEach-Object {
        $driveType = switch ($_.DriveType) {
            2 { "Removable" }
            3 { "Local Disk" }
            default { "Unknown" }
        }

        $drives += @{
            Name = "$($_.DeviceID) - $($_.VolumeName) ($driveType)"
            Letter = $_.DeviceID
            FreeSpace = $_.FreeSpace
            FreeSpaceGB = [math]::Round($_.FreeSpace / 1GB, 2)
            TotalSpace = $_.Size
            DriveType = $driveType
            Path = $_.DeviceID
        }
    }

    return $drives
}

function Get-BackupFolders {
    <#
    .SYNOPSIS
        Scans a location for existing backup folders
    #>
    param([string]$BasePath)

    $backups = @()

    if (-not (Test-Path $BasePath)) {
        return $backups
    }

    # Look for folders matching pattern HOSTNAME-DATE_TIME or containing manifest.json
    Get-ChildItem -Path $BasePath -Directory -ErrorAction SilentlyContinue | ForEach-Object {
        $manifestPath = Join-Path $_.FullName "manifest.json"
        if (Test-Path $manifestPath) {
            try {
                $manifest = Get-Content $manifestPath -Raw | ConvertFrom-Json
                $backups += @{
                    Name = "$($manifest.ComputerName) - $($manifest.BackupDate)"
                    Path = $_.FullName
                    Date = [DateTime]::Parse($manifest.BackupDate)
                    ComputerName = $manifest.ComputerName
                    WindowsVersion = $manifest.WindowsVersion
                }
            } catch {
                # Try to parse folder name
                if ($_.Name -match "^(.+)-(\d{4}-\d{2}-\d{2})_(\d{6})$") {
                    $backups += @{
                        Name = $_.Name
                        Path = $_.FullName
                        Date = [DateTime]::ParseExact("$($Matches[2]) $($Matches[3])", "yyyy-MM-dd HHmmss", $null)
                        ComputerName = $Matches[1]
                        WindowsVersion = "Unknown"
                    }
                }
            }
        }
        # Also check for backup_timestamp.txt (legacy format)
        elseif (Test-Path (Join-Path $_.FullName "backup_timestamp.txt")) {
            $timestamp = Get-Content (Join-Path $_.FullName "backup_timestamp.txt") -Raw
            $backups += @{
                Name = "$($_.Name) (Legacy)"
                Path = $_.FullName
                Date = if ($timestamp) { try { [DateTime]::Parse($timestamp.Trim()) } catch { Get-Date } } else { Get-Date }
                ComputerName = $_.Name
                WindowsVersion = "Unknown"
            }
        }
    }

    # Sort by date, most recent first, return top 3
    return $backups | Sort-Object -Property Date -Descending | Select-Object -First 3
}

function Get-FolderSize {
    <#
    .SYNOPSIS
        Calculates the total size of a folder
    #>
    param([string]$Path)

    if (-not $Path) {
        return 0
    }

    # Test if path exists using LiteralPath (handles special characters)
    if (-not (Test-Path -LiteralPath $Path -ErrorAction SilentlyContinue)) {
        return 0
    }

    try {
        $files = Get-ChildItem -LiteralPath $Path -Recurse -File -ErrorAction SilentlyContinue
        if ($files) {
            $size = ($files | Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
            if ($size) { return $size }
        }
        return 0
    } catch {
        return 0
    }
}

function Format-Size {
    <#
    .SYNOPSIS
        Formats a byte size into human-readable format
    #>
    param([int64]$Bytes)

    if ($Bytes -gt 1TB) { return "{0:N2} TB" -f ($Bytes / 1TB) }
    if ($Bytes -gt 1GB) { return "{0:N2} GB" -f ($Bytes / 1GB) }
    if ($Bytes -gt 1MB) { return "{0:N2} MB" -f ($Bytes / 1MB) }
    if ($Bytes -gt 1KB) { return "{0:N2} KB" -f ($Bytes / 1KB) }
    return "$Bytes bytes"
}

function Test-AdminPrivileges {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

#endregion UTILITY FUNCTIONS

#region APPLICATION INVENTORY FUNCTIONS

function Get-InstalledApplications {
    <#
    .SYNOPSIS
        Gets list of installed applications from registry (Programs and Features)
    #>

    $apps = @()

    # Registry paths for installed applications
    $regPaths = @(
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )

    foreach ($path in $regPaths) {
        try {
            Get-ItemProperty $path -ErrorAction SilentlyContinue |
            Where-Object { $_.DisplayName -and $_.DisplayName.Trim() } |
            ForEach-Object {
                $apps += @{
                    Name = $_.DisplayName.Trim()
                    Version = if ($_.DisplayVersion) { $_.DisplayVersion } else { "" }
                    Publisher = if ($_.Publisher) { $_.Publisher } else { "" }
                    InstallDate = if ($_.InstallDate) { $_.InstallDate } else { "" }
                    UninstallString = if ($_.UninstallString) { $_.UninstallString } else { "" }
                }
            }
        } catch { }
    }

    # Remove duplicates and sort
    $apps = $apps | Sort-Object -Property Name -Unique

    return $apps
}

function Test-WingetAvailable {
    <#
    .SYNOPSIS
        Checks if Winget is available on the system
    #>
    try {
        $null = Get-Command winget -ErrorAction Stop
        return $true
    } catch {
        return $false
    }
}

function Test-ChocolateyAvailable {
    <#
    .SYNOPSIS
        Checks if Chocolatey is available on the system
    #>
    try {
        $null = Get-Command choco -ErrorAction Stop
        return $true
    } catch {
        return $false
    }
}

function Install-Chocolatey {
    <#
    .SYNOPSIS
        Installs Chocolatey package manager
    #>
    Write-Info "Installing Chocolatey..."

    try {
        Set-ExecutionPolicy Bypass -Scope Process -Force
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
        Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
        return $true
    } catch {
        Write-Error "Failed to install Chocolatey: $_"
        return $false
    }
}

function Find-WingetPackage {
    <#
    .SYNOPSIS
        Searches for a package in Winget repository
    #>
    param([string]$AppName)

    try {
        # Clean up app name for search
        $searchName = $AppName -replace '\s*\([^)]*\)\s*$', ''  # Remove version in parentheses
        $searchName = $searchName -replace '\s+x64\s*$', ''      # Remove x64
        $searchName = $searchName -replace '\s+x86\s*$', ''      # Remove x86
        $searchName = $searchName.Trim()

        if ($searchName.Length -lt 3) { return $null }

        $result = winget search --name $searchName --accept-source-agreements 2>$null |
                  Select-String -Pattern "^\S+" |
                  Select-Object -First 5

        if ($result) {
            # Parse the first match that looks like a valid package ID
            foreach ($line in $result) {
                $parts = $line.Line -split '\s{2,}'
                if ($parts.Count -ge 2) {
                    $packageId = $parts[1].Trim()
                    if ($packageId -match '^\S+\.\S+') {
                        return @{
                            Name = $parts[0].Trim()
                            Id = $packageId
                            Source = "winget"
                        }
                    }
                }
            }
        }
        return $null
    } catch {
        return $null
    }
}

function Get-ApplicationInventoryWithPackages {
    <#
    .SYNOPSIS
        Gets installed applications and matches them to Winget/Chocolatey packages
    #>
    param(
        [switch]$SkipPackageMatching
    )

    Write-Info "Gathering installed applications..."
    $apps = Get-InstalledApplications

    if ($SkipPackageMatching) {
        return $apps
    }

    $hasWinget = Test-WingetAvailable
    $hasChoco = Test-ChocolateyAvailable

    if (-not $hasWinget -and -not $hasChoco) {
        Write-Warning "Neither Winget nor Chocolatey is available. Package matching skipped."
        return $apps
    }

    Write-Info "Matching applications to package managers (this may take a few minutes)..."

    $total = $apps.Count
    $current = 0

    foreach ($app in $apps) {
        $current++
        $percent = [math]::Round(($current / $total) * 100)
        Show-ProgressBar -Percent $percent -Status "Checking: $($app.Name.Substring(0, [Math]::Min(30, $app.Name.Length)))..."

        $app.WingetId = $null
        $app.ChocoId = $null

        if ($hasWinget) {
            $wingetResult = Find-WingetPackage -AppName $app.Name
            if ($wingetResult) {
                $app.WingetId = $wingetResult.Id
            }
        }
    }

    Write-Host ""  # Clear progress line
    return $apps
}

#endregion APPLICATION INVENTORY FUNCTIONS

#region PERSONALIZATION FUNCTIONS

function Get-CurrentWallpaper {
    <#
    .SYNOPSIS
        Gets the current desktop wallpaper path
    #>
    try {
        $wallpaperPath = (Get-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name Wallpaper).Wallpaper
        if ($wallpaperPath -and (Test-Path $wallpaperPath)) {
            return $wallpaperPath
        }
    } catch { }
    return $null
}

function Backup-Wallpaper {
    <#
    .SYNOPSIS
        Backs up the current desktop wallpaper
    #>
    param([string]$DestinationFolder)

    $wallpaper = Get-CurrentWallpaper
    if ($wallpaper) {
        $destPath = Join-Path $DestinationFolder "wallpaper$([System.IO.Path]::GetExtension($wallpaper))"
        try {
            Copy-Item -Path $wallpaper -Destination $destPath -Force
            return @{
                Success = $true
                SourcePath = $wallpaper
                DestPath = $destPath
            }
        } catch {
            return @{ Success = $false; Error = $_.Exception.Message }
        }
    }
    return @{ Success = $false; Error = "No wallpaper found" }
}

function Restore-Wallpaper {
    <#
    .SYNOPSIS
        Restores desktop wallpaper from backup
    #>
    param([string]$WallpaperPath)

    if (-not (Test-Path $WallpaperPath)) {
        return @{ Success = $false; Error = "Wallpaper file not found" }
    }

    try {
        # Copy to user's Pictures folder
        $destFolder = [Environment]::GetFolderPath("MyPictures")
        $destPath = Join-Path $destFolder "RestoredWallpaper$([System.IO.Path]::GetExtension($WallpaperPath))"
        Copy-Item -Path $WallpaperPath -Destination $destPath -Force

        # Set as wallpaper using SystemParametersInfo
        Add-Type @"
using System;
using System.Runtime.InteropServices;
public class Wallpaper {
    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);
}
"@
        [Wallpaper]::SystemParametersInfo(0x0014, 0, $destPath, 0x0003)

        return @{ Success = $true; Path = $destPath }
    } catch {
        return @{ Success = $false; Error = $_.Exception.Message }
    }
}

function Get-TaskbarPins {
    <#
    .SYNOPSIS
        Gets taskbar pinned items
    #>
    $taskbarPath = "$env:APPDATA\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar"

    if (-not (Test-Path $taskbarPath)) {
        return @()
    }

    $pins = @()
    Get-ChildItem -Path $taskbarPath -Filter "*.lnk" -ErrorAction SilentlyContinue | ForEach-Object {
        $pins += @{
            Name = $_.BaseName
            Path = $_.FullName
        }
    }

    return $pins
}

function Backup-TaskbarPins {
    <#
    .SYNOPSIS
        Backs up taskbar pinned shortcuts
    #>
    param([string]$DestinationFolder)

    $taskbarPath = "$env:APPDATA\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar"
    $destPath = Join-Path $DestinationFolder "TaskbarPins"

    if (-not (Test-Path $taskbarPath)) {
        return @{ Success = $false; Error = "Taskbar pins folder not found" }
    }

    try {
        if (-not (Test-Path $destPath)) {
            New-Item -Path $destPath -ItemType Directory -Force | Out-Null
        }

        $shortcuts = Get-ChildItem -Path $taskbarPath -Filter "*.lnk" -ErrorAction SilentlyContinue
        $copied = 0

        foreach ($shortcut in $shortcuts) {
            Copy-Item -Path $shortcut.FullName -Destination $destPath -Force
            $copied++
        }

        return @{ Success = $true; Count = $copied; Path = $destPath }
    } catch {
        return @{ Success = $false; Error = $_.Exception.Message }
    }
}

function Restore-TaskbarPins {
    <#
    .SYNOPSIS
        Restores taskbar pinned shortcuts
    #>
    param([string]$SourceFolder)

    $taskbarPath = "$env:APPDATA\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar"

    if (-not (Test-Path $SourceFolder)) {
        return @{ Success = $false; Error = "Source folder not found" }
    }

    try {
        if (-not (Test-Path $taskbarPath)) {
            New-Item -Path $taskbarPath -ItemType Directory -Force | Out-Null
        }

        $shortcuts = Get-ChildItem -Path $SourceFolder -Filter "*.lnk" -ErrorAction SilentlyContinue
        $restored = 0
        $skipped = 0

        foreach ($shortcut in $shortcuts) {
            $destPath = Join-Path $taskbarPath $shortcut.Name

            # Create .bak of existing if it exists
            if (Test-Path $destPath) {
                Copy-Item -Path $destPath -Destination "$destPath.bak" -Force
            }

            Copy-Item -Path $shortcut.FullName -Destination $destPath -Force
            $restored++
        }

        return @{ Success = $true; Restored = $restored; Skipped = $skipped }
    } catch {
        return @{ Success = $false; Error = $_.Exception.Message }
    }
}

#endregion PERSONALIZATION FUNCTIONS

#region APP CONFIG FUNCTIONS

function Backup-BrowserBookmarks {
    <#
    .SYNOPSIS
        Backs up browser bookmarks and creates extension list
    #>
    param(
        [string]$DestinationFolder,
        [hashtable]$Config
    )

    $userProfile = [Environment]::GetFolderPath("UserProfile")
    $sourcePath = Join-Path $userProfile $Config.Source
    $results = @{ Success = $false; Files = @(); Extensions = @() }

    if (-not (Test-Path $sourcePath)) {
        $results.Error = "Browser profile not found at $sourcePath"
        return $results
    }

    $destPath = Join-Path $DestinationFolder $Config.Name.Replace(" ", "")
    if (-not (Test-Path $destPath)) {
        New-Item -Path $destPath -ItemType Directory -Force | Out-Null
    }

    try {
        foreach ($item in $Config.Items) {
            $itemPath = Join-Path $sourcePath $item
            if (Test-Path $itemPath) {
                $itemDest = Join-Path $destPath $item

                if ((Get-Item $itemPath).PSIsContainer) {
                    # It's a directory
                    Copy-Item -Path $itemPath -Destination $itemDest -Recurse -Force
                } else {
                    Copy-Item -Path $itemPath -Destination $itemDest -Force
                }
                $results.Files += $item
            }
        }

        # If this browser has extensions, create a manifest
        if ($Config.ExtensionManifest) {
            $extensionsPath = Join-Path $sourcePath "Extensions"
            if (Test-Path $extensionsPath) {
                $extensions = @()
                Get-ChildItem -Path $extensionsPath -Directory -ErrorAction SilentlyContinue | ForEach-Object {
                    $extId = $_.Name
                    # Try to get extension name from manifest
                    $manifestFiles = Get-ChildItem -Path $_.FullName -Filter "manifest.json" -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
                    if ($manifestFiles) {
                        try {
                            $manifest = Get-Content $manifestFiles.FullName -Raw | ConvertFrom-Json
                            $extensions += @{
                                Id = $extId
                                Name = if ($manifest.name) { $manifest.name } else { $extId }
                            }
                        } catch {
                            $extensions += @{ Id = $extId; Name = $extId }
                        }
                    }
                }

                if ($extensions.Count -gt 0) {
                    $results.Extensions = $extensions
                    $extensions | ConvertTo-Json | Out-File (Join-Path $destPath "extensions_list.json") -Force
                }
            }
        }

        $results.Success = $true
        $results.Path = $destPath
    } catch {
        $results.Error = $_.Exception.Message
    }

    return $results
}

function Backup-FirefoxProfile {
    <#
    .SYNOPSIS
        Backs up Firefox profile (special handling for profile folder structure)
    #>
    param(
        [string]$DestinationFolder,
        [hashtable]$Config
    )

    $userProfile = [Environment]::GetFolderPath("UserProfile")
    $profilesPath = Join-Path $userProfile $Config.Source
    $results = @{ Success = $false; Files = @(); Extensions = @() }

    if (-not (Test-Path $profilesPath)) {
        $results.Error = "Firefox profiles not found"
        return $results
    }

    $destPath = Join-Path $DestinationFolder "MozillaFirefox"
    if (-not (Test-Path $destPath)) {
        New-Item -Path $destPath -ItemType Directory -Force | Out-Null
    }

    try {
        # Find the default profile (usually ends with .default or .default-release)
        $profiles = Get-ChildItem -Path $profilesPath -Directory -ErrorAction SilentlyContinue |
                    Where-Object { $_.Name -match "\.default" }

        foreach ($profile in $profiles) {
            $profileDest = Join-Path $destPath $profile.Name
            if (-not (Test-Path $profileDest)) {
                New-Item -Path $profileDest -ItemType Directory -Force | Out-Null
            }

            foreach ($item in $Config.Items) {
                $itemPath = Join-Path $profile.FullName $item
                if (Test-Path $itemPath) {
                    Copy-Item -Path $itemPath -Destination $profileDest -Force
                    $results.Files += "$($profile.Name)\$item"
                }
            }
        }

        $results.Success = $true
        $results.Path = $destPath
    } catch {
        $results.Error = $_.Exception.Message
    }

    return $results
}

function Backup-VSCodeConfig {
    <#
    .SYNOPSIS
        Backs up VS Code settings and exports extension list
    #>
    param(
        [string]$DestinationFolder,
        [hashtable]$Config
    )

    $userProfile = [Environment]::GetFolderPath("UserProfile")
    $sourcePath = Join-Path $userProfile $Config.Source
    $results = @{ Success = $false; Files = @(); Extensions = @() }

    if (-not (Test-Path $sourcePath)) {
        $results.Error = "VS Code config not found"
        return $results
    }

    $destPath = Join-Path $DestinationFolder "VSCode"
    if (-not (Test-Path $destPath)) {
        New-Item -Path $destPath -ItemType Directory -Force | Out-Null
    }

    try {
        foreach ($item in $Config.Items) {
            $itemPath = Join-Path $sourcePath $item
            if (Test-Path $itemPath) {
                $itemDest = Join-Path $destPath $item
                if ((Get-Item $itemPath).PSIsContainer) {
                    Copy-Item -Path $itemPath -Destination $itemDest -Recurse -Force
                } else {
                    Copy-Item -Path $itemPath -Destination $itemDest -Force
                }
                $results.Files += $item
            }
        }

        # Export extension list using VS Code CLI
        if ($Config.ExtensionList) {
            try {
                $extensions = & code --list-extensions 2>$null
                if ($extensions) {
                    $results.Extensions = $extensions
                    $extensions | Out-File (Join-Path $destPath "extensions_list.txt") -Force
                }
            } catch {
                # VS Code CLI not available
            }
        }

        $results.Success = $true
        $results.Path = $destPath
    } catch {
        $results.Error = $_.Exception.Message
    }

    return $results
}

function Backup-GenericAppConfig {
    <#
    .SYNOPSIS
        Backs up a generic application config based on whitelist entry
    #>
    param(
        [string]$DestinationFolder,
        [hashtable]$Config
    )

    $userProfile = [Environment]::GetFolderPath("UserProfile")
    $sourcePath = Join-Path $userProfile $Config.Source
    $results = @{ Success = $false; Files = @() }

    if (-not (Test-Path $sourcePath)) {
        $results.Error = "Config path not found: $($Config.Source)"
        return $results
    }

    $destPath = Join-Path $DestinationFolder ($Config.Name -replace '\s+', '')
    if (-not (Test-Path $destPath)) {
        New-Item -Path $destPath -ItemType Directory -Force | Out-Null
    }

    try {
        foreach ($item in $Config.Items) {
            $itemPath = Join-Path $sourcePath $item
            if (Test-Path $itemPath) {
                $itemDest = Join-Path $destPath $item
                if ((Get-Item $itemPath).PSIsContainer) {
                    Copy-Item -Path $itemPath -Destination $itemDest -Recurse -Force
                } else {
                    Copy-Item -Path $itemPath -Destination $itemDest -Force
                }
                $results.Files += $item
            }
        }

        $results.Success = $true
        $results.Path = $destPath
    } catch {
        $results.Error = $_.Exception.Message
    }

    return $results
}

function Backup-AllAppConfigs {
    <#
    .SYNOPSIS
        Backs up all enabled application configurations
    #>
    param([string]$DestinationFolder)

    $results = @{}

    foreach ($config in $script:AppConfigWhitelist) {
        if (-not $config.Enabled) { continue }

        Write-Info "Backing up $($config.Name) configuration..."

        switch ($config.Name) {
            "Google Chrome" {
                $results[$config.Name] = Backup-BrowserBookmarks -DestinationFolder $DestinationFolder -Config $config
            }
            "Microsoft Edge" {
                $results[$config.Name] = Backup-BrowserBookmarks -DestinationFolder $DestinationFolder -Config $config
            }
            "Mozilla Firefox" {
                $results[$config.Name] = Backup-FirefoxProfile -DestinationFolder $DestinationFolder -Config $config
            }
            "VS Code" {
                $results[$config.Name] = Backup-VSCodeConfig -DestinationFolder $DestinationFolder -Config $config
            }
            default {
                $results[$config.Name] = Backup-GenericAppConfig -DestinationFolder $DestinationFolder -Config $config
            }
        }

        if ($results[$config.Name].Success) {
            Write-Success "$($config.Name): $($results[$config.Name].Files.Count) items backed up"
        } else {
            Write-Warning "$($config.Name): $($results[$config.Name].Error)"
        }
    }

    return $results
}

function Restore-AllAppConfigs {
    <#
    .SYNOPSIS
        Restores all application configurations from backup
    #>
    param([string]$SourceFolder)

    $results = @{}
    $userProfile = [Environment]::GetFolderPath("UserProfile")

    foreach ($config in $script:AppConfigWhitelist) {
        if (-not $config.Enabled) { continue }

        $appFolder = Join-Path $SourceFolder ($config.Name -replace '\s+', '')

        if (-not (Test-Path $appFolder)) {
            Write-Warning "$($config.Name): No backup found"
            continue
        }

        Write-Info "Restoring $($config.Name) configuration..."

        $destPath = Join-Path $userProfile $config.Source
        $results[$config.Name] = @{ Success = $false; Restored = @() }

        # Handle Firefox specially due to profile folder structure
        if ($config.Name -eq "Mozilla Firefox") {
            if (-not (Test-Path $destPath)) {
                Write-Warning "$($config.Name): Firefox not installed, skipping"
                $results[$config.Name].Error = "Firefox not installed"
                continue
            }

            # Find profiles in both source and destination
            $sourceProfiles = Get-ChildItem -Path $appFolder -Directory -ErrorAction SilentlyContinue
            $destProfiles = Get-ChildItem -Path $destPath -Directory -ErrorAction SilentlyContinue |
                           Where-Object { $_.Name -match "\.default" }

            foreach ($sourceProfile in $sourceProfiles) {
                $destProfile = $destProfiles | Select-Object -First 1
                if ($destProfile) {
                    foreach ($item in (Get-ChildItem -Path $sourceProfile.FullName)) {
                        $destItem = Join-Path $destProfile.FullName $item.Name
                        if (Test-Path $destItem) {
                            Copy-Item -Path $destItem -Destination "$destItem.bak" -Force
                        }
                        Copy-Item -Path $item.FullName -Destination $destItem -Force
                        $results[$config.Name].Restored += $item.Name
                    }
                }
            }

            $results[$config.Name].Success = $true
        }
        else {
            # Generic restore
            try {
                if (-not (Test-Path $destPath)) {
                    New-Item -Path $destPath -ItemType Directory -Force | Out-Null
                }

                foreach ($item in (Get-ChildItem -Path $appFolder)) {
                    $destItem = Join-Path $destPath $item.Name

                    # Backup existing
                    if (Test-Path $destItem) {
                        if ((Get-Item $destItem).PSIsContainer) {
                            # For directories, rename to .bak
                            if (Test-Path "$destItem.bak") {
                                Remove-Item "$destItem.bak" -Recurse -Force
                            }
                            Rename-Item -Path $destItem -NewName "$($item.Name).bak" -Force
                        } else {
                            Copy-Item -Path $destItem -Destination "$destItem.bak" -Force
                        }
                    }

                    if ($item.PSIsContainer) {
                        Copy-Item -Path $item.FullName -Destination $destItem -Recurse -Force
                    } else {
                        Copy-Item -Path $item.FullName -Destination $destItem -Force
                    }
                    $results[$config.Name].Restored += $item.Name
                }

                $results[$config.Name].Success = $true
            } catch {
                $results[$config.Name].Error = $_.Exception.Message
            }
        }

        if ($results[$config.Name].Success) {
            Write-Success "$($config.Name): $($results[$config.Name].Restored.Count) items restored"
        } else {
            Write-Warning "$($config.Name): $($results[$config.Name].Error)"
        }
    }

    return $results
}

#endregion APP CONFIG FUNCTIONS

#region ROBOCOPY FUNCTIONS

function Invoke-RobocopyWithProgress {
    <#
    .SYNOPSIS
        Runs RoboCopy with progress monitoring
    #>
    param(
        [string]$Source,
        [string]$Destination,
        [string]$LogFile,
        [bool]$Mirror = $true,
        [array]$ExcludeFiles = @(),
        [array]$ExcludeDirs = @()
    )

    $results = @{
        Success = $false
        FilesCopied = 0
        BytesCopied = 0
        Errors = @()
        Skipped = @()
    }

    if (-not (Test-Path $Source)) {
        $results.Errors += "Source path does not exist: $Source"
        return $results
    }

    # Create destination if it doesn't exist
    if (-not (Test-Path $Destination)) {
        try {
            New-Item -Path $Destination -ItemType Directory -Force | Out-Null
        } catch {
            $results.Errors += "Failed to create destination: $_"
            return $results
        }
    }

    # Count total files for progress
    Write-Host "  Analyzing folder..." -ForegroundColor DarkGray -NoNewline
    $totalFiles = 0
    $totalSize = 0
    try {
        $files = Get-ChildItem -Path $Source -Recurse -File -ErrorAction SilentlyContinue
        $totalFiles = $files.Count
        $totalSize = ($files | Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
    } catch { }
    Write-Host " $totalFiles files ($(Format-Size $totalSize))" -ForegroundColor DarkGray

    # Build RoboCopy arguments - using /E instead of /MIR to avoid deleting files
    $robocopyArgs = @(
        "`"$Source`"",
        "`"$Destination`"",
        "/E",                    # Copy subdirectories including empty ones
        "/MT:$script:Threads",   # Multithreaded
        "/R:$script:RetryCount", # Retry count
        "/W:$script:RetryWait",  # Wait between retries
        "/XJ",                   # Exclude junction points
        "/DCOPY:T",              # Copy directory timestamps
        "/COPY:DAT",             # Copy Data, Attributes, Timestamps
        "/NP",                   # No progress percentage (we handle our own)
        "/NDL",                  # No directory list
        "/NC",                   # No file class
        "/BYTES",                # Print sizes as bytes
        "/LOG+:`"$LogFile`""     # Append to log file
    )

    # Add file exclusions
    foreach ($file in ($script:ExcludeFiles + $ExcludeFiles)) {
        $robocopyArgs += "/XF"
        $robocopyArgs += "`"$file`""
    }

    # Add directory exclusions
    foreach ($dir in ($script:ExcludeDirs + $ExcludeDirs)) {
        $robocopyArgs += "/XD"
        $robocopyArgs += "`"$dir`""
    }

    # Start RoboCopy process
    $processInfo = New-Object System.Diagnostics.ProcessStartInfo
    $processInfo.FileName = "robocopy.exe"
    $processInfo.Arguments = $robocopyArgs -join " "
    $processInfo.RedirectStandardOutput = $true
    $processInfo.RedirectStandardError = $true
    $processInfo.UseShellExecute = $false
    $processInfo.CreateNoWindow = $true

    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $processInfo

    # Start RoboCopy and show simple spinner (minimal overhead for maximum speed)
    $startTime = Get-Date
    $spinChars = @('|', '/', '-', '\')
    $spinIndex = 0

    [void]$process.Start()

    # Simple spinner - no file counting, no output parsing, just elapsed time
    while (-not $process.HasExited) {
        $elapsed = (Get-Date) - $startTime
        $spinChar = $spinChars[$spinIndex % 4]
        $spinIndex++
        Write-Host "`r  $spinChar Copying $totalFiles files... | Elapsed: $($elapsed.ToString('mm\:ss'))     " -NoNewline
        Start-Sleep -Milliseconds 200
    }

    # Get output for result parsing
    $output = $process.StandardOutput.ReadToEnd()
    $exitCode = $process.ExitCode
    $process.Close()

    Write-Host ""  # New line after progress

    # RoboCopy exit codes: 0-7 are success/info, 8+ are errors
    if ($exitCode -lt 8) {
        $results.Success = $true

        # Note: When using /LOG+:, RoboCopy output goes to log file, not stdout
        # So we use the pre-calculated source file count and size as the result
        # The exit code confirms the copy succeeded
        $results.FilesCopied = $totalFiles
        $results.BytesCopied = if ($totalSize) { $totalSize } else { 0 }

        # Try to parse stdout if available (may have partial info without /LOG)
        if ($output -and $output.Length -gt 100) {
            # Try to extract actual copied count from RoboCopy summary
            if ($output -match "Files\s*:\s*\d+\s+(\d+)") {
                $results.FilesCopied = [int]$Matches[1]
            }
            if ($output -match "Bytes\s*:\s*\d+\s+(\d+)") {
                $results.BytesCopied = [int64]$Matches[1]
            }
        }
    } else {
        $results.Success = $false
        if ($exitCode -band 8) {
            $results.Errors += "Some files could not be copied (copy errors)"
        }
        if ($exitCode -band 16) {
            $results.Errors += "Serious error - robocopy did not copy any files"
        }
    }

    return $results
}

#endregion ROBOCOPY FUNCTIONS

#region BACKUP OPERATION

function Invoke-BackupOperation {
    Write-Header "BACKUP OPERATION"

    $userProfile = [Environment]::GetFolderPath("UserProfile")
    $hostname = $env:COMPUTERNAME

    # State variables for each step (preserved when going back)
    $basePath = $null
    $selectedFolders = $null
    $selectedOptions = $null
    $folderItems = $null  # Cache folder list to avoid recalculating

    # Step-based navigation
    $currentStep = 1
    $maxStep = 4

    while ($currentStep -le $maxStep) {
        switch ($currentStep) {
            1 {
                # Step 1: Select destination drive
                Write-Header "BACKUP - Step 1 of 3: Select Destination"

                $drives = Get-RemovableDrives
                if ($drives.Count -eq 0) {
                    Write-Error "No available drives found for backup."
                    return
                }

                $driveOptions = $drives | ForEach-Object { @{ Name = "$($_.Name) - $($_.FreeSpaceGB) GB free"; Value = $_ } }
                $driveOptions += @{ Name = "Enter path manually..."; Value = "manual" }

                $selectedDrive = Show-InteractiveMenu -Title "Select Backup Drive" -Items $driveOptions -AllowBack $false

                if ($null -eq $selectedDrive) {
                    Write-Warning "Operation cancelled."
                    return
                }

                if ($selectedDrive.Value -eq "manual") {
                    Write-Host ""
                    $basePath = Read-Host "Enter full path for backup destination"
                    if (-not $basePath) {
                        Write-Warning "No path entered. Operation cancelled."
                        return
                    }
                } else {
                    $basePath = $selectedDrive.Value.Letter
                    Write-Host ""
                    $subFolder = Read-Host "Enter subfolder name (press Enter for drive root)"
                    if ($subFolder) {
                        $basePath = Join-Path $basePath $subFolder
                    }
                }

                # Validate the base path exists or can be created
                if (-not (Test-Path $basePath)) {
                    try {
                        New-Item -Path $basePath -ItemType Directory -Force | Out-Null
                        Write-Success "Created directory: $basePath"
                        Start-Sleep -Milliseconds 500
                    } catch {
                        Write-Error "Failed to create directory: $_"
                        Start-Sleep -Seconds 2
                        continue  # Retry step 1
                    }
                }

                $currentStep = 2
            }

            2 {
                # Step 2: Select folders to backup
                Write-Header "BACKUP - Step 2 of 3: Select Folders"

                # Get available folders (cache if not already loaded)
                if ($null -eq $folderItems) {
                    $folderItems = Get-UserFoldersForBackup
                }

                if ($folderItems.Count -eq 0) {
                    Write-Error "No user folders found to backup."
                    return
                }

                $result = Show-InteractiveMenu -Title "Select Folders to Backup" -Items $folderItems -MultiSelect $true -ShowSize $true -AllowBack $true

                if (Test-GoBack $result) {
                    $currentStep = 1
                    continue
                }

                if ($null -eq $result -or $result.Count -eq 0) {
                    Write-Warning "No folders selected. Operation cancelled."
                    return
                }

                $selectedFolders = $result
                $currentStep = 3
            }

            3 {
                # Step 3: Select additional options
                Write-Header "BACKUP - Step 3 of 3: Additional Options"

                $options = @(
                    @{ Name = "Backup application configs (browsers, VS Code, etc.)"; Selected = $true },
                    @{ Name = "Backup desktop wallpaper"; Selected = $true },
                    @{ Name = "Backup taskbar pinned items"; Selected = $true },
                    @{ Name = "Generate application inventory"; Selected = $true },
                    @{ Name = "Enable dry run (no actual copy)"; Selected = $false }
                )

                $result = Show-InteractiveMenu -Title "Additional Backup Options" -Items $options -MultiSelect $true -AllowBack $true

                if (Test-GoBack $result) {
                    $currentStep = 2
                    continue
                }

                $selectedOptions = $result
                $currentStep = 4
            }

            4 {
                # Step 4: Summary and confirmation - exit the loop
                $currentStep = $maxStep + 1  # Exit the while loop
            }
        }
    }

    # Parse selected options
    $backupAppConfigs = ($selectedOptions | Where-Object { $_.Name -match "application configs" }).Count -gt 0
    $backupWallpaper = ($selectedOptions | Where-Object { $_.Name -match "wallpaper" }).Count -gt 0
    $backupTaskbar = ($selectedOptions | Where-Object { $_.Name -match "taskbar" }).Count -gt 0
    $generateAppInventory = ($selectedOptions | Where-Object { $_.Name -match "inventory" }).Count -gt 0
    $dryRun = ($selectedOptions | Where-Object { $_.Name -match "dry run" }).Count -gt 0

    # Now create backup folder (after all selections confirmed)
    $timestamp = Get-Date -Format "yyyy-MM-dd_HHmmss"
    $backupFolder = Join-Path $basePath "$hostname-$timestamp"

    # Calculate sizes for summary
    Write-Section "Calculating Sizes"

    $totalSize = 0
    $folderSizes = @{}

    foreach ($folder in $selectedFolders) {
        $folderPath = $folder.Path

        if (-not $folderPath) {
            Write-Warning "  $($folder.Name): No path found, skipping"
            continue
        }

        Write-Host "  $($folder.Name)..." -NoNewline -ForegroundColor Cyan
        $size = Get-FolderSize -Path $folderPath
        $folderSizes[$folder.Name] = $size
        $totalSize += $size
        Write-Host " $(Format-Size $size)" -ForegroundColor DarkGray
    }

    # Check destination space
    $destDrive = Split-Path $basePath -Qualifier
    $destInfo = Get-WmiObject Win32_LogicalDisk -Filter "DeviceID='$destDrive'"
    $freeSpace = $destInfo.FreeSpace

    # Show summary
    Write-Header "BACKUP SUMMARY"

    Write-Host "  Destination:  $backupFolder" -ForegroundColor White
    Write-Host ""
    Write-Host "  Folders to backup:" -ForegroundColor Yellow
    foreach ($folder in $selectedFolders) {
        Write-Host "    - $($folder.Name) ($(Format-Size $folderSizes[$folder.Name]))" -ForegroundColor White
        Write-Host "      $($folder.Path)" -ForegroundColor DarkGray
    }
    Write-Host ""
    Write-Host "  Total size:    $(Format-Size $totalSize)" -ForegroundColor Yellow
    Write-Host "  Free space:    $(Format-Size $freeSpace)" -ForegroundColor Yellow

    if ($totalSize -gt $freeSpace) {
        Write-Host ""
        Write-Error "Not enough space on destination drive!"
        Write-Host "  Need: $(Format-Size $totalSize) | Have: $(Format-Size $freeSpace)"
        return
    }

    Write-Host ""
    Write-Host "  Additional operations:" -ForegroundColor Yellow
    if ($backupAppConfigs) { Write-Host "    - Backup application configs" }
    if ($backupWallpaper) { Write-Host "    - Backup desktop wallpaper" }
    if ($backupTaskbar) { Write-Host "    - Backup taskbar pins" }
    if ($generateAppInventory) { Write-Host "    - Generate application inventory" }
    if ($dryRun) { Write-Host "    - DRY RUN MODE (no actual copy)" -ForegroundColor Magenta }

    Write-Host ""
    Write-Host "  [BACKSPACE] Go back to options   [ENTER] Proceed   [ESC] Cancel" -ForegroundColor DarkGray
    Write-Host ""

    # Allow going back from summary
    $confirmKey = [Console]::ReadKey($true)
    if ($confirmKey.Key -eq "Escape") {
        Write-Warning "Backup cancelled."
        return
    }
    if ($confirmKey.Key -eq "Backspace" -or $confirmKey.Key -eq "B") {
        # Go back to step 3 - need to re-run the function with preserved state
        # For simplicity, just restart the whole operation
        Write-Host "Going back..." -ForegroundColor Yellow
        Start-Sleep -Milliseconds 500
        Invoke-BackupOperation
        return
    }

    # Create backup folder structure now that user confirmed
    try {
        New-Item -Path $backupFolder -ItemType Directory -Force | Out-Null
        Write-Success "Created backup folder: $backupFolder"
    } catch {
        Write-Error "Failed to create backup folder: $_"
        return
    }

    $userDataFolder = Join-Path $backupFolder "UserData"
    $appConfigFolder = Join-Path $backupFolder "AppConfigs"
    $personalizationFolder = Join-Path $backupFolder "Personalization"
    $logsFolder = Join-Path $backupFolder "Logs"

    New-Item -Path $userDataFolder -ItemType Directory -Force | Out-Null
    New-Item -Path $appConfigFolder -ItemType Directory -Force | Out-Null
    New-Item -Path $personalizationFolder -ItemType Directory -Force | Out-Null
    New-Item -Path $logsFolder -ItemType Directory -Force | Out-Null

    # Initialize logging - separate files for readable script log and RoboCopy technical log
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $logFile = Join-Path $logsFolder "backup_$timestamp.log"
    $robocopyLogFile = Join-Path $logsFolder "backup_robocopy_$timestamp.log"
    $errorLog = @()
    $skipLog = @()

    # Initialize script log
    "========================================" | Out-File -FilePath $logFile -Force -Encoding UTF8
    "Backup Log - $(Get-Date)" | Out-File -FilePath $logFile -Append -Encoding UTF8
    "Computer: $hostname" | Out-File -FilePath $logFile -Append -Encoding UTF8
    "Destination: $backupFolder" | Out-File -FilePath $logFile -Append -Encoding UTF8
    "========================================" | Out-File -FilePath $logFile -Append -Encoding UTF8
    "" | Out-File -FilePath $logFile -Append -Encoding UTF8

    # Step 5: Execute backup
    Write-Header "EXECUTING BACKUP"

    $manifest = @{
        BackupDate = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        ComputerName = $hostname
        WindowsVersion = (Get-WmiObject Win32_OperatingSystem).Caption
        UserName = $env:USERNAME
        Folders = @()
        AppConfigs = @()
        Personalization = @{}
        Applications = @()
    }

    # Backup user folders
    if (-not $dryRun) {
        foreach ($folder in $selectedFolders) {
            # Source path is absolute, destination uses folder name for consistent naming
            $sourcePath = $folder.Path
            $destFolderName = if ($folder.RelativeName) { $folder.RelativeName } else { $folder.Name }
            $destPath = Join-Path $userDataFolder $destFolderName

            Write-Section "Backing up $($folder.Name)"
            Write-Host "  Source: $sourcePath" -ForegroundColor DarkGray

            "Processing: $($folder.Name)" | Out-File -FilePath $logFile -Append -Encoding UTF8
            "  Source: $sourcePath" | Out-File -FilePath $logFile -Append -Encoding UTF8
            "  Destination: $destPath" | Out-File -FilePath $logFile -Append -Encoding UTF8

            $result = Invoke-RobocopyWithProgress -Source $sourcePath -Destination $destPath -LogFile $robocopyLogFile

            if ($result.Success) {
                Write-Success "$($folder.Name) - $(Format-Size $result.BytesCopied) copied"
                "SUCCESS: $($folder.Name) - $($result.FilesCopied) files, $(Format-Size $result.BytesCopied)" | Out-File -FilePath $logFile -Append -Encoding UTF8
                $manifest.Folders += @{
                    Name = $folder.Name
                    SourcePath = $folder.Path
                    BackupPath = $destFolderName
                    FilesCopied = $result.FilesCopied
                    BytesCopied = $result.BytesCopied
                }
            } else {
                Write-Warning "$($folder.Name) completed with errors"
                "WARNING: $($folder.Name) completed with errors" | Out-File -FilePath $logFile -Append -Encoding UTF8
                foreach ($err in $result.Errors) {
                    "  - $err" | Out-File -FilePath $logFile -Append -Encoding UTF8
                }
                $errorLog += $result.Errors
            }
        }
    } else {
        Write-Info "DRY RUN: Skipping folder copy"
        foreach ($folder in $selectedFolders) {
            $destFolderName = if ($folder.RelativeName) { $folder.RelativeName } else { $folder.Name }
            $manifest.Folders += @{
                Name = $folder.Name
                SourcePath = $folder.Path
                BackupPath = $destFolderName
                FilesCopied = 0
                BytesCopied = $folderSizes[$folder.Name]
                DryRun = $true
            }
        }
    }

    # Backup app configs
    if ($backupAppConfigs) {
        Write-Section "Backing up Application Configurations"
        if (-not $dryRun) {
            $configResults = Backup-AllAppConfigs -DestinationFolder $appConfigFolder
            foreach ($app in $configResults.Keys) {
                $manifest.AppConfigs += @{
                    Name = $app
                    Success = $configResults[$app].Success
                    Files = $configResults[$app].Files
                }
            }
        } else {
            Write-Info "DRY RUN: Skipping app config backup"
        }
    }

    # Backup personalization
    if ($backupWallpaper) {
        Write-Section "Backing up Wallpaper"
        if (-not $dryRun) {
            $wallpaperResult = Backup-Wallpaper -DestinationFolder $personalizationFolder
            $manifest.Personalization.Wallpaper = $wallpaperResult
            if ($wallpaperResult.Success) {
                Write-Success "Wallpaper backed up"
            } else {
                Write-Warning "Wallpaper backup: $($wallpaperResult.Error)"
            }
        } else {
            Write-Info "DRY RUN: Skipping wallpaper backup"
        }
    }

    if ($backupTaskbar) {
        Write-Section "Backing up Taskbar Pins"
        if (-not $dryRun) {
            $taskbarResult = Backup-TaskbarPins -DestinationFolder $personalizationFolder
            $manifest.Personalization.TaskbarPins = $taskbarResult
            if ($taskbarResult.Success) {
                Write-Success "Taskbar pins backed up ($($taskbarResult.Count) items)"
            } else {
                Write-Warning "Taskbar backup: $($taskbarResult.Error)"
            }
        } else {
            Write-Info "DRY RUN: Skipping taskbar backup"
        }
    }

    # Generate application inventory
    if ($generateAppInventory) {
        Write-Section "Generating Application Inventory"

        $apps = Get-ApplicationInventoryWithPackages -SkipPackageMatching:$dryRun
        $manifest.Applications = $apps

        # Save CSV
        $csvPath = Join-Path $backupFolder "AppInventory.csv"
        $apps | ForEach-Object {
            [PSCustomObject]@{
                Name = $_.Name
                Version = $_.Version
                Publisher = $_.Publisher
                WingetId = if ($_.WingetId) { $_.WingetId } else { "" }
            }
        } | Export-Csv -Path $csvPath -NoTypeInformation
        Write-Success "Application inventory saved to AppInventory.csv"

        # Generate reinstall script
        $appsWithWinget = $apps | Where-Object { $_.WingetId }
        if ($appsWithWinget.Count -gt 0) {
            $reinstallScript = @"
# AppReinstall.ps1
# Generated by Enhanced-DataMigration on $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
# Source Computer: $hostname
#
# This script will reinstall applications using Winget.
# Review the list and remove any apps you don't want to install.
#
# Usage: Run this script in an elevated PowerShell session
#        .\AppReinstall.ps1

`$ErrorActionPreference = 'Continue'

# Check for Winget
if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
    Write-Host "Winget is not available. Please install it first." -ForegroundColor Red
    exit 1
}

Write-Host "Starting application installation..." -ForegroundColor Cyan
Write-Host "This may take a while. You can cancel with Ctrl+C" -ForegroundColor Yellow
Write-Host ""

`$apps = @(
"@
            foreach ($app in $appsWithWinget) {
                $reinstallScript += "`n    @{ Name = `"$($app.Name)`"; Id = `"$($app.WingetId)`" }"
                if ($app -ne $appsWithWinget[-1]) { $reinstallScript += "," }
            }

            $reinstallScript += @"

)

`$total = `$apps.Count
`$current = 0

foreach (`$app in `$apps) {
    `$current++
    Write-Host "[`$current/`$total] Installing `$(`$app.Name)..." -ForegroundColor Yellow

    try {
        winget install --id `$app.Id --accept-package-agreements --accept-source-agreements --silent
        if (`$LASTEXITCODE -eq 0) {
            Write-Host "  [OK] Installed successfully" -ForegroundColor Green
        } else {
            Write-Host "  [WARNING] Installation may have issues (exit code: `$LASTEXITCODE)" -ForegroundColor Yellow
        }
    } catch {
        Write-Host "  [ERROR] Failed: `$_" -ForegroundColor Red
    }

    Write-Host ""
}

Write-Host "Installation complete!" -ForegroundColor Cyan
Write-Host "Some applications may require a restart to complete installation." -ForegroundColor Yellow
"@

            $reinstallPath = Join-Path $backupFolder "AppReinstall.ps1"
            $reinstallScript | Out-File -FilePath $reinstallPath -Encoding UTF8
            Write-Success "Reinstall script saved to AppReinstall.ps1 ($($appsWithWinget.Count) apps)"
        }
    }

    # Save manifest
    $manifestPath = Join-Path $backupFolder "manifest.json"
    $manifest | ConvertTo-Json -Depth 10 | Out-File -FilePath $manifestPath -Encoding UTF8

    # Save error log if any
    if ($errorLog.Count -gt 0) {
        $errorLogPath = Join-Path $logsFolder "errors.log"
        $errorLog | Out-File -FilePath $errorLogPath -Encoding UTF8
        Write-Warning "Some errors occurred. See: $errorLogPath"
    }

    # Write completion to log
    "" | Out-File -FilePath $logFile -Append -Encoding UTF8
    "========================================" | Out-File -FilePath $logFile -Append -Encoding UTF8
    "Backup completed: $(Get-Date)" | Out-File -FilePath $logFile -Append -Encoding UTF8
    "========================================" | Out-File -FilePath $logFile -Append -Encoding UTF8

    # Final summary
    Write-Header "BACKUP COMPLETE"

    Write-Host "  Backup Location: $backupFolder" -ForegroundColor Green
    Write-Host ""
    Write-Host "  Contents:" -ForegroundColor Yellow
    Write-Host "    - UserData\        : User profile folders"
    Write-Host "    - AppConfigs\      : Application configurations"
    Write-Host "    - Personalization\ : Wallpaper, taskbar pins"
    Write-Host "    - Logs\            : Backup logs"
    Write-Host "    - manifest.json    : Backup metadata"
    Write-Host "    - AppInventory.csv : Installed applications list"
    if (Test-Path (Join-Path $backupFolder "AppReinstall.ps1")) {
        Write-Host "    - AppReinstall.ps1 : Application reinstall script"
    }

    Write-Host ""
    Write-Host "  Log files:" -ForegroundColor DarkGray
    Write-Host "    Script log:   $logFile" -ForegroundColor DarkGray
    Write-Host "    RoboCopy log: $robocopyLogFile" -ForegroundColor DarkGray

    Write-Host ""
    Write-Host "  Next Steps:" -ForegroundColor Cyan
    Write-Host "    1. Safely eject the external drive"
    Write-Host "    2. Re-image or replace the computer"
    Write-Host "    3. Run this script in RESTORE mode"
    Write-Host "    4. Run AppReinstall.ps1 to reinstall applications"
    Write-Host ""
}

#endregion BACKUP OPERATION

#region RESTORE OPERATION

function Invoke-RestoreOperation {
    Write-Header "RESTORE OPERATION"

    $userProfile = [Environment]::GetFolderPath("UserProfile")

    # Step 1: Select restore source
    Write-Section "Select Restore Source"

    $drives = Get-RemovableDrives
    if ($drives.Count -eq 0) {
        Write-Error "No available drives found."
        return
    }

    $driveOptions = $drives | ForEach-Object { @{ Name = "$($_.Name) - $($_.FreeSpaceGB) GB free"; Value = $_ } }
    $driveOptions += @{ Name = "Enter path manually..."; Value = "manual" }

    $selectedDrive = Show-InteractiveMenu -Title "Select Drive with Backup" -Items $driveOptions

    if ($null -eq $selectedDrive) {
        Write-Warning "Operation cancelled."
        return
    }

    $searchPath = ""
    if ($selectedDrive.Value -eq "manual") {
        Write-Host ""
        $searchPath = Read-Host "Enter full path to backup folder"
    } else {
        $searchPath = $selectedDrive.Value.Letter
        Write-Host ""
        $subFolder = Read-Host "Enter subfolder to search (press Enter for drive root)"
        if ($subFolder) {
            $searchPath = Join-Path $searchPath $subFolder
        }
    }

    # Find backups
    Write-Info "Searching for backups..."
    $backups = Get-BackupFolders -BasePath $searchPath

    if ($backups.Count -eq 0) {
        # Check if the path itself is a backup
        $manifestPath = Join-Path $searchPath "manifest.json"
        if (Test-Path $manifestPath) {
            $backups = @(@{
                Name = (Split-Path $searchPath -Leaf)
                Path = $searchPath
                Date = Get-Date
            })
        } else {
            Write-Error "No backups found at $searchPath"
            Write-Host ""
            $manualPath = Read-Host "Enter full path to specific backup folder"
            if ($manualPath -and (Test-Path $manualPath)) {
                $backups = @(@{
                    Name = (Split-Path $manualPath -Leaf)
                    Path = $manualPath
                    Date = Get-Date
                })
            } else {
                Write-Warning "No valid backup path provided."
                return
            }
        }
    }

    # Select backup to restore
    $backupOptions = $backups | ForEach-Object {
        @{
            Name = "$($_.Name) ($($_.Date.ToString('yyyy-MM-dd HH:mm')))"
            Value = $_
        }
    }
    $backupOptions += @{ Name = "Enter path manually..."; Value = "manual" }

    $selectedBackup = Show-InteractiveMenu -Title "Select Backup to Restore" -Items $backupOptions

    if ($null -eq $selectedBackup) {
        Write-Warning "Operation cancelled."
        return
    }

    $backupPath = ""
    if ($selectedBackup.Value -eq "manual") {
        Write-Host ""
        $backupPath = Read-Host "Enter full path to backup folder"
    } else {
        $backupPath = $selectedBackup.Value.Path
    }

    if (-not (Test-Path $backupPath)) {
        Write-Error "Backup path not found: $backupPath"
        return
    }

    # Load manifest if available
    $manifest = $null
    $manifestPath = Join-Path $backupPath "manifest.json"
    if (Test-Path $manifestPath) {
        try {
            $manifest = Get-Content $manifestPath -Raw | ConvertFrom-Json
            Write-Info "Backup from: $($manifest.ComputerName) on $($manifest.BackupDate)"
            Write-Info "Windows version: $($manifest.WindowsVersion)"
        } catch { }
    }

    # Step 2: Select what to restore
    Write-Section "Select Items to Restore"

    $userDataPath = Join-Path $backupPath "UserData"
    $appConfigPath = Join-Path $backupPath "AppConfigs"
    $personalizationPath = Join-Path $backupPath "Personalization"

    $restoreOptions = @()

    # Find available folders and resolve correct destination paths
    if (Test-Path $userDataPath) {
        Get-ChildItem -Path $userDataPath -Directory | ForEach-Object {
            $folderName = $_.Name
            $destPath = $null

            # First, try to get the original SourcePath from manifest
            if ($manifest -and $manifest.Folders) {
                $manifestEntry = $manifest.Folders | Where-Object {
                    $_.Name -eq $folderName -or $_.BackupPath -eq $folderName
                } | Select-Object -First 1

                if ($manifestEntry -and $manifestEntry.SourcePath) {
                    # Check if the original path still exists (same machine or similar setup)
                    if (Test-Path $manifestEntry.SourcePath -ErrorAction SilentlyContinue) {
                        $destPath = $manifestEntry.SourcePath
                    }
                }
            }

            # If no manifest path, try to resolve using Windows special folders
            if (-not $destPath) {
                # Map folder names to special folder detection
                $folderDef = $script:DefaultFolders | Where-Object { $_.Name -eq $folderName } | Select-Object -First 1
                if ($folderDef) {
                    $resolvedPath = Get-SpecialFolderPath -FolderName $folderDef.Name -SpecialFolder $folderDef.SpecialFolder -KnownFolder $folderDef.KnownFolder
                    if ($resolvedPath) {
                        $destPath = $resolvedPath
                    }
                }
            }

            # Final fallback: use user profile + folder name
            if (-not $destPath) {
                $destPath = Join-Path $userProfile $folderName
            }

            $restoreOptions += @{
                Name = "User Data: $folderName"
                Selected = $true
                Type = "UserData"
                Path = $_.FullName
                FolderName = $folderName
                DestPath = $destPath
            }
        }
    }

    if (Test-Path $appConfigPath) {
        $restoreOptions += @{
            Name = "Application Configurations"
            Selected = $true
            Type = "AppConfigs"
            Path = $appConfigPath
        }
    }

    if (Test-Path (Join-Path $personalizationPath "TaskbarPins")) {
        $restoreOptions += @{
            Name = "Taskbar Pinned Items"
            Selected = $true
            Type = "TaskbarPins"
            Path = Join-Path $personalizationPath "TaskbarPins"
        }
    }

    # Check for wallpaper
    $wallpaperFile = Get-ChildItem -Path $personalizationPath -Filter "wallpaper.*" -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($wallpaperFile) {
        $restoreOptions += @{
            Name = "Desktop Wallpaper"
            Selected = $true
            Type = "Wallpaper"
            Path = $wallpaperFile.FullName
        }
    }

    if ($restoreOptions.Count -eq 0) {
        Write-Error "No restorable items found in backup."
        return
    }

    $selectedItems = Show-InteractiveMenu -Title "Select Items to Restore" -Items $restoreOptions -MultiSelect $true

    if ($null -eq $selectedItems -or $selectedItems.Count -eq 0) {
        Write-Warning "No items selected. Operation cancelled."
        return
    }

    # Step 3: Restore options
    Write-Section "Restore Options"

    $restoreOptionsMenu = @(
        @{ Name = "Restore to original locations"; Selected = $true },
        @{ Name = "Restore to alternate folder (for testing)"; Selected = $false },
        @{ Name = "Dry run (show what would happen, no actual copy)"; Selected = $false }
    )

    $selectedRestoreOptions = Show-InteractiveMenu -Title "Restore Options" -Items $restoreOptionsMenu -MultiSelect $true

    if ($null -eq $selectedRestoreOptions) {
        Write-Warning "Operation cancelled."
        return
    }

    $dryRun = ($selectedRestoreOptions | Where-Object { $_.Name -match "Dry run" }).Count -gt 0
    $useAlternateLocation = ($selectedRestoreOptions | Where-Object { $_.Name -match "alternate folder" }).Count -gt 0
    $alternatePath = $null

    if ($useAlternateLocation) {
        Write-Host ""
        $alternatePath = Read-Host "Enter alternate restore destination (e.g., C:\RestoreTest)"
        if (-not $alternatePath) {
            Write-Warning "No path entered. Operation cancelled."
            return
        }

        # Validate/create the alternate path
        if (-not (Test-Path $alternatePath)) {
            try {
                New-Item -Path $alternatePath -ItemType Directory -Force | Out-Null
                Write-Success "Created test folder: $alternatePath"
            } catch {
                Write-Error "Failed to create folder: $_"
                return
            }
        }

        # Update all destination paths to use alternate location
        foreach ($item in $selectedItems) {
            if ($item.Type -eq "UserData" -and $item.DestPath) {
                $item.OriginalDestPath = $item.DestPath
                $item.DestPath = Join-Path $alternatePath $item.FolderName
            }
        }
    }

    # Show summary
    Write-Header "RESTORE SUMMARY"

    if ($dryRun) {
        Write-Host "  *** DRY RUN MODE - No files will be copied ***" -ForegroundColor Magenta
        Write-Host ""
    }

    if ($useAlternateLocation) {
        Write-Host "  *** TEST MODE - Restoring to: $alternatePath ***" -ForegroundColor Cyan
        Write-Host ""
    }

    Write-Host "  Source: $backupPath" -ForegroundColor White
    Write-Host ""
    Write-Host "  Items to restore:" -ForegroundColor Yellow
    foreach ($item in $selectedItems) {
        Write-Host "    - $($item.Name)" -ForegroundColor White
        if ($item.DestPath) {
            Write-Host "      -> $($item.DestPath)" -ForegroundColor DarkGray
        }
    }

    Write-Host ""
    if (-not $dryRun -and -not $useAlternateLocation) {
        Write-Host "  [WARNING] Existing files will be backed up with .bak extension" -ForegroundColor Yellow
        Write-Host ""
    }

    $confirmMessage = if ($dryRun) { "Proceed with dry run?" } else { "Proceed with restore?" }
    if (-not (Show-Confirmation -Message $confirmMessage)) {
        Write-Warning "Restore cancelled."
        return
    }

    # Execute restore
    Write-Header "EXECUTING RESTORE"

    # Create logs folder
    $logsFolder = Join-Path $userProfile "Logs"
    if (-not (Test-Path $logsFolder)) {
        try {
            New-Item -Path $logsFolder -ItemType Directory -Force | Out-Null
            Write-Info "Created logs folder: $logsFolder"
        } catch {
            Write-Warning "Could not create logs folder: $logsFolder"
            $logsFolder = $env:TEMP
        }
    }

    # Use separate log files - script log (readable) and RoboCopy log (technical)
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $logFile = Join-Path $logsFolder "restore_$timestamp.log"
    $robocopyLogFile = Join-Path $logsFolder "restore_robocopy_$timestamp.log"

    # Initialize script log file with UTF8 encoding for readability
    "========================================" | Out-File -FilePath $logFile -Force -Encoding UTF8
    "Restore Log - $(Get-Date)" | Out-File -FilePath $logFile -Append -Encoding UTF8
    "========================================" | Out-File -FilePath $logFile -Append -Encoding UTF8
    "Backup source: $backupPath" | Out-File -FilePath $logFile -Append -Encoding UTF8
    "" | Out-File -FilePath $logFile -Append -Encoding UTF8

    $restoreResults = @()

    # Verify selected items have required properties
    Write-Host ""
    Write-Host "  Processing $($selectedItems.Count) items..." -ForegroundColor DarkGray

    foreach ($item in $selectedItems) {
        # Check for missing Type property (indicates menu didn't preserve properties)
        if (-not $item.Type) {
            Write-Warning "Item '$($item.Name)' is missing Type property - skipping"
            "ERROR: Item missing Type property: $($item.Name)" | Out-File -FilePath $logFile -Append -Encoding UTF8
            continue
        }

        switch ($item.Type) {
            "UserData" {
                $folderName = if ($item.FolderName) { $item.FolderName } else { $item.Name -replace "^User Data: ", "" }
                $destPath = $item.DestPath

                Write-Section "Restoring $folderName"

                # Validate we have required paths
                if (-not $item.Path) {
                    Write-Error "Source path is missing for $folderName"
                    "ERROR: Missing source path for $folderName" | Out-File -FilePath $logFile -Append -Encoding UTF8
                    $restoreResults += @{ Name = $folderName; Success = $false; Errors = @("Missing source path") }
                    continue
                }

                if (-not $destPath) {
                    Write-Error "Destination path is missing for $folderName"
                    "ERROR: Missing destination path for $folderName" | Out-File -FilePath $logFile -Append -Encoding UTF8
                    $restoreResults += @{ Name = $folderName; Success = $false; Errors = @("Missing destination path") }
                    continue
                }

                Write-Host "  Source:      $($item.Path)" -ForegroundColor DarkGray
                Write-Host "  Destination: $destPath" -ForegroundColor DarkGray
                "Processing: $folderName" | Out-File -FilePath $logFile -Append -Encoding UTF8
                "  Source: $($item.Path)" | Out-File -FilePath $logFile -Append -Encoding UTF8
                "  Destination: $destPath" | Out-File -FilePath $logFile -Append -Encoding UTF8

                if ($dryRun) {
                    # Count files that would be copied
                    $fileCount = (Get-ChildItem -LiteralPath $item.Path -Recurse -File -ErrorAction SilentlyContinue | Measure-Object).Count
                    $size = Get-FolderSize -Path $item.Path
                    Write-Info "[DRY RUN] Would copy $fileCount files ($(Format-Size $size))"
                    "[DRY RUN] $folderName : $fileCount files, $(Format-Size $size)" | Out-File -FilePath $logFile -Append -Encoding UTF8
                    $restoreResults += @{ Name = $folderName; Success = $true; DryRun = $true; FileCount = $fileCount }
                } else {
                    # Create destination directory if it doesn't exist
                    if (-not (Test-Path $destPath)) {
                        try {
                            New-Item -Path $destPath -ItemType Directory -Force | Out-Null
                            Write-Info "Created destination folder: $destPath"
                        } catch {
                            Write-Error "Failed to create destination folder: $_"
                            "ERROR: Failed to create destination: $_" | Out-File -FilePath $logFile -Append -Encoding UTF8
                            $restoreResults += @{ Name = $folderName; Success = $false; Errors = @("Failed to create destination: $_") }
                            continue
                        }
                    }

                    # Use separate RoboCopy log file to avoid encoding issues
                    $result = Invoke-RobocopyWithProgress -Source $item.Path -Destination $destPath -LogFile $robocopyLogFile -Mirror $false

                    if ($result.Success) {
                        Write-Success "$folderName restored successfully ($($result.FilesCopied) files, $(Format-Size $result.BytesCopied))"
                        "SUCCESS: $folderName - $($result.FilesCopied) files, $(Format-Size $result.BytesCopied)" | Out-File -FilePath $logFile -Append -Encoding UTF8
                        $restoreResults += @{ Name = $folderName; Success = $true; FilesCopied = $result.FilesCopied }
                    } else {
                        Write-Warning "$folderName restore had errors"
                        "WARNING: $folderName restore had errors" | Out-File -FilePath $logFile -Append -Encoding UTF8
                        foreach ($err in $result.Errors) {
                            Write-Host "    $err" -ForegroundColor Red
                            "  - $err" | Out-File -FilePath $logFile -Append -Encoding UTF8
                        }
                        $restoreResults += @{ Name = $folderName; Success = $false; Errors = $result.Errors }
                    }
                }
            }
            "AppConfigs" {
                Write-Section "Restoring Application Configurations"
                if ($dryRun) {
                    Write-Info "[DRY RUN] Would restore application configurations"
                    $restoreResults += @{ Name = "App Configs"; Success = $true; DryRun = $true }
                } else {
                    $configResults = Restore-AllAppConfigs -SourceFolder $item.Path
                    $restoreResults += @{ Name = "App Configs"; Success = $true; Details = $configResults }
                }
            }
            "TaskbarPins" {
                Write-Section "Restoring Taskbar Pins"
                if ($dryRun) {
                    $pinCount = (Get-ChildItem -Path $item.Path -Filter "*.lnk" -ErrorAction SilentlyContinue | Measure-Object).Count
                    Write-Info "[DRY RUN] Would restore $pinCount taskbar pins"
                    $restoreResults += @{ Name = "Taskbar Pins"; Success = $true; DryRun = $true }
                } else {
                    $result = Restore-TaskbarPins -SourceFolder $item.Path
                    if ($result.Success) {
                        Write-Success "Taskbar pins restored ($($result.Restored) items)"
                        $restoreResults += @{ Name = "Taskbar Pins"; Success = $true }
                    } else {
                        Write-Warning "Taskbar pins: $($result.Error)"
                        $restoreResults += @{ Name = "Taskbar Pins"; Success = $false }
                    }
                }
            }
            "Wallpaper" {
                Write-Section "Restoring Wallpaper"
                if ($dryRun) {
                    Write-Info "[DRY RUN] Would restore wallpaper from: $($item.Path)"
                    $restoreResults += @{ Name = "Wallpaper"; Success = $true; DryRun = $true }
                } else {
                    $result = Restore-Wallpaper -WallpaperPath $item.Path
                    if ($result.Success) {
                        Write-Success "Wallpaper restored and applied"
                        $restoreResults += @{ Name = "Wallpaper"; Success = $true }
                    } else {
                        Write-Warning "Wallpaper: $($result.Error)"
                        $restoreResults += @{ Name = "Wallpaper"; Success = $false }
                    }
                }
            }
        }
    }

    # Final summary
    if ($dryRun) {
        Write-Header "DRY RUN COMPLETE"
        Write-Host "  No files were actually copied." -ForegroundColor Magenta
        Write-Host ""
    } else {
        Write-Header "RESTORE COMPLETE"
    }

    Write-Host "  Results:" -ForegroundColor Yellow
    foreach ($result in $restoreResults) {
        $status = if ($result.Success) { "[OK]" } else { "[WARN]" }
        $color = if ($result.Success) { "Green" } else { "Yellow" }
        $suffix = if ($result.DryRun) { " (dry run)" } else { "" }
        Write-Host "    $status $($result.Name)$suffix" -ForegroundColor $color
    }

    # Write completion to log
    "" | Out-File -FilePath $logFile -Append -Encoding UTF8
    "========================================" | Out-File -FilePath $logFile -Append -Encoding UTF8
    "Restore completed: $(Get-Date)" | Out-File -FilePath $logFile -Append -Encoding UTF8
    "========================================" | Out-File -FilePath $logFile -Append -Encoding UTF8

    if (-not $dryRun) {
        Write-Host ""
        Write-Host "  Log files:" -ForegroundColor DarkGray
        Write-Host "    Script log:   $logFile" -ForegroundColor DarkGray
        Write-Host "    RoboCopy log: $robocopyLogFile" -ForegroundColor DarkGray
        Write-Host ""
        Write-Host "  Next Steps:" -ForegroundColor Cyan
        Write-Host "    1. Review any warnings above"

        $reinstallScript = Join-Path $backupPath "AppReinstall.ps1"
        if (Test-Path $reinstallScript) {
            Write-Host "    2. Run AppReinstall.ps1 to reinstall applications:"
            Write-Host "       $reinstallScript" -ForegroundColor DarkGray
            Write-Host "    3. After apps are installed, taskbar pins should work"
        }

        Write-Host "    4. Sign in to browsers to sync remaining data"
        Write-Host "    5. Restart if prompted by any applications"
    } else {
        Write-Host ""
        Write-Host "  Log file: $logFile" -ForegroundColor DarkGray
        Write-Host ""
        Write-Host "  To perform the actual restore, run again without dry run option." -ForegroundColor Cyan
    }
    Write-Host ""
}

#endregion RESTORE OPERATION

#region MAIN SCRIPT

# Clear screen and show header
Clear-Host
Write-Host ""
Write-Host "  ============================================" -ForegroundColor Cyan
Write-Host "       ENHANCED DATA MIGRATION TOOL v$script:Version" -ForegroundColor Cyan
Write-Host "  ============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Comprehensive backup and restore for PC migration" -ForegroundColor DarkGray
Write-Host "  Supports: User data, app configs, personalization" -ForegroundColor DarkGray
Write-Host ""

# Main menu
$mainMenuItems = @(
    @{ Name = "BACKUP - Save data from this PC"; Value = "backup" },
    @{ Name = "RESTORE - Restore data to this PC"; Value = "restore" },
    @{ Name = "EXIT"; Value = "exit" }
)

$selection = Show-InteractiveMenu -Title "Select Operation" -Items $mainMenuItems

if ($null -eq $selection -or $selection.Value -eq "exit") {
    Write-Host ""
    Write-Host "  Goodbye!" -ForegroundColor Cyan
    Write-Host ""
    exit 0
}

switch ($selection.Value) {
    "backup" {
        Invoke-BackupOperation
    }
    "restore" {
        Invoke-RestoreOperation
    }
}

Write-Host ""
Write-Host "Press any key to exit..." -ForegroundColor DarkGray
$null = [Console]::ReadKey($true)

#endregion MAIN SCRIPT
