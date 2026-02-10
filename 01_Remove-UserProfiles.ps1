<#
.SYNOPSIS
    Deletes all user profiles under C:\Users except for Default, Public, MISAdmin, and other exclusions.

.DESCRIPTION
    This script enumerates all Win32_UserProfile instances whose LocalPath begins with C:\Users. 
    Any profile whose folder name is not in the exclusion list will be passed to Remove-CimInstance, 
    which calls Delete() on the underlying WMI object. It includes:
      • A check to ensure the script is running elevated.
      • An exclusion list: Default, Public, MISAdmin, dvswaney (and any others you add).
      • Error handling/logging for each deletion attempt.

.NOTES
    • Run this script as Administrator.
    • If a profile is in use or protected, removal will fail and be logged as a warning.
    • Adjust or expand the $ExcludedProfiles array to preserve additional profiles.
#>

# Ensure the script is running as Administrator
If (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
        [Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Error "Please run this script as Administrator."
    exit 1
}

# Define the profiles to keep (folder names under C:\Users)
$ExcludedProfiles = @(
    'Default',
    'Public',
    'MISAdmin',
    'dswaney'
)

Write-Host "Starting profile cleanup..." -ForegroundColor Cyan

# Retrieve all user profiles from Win32_UserProfile whose LocalPath is under C:\Users
$AllUserProfiles = Get-CimInstance -Class Win32_UserProfile | Where-Object {
    $_.LocalPath -like 'C:\Users\*'
}

foreach ($Profile in $AllUserProfiles) {
    $ProfilePath = $Profile.LocalPath
    $ProfileName = Split-Path $ProfilePath -Leaf

    # Skip any profile whose folder name matches an excluded profile
    if ($ExcludedProfiles -contains $ProfileName) {
        Write-Host "Skipping excluded profile: $ProfileName" -ForegroundColor Yellow
        continue
    }

    # Attempt deletion via Remove-CimInstance (calls Delete() under the hood)
    try {
        Write-Host "Deleting profile: $ProfileName ($ProfilePath)..." -ForegroundColor Cyan
        Remove-CimInstance -InputObject $Profile -ErrorAction Stop
        Write-Host "✔ Successfully deleted profile: $ProfileName" -ForegroundColor Green
    }
    catch {
        Write-Warning "⚠ Failed to delete profile: $ProfileName. Error: $_"
    }
}

Write-Host "`nProfile cleanup COMPLETE." -ForegroundColor Cyan
