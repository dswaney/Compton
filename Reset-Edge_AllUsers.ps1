<#
====================================================================================================
Reset-Edge_AllUsers.ps1
Version:        1.1
Created:        2026-05-26
Updated:        2026-05-26
Purpose:        Clear Microsoft Edge cache, extension data, add-ons/extensions, and optional policies
                for all local user profiles.

Run As:         Administrator

Usage:
    .\Reset-Edge_AllUsers.ps1

Optional:
    .\Reset-Edge_AllUsers.ps1 -RemoveExtensionPolicies
    .\Reset-Edge_AllUsers.ps1 -FullEdgeProfileReset

Notes:
    - Default mode removes Edge cache, cookies, history, sessions, extension data, and installed extensions.
    - -RemoveExtensionPolicies also removes Edge extension force/allow/block policy keys.
    - -FullEdgeProfileReset removes the entire Edge User Data folder for each local user.
====================================================================================================
#>

[CmdletBinding()]
param(
    [switch]$RemoveExtensionPolicies,
    [switch]$FullEdgeProfileReset
)

$ScriptVersion = "1.1"
$StartTime = Get-Date
$LogRoot = "C:\Logs"

if (-not (Test-Path $LogRoot)) {
    New-Item -Path $LogRoot -ItemType Directory -Force | Out-Null
}

$LogFile = Join-Path $LogRoot ("Reset-Edge-AllUsers-{0}.log" -f (Get-Date -Format "yyyy-MM-dd_HH-mm-ss"))

function Write-Log {
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [ValidateSet("INFO", "OK", "WARN", "ERROR", "SECTION")]
        [string]$Level = "INFO"
    )

    $Time = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

    switch ($Level) {
        "SECTION" { $Color = "Magenta"; $Prefix = "[SECTION]" }
        "INFO"    { $Color = "Cyan";    $Prefix = "[INFO]" }
        "OK"      { $Color = "Green";   $Prefix = "[OK]" }
        "WARN"    { $Color = "Yellow";  $Prefix = "[WARN]" }
        "ERROR"   { $Color = "Red";     $Prefix = "[ERROR]" }
    }

    $Line = "$Time $Prefix $Message"

    Write-Host $Line -ForegroundColor $Color
    Add-Content -Path $LogFile -Value $Line
}

function Remove-SafeItem {
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter(Mandatory)]
        [string]$Description
    )

    if (-not (Test-Path $Path)) {
        Write-Log "${Description} does not exist: $Path" "INFO"
        return
    }

    try {
        Remove-Item -Path $Path -Recurse -Force -ErrorAction Stop
        Write-Log "Removed ${Description}: $Path" "OK"
    }
    catch {
        Write-Log "Failed removing ${Description}: $Path | $($_.Exception.Message)" "WARN"
    }
}

function Stop-ProcessSafe {
    param(
        [Parameter(Mandatory)]
        [string]$ProcessName
    )

    try {
        $Processes = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue

        if ($Processes) {
            $Processes | Stop-Process -Force -ErrorAction SilentlyContinue
            Write-Log "Stopped process: $ProcessName" "OK"
        }
        else {
            Write-Log "Process not running: $ProcessName" "INFO"
        }
    }
    catch {
        Write-Log "Could not stop process ${ProcessName}: $($_.Exception.Message)" "WARN"
    }
}

Clear-Host

Write-Host ""
Write-Host "============================================================" -ForegroundColor Magenta
Write-Host " Reset-Edge_AllUsers.ps1" -ForegroundColor White
Write-Host " Version $ScriptVersion" -ForegroundColor White
Write-Host "============================================================" -ForegroundColor Magenta
Write-Host ""

Write-Log "Starting Microsoft Edge cleanup for all users." "SECTION"
Write-Log "Log File: $LogFile" "INFO"

# --------------------------------------------------------------------------------------------------
# ADMIN CHECK
# --------------------------------------------------------------------------------------------------

$IsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltInRole]::Administrator
)

if (-not $IsAdmin) {
    Write-Log "This script must be run as Administrator." "ERROR"
    exit 1
}

# --------------------------------------------------------------------------------------------------
# STOP EDGE PROCESSES
# --------------------------------------------------------------------------------------------------

Write-Log "Stopping Microsoft Edge and WebView2 processes..." "SECTION"

$ProcessesToStop = @(
    "msedge",
    "msedgewebview2",
    "MicrosoftEdgeUpdate",
    "MicrosoftEdgeCP",
    "MicrosoftEdgeSH"
)

foreach ($ProcessName in $ProcessesToStop) {
    Stop-ProcessSafe -ProcessName $ProcessName
}

Start-Sleep -Seconds 3

# --------------------------------------------------------------------------------------------------
# ENUMERATE USERS
# --------------------------------------------------------------------------------------------------

Write-Log "Enumerating local user profiles..." "SECTION"

$ExcludedProfiles = @(
    "Public",
    "Default",
    "Default User",
    "All Users",
    "defaultuser0"
)

$UserProfiles = Get-ChildItem "C:\Users" -Directory -ErrorAction SilentlyContinue |
Where-Object {
    $_.Name -notin $ExcludedProfiles
}

Write-Log "Found $($UserProfiles.Count) user profile(s)." "INFO"

# --------------------------------------------------------------------------------------------------
# EDGE CLEANUP TARGETS
# --------------------------------------------------------------------------------------------------

$EdgeRelativeTargets = @(
    "Cache",
    "Code Cache",
    "GPUCache",
    "DawnCache",
    "GrShaderCache",
    "ShaderCache",
    "Service Worker",
    "Session Storage",
    "Local Storage",
    "IndexedDB",
    "Extension State",
    "Extensions",
    "History",
    "History-journal",
    "Cookies",
    "Cookies-journal",
    "Web Data",
    "Web Data-journal",
    "Login Data",
    "Login Data-journal",
    "Network",
    "Sessions",
    "Storage",
    "Crashpad",
    "Media Cache",
    "File System",
    "JumpListIconsRecentClosed",
    "JumpListIconsTopSites"
)

$RootTargets = @(
    "Crashpad",
    "ShaderCache",
    "GrShaderCache",
    "DawnCache",
    "BrowserMetrics",
    "CertificateRevocation"
)

# --------------------------------------------------------------------------------------------------
# CLEAN EACH USER PROFILE
# --------------------------------------------------------------------------------------------------

foreach ($UserProfile in $UserProfiles) {

    Write-Log "Processing user profile: $($UserProfile.Name)" "SECTION"

    $EdgeUserDataRoot = Join-Path $UserProfile.FullName "AppData\Local\Microsoft\Edge\User Data"

    if (-not (Test-Path $EdgeUserDataRoot)) {
        Write-Log "Edge User Data path not found for $($UserProfile.Name): $EdgeUserDataRoot" "INFO"
        continue
    }

    if ($FullEdgeProfileReset) {
        Remove-SafeItem -Path $EdgeUserDataRoot -Description "FULL Edge profile reset for $($UserProfile.Name)"
        continue
    }

    $EdgeProfileFolders = Get-ChildItem -Path $EdgeUserDataRoot -Directory -ErrorAction SilentlyContinue |
    Where-Object {
        $_.Name -eq "Default" -or
        $_.Name -like "Profile *" -or
        $_.Name -eq "Guest Profile"
    }

    foreach ($EdgeProfile in $EdgeProfileFolders) {

        Write-Log "Cleaning Edge profile: $($UserProfile.Name)\$($EdgeProfile.Name)" "INFO"

        foreach ($Target in $EdgeRelativeTargets) {
            $TargetPath = Join-Path $EdgeProfile.FullName $Target
            Remove-SafeItem -Path $TargetPath -Description "$Target for $($UserProfile.Name)\$($EdgeProfile.Name)"
        }
    }

    foreach ($RootTarget in $RootTargets) {
        $RootTargetPath = Join-Path $EdgeUserDataRoot $RootTarget
        Remove-SafeItem -Path $RootTargetPath -Description "$RootTarget root cache for $($UserProfile.Name)"
    }
}

# --------------------------------------------------------------------------------------------------
# REMOVE EDGE EXTENSION POLICIES IF REQUESTED
# --------------------------------------------------------------------------------------------------

if ($RemoveExtensionPolicies) {

    Write-Log "Removing Edge extension install policies..." "SECTION"

    $PolicyPaths = @(
        "HKLM:\SOFTWARE\Policies\Microsoft\Edge\ExtensionInstallForcelist",
        "HKLM:\SOFTWARE\Policies\Microsoft\Edge\ExtensionInstallAllowlist",
        "HKLM:\SOFTWARE\Policies\Microsoft\Edge\ExtensionInstallBlocklist",
        "HKLM:\SOFTWARE\WOW6432Node\Policies\Microsoft\Edge\ExtensionInstallForcelist",
        "HKLM:\SOFTWARE\WOW6432Node\Policies\Microsoft\Edge\ExtensionInstallAllowlist",
        "HKLM:\SOFTWARE\WOW6432Node\Policies\Microsoft\Edge\ExtensionInstallBlocklist"
    )

    foreach ($PolicyPath in $PolicyPaths) {
        Remove-SafeItem -Path $PolicyPath -Description "Edge extension policy registry key"
    }

    foreach ($UserProfile in $UserProfiles) {

        $TempHiveName = "TempEdgePolicy_$($UserProfile.Name -replace '[^a-zA-Z0-9]', '_')"
        $TempHivePath = "Registry::HKEY_USERS\$TempHiveName"
        $HiveLoaded = $false

        try {
            $NtUserDat = Join-Path $UserProfile.FullName "NTUSER.DAT"

            if (-not (Test-Path $NtUserDat)) {
                Write-Log "NTUSER.DAT not found for $($UserProfile.Name). Skipping per-user policies." "INFO"
                continue
            }

            reg.exe load "HKU\$TempHiveName" "$NtUserDat" | Out-Null
            $HiveLoaded = $true

            $UserPolicyPaths = @(
                "$TempHivePath\SOFTWARE\Policies\Microsoft\Edge\ExtensionInstallForcelist",
                "$TempHivePath\SOFTWARE\Policies\Microsoft\Edge\ExtensionInstallAllowlist",
                "$TempHivePath\SOFTWARE\Policies\Microsoft\Edge\ExtensionInstallBlocklist"
            )

            foreach ($UserPolicyPath in $UserPolicyPaths) {
                Remove-SafeItem -Path $UserPolicyPath -Description "per-user Edge extension policy for $($UserProfile.Name)"
            }
        }
        catch {
            Write-Log "Could not process per-user Edge policies for $($UserProfile.Name): $($_.Exception.Message)" "WARN"
        }
        finally {
            if ($HiveLoaded) {
                try {
                    [gc]::Collect()
                    Start-Sleep -Milliseconds 500
                    reg.exe unload "HKU\$TempHiveName" | Out-Null
                }
                catch {
                    Write-Log "Could not unload temporary registry hive HKU\${TempHiveName}: $($_.Exception.Message)" "WARN"
                }
            }
        }
    }
}
else {
    Write-Log "Skipping Edge extension policy removal. Use -RemoveExtensionPolicies to enable it." "WARN"
}

# --------------------------------------------------------------------------------------------------
# SUMMARY
# --------------------------------------------------------------------------------------------------

$Runtime = New-TimeSpan -Start $StartTime -End (Get-Date)

Write-Host ""
Write-Host "============================================================" -ForegroundColor Magenta
Write-Host " Edge Cleanup Completed" -ForegroundColor White
Write-Host "============================================================" -ForegroundColor Magenta
Write-Host ""

Write-Log ("Total Runtime: {0:hh\:mm\:ss}" -f $Runtime) "INFO"
Write-Log "Log File: $LogFile" "INFO"

Write-Host ""
Write-Host "Recommended: restart the computer before testing Edge/Explorer again." -ForegroundColor Yellow
Write-Host ""
