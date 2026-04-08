#requires -version 5.1
<#
.SYNOPSIS
    Removes common Sysprep-blocking AppX packages, plus Microsoft.Zune* and Microsoft.WinGet* packages.

.DESCRIPTION
    - Removes installed AppX packages for all users when supported
    - Falls back to current-user and per-user SID removal
    - Removes matching provisioned packages
    - Targets:
        * Microsoft.Xbox.TCUI
        * Microsoft.XboxGameCallableUI
        * Microsoft.XboxGamingOverlay
        * Microsoft.XboxIdentityProvider
        * Microsoft.XboxSpeechToTextOverlay
        * Microsoft.BingWeather
        * Microsoft.Zune*
        * Microsoft.WinGet*
    - Logs to C:\Logs\Sysprep_Appx_Blockers_Removal.log

.NOTES
    Run as Administrator.
#>

[CmdletBinding()]
param(
    [string]$LogPath = 'C:\Logs\Sysprep_Appx_Blockers_Removal.log'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Ensure-Directory {
    param([Parameter(Mandatory = $true)][string]$Path)
    $dir = Split-Path -Path $Path -Parent
    if (-not [string]::IsNullOrWhiteSpace($dir) -and -not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
}

function Write-Log {
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR','SUCCESS')][string]$Level = 'INFO'
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[$timestamp] [$Level] $Message"
    Add-Content -Path $LogPath -Value $line

    switch ($Level) {
        'INFO'    { Write-Host $line }
        'WARN'    { Write-Host $line -ForegroundColor Yellow }
        'ERROR'   { Write-Host $line -ForegroundColor Red }
        'SUCCESS' { Write-Host $line -ForegroundColor Green }
    }
}

function Test-IsAdministrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Get-TargetPackagePatterns {
    return @(
        'Microsoft.Xbox.TCUI',
        'Microsoft.XboxGameCallableUI',
        'Microsoft.XboxGamingOverlay',
        'Microsoft.XboxIdentityProvider',
        'Microsoft.XboxSpeechToTextOverlay',
        'Microsoft.BingWeather',
        'Microsoft.Zune*',
        'Microsoft.WinGet*'
    )
}

function Get-InstalledPackagesByPattern {
    param([Parameter(Mandatory = $true)][string]$Pattern)

    $found = New-Object System.Collections.Generic.List[object]

    try {
        $pkgs = @(Get-AppxPackage -AllUsers | Where-Object {
            $_.Name -like $Pattern -or $_.PackageFullName -like "$Pattern*"
        })
    }
    catch {
        $pkgs = @()
    }

    foreach ($pkg in $pkgs) {
        if ($null -ne $pkg) {
            [void]$found.Add($pkg)
        }
    }

    return @($found | Group-Object PackageFullName | ForEach-Object { $_.Group[0] })
}

function Get-ProvisionedPackagesByPattern {
    param([Parameter(Mandatory = $true)][string]$Pattern)

    try {
        return @(Get-AppxProvisionedPackage -Online -ErrorAction Stop | Where-Object {
            $_.DisplayName -like $Pattern -or $_.PackageName -like "$Pattern*"
        })
    }
    catch {
        Write-Log "Unable to enumerate provisioned packages for pattern ${Pattern}: $($_.Exception.Message)" 'WARN'
        return @()
    }
}

function Remove-InstalledPackageSafe {
    param([Parameter(Mandatory = $true)]$Package)

    $full = $Package.PackageFullName
    Write-Log "Found installed package ${full}" 'INFO'

    try {
        Remove-AppxPackage -Package $full -AllUsers -ErrorAction Stop
        Write-Log "Removed package for all users ${full}" 'SUCCESS'
        return
    }
    catch {
        Write-Log "AllUsers removal failed for ${full}: $($_.Exception.Message)" 'WARN'
    }

    try {
        Remove-AppxPackage -Package $full -ErrorAction Stop
        Write-Log "Removed package for current user ${full}" 'SUCCESS'
    }
    catch {
        Write-Log "Current-user removal failed for ${full}: $($_.Exception.Message)" 'WARN'
    }

    try {
        foreach ($ui in @($Package.PackageUserInformation)) {
            $sid = $null
            try { $sid = [string]$ui.UserSecurityId.Sid } catch {}
            if ([string]::IsNullOrWhiteSpace($sid)) { continue }

            try {
                Remove-AppxPackage -Package $full -User $sid -ErrorAction Stop
                Write-Log "Removed package for SID ${sid}: ${full}" 'SUCCESS'
            }
            catch {
                Write-Log "Per-user removal failed for SID ${sid} on ${full}: $($_.Exception.Message)" 'WARN'
            }
        }
    }
    catch {
        Write-Log "Unable to enumerate PackageUserInformation for ${full}: $($_.Exception.Message)" 'WARN'
    }
}

function Remove-ProvisionedPackageSafe {
    param([Parameter(Mandatory = $true)]$ProvisionedPackage)

    try {
        Write-Log "Removing provisioned package $($ProvisionedPackage.PackageName)" 'INFO'
        Remove-AppxProvisionedPackage -Online -PackageName $ProvisionedPackage.PackageName -ErrorAction Stop | Out-Null
        Write-Log "Removed provisioned package $($ProvisionedPackage.PackageName)" 'SUCCESS'
    }
    catch {
        Write-Log "Provisioned removal failed for $($ProvisionedPackage.PackageName): $($_.Exception.Message)" 'WARN'
    }
}

function Test-RemainingStateForPattern {
    param([Parameter(Mandatory = $true)][string]$Pattern)

    $installed = @()
    $provisioned = @()

    try {
        $installed = @(Get-AppxPackage -AllUsers | Where-Object {
            $_.Name -like $Pattern -or $_.PackageFullName -like "$Pattern*"
        })
    }
    catch {}

    try {
        $provisioned = @(Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue | Where-Object {
            $_.DisplayName -like $Pattern -or $_.PackageName -like "$Pattern*"
        })
    }
    catch {}

    [pscustomobject]@{
        Pattern          = $Pattern
        InstalledCount   = @($installed).Count
        ProvisionedCount = @($provisioned).Count
    }
}

function Show-ValidationSummary {
    param([Parameter(Mandatory = $true)][object[]]$States)

    Write-Host ""
    Write-Host "==================== REMAINING APPX STATE ====================" -ForegroundColor Cyan

    foreach ($state in $States) {
        if ($state.InstalledCount -eq 0 -and $state.ProvisionedCount -eq 0) {
            Write-Host ("[CLEAR] {0} -> Installed={1}, Provisioned={2}" -f $state.Pattern, $state.InstalledCount, $state.ProvisionedCount) -ForegroundColor Green
        }
        else {
            Write-Host ("[REMAINING] {0} -> Installed={1}, Provisioned={2}" -f $state.Pattern, $state.InstalledCount, $state.ProvisionedCount) -ForegroundColor Yellow
        }
    }

    Write-Host "==============================================================" -ForegroundColor Cyan
    Write-Host ""
}

Ensure-Directory -Path $LogPath
Clear-Content -Path $LogPath -ErrorAction SilentlyContinue

try {
    if (-not (Test-IsAdministrator)) {
        throw 'This script must be run as Administrator.'
    }

    Write-Log 'Starting AppX cleanup for Sysprep blockers + Microsoft.Zune* + Microsoft.WinGet*' 'INFO'

    $patterns = Get-TargetPackagePatterns

    foreach ($pattern in $patterns) {
        Write-Log "Processing pattern ${pattern}" 'INFO'

        $installed = Get-InstalledPackagesByPattern -Pattern $pattern
        if (@($installed).Count -eq 0) {
            Write-Log "No installed packages found for pattern ${pattern}" 'INFO'
        }
        else {
            foreach ($pkg in $installed) {
                Remove-InstalledPackageSafe -Package $pkg
            }
        }

        $prov = Get-ProvisionedPackagesByPattern -Pattern $pattern
        if (@($prov).Count -eq 0) {
            Write-Log "No provisioned packages found for pattern ${pattern}" 'INFO'
        }
        else {
            foreach ($p in $prov) {
                Remove-ProvisionedPackageSafe -ProvisionedPackage $p
            }
        }
    }

    $states = @()
    foreach ($pattern in $patterns) {
        $states += Test-RemainingStateForPattern -Pattern $pattern
    }

    Show-ValidationSummary -States $states
    Write-Log 'Cleanup complete. Review remaining state above before running Sysprep.' 'SUCCESS'
}
catch {
    Write-Log "Script failed: $($_.Exception.Message)" 'ERROR'
    throw
}
