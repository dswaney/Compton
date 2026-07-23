<#
.SYNOPSIS
    Enables System Restore on Windows 11 and creates a restore point.

.NOTES
    ScriptVersion: 1.1.0
    Designed for Windows PowerShell 5.1 on Windows 11.
#>

[CmdletBinding()]
param(
    [ValidateNotNullOrEmpty()]
    [string]$RestorePointDescription = "Compton IT - Weekly Restore Point",

    [ValidateRange(1, 30)]
    [int]$VerificationTimeoutMinutes = 5
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$ScriptVersion = '1.1.0'
$LogDirectory = 'C:\Logs'
$LogPath = Join-Path $LogDirectory 'Enable-SystemRestore-And-Create-RestorePoint.log'

function Write-Log {
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [ValidateSet('INFO', 'OK', 'WARN', 'ERROR')]
        [string]$Level = 'INFO'
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = '[{0}] [{1,-5}] {2}' -f $timestamp, $Level, $Message

    switch ($Level) {
        'OK'    { Write-Host $line -ForegroundColor Green }
        'WARN'  { Write-Host $line -ForegroundColor Yellow }
        'ERROR' { Write-Host $line -ForegroundColor Red }
        default { Write-Host $line }
    }

    try {
        if (-not (Test-Path -LiteralPath $LogDirectory)) {
            New-Item -Path $LogDirectory -ItemType Directory -Force | Out-Null
        }

        Add-Content -LiteralPath $LogPath -Value $line -Encoding UTF8
    }
    catch {
        Write-Warning "Unable to write to the log file: $($_.Exception.Message)"
    }
}

function Test-IsAdministrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Invoke-SelfElevation {
    if (Test-IsAdministrator) {
        return
    }

    if (-not $PSCommandPath) {
        throw 'Administrative privileges are required. Open Windows PowerShell as Administrator and run the script again.'
    }

    Write-Host 'Administrative privileges are required. Requesting elevation...' -ForegroundColor Yellow

    $arguments = @(
        '-NoProfile'
        '-ExecutionPolicy Bypass'
        '-File "{0}"' -f $PSCommandPath
        '-RestorePointDescription "{0}"' -f $RestorePointDescription.Replace('"', '\"')
        '-VerificationTimeoutMinutes {0}' -f $VerificationTimeoutMinutes
    ) -join ' '

    Start-Process `
        -FilePath "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe" `
        -ArgumentList $arguments `
        -Verb RunAs | Out-Null

    exit 0
}

function Get-OperatingSystemDrive {
    $os = Get-CimInstance -ClassName Win32_OperatingSystem
    if (-not $os.SystemDrive) {
        throw 'Unable to determine the Windows operating-system drive.'
    }

    return ('{0}\' -f $os.SystemDrive.TrimEnd('\'))
}

function Enable-SystemRestoreProtection {
    param(
        [Parameter(Mandatory)]
        [string]$Drive
    )

    Write-Log "Ensuring System Restore is enabled for $Drive"

    Enable-ComputerRestore -Drive $Drive -ErrorAction Stop
    Write-Log "System Restore protection is enabled for $Drive" 'OK'
}

function Enable-RestorePointOnEveryRun {
    $registryPath = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SystemRestore'
    $valueName = 'SystemRestorePointCreationFrequency'

    if (-not (Test-Path -LiteralPath $registryPath)) {
        New-Item -Path $registryPath -Force | Out-Null
    }

    $currentValue = $null
    try {
        $currentValue = (Get-ItemProperty -LiteralPath $registryPath -Name $valueName -ErrorAction Stop).$valueName
    }
    catch {
        $currentValue = $null
    }

    if ($currentValue -ne 0) {
        New-ItemProperty `
            -LiteralPath $registryPath `
            -Name $valueName `
            -PropertyType DWord `
            -Value 0 `
            -Force | Out-Null

        Write-Log 'Configured Windows to allow a restore point on every script run.' 'OK'
    }
    else {
        Write-Log 'Restore-point creation frequency is already configured for every run.'
    }
}

function Initialize-ShadowCopyServices {
    foreach ($serviceName in @('VSS', 'swprv')) {
        $service = Get-Service -Name $serviceName -ErrorAction SilentlyContinue

        if (-not $service) {
            Write-Log "Service $serviceName was not found." 'WARN'
            continue
        }

        if ($service.Status -ne 'Running') {
            Write-Log "Starting required service $serviceName..."
            try {
                Start-Service -Name $serviceName -ErrorAction Stop
                $service.WaitForStatus(
                    [System.ServiceProcess.ServiceControllerStatus]::Running,
                    [TimeSpan]::FromSeconds(30)
                )
                Write-Log "Service $serviceName is running." 'OK'
            }
            catch {
                Write-Log "Service $serviceName could not be started: $($_.Exception.Message)" 'WARN'
            }
        }
        else {
            Write-Log "Required service $serviceName is already running."
        }
    }
}

function Get-LatestRestorePoint {
    try {
        return Get-ComputerRestorePoint -ErrorAction Stop |
            Sort-Object -Property CreationTime -Descending |
            Select-Object -First 1
    }
    catch {
        return $null
    }
}

function Convert-RestorePointTime {
    param(
        [Parameter(Mandatory)]
        [string]$CreationTime
    )

    try {
        return [Management.ManagementDateTimeConverter]::ToDateTime($CreationTime)
    }
    catch {
        return $null
    }
}


function Initialize-SystemRestoreNativeApi {
    if ('SystemRestore.NativeMethods' -as [type]) {
        return
    }

    $source = @'
using System;
using System.Runtime.InteropServices;

namespace SystemRestore
{
    public static class NativeMethods
    {
        [DllImport("SrClient.dll", SetLastError = true)]
        public static extern uint SRRemoveRestorePoint(uint restorePointSequenceNumber);
    }
}
'@

    Add-Type -TypeDefinition $source -Language CSharp -ErrorAction Stop
}

function Remove-RestorePointBySequenceNumber {
    param(
        [Parameter(Mandatory)]
        [uint32]$SequenceNumber
    )

    Initialize-SystemRestoreNativeApi

    $result = [SystemRestore.NativeMethods]::SRRemoveRestorePoint($SequenceNumber)
    if ($result -eq 0) {
        return
    }

    $message = (New-Object ComponentModel.Win32Exception([int]$result)).Message
    throw "SRRemoveRestorePoint failed for sequence $SequenceNumber. Win32Result=$result; Message=$message"
}

function Remove-ObsoleteRestorePoints {
    param(
        [Parameter(Mandatory)]
        [uint32]$CurrentSequenceNumber,

        [Nullable[uint32]]$PreviousSequenceNumber
    )

    $allRestorePoints = @(Get-ComputerRestorePoint -ErrorAction Stop |
        Sort-Object -Property SequenceNumber -Descending)

    $keep = New-Object 'System.Collections.Generic.HashSet[uint32]'
    [void]$keep.Add($CurrentSequenceNumber)

    if ($null -ne $PreviousSequenceNumber) {
        [void]$keep.Add([uint32]$PreviousSequenceNumber)
    }

    $toRemove = @($allRestorePoints | Where-Object {
        -not $keep.Contains([uint32]$_.SequenceNumber)
    })

    Write-Log (
        "Restore-point retention: Total={0}; Keeping={1}; Removing={2}" -f
        $allRestorePoints.Count,
        $keep.Count,
        $toRemove.Count
    )

    foreach ($restorePoint in $toRemove) {
        $created = Convert-RestorePointTime -CreationTime ([string]$restorePoint.CreationTime)
        $createdText = if ($created) { $created.ToString('yyyy-MM-dd HH:mm:ss') } else { 'Unknown' }

        Write-Log (
            "Deleting restore point: SequenceNumber={0}; Created={1}; Description={2}" -f
            $restorePoint.SequenceNumber,
            $createdText,
            $restorePoint.Description
        )

        try {
            Remove-RestorePointBySequenceNumber -SequenceNumber ([uint32]$restorePoint.SequenceNumber)
            Write-Log "Deleted restore point sequence $($restorePoint.SequenceNumber)." 'OK'
        }
        catch {
            Write-Log "Unable to delete restore point sequence $($restorePoint.SequenceNumber): $($_.Exception.Message)" 'WARN'
        }
    }

    $remaining = @(Get-ComputerRestorePoint -ErrorAction SilentlyContinue |
        Sort-Object -Property SequenceNumber -Descending)

    Write-Log "Restore-point retention completed. Remaining restore points: $($remaining.Count)" 'OK'

    foreach ($restorePoint in $remaining) {
        $created = Convert-RestorePointTime -CreationTime ([string]$restorePoint.CreationTime)
        $createdText = if ($created) { $created.ToString('yyyy-MM-dd HH:mm:ss') } else { 'Unknown' }
        Write-Log (
            "RESTORE_POINT_RETAINED|SequenceNumber={0}|Created={1}|Description={2}" -f
            $restorePoint.SequenceNumber,
            $createdText,
            $restorePoint.Description
        )
    }
}

function New-VerifiedRestorePoint {
    param(
        [Parameter(Mandatory)]
        [string]$Description,

        [Parameter(Mandatory)]
        [int]$TimeoutMinutes
    )

    $startedAt = Get-Date
    $safeDescription = $Description.Trim()

    if ($safeDescription.Length -gt 256) {
        $safeDescription = $safeDescription.Substring(0, 256)
    }

    Write-Log "Creating restore point: $safeDescription"

    Checkpoint-Computer `
        -Description $safeDescription `
        -RestorePointType MODIFY_SETTINGS `
        -ErrorAction Stop

    $deadline = (Get-Date).AddMinutes($TimeoutMinutes)

    do {
        Start-Sleep -Seconds 5
        $latest = Get-LatestRestorePoint

        if ($latest) {
            $creationDate = Convert-RestorePointTime -CreationTime ([string]$latest.CreationTime)

            if (
                $latest.Description -eq $safeDescription -and
                $creationDate -and
                $creationDate -ge $startedAt.AddMinutes(-1)
            ) {
                Write-Log (
                    "Restore point verified: SequenceNumber={0}; Created={1}; Description={2}" -f
                    $latest.SequenceNumber,
                    $creationDate.ToString('yyyy-MM-dd HH:mm:ss'),
                    $latest.Description
                ) 'OK'

                return $latest
            }
        }
    }
    while ((Get-Date) -lt $deadline)

    throw "The restore point could not be verified within $TimeoutMinutes minute(s)."
}

Invoke-SelfElevation

Write-Log "===== System Restore script v$ScriptVersion started ====="

try {
    $os = Get-CimInstance -ClassName Win32_OperatingSystem

    if ($os.ProductType -ne 1) {
        throw 'System Restore checkpoints are supported on Windows client operating systems, not Windows Server.'
    }

    $osDrive = Get-OperatingSystemDrive
    Write-Log "Detected operating-system drive: $osDrive"

    Enable-SystemRestoreProtection -Drive $osDrive
    Enable-RestorePointOnEveryRun
    Initialize-ShadowCopyServices

    $pointsBeforeCreation = @(Get-ComputerRestorePoint -ErrorAction SilentlyContinue |
        Sort-Object -Property SequenceNumber -Descending)

    $previousManagedPoint = $pointsBeforeCreation |
        Where-Object { $_.Description -eq $RestorePointDescription } |
        Select-Object -First 1

    if (-not $previousManagedPoint) {
        $previousManagedPoint = $pointsBeforeCreation | Select-Object -First 1
        if ($previousManagedPoint) {
            Write-Log (
                "No prior managed weekly restore point was found. Preserving newest existing restore point as the initial fallback: SequenceNumber={0}; Description={1}" -f
                $previousManagedPoint.SequenceNumber,
                $previousManagedPoint.Description
            ) 'WARN'
        }
    }

    $newRestorePoint = New-VerifiedRestorePoint `
        -Description $RestorePointDescription `
        -TimeoutMinutes $VerificationTimeoutMinutes

    $previousSequence = $null
    if ($previousManagedPoint) {
        $previousSequence = [Nullable[uint32]]([uint32]$previousManagedPoint.SequenceNumber)
    }

    Remove-ObsoleteRestorePoints `
        -CurrentSequenceNumber ([uint32]$newRestorePoint.SequenceNumber) `
        -PreviousSequenceNumber $previousSequence

    Write-Log 'System Restore is enabled, the new restore point was created, and retention cleanup completed.' 'OK'
    Write-Log '===== System Restore script completed successfully =====' 'OK'
    exit 0
}
catch {
    Write-Log "System Restore operation failed: $($_.Exception.Message)" 'ERROR'
    Write-Log '===== System Restore script completed with errors =====' 'ERROR'
    exit 1
}
