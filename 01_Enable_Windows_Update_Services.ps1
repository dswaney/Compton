# =====================================================================
# ScriptName: 01_Enable_Windows_Update_Services.ps1
# ScriptVersion: 1.4
# LastUpdated: 2026-04-16
# Purpose: Restore Windows Update services, tasks, and policy settings
#          on Windows 11, verify required services are running, retry
#          startup failures up to 4 total attempts, and force a reboot
#          if critical services still refuse to start.
# =====================================================================

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

function Write-Status {
    param(
        [string]$Message,
        [ValidateSet('INFO','OK','WARN','ERROR')]
        [string]$Level = 'INFO'
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    switch ($Level) {
        'INFO'  { Write-Host "[$timestamp] [INFO ] $Message" -ForegroundColor Cyan }
        'OK'    { Write-Host "[$timestamp] [ OK  ] $Message" -ForegroundColor Green }
        'WARN'  { Write-Host "[$timestamp] [WARN ] $Message" -ForegroundColor Yellow }
        'ERROR' { Write-Host "[$timestamp] [ERROR] $Message" -ForegroundColor Red }
    }
}

function Test-IsAdmin {
    $currentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentIdentity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Set-ServiceStartRegistry {
    param(
        [Parameter(Mandatory)]
        [string]$ServiceName,

        [Parameter(Mandatory)]
        [int]$StartValue
    )

    $paths = @(
        "HKLM:\SYSTEM\CurrentControlSet\Services\$ServiceName",
        "HKLM:\SYSTEM\ControlSet001\Services\$ServiceName"
    )

    foreach ($path in $paths) {
        if (Test-Path $path) {
            try {
                Set-ItemProperty -Path $path -Name Start -Value $StartValue -Type DWord -ErrorAction Stop
                Write-Status "Set registry Start=$StartValue for $ServiceName at $path" 'OK'
            }
            catch {
                Write-Status "Failed setting Start for $ServiceName at $path : $($_.Exception.Message)" 'WARN'
            }
        }
        else {
            Write-Status "Registry path not found for $ServiceName at $path" 'WARN'
        }
    }
}

function Set-ServiceStartupAndStart {
    param(
        [Parameter(Mandatory)]
        [string]$Name,

        [Parameter(Mandatory)]
        [ValidateSet('Automatic','Manual')]
        [string]$StartupType
    )

    try {
        $svc = Get-Service -Name $Name -ErrorAction Stop

        try {
            Set-Service -Name $Name -StartupType $StartupType -ErrorAction Stop
            Write-Status "Set startup type for $Name to $StartupType" 'OK'
        }
        catch {
            Write-Status "Set-Service failed for $Name. Trying sc.exe config..." 'WARN'
            $startValue = if ($StartupType -eq 'Automatic') { 'auto' } else { 'demand' }
            & sc.exe config $Name start= $startValue | Out-Null
            Write-Status "Configured startup type for $Name via sc.exe" 'OK'
        }

        try {
            Start-Service -Name $Name -ErrorAction Stop
            Write-Status "Started service: $Name" 'OK'
        }
        catch {
            Write-Status "Could not start service $Name immediately: $($_.Exception.Message)" 'WARN'
        }
    }
    catch {
        Write-Status "Service not found or inaccessible: $Name" 'WARN'
    }
}

function Get-ServiceStateSafe {
    param(
        [Parameter(Mandatory)]
        [string]$Name
    )

    try {
        return (Get-Service -Name $Name -ErrorAction Stop).Status
    }
    catch {
        return $null
    }
}

function Wait-ForServiceRunning {
    param(
        [Parameter(Mandatory)]
        [string]$Name,

        [int]$TimeoutSeconds = 15
    )

    $stopWatch = [System.Diagnostics.Stopwatch]::StartNew()

    while ($stopWatch.Elapsed.TotalSeconds -lt $TimeoutSeconds) {
        $status = Get-ServiceStateSafe -Name $Name
        if ($status -eq 'Running') {
            return $true
        }

        Start-Sleep -Seconds 2
    }

    return $false
}

function Ensure-ServiceRunningWithRetry {
    param(
        [Parameter(Mandatory)]
        [string]$Name,

        [Parameter(Mandatory)]
        [ValidateSet('Automatic','Manual')]
        [string]$StartupType,

        [int]$MaxAttempts = 4,

        [int]$WaitPerAttemptSeconds = 15
    )

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        $currentState = Get-ServiceStateSafe -Name $Name

        if ($currentState -eq 'Running') {
            Write-Status "Service $Name is already running." 'OK'
            return $true
        }

        Write-Status "Attempt $attempt of $MaxAttempts to start service $Name..." 'INFO'

        try {
            Set-ServiceStartupAndStart -Name $Name -StartupType $StartupType
        }
        catch {
            Write-Status "Unexpected error while attempting to start $Name : $($_.Exception.Message)" 'WARN'
        }

        if (Wait-ForServiceRunning -Name $Name -TimeoutSeconds $WaitPerAttemptSeconds) {
            Write-Status "Verified service is running: $Name" 'OK'
            return $true
        }

        $stateAfterWait = Get-ServiceStateSafe -Name $Name
        Write-Status "Service $Name did not reach Running state after attempt $attempt. Current state: $stateAfterWait" 'WARN'

        if ($attempt -lt $MaxAttempts) {
            Start-Sleep -Seconds 5
        }
    }

    Write-Status "Service $Name failed to reach Running state after $MaxAttempts attempts." 'ERROR'
    return $false
}

function Force-RebootNow {
    param(
        [string]$Reason = 'Required Windows Update services failed to start after multiple attempts.'
    )

    Write-Status "FORCING REBOOT: $Reason" 'ERROR'

    try {
        shutdown.exe /r /f /t 30 /c "$Reason" | Out-Null
        Write-Status "Forced reboot command issued successfully. System will restart in 30 seconds." 'ERROR'
    }
    catch {
        Write-Status "Failed to issue shutdown.exe reboot command: $($_.Exception.Message)" 'ERROR'
    }

    exit 1
}

function Enable-ScheduledTaskSafe {
    param(
        [Parameter(Mandatory)]
        [string]$TaskPath,

        [Parameter(Mandatory)]
        [string]$TaskName
    )

    try {
        $task = Get-ScheduledTask -TaskPath $TaskPath -TaskName $TaskName -ErrorAction Stop
        if ($task.State -eq 'Disabled') {
            Enable-ScheduledTask -TaskPath $TaskPath -TaskName $TaskName -ErrorAction Stop | Out-Null
            Write-Status "Enabled scheduled task: $TaskPath$TaskName" 'OK'
        }
        else {
            Write-Status "Scheduled task already enabled or available: $TaskPath$TaskName" 'INFO'
        }
    }
    catch {
        Write-Status "Scheduled task not found or could not be enabled: $TaskPath$TaskName" 'WARN'
    }
}

function Remove-RegistryValueSafe {
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter(Mandatory)]
        [string]$Name
    )

    try {
        if (Test-Path $Path) {
            $prop = Get-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue
            if ($null -ne $prop) {
                Remove-ItemProperty -Path $Path -Name $Name -ErrorAction Stop
                Write-Status "Removed $Path\$Name" 'OK'
            }
            else {
                Write-Status "Registry value not present: $Path\$Name" 'INFO'
            }
        }
        else {
            Write-Status "Registry path not present: $Path" 'INFO'
        }
    }
    catch {
        Write-Status "Failed to remove $Path\$Name : $($_.Exception.Message)" 'WARN'
    }
}

function Set-RegistryDwordSafe {
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter(Mandatory)]
        [string]$Name,

        [Parameter(Mandatory)]
        [int]$Value
    )

    try {
        if (-not (Test-Path $Path)) {
            New-Item -Path $Path -Force | Out-Null
        }

        New-ItemProperty -Path $Path -Name $Name -PropertyType DWord -Value $Value -Force | Out-Null
        Write-Status "Set $Path\$Name = $Value" 'OK'
    }
    catch {
        Write-Status "Failed to set $Path\$Name : $($_.Exception.Message)" 'ERROR'
    }
}

if (-not (Test-IsAdmin)) {
    Write-Host ""
    Write-Host "This script must be run as Administrator." -ForegroundColor Red
    exit 1
}

Write-Status "Initializing script..." 'INFO'
# Restore registry startup values first
# Common defaults used for Windows Update-related services:
# wuauserv = Manual (3)
# bits = Manual (3)
# dosvc = Automatic family (2)
# UsoSvc = Automatic (2)
# WaaSMedicSvc = Manual/triggered on many systems (3)

Set-ServiceStartRegistry -ServiceName 'wuauserv'     -StartValue 3
Set-ServiceStartRegistry -ServiceName 'bits'         -StartValue 3
Set-ServiceStartRegistry -ServiceName 'dosvc'        -StartValue 2
Set-ServiceStartRegistry -ServiceName 'UsoSvc'       -StartValue 2
Set-ServiceStartRegistry -ServiceName 'WaaSMedicSvc' -StartValue 3

# Initial restore and startup
Set-ServiceStartupAndStart -Name 'wuauserv' -StartupType Manual
Set-ServiceStartupAndStart -Name 'bits'     -StartupType Manual
Set-ServiceStartupAndStart -Name 'dosvc'    -StartupType Automatic
Set-ServiceStartupAndStart -Name 'UsoSvc'   -StartupType Automatic

# WaaSMedicSvc can be protected; set registry above, then try starting via sc.exe
try {
    & sc.exe config WaaSMedicSvc start= demand | Out-Null
    Write-Status "Configured WaaSMedicSvc startup via sc.exe" 'OK'
}
catch {
    Write-Status "Could not configure WaaSMedicSvc via sc.exe" 'WARN'
}

try {
    & sc.exe start WaaSMedicSvc | Out-Null
    Write-Status "Attempted to start WaaSMedicSvc" 'INFO'
}
catch {
    Write-Status "Could not start WaaSMedicSvc directly" 'WARN'
}

# Restore Automatic Updates policy
$wuPolicyPath = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate'
$auPolicyPath = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU'

Set-RegistryDwordSafe -Path $auPolicyPath -Name 'NoAutoUpdate' -Value 0
Set-RegistryDwordSafe -Path $auPolicyPath -Name 'AUOptions'    -Value 3

# Remove common WSUS redirection values if they were previously set
Remove-RegistryValueSafe -Path $wuPolicyPath -Name 'WUServer'
Remove-RegistryValueSafe -Path $wuPolicyPath -Name 'WUStatusServer'
Remove-RegistryValueSafe -Path $wuPolicyPath -Name 'UpdateServiceUrlAlternate'
Remove-RegistryValueSafe -Path $wuPolicyPath -Name 'SetProxyBehaviorForUpdateDetection'
Remove-RegistryValueSafe -Path $wuPolicyPath -Name 'DisableWindowsUpdateAccess'
Remove-RegistryValueSafe -Path $auPolicyPath -Name 'UseWUServer'

# Re-enable common update scheduled tasks
$tasks = @(
    @{ Path = '\Microsoft\Windows\WindowsUpdate\';      Name = 'Scheduled Start' },
    @{ Path = '\Microsoft\Windows\UpdateOrchestrator\'; Name = 'Schedule Scan' },
    @{ Path = '\Microsoft\Windows\UpdateOrchestrator\'; Name = 'Schedule Scan Static Task' },
    @{ Path = '\Microsoft\Windows\UpdateOrchestrator\'; Name = 'USO_UxBroker' },
    @{ Path = '\Microsoft\Windows\UpdateOrchestrator\'; Name = 'Reboot' },
    @{ Path = '\Microsoft\Windows\UpdateOrchestrator\'; Name = 'Maintenance Install' },
    @{ Path = '\Microsoft\Windows\UpdateOrchestrator\'; Name = 'Refresh Settings' },
    @{ Path = '\Microsoft\Windows\WaaSMedic\';          Name = 'PerformRemediation' }
)

foreach ($task in $tasks) {
    Enable-ScheduledTaskSafe -TaskPath $task.Path -TaskName $task.Name
}

# Restart key services in a sensible order
$restartOrder = @('bits', 'dosvc', 'wuauserv', 'UsoSvc')
foreach ($svc in $restartOrder) {
    try {
        Restart-Service -Name $svc -Force -ErrorAction Stop
        Write-Status "Restarted service: $svc" 'OK'
    }
    catch {
        Write-Status "Could not restart $svc : $($_.Exception.Message)" 'WARN'
    }
}

# Verify and retry critical services
$requiredServices = @(
    @{ Name = 'wuauserv'; StartupType = 'Manual' },
    @{ Name = 'bits';     StartupType = 'Manual' },
    @{ Name = 'dosvc';    StartupType = 'Automatic' },
    @{ Name = 'UsoSvc';   StartupType = 'Automatic' }
)

$failedServices = @()

foreach ($requiredService in $requiredServices) {
    $serviceStarted = Ensure-ServiceRunningWithRetry -Name $requiredService.Name -StartupType $requiredService.StartupType -MaxAttempts 4 -WaitPerAttemptSeconds 15
    if (-not $serviceStarted) {
        $failedServices += $requiredService.Name
    }
}

if ($failedServices.Count -gt 0) {
    $failedList = $failedServices -join ', '
    Write-Status "One or more critical Windows Update services failed to start: $failedList" 'ERROR'
    Force-RebootNow -Reason "Windows Update service recovery failed. Services not running: $failedList"
}

Write-Status "Windows Update settings have been restored and critical services are running." 'OK'
Write-Status "No reboot required. Continuing normally." 'INFO'
