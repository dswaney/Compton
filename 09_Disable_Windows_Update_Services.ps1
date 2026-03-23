# ScriptVersion: 1.0
# LastUpdated: 2026-03-23

# Purpose: Disable Windows Update background services and scheduled tasks on Windows 11

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

function Stop-AndDisableService {
    param(
        [Parameter(Mandatory)]
        [string]$Name
    )

    try {
        $svc = Get-Service -Name $Name -ErrorAction Stop

        if ($svc.Status -ne 'Stopped') {
            try {
                Stop-Service -Name $Name -Force -ErrorAction Stop
                Write-Status "Stopped service: $Name" 'OK'
            }
            catch {
                Write-Status "Could not stop service $Name with Stop-Service. Trying sc.exe..." 'WARN'
                & sc.exe stop $Name | Out-Null
                Start-Sleep -Seconds 2
            }
        }
        else {
            Write-Status "Service already stopped: $Name" 'INFO'
        }

        try {
            Set-Service -Name $Name -StartupType Disabled -ErrorAction Stop
            Write-Status "Disabled startup type for service: $Name" 'OK'
        }
        catch {
            Write-Status "Set-Service could not disable $Name. Trying sc.exe config..." 'WARN'
            & sc.exe config $Name start= disabled | Out-Null
            Write-Status "Applied disabled startup type via sc.exe: $Name" 'OK'
        }
    }
    catch {
        Write-Status "Service not found or inaccessible: $Name" 'WARN'
    }
}

function Set-ServiceStartRegistry {
    param(
        [Parameter(Mandatory)]
        [string]$ServiceName,

        [Parameter(Mandatory)]
        [ValidateSet(2,3,4)]
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
    }
}

function Disable-ScheduledTaskSafe {
    param(
        [Parameter(Mandatory)]
        [string]$TaskPath,

        [Parameter(Mandatory)]
        [string]$TaskName
    )

    try {
        $task = Get-ScheduledTask -TaskPath $TaskPath -TaskName $TaskName -ErrorAction Stop
        if ($task.State -ne 'Disabled') {
            Disable-ScheduledTask -TaskPath $TaskPath -TaskName $TaskName -ErrorAction Stop | Out-Null
            Write-Status "Disabled scheduled task: $TaskPath$TaskName" 'OK'
        }
        else {
            Write-Status "Scheduled task already disabled: $TaskPath$TaskName" 'INFO'
        }
    }
    catch {
        Write-Status "Scheduled task not found or could not be disabled: $TaskPath$TaskName" 'WARN'
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

# Main update-related services
$servicesToDisable = @(
    'wuauserv',    # Windows Update
    'bits',        # Background Intelligent Transfer Service
    'dosvc',       # Delivery Optimization
    'UsoSvc'       # Update Orchestrator Service
)

foreach ($svc in $servicesToDisable) {
    Stop-AndDisableService -Name $svc
}

# Windows Update Medic Service usually needs registry change as well
Set-ServiceStartRegistry -ServiceName 'WaaSMedicSvc' -StartValue 4
try {
    & sc.exe stop WaaSMedicSvc | Out-Null
    Write-Status "Attempted to stop WaaSMedicSvc" 'INFO'
}
catch {
    Write-Status "Could not stop WaaSMedicSvc directly" 'WARN'
}

# Also enforce registry startup state for the others
Set-ServiceStartRegistry -ServiceName 'wuauserv'    -StartValue 4
Set-ServiceStartRegistry -ServiceName 'bits'        -StartValue 4
Set-ServiceStartRegistry -ServiceName 'dosvc'       -StartValue 4
Set-ServiceStartRegistry -ServiceName 'UsoSvc'      -StartValue 4

# Group Policy-style registry settings to turn off automatic updates
$wuPolicyPath = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate'
$auPolicyPath = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU'

Set-RegistryDwordSafe -Path $auPolicyPath -Name 'NoAutoUpdate' -Value 1
Set-RegistryDwordSafe -Path $auPolicyPath -Name 'AUOptions' -Value 1

# Disable common update scheduled tasks
$tasks = @(
    @{ Path = '\Microsoft\Windows\WindowsUpdate\';          Name = 'Scheduled Start' },
    @{ Path = '\Microsoft\Windows\UpdateOrchestrator\';     Name = 'Schedule Scan' },
    @{ Path = '\Microsoft\Windows\UpdateOrchestrator\';     Name = 'Schedule Scan Static Task' },
    @{ Path = '\Microsoft\Windows\UpdateOrchestrator\';     Name = 'USO_UxBroker' },
    @{ Path = '\Microsoft\Windows\UpdateOrchestrator\';     Name = 'Reboot' },
    @{ Path = '\Microsoft\Windows\UpdateOrchestrator\';     Name = 'Maintenance Install' },
    @{ Path = '\Microsoft\Windows\UpdateOrchestrator\';     Name = 'Refresh Settings' },
    @{ Path = '\Microsoft\Windows\WaaSMedic\';              Name = 'PerformRemediation' }
)

foreach ($task in $tasks) {
    Disable-ScheduledTaskSafe -TaskPath $task.Path -TaskName $task.Name
}

Write-Status "Windows Update background components have been disabled as much as possible." 'OK'
Write-Status "A reboot is recommended." 'INFO'