# ScriptVersion: 1.1
# LastUpdated: 2026-03-23

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
# dosvc = Automatic Delayed/Automatic family; registry restore to Automatic (2)
# UsoSvc = Automatic (2)
# WaaSMedicSvc = Manual/triggered on many systems; registry restore to Manual (3)

Set-ServiceStartRegistry -ServiceName 'wuauserv'    -StartValue 3
Set-ServiceStartRegistry -ServiceName 'bits'        -StartValue 3
Set-ServiceStartRegistry -ServiceName 'dosvc'       -StartValue 2
Set-ServiceStartRegistry -ServiceName 'UsoSvc'      -StartValue 2
Set-ServiceStartRegistry -ServiceName 'WaaSMedicSvc' -StartValue 3

# Restore service startup and start the key update services
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

# Either remove the policy values entirely or set them to enabled defaults.
# This script sets explicit values for clarity:
Set-RegistryDwordSafe -Path $auPolicyPath -Name 'NoAutoUpdate' -Value 0
Set-RegistryDwordSafe -Path $auPolicyPath -Name 'AUOptions' -Value 3

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

Write-Status "Windows Update settings have been restored." 'OK'
Write-Status "A reboot is recommended, then check Settings > Windows Update." 'INFO'
