# ScriptVersion: 2.0
# LastUpdated: 2026-04-16

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Low')]
param(
    [string]$ScriptsRoot = 'C:\Scripts',
    [string]$LogPath = "$env:SystemDrive\Temp\Register-Tasks_SYSTEM.log",
    [switch]$IncludeMicrosoftTasks = $false
)

$ErrorActionPreference = 'Stop'
$WindowsPowerShellExe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','OK','WARN','ERROR')]
        [string]$Level = 'INFO'
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[$timestamp] [$('{0,-5}' -f $Level)] $Message"

    switch ($Level) {
        'INFO'  { Write-Host $line -ForegroundColor Cyan }
        'OK'    { Write-Host $line -ForegroundColor Green }
        'WARN'  { Write-Host $line -ForegroundColor Yellow }
        'ERROR' { Write-Host $line -ForegroundColor Red }
    }

    try {
        $logDir = Split-Path -Path $LogPath -Parent
        if (-not [string]::IsNullOrWhiteSpace($logDir) -and -not (Test-Path -LiteralPath $logDir)) {
            New-Item -Path $logDir -ItemType Directory -Force | Out-Null
        }
        Add-Content -Path $LogPath -Value $line -Encoding UTF8
    }
    catch {
    }
}

function Test-IsAdministrator {
    try {
        $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object Security.Principal.WindowsPrincipal($identity)
        return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    }
    catch {
        return $false
    }
}

function Get-TaskPrincipal {
    return New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
}

function Get-TaskSettings {
    return New-ScheduledTaskSettingsSet `
        -AllowStartIfOnBatteries `
        -DontStopIfGoingOnBatteries `
        -StartWhenAvailable `
        -ExecutionTimeLimit (New-TimeSpan -Hours 12)
}

function Remove-ExistingScheduledTasks {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [switch]$DeleteMicrosoftTasks
    )

    try {
        Import-Module ScheduledTasks -ErrorAction Stop | Out-Null
    }
    catch {
        throw "The ScheduledTasks module is not available. $($_.Exception.Message)"
    }

    $tasks = Get-ScheduledTask | Where-Object {
        if ($DeleteMicrosoftTasks) {
            $true
        }
        else {
            $_.TaskPath -notlike '\Microsoft\*'
        }
    } | Sort-Object TaskPath, TaskName

    if (-not $tasks) {
        Write-Log 'No existing scheduled tasks matched the deletion scope.' 'INFO'
        return
    }

    Write-Log "Removing $($tasks.Count) existing scheduled task(s)." 'WARN'

    foreach ($task in $tasks) {
        $fullName = "$($task.TaskPath)$($task.TaskName)"

        if ($PSCmdlet.ShouldProcess($fullName, 'Unregister scheduled task')) {
            try {
                Unregister-ScheduledTask -TaskName $task.TaskName -TaskPath $task.TaskPath -Confirm:$false -ErrorAction Stop
                Write-Log "Removed scheduled task: $fullName" 'OK'
            }
            catch {
                Write-Log "Failed to remove scheduled task ${fullName}: $($_.Exception.Message)" 'WARN'
            }
        }
    }
}

function Register-WeeklyPowerShellTask {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory)][string]$TaskName,
        [Parameter(Mandatory)][string]$ScriptPath,
        [Parameter(Mandatory)][string]$StartTime,
        [string]$ExtraArguments = ''
    )

    if (-not (Test-Path -LiteralPath $WindowsPowerShellExe)) {
        throw "Windows PowerShell executable not found: $WindowsPowerShellExe"
    }

    if (-not (Test-Path -LiteralPath $ScriptPath)) {
        Write-Log "Script path does not currently exist, but the task will still be created: $ScriptPath" 'WARN'
    }

    $argumentString = "-NoProfile -ExecutionPolicy Bypass -File `"$ScriptPath`""
    if (-not [string]::IsNullOrWhiteSpace($ExtraArguments)) {
        $argumentString = "$argumentString $ExtraArguments"
    }

    $displayCommand = "$WindowsPowerShellExe $argumentString"
    Write-Log "Registering task '$TaskName' with action: $displayCommand" 'INFO'

    try {
        $action = New-ScheduledTaskAction -Execute $WindowsPowerShellExe -Argument $argumentString
        $startAt = [datetime]::Today.Add([timespan]::Parse($StartTime))
        $trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Sunday -WeeksInterval 1 -At $startAt
        $principal = Get-TaskPrincipal
        $settings = Get-TaskSettings

        if ($PSCmdlet.ShouldProcess($TaskName, 'Register weekly scheduled task')) {
            Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Force | Out-Null
            Write-Log "Created task '$TaskName' for Sunday at $StartTime." 'OK'
        }
    }
    catch {
        throw "Failed to create weekly task '$TaskName'. $($_.Exception.Message)"
    }
}

function Ensure-TimeSyncHelperScript {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ScriptsDirectory
    )

    try {
        if (-not (Test-Path -LiteralPath $ScriptsDirectory)) {
            New-Item -Path $ScriptsDirectory -ItemType Directory -Force | Out-Null
            Write-Log "Created scripts directory: $ScriptsDirectory" 'OK'
        }

        $helperPath = Join-Path $ScriptsDirectory '10_Sync_System_Time.ps1'
        $helperContent = @'
$ErrorActionPreference = 'SilentlyContinue'

try {
    Start-Service -Name 'w32time' -ErrorAction SilentlyContinue
}
catch {
}

try {
    & "$env:SystemRoot\System32\w32tm.exe" /resync /force *> $null
}
catch {
}

exit 0
'@

        Set-Content -Path $helperPath -Value $helperContent -Encoding UTF8 -Force
        Write-Log "Ensured time sync helper script: $helperPath" 'OK'
        return $helperPath
    }
    catch {
        throw "Failed to create time sync helper script. $($_.Exception.Message)"
    }
}

function Register-TimeSyncTask {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory)][string]$ScriptsDirectory,
        [string]$TaskName = '00. Sync System Time Every 4 Hours'
    )

    if (-not (Test-Path -LiteralPath $WindowsPowerShellExe)) {
        throw "Windows PowerShell executable not found: $WindowsPowerShellExe"
    }

    $helperScriptPath = Ensure-TimeSyncHelperScript -ScriptsDirectory $ScriptsDirectory
    $argumentString = "-NoProfile -ExecutionPolicy Bypass -File `"$helperScriptPath`""
    $displayCommand = "$WindowsPowerShellExe $argumentString"
    Write-Log "Registering task '$TaskName' with action: $displayCommand" 'INFO'

    try {
        $action = New-ScheduledTaskAction -Execute $WindowsPowerShellExe -Argument $argumentString
        $triggerTimes = @('00:00','04:00','08:00','12:00','16:00','20:00')
        $triggers = foreach ($timeText in $triggerTimes) {
            $atTime = [datetime]::Today.Add([timespan]::Parse($timeText))
            New-ScheduledTaskTrigger -Daily -At $atTime -DaysInterval 1
        }
        $principal = Get-TaskPrincipal
        $settings = Get-TaskSettings

        if ($PSCmdlet.ShouldProcess($TaskName, 'Register time sync scheduled task')) {
            Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $triggers -Principal $principal -Settings $settings -Force | Out-Null
            Write-Log "Created task '$TaskName' to sync time every 4 hours using helper script '$helperScriptPath'." 'OK'
        }
    }
    catch {
        throw "Failed to create time sync task '$TaskName'. $($_.Exception.Message)"
    }
}

if (-not (Test-IsAdministrator)) {
    Write-Error 'This script must be run as Administrator.'
    exit 1
}

$taskDefinitions = @(
    [pscustomobject]@{ Name = '01. Check for Updated Scripts';          Script = (Join-Path $ScriptsRoot '00_Update-Scripts-FromShare.ps1');       Time = '01:15'; Args = '' },
    [pscustomobject]@{ Name = '02. Enable Windows Update Services';     Script = (Join-Path $ScriptsRoot '01_Enable_Windows_Update_Services.ps1');  Time = '01:20'; Args = '' },
    [pscustomobject]@{ Name = '03. Remove User Profiles Weekly';        Script = (Join-Path $ScriptsRoot '02_Remove_User_Profiles.ps1');            Time = '01:30'; Args = '' },
    [pscustomobject]@{ Name = '04. Weekend Apps Update';                Script = (Join-Path $ScriptsRoot '03_Weekend_Apps_Update.ps1');             Time = '02:00'; Args = '' },
    [pscustomobject]@{ Name = '05. Update Edge Silent';                 Script = (Join-Path $ScriptsRoot '04_Update_Edge_Silent.ps1');              Time = '02:45'; Args = '-KillEdgeProcesses' },
    [pscustomobject]@{ Name = '06. Weekend HP Drivers Update';          Script = (Join-Path $ScriptsRoot '05_Weekend_HP_Drivers_Update.ps1');       Time = '03:00'; Args = '' },
    [pscustomobject]@{ Name = '07. Weekend Windows Updates - 1st Pass'; Script = (Join-Path $ScriptsRoot '06_Weekend_Windows_Updates.ps1');         Time = '04:00'; Args = '' },
    [pscustomobject]@{ Name = '08. Force Reboot Install Updates';       Script = (Join-Path $ScriptsRoot '07_Force_Reboot_Install_Updates.ps1');    Time = '05:00'; Args = '' },
    [pscustomobject]@{ Name = '09. Weekend Windows Updates - 2nd Pass'; Script = (Join-Path $ScriptsRoot '06_Weekend_Windows_Updates.ps1');         Time = '05:30'; Args = '' },
    [pscustomobject]@{ Name = '10. Disable Windows Update Services';    Script = (Join-Path $ScriptsRoot '09_Disable_Windows_Update_Services.ps1'); Time = '06:00'; Args = '' },
    [pscustomobject]@{ Name = '11. Force Reboot Install Updates 2';     Script = (Join-Path $ScriptsRoot '07_Force_Reboot_Install_Updates.ps1');    Time = '06:05'; Args = '' },
    [pscustomobject]@{ Name = '12. System Repair';                      Script = (Join-Path $ScriptsRoot '08_System_Repair.ps1');                   Time = '06:15'; Args = '' },
    [pscustomobject]@{ Name = '13. Force Reboot Install Updates 3';     Script = (Join-Path $ScriptsRoot '07_Force_Reboot_Install_Updates.ps1');    Time = '07:00'; Args = '' }
)

Write-Log 'Initializing scheduled task rebuild script...' 'INFO'
Write-Log "Scripts root: $ScriptsRoot" 'INFO'
Write-Log "Delete Microsoft tasks: $IncludeMicrosoftTasks" 'INFO'
Write-Log 'Weekly tasks will be registered for Sunday, matching the batch file commands.' 'INFO'

try {
    Import-Module ScheduledTasks -ErrorAction Stop | Out-Null

    Remove-ExistingScheduledTasks -DeleteMicrosoftTasks:$IncludeMicrosoftTasks

    foreach ($task in $taskDefinitions) {
        Register-WeeklyPowerShellTask -TaskName $task.Name -ScriptPath $task.Script -StartTime $task.Time -ExtraArguments $task.Args
    }

    Register-TimeSyncTask -ScriptsDirectory $ScriptsRoot

    Write-Log 'Scheduled task rebuild completed successfully.' 'OK'
    exit 0
}
catch {
    Write-Log "Script failed: $($_.Exception.Message)" 'ERROR'
    exit 2
}
