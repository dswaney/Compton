# =====================================================================
# Script Name   : Register-Tasks_SYSTEM.ps1
# ScriptVersion : 3.0
# LastUpdated   : 2026-07-21
# Purpose       : Reconcile managed scheduled tasks under SYSTEM.
#
# Behavior:
#   - Creates missing managed tasks.
#   - Updates managed tasks only when their action, trigger, principal, or
#     important settings differ.
#   - Leaves unrelated and Microsoft scheduled tasks untouched.
#   - Removes no scheduled tasks automatically.
# =====================================================================

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Low')]
param(
    [string]$ScriptsRoot = 'C:\Scripts',
    [string]$TaskPath = '\',
    [string]$LogPath = 'C:\Logs\Register-Tasks_SYSTEM.log'
)

$ErrorActionPreference = 'Stop'
$WindowsPowerShellExe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','ACTION','OK','WARN','ERROR')]
        [string]$Level = 'INFO'
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[$timestamp] [$('{0,-6}' -f $Level)] $Message"

    $color = switch ($Level) {
        'ACTION' { 'Yellow' }
        'OK'     { 'Green' }
        'WARN'   { 'DarkYellow' }
        'ERROR'  { 'Red' }
        default  { 'Cyan' }
    }

    Write-Host $line -ForegroundColor $color

    try {
        $logDirectory = Split-Path -Path $LogPath -Parent
        if (-not (Test-Path -LiteralPath $logDirectory -PathType Container)) {
            New-Item -Path $logDirectory -ItemType Directory -Force | Out-Null
        }

        Add-Content -LiteralPath $LogPath -Value $line -Encoding UTF8
    }
    catch {
    }
}

function Test-IsAdministrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Get-DesiredActionArguments {
    param(
        [Parameter(Mandatory)][string]$ScriptPath,
        [string]$ExtraArguments = ''
    )

    $arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$ScriptPath`""

    if (-not [string]::IsNullOrWhiteSpace($ExtraArguments)) {
        $arguments = "$arguments $ExtraArguments"
    }

    return $arguments
}

function Get-TimeText {
    param([Parameter(Mandatory)]$DateValue)

    try {
        return ([datetime]$DateValue).ToString('HH:mm')
    }
    catch {
        return $null
    }
}

function Test-CommonTaskProperties {
    param(
        [Parameter(Mandatory)]$ExistingTask,
        [Parameter(Mandatory)][string]$ExpectedArguments
    )

    $issues = New-Object System.Collections.Generic.List[string]

    $action = @($ExistingTask.Actions)[0]
    if (-not $action) {
        [void]$issues.Add('PowerShell action is missing.')
    }
    else {
        if ([string]$action.Execute -ine $WindowsPowerShellExe) {
            [void]$issues.Add("Executable differs: $($action.Execute)")
        }

        if ([string]$action.Arguments -cne $ExpectedArguments) {
            [void]$issues.Add("Arguments differ: $($action.Arguments)")
        }
    }

    if ([string]$ExistingTask.Principal.UserId -notmatch '^(SYSTEM|S-1-5-18)$') {
        [void]$issues.Add("Principal differs: $($ExistingTask.Principal.UserId)")
    }

    if ([string]$ExistingTask.Principal.RunLevel -ne 'Highest') {
        [void]$issues.Add("Run level differs: $($ExistingTask.Principal.RunLevel)")
    }

    if ($ExistingTask.Settings.StartWhenAvailable) {
        [void]$issues.Add('Run-as-soon-as-possible-after-missed-start is enabled.')
    }

    if ($ExistingTask.Settings.DisallowStartIfOnBatteries) {
        [void]$issues.Add('Task is blocked while on battery power.')
    }

    if ($ExistingTask.Settings.StopIfGoingOnBatteries) {
        [void]$issues.Add('Task stops when switching to battery power.')
    }

    return @($issues)
}

function Test-WeeklyTaskMatches {
    param(
        [Parameter(Mandatory)]$ExistingTask,
        [Parameter(Mandatory)][string]$ExpectedArguments,
        [Parameter(Mandatory)][string]$StartTime
    )

    $issues = New-Object System.Collections.Generic.List[string]
    foreach ($issue in @(Test-CommonTaskProperties -ExistingTask $ExistingTask -ExpectedArguments $ExpectedArguments)) {
        [void]$issues.Add($issue)
    }

    $triggers = @($ExistingTask.Triggers)
    if ($triggers.Count -ne 1) {
        [void]$issues.Add("Expected one weekly trigger; found $($triggers.Count).")
    }
    else {
        $trigger = $triggers[0]

        if ([string]$trigger.CimClass.CimClassName -notmatch 'Weekly') {
            [void]$issues.Add('Trigger is not weekly.')
        }

        # Sunday is bit value 1 for MSFT_TaskWeeklyTrigger.DaysOfWeek.
        if ([int]$trigger.DaysOfWeek -ne 1) {
            [void]$issues.Add("Weekly day differs: $($trigger.DaysOfWeek)")
        }

        if ([int]$trigger.WeeksInterval -ne 1) {
            [void]$issues.Add("Weeks interval differs: $($trigger.WeeksInterval)")
        }

        if ((Get-TimeText -DateValue $trigger.StartBoundary) -ne $StartTime) {
            [void]$issues.Add("Start time differs: $(Get-TimeText -DateValue $trigger.StartBoundary)")
        }
    }

    return @($issues)
}

function Test-TimeSyncTaskMatches {
    param(
        [Parameter(Mandatory)]$ExistingTask,
        [Parameter(Mandatory)][string]$ExpectedArguments,
        [Parameter(Mandatory)][string[]]$ExpectedTimes
    )

    $issues = New-Object System.Collections.Generic.List[string]
    foreach ($issue in @(Test-CommonTaskProperties -ExistingTask $ExistingTask -ExpectedArguments $ExpectedArguments)) {
        [void]$issues.Add($issue)
    }

    $actualTimes = @(
        $ExistingTask.Triggers |
        ForEach-Object { Get-TimeText -DateValue $_.StartBoundary } |
        Where-Object { $_ } |
        Sort-Object
    )

    $wantedTimes = @($ExpectedTimes | Sort-Object)

    if (($actualTimes -join ',') -ne ($wantedTimes -join ',')) {
        [void]$issues.Add("Trigger times differ. Actual: $($actualTimes -join ', ')")
    }

    return @($issues)
}

function New-ManagedTaskSettings {
    return New-ScheduledTaskSettingsSet `
        -AllowStartIfOnBatteries `
        -DontStopIfGoingOnBatteries `
        -ExecutionTimeLimit (New-TimeSpan -Hours 12)
}

function New-ManagedTaskPrincipal {
    return New-ScheduledTaskPrincipal `
        -UserId 'SYSTEM' `
        -LogonType ServiceAccount `
        -RunLevel Highest
}

function Ensure-WeeklyTask {
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$ScriptPath,
        [Parameter(Mandatory)][string]$StartTime,
        [string]$ExtraArguments = ''
    )

    if (-not (Test-Path -LiteralPath $ScriptPath -PathType Leaf)) {
        Write-Log "The script is currently missing, but its task definition will still be reconciled: $ScriptPath" 'WARN'
    }

    $arguments = Get-DesiredActionArguments -ScriptPath $ScriptPath -ExtraArguments $ExtraArguments
    $existing = Get-ScheduledTask -TaskName $Name -TaskPath $TaskPath -ErrorAction SilentlyContinue

    $needsUpdate = $true
    $issues = @()

    if ($existing) {
        $issues = @(Test-WeeklyTaskMatches -ExistingTask $existing -ExpectedArguments $arguments -StartTime $StartTime)
        $needsUpdate = $issues.Count -gt 0
    }

    if (-not $needsUpdate) {
        Write-Log "Task is already correct: $TaskPath$Name" 'OK'
        return
    }

    if ($existing) {
        Write-Log "Task requires an update: $TaskPath$Name" 'ACTION'
        foreach ($issue in $issues) {
            Write-Log " - $issue" 'WARN'
        }
    }
    else {
        Write-Log "Task is missing and will be created: $TaskPath$Name" 'ACTION'
    }

    $action = New-ScheduledTaskAction -Execute $WindowsPowerShellExe -Argument $arguments
    $startAt = [datetime]::Today.Add([timespan]::Parse($StartTime))
    $trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Sunday -WeeksInterval 1 -At $startAt
    $principal = New-ManagedTaskPrincipal
    $settings = New-ManagedTaskSettings

    if ($PSCmdlet.ShouldProcess("$TaskPath$Name", 'Create or update managed weekly task')) {
        Register-ScheduledTask `
            -TaskName $Name `
            -TaskPath $TaskPath `
            -Action $action `
            -Trigger $trigger `
            -Principal $principal `
            -Settings $settings `
            -Force | Out-Null

        Write-Log "Task reconciled: $TaskPath$Name at Sunday $StartTime" 'OK'
    }
}

function Ensure-TimeSyncTask {
    param(
        [Parameter(Mandatory)][string]$ScriptPath,
        [string]$Name = '00. Sync System Time Every 4 Hours'
    )

    $triggerTimes = @('00:00','04:00','08:00','12:00','16:00','20:00')
    $arguments = Get-DesiredActionArguments -ScriptPath $ScriptPath
    $existing = Get-ScheduledTask -TaskName $Name -TaskPath $TaskPath -ErrorAction SilentlyContinue

    $needsUpdate = $true
    $issues = @()

    if ($existing) {
        $issues = @(Test-TimeSyncTaskMatches -ExistingTask $existing -ExpectedArguments $arguments -ExpectedTimes $triggerTimes)
        $needsUpdate = $issues.Count -gt 0
    }

    if (-not $needsUpdate) {
        Write-Log "Task is already correct: $TaskPath$Name" 'OK'
        return
    }

    if ($existing) {
        Write-Log "Time-sync task requires an update: $TaskPath$Name" 'ACTION'
        foreach ($issue in $issues) {
            Write-Log " - $issue" 'WARN'
        }
    }
    else {
        Write-Log "Time-sync task is missing and will be created: $TaskPath$Name" 'ACTION'
    }

    $action = New-ScheduledTaskAction -Execute $WindowsPowerShellExe -Argument $arguments
    $triggers = foreach ($timeText in $triggerTimes) {
        $atTime = [datetime]::Today.Add([timespan]::Parse($timeText))
        New-ScheduledTaskTrigger -Daily -At $atTime -DaysInterval 1
    }

    if ($PSCmdlet.ShouldProcess("$TaskPath$Name", 'Create or update managed time-sync task')) {
        Register-ScheduledTask `
            -TaskName $Name `
            -TaskPath $TaskPath `
            -Action $action `
            -Trigger $triggers `
            -Principal (New-ManagedTaskPrincipal) `
            -Settings (New-ManagedTaskSettings) `
            -Force | Out-Null

        Write-Log "Time-sync task reconciled: $TaskPath$Name" 'OK'
    }
}

if (-not (Test-IsAdministrator)) {
    Write-Error 'This script must be run as Administrator or SYSTEM.'
    exit 1
}

# These are previous names from the transition that would otherwise leave
# duplicate 07:00 or 07:30 tasks. Only these explicitly managed obsolete names
# are removed; unrelated tasks are never touched.
$obsoleteManagedTaskNames = @(
    '13. Force Reboot Install Updates 3',
    '14. Maintain SHARP Driver and PaperCut'
)

$taskDefinitions = @(
    [pscustomobject]@{ Name = '01. Check for Updated Scripts';          Script = (Join-Path $ScriptsRoot '00_Update-Scripts-FromShare.ps1');        Time = '01:15'; Args = '' },
    [pscustomobject]@{ Name = '02. Enable Windows Update Services';     Script = (Join-Path $ScriptsRoot '01_Enable_Windows_Update_Services.ps1'); Time = '01:20'; Args = '' },
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
    [pscustomobject]@{ Name = '13. Maintain SHARP Driver and PaperCut'; Script = (Join-Path $ScriptsRoot '11_Install_SharpDriver_And_PaperCut.ps1'); Time = '07:00'; Args = '' },
    [pscustomobject]@{ Name = '14. Force Reboot Install Updates 3';     Script = (Join-Path $ScriptsRoot '07_Force_Reboot_Install_Updates.ps1');    Time = '07:30'; Args = '' }
)

try {
    Import-Module ScheduledTasks -ErrorAction Stop

    Write-Log "Reconciling managed tasks under $TaskPath" 'INFO'

    $desiredTaskNames = @($taskDefinitions.Name) + '00. Sync System Time Every 4 Hours'

    foreach ($obsoleteName in $obsoleteManagedTaskNames) {
        if ($obsoleteName -in $desiredTaskNames) {
            continue
        }

        $obsoleteTask = Get-ScheduledTask `
            -TaskName $obsoleteName `
            -TaskPath $TaskPath `
            -ErrorAction SilentlyContinue

        if ($obsoleteTask) {
            Write-Log "Removing superseded managed task: $TaskPath$obsoleteName" 'ACTION'

            if ($PSCmdlet.ShouldProcess("$TaskPath$obsoleteName", 'Remove superseded managed task')) {
                Unregister-ScheduledTask `
                    -TaskName $obsoleteName `
                    -TaskPath $TaskPath `
                    -Confirm:$false `
                    -ErrorAction Stop

                Write-Log "Removed superseded managed task: $TaskPath$obsoleteName" 'OK'
            }
        }
    }

    foreach ($task in $taskDefinitions) {
        Ensure-WeeklyTask `
            -Name $task.Name `
            -ScriptPath $task.Script `
            -StartTime $task.Time `
            -ExtraArguments $task.Args
    }

    Ensure-TimeSyncTask -ScriptPath (Join-Path $ScriptsRoot '10_Sync_System_Time.ps1')

    Write-Log 'All managed scheduled tasks were reconciled successfully.' 'OK'
    exit 0
}
catch {
    Write-Log "Task reconciliation failed: $($_.Exception.Message)" 'ERROR'
    exit 2
}
