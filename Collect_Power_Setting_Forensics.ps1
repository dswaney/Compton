#requires -version 5.1
<#
.SYNOPSIS
    Collects forensic data to determine what is changing Windows power settings.

.DESCRIPTION
    This script enables the most useful logging and captures a full forensic bundle for
    power setting drift issues, especially when monitor or sleep timeouts are being reset.

    It can:
      - Enable Kernel-Power diagnostic logging
      - Enable process creation auditing and command-line capture
      - Snapshot current power settings before and after monitoring
      - Watch for monitor/sleep timeout changes in real time
      - Export recent Security and Kernel-Power events
      - Enumerate scheduled tasks and services that may be changing power settings
      - Optionally create a self-heal scheduled task to re-apply desired power settings

.NOTES
    File Name   : Collect_Power_Setting_Forensics.ps1
    Version     : 1.0.0
    Last Updated: 2026-04-22
    Run As      : Administrator
#>

[CmdletBinding()]
param(
    [string]$OutputRoot = 'C:\Logs\PowerSettingForensics',
    [int]$WatchMinutes = 60,
    [int]$PollSeconds = 30,
    [switch]$InstallSelfHealTask,
    [string]$SelfHealTaskName = 'Enforce Desktop Power Settings',
    [switch]$DisableSelfHealTask,
    [switch]$ForceDesiredSettingsNow,
    [int]$DesiredDisplayMinutes = 60,
    [ValidateSet('Never','0','15','30','60','120')]
    [string]$DesiredSleepMinutes = 'Never'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$ScriptVersion = '1.0.0'
$RunTimestamp = Get-Date
$ComputerName = $env:COMPUTERNAME
$RunId = '{0}_{1:yyyy-MM-dd_HHmmss}' -f $ComputerName, $RunTimestamp
$OutputDir = Join-Path $OutputRoot $RunId
$LogPath = Join-Path $OutputDir ("{0}-PowerSettingForensics-{1:yyyy-MM-dd_HHmmss}.log" -f $ComputerName, $RunTimestamp)
$YamlPath = Join-Path $OutputDir ("{0}-PowerSettingForensics-{1:yyyy-MM-dd_HHmmss}.yml" -f $ComputerName, $RunTimestamp)
$BeforePath = Join-Path $OutputDir 'power_before_qh.txt'
$AfterPath = Join-Path $OutputDir 'power_after_qh.txt'
$PowerCfgChangesPath = Join-Path $OutputDir 'powercfg_diff.txt'
$KernelPowerEventsPath = Join-Path $OutputDir 'KernelPower_Diagnostic_Recent.txt'
$SecurityPowerCfgEventsPath = Join-Path $OutputDir 'Security_ProcessCreation_powercfg.txt'
$TaskInventoryPath = Join-Path $OutputDir 'ScheduledTasks_PowerRelated.txt'
$ServiceInventoryPath = Join-Path $OutputDir 'Services_PowerRelated.txt'
$WatcherPath = Join-Path $OutputDir 'WatcherTimeline.txt'
$SummaryPath = Join-Path $OutputDir 'Summary.txt'

$script:YamlEntries = New-Object System.Collections.Generic.List[string]

function Ensure-Directory {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','OK','WARN','ERROR')][string]$Level = 'INFO'
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[$timestamp] [$Level] $Message"
    Write-Host $line
    Add-Content -LiteralPath $LogPath -Value $line

    $yamlSafe = $Message.Replace("'", "''")
    $script:YamlEntries.Add(("  - time: '{0}'`n    level: '{1}'`n    message: '{2}'" -f $timestamp, $Level, $yamlSafe))
}

function Save-YamlLog {
    $content = @(
        'script: PowerSettingForensics'
        "version: '$ScriptVersion'"
        "computer: '$ComputerName'"
        ("started: '{0:yyyy-MM-dd HH:mm:ss}'" -f $RunTimestamp)
        "output_dir: '$($OutputDir.Replace("'","''"))'"
        'entries:'
    ) + $script:YamlEntries

    Set-Content -LiteralPath $YamlPath -Value $content -Encoding UTF8
}

function Test-IsAdministrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = [Security.Principal.WindowsPrincipal]::new($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Invoke-Native {
    param(
        [Parameter(Mandatory)][string]$FilePath,
        [string[]]$ArgumentList = @(),
        [switch]$IgnoreExitCode
    )

    $psi = [System.Diagnostics.ProcessStartInfo]::new()
    $psi.FileName = $FilePath
    foreach ($arg in $ArgumentList) {
        [void]$psi.ArgumentList.Add($arg)
    }
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true

    $process = [System.Diagnostics.Process]::new()
    $process.StartInfo = $psi
    [void]$process.Start()
    $stdout = $process.StandardOutput.ReadToEnd()
    $stderr = $process.StandardError.ReadToEnd()
    $process.WaitForExit()

    if (-not $IgnoreExitCode -and $process.ExitCode -ne 0) {
        throw "Command failed: $FilePath $($ArgumentList -join ' ') | ExitCode=$($process.ExitCode) | STDERR=$stderr"
    }

    [pscustomobject]@{
        ExitCode = $process.ExitCode
        StdOut   = $stdout
        StdErr   = $stderr
    }
}

function Get-ActiveSchemeGuid {
    $result = Invoke-Native -FilePath 'powercfg.exe' -ArgumentList @('/getactivescheme')
    if ($result.StdOut -match '([A-Fa-f0-9\-]{36})') {
        return $matches[1]
    }
    throw 'Unable to determine active power scheme GUID.'
}

function Get-SettingSeconds {
    param(
        [Parameter(Mandatory)][string]$SubGroup,
        [Parameter(Mandatory)][string]$Setting
    )

    $scheme = Get-ActiveSchemeGuid
    $query = Invoke-Native -FilePath 'powercfg.exe' -ArgumentList @('/query', $scheme, $SubGroup, $Setting)
    $ac = $null
    $dc = $null

    foreach ($line in ($query.StdOut -split "`r?`n")) {
        if ($line -match 'Current AC Power Setting Index:\s*0x([0-9A-Fa-f]+)') {
            $ac = [Convert]::ToInt32($matches[1], 16)
        }
        elseif ($line -match 'Current DC Power Setting Index:\s*0x([0-9A-Fa-f]+)') {
            $dc = [Convert]::ToInt32($matches[1], 16)
        }
    }

    [pscustomobject]@{
        SchemeGuid = $scheme
        ACSeconds  = $ac
        DCSeconds  = $dc
    }
}

function Get-PowerStateSummary {
    $display = Get-SettingSeconds -SubGroup 'SUB_VIDEO' -Setting 'VIDEOIDLE'
    $sleep = Get-SettingSeconds -SubGroup 'SUB_SLEEP' -Setting 'STANDBYIDLE'

    [pscustomobject]@{
        SchemeGuid          = $display.SchemeGuid
        DisplayACSeconds    = $display.ACSeconds
        DisplayDCSeconds    = $display.DCSeconds
        SleepACSeconds      = $sleep.ACSeconds
        SleepDCSeconds      = $sleep.DCSeconds
        DisplayACMinutes    = if ($display.ACSeconds -eq 0) { 'Never' } else { [math]::Round($display.ACSeconds / 60, 2) }
        DisplayDCMinutes    = if ($display.DCSeconds -eq 0) { 'Never' } else { [math]::Round($display.DCSeconds / 60, 2) }
        SleepACMinutes      = if ($sleep.ACSeconds -eq 0) { 'Never' } else { [math]::Round($sleep.ACSeconds / 60, 2) }
        SleepDCMinutes      = if ($sleep.DCSeconds -eq 0) { 'Never' } else { [math]::Round($sleep.DCSeconds / 60, 2) }
    }
}

function Export-PowerSnapshot {
    param([Parameter(Mandatory)][string]$Path)
    $result = Invoke-Native -FilePath 'powercfg.exe' -ArgumentList @('/qh')
    Set-Content -LiteralPath $Path -Value $result.StdOut -Encoding UTF8
}

function Enable-KernelPowerDiagnosticLogging {
    Write-Log 'Enabling Microsoft-Windows-Kernel-Power/Diagnostic log...' 'INFO'
    Invoke-Native -FilePath 'wevtutil.exe' -ArgumentList @('set-log', 'Microsoft-Windows-Kernel-Power/Diagnostic', '/enabled:true') | Out-Null
    Write-Log 'Kernel-Power diagnostic log enabled.' 'OK'
}

function Enable-ProcessCreationAuditing {
    Write-Log 'Enabling Process Creation auditing...' 'INFO'
    Invoke-Native -FilePath 'auditpol.exe' -ArgumentList @('/set', '/subcategory:Process Creation', '/success:enable') | Out-Null
    Invoke-Native -FilePath 'reg.exe' -ArgumentList @('add', 'HKLM\Software\Microsoft\Windows\CurrentVersion\Policies\System\Audit', '/v', 'ProcessCreationIncludeCmdLine_Enabled', '/t', 'REG_DWORD', '/d', '1', '/f') | Out-Null
    Write-Log 'Process creation auditing and command-line capture enabled.' 'OK'
}

function Export-RecentKernelPowerEvents {
    Write-Log 'Exporting recent Kernel-Power diagnostic events...' 'INFO'
    $events = Get-WinEvent -LogName 'Microsoft-Windows-Kernel-Power/Diagnostic' -ErrorAction SilentlyContinue |
        Sort-Object TimeCreated -Descending |
        Select-Object -First 200 TimeCreated, Id, ProviderName, LevelDisplayName, Message

    if ($events) {
        $events | Format-List | Out-String | Set-Content -LiteralPath $KernelPowerEventsPath -Encoding UTF8
        Write-Log "Saved Kernel-Power diagnostic events to $KernelPowerEventsPath" 'OK'
    }
    else {
        Set-Content -LiteralPath $KernelPowerEventsPath -Value 'No Kernel-Power diagnostic events found.' -Encoding UTF8
        Write-Log 'No Kernel-Power diagnostic events found.' 'WARN'
    }
}

function Export-RecentPowerCfgProcessCreationEvents {
    Write-Log 'Exporting recent Security process creation events containing powercfg or power-related commands...' 'INFO'
    $filterStart = (Get-Date).AddDays(-2)
    $events = Get-WinEvent -FilterHashtable @{ LogName = 'Security'; Id = 4688; StartTime = $filterStart } -ErrorAction SilentlyContinue |
        Where-Object {
            $_.Message -match 'powercfg|standby-timeout|monitor-timeout|VIDEOIDLE|STANDBYIDLE|SUB_SLEEP|SUB_VIDEO'
        } |
        Sort-Object TimeCreated -Descending |
        Select-Object -First 300 TimeCreated, Id, ProviderName, Message

    if ($events) {
        $events | Format-List | Out-String | Set-Content -LiteralPath $SecurityPowerCfgEventsPath -Encoding UTF8
        Write-Log "Saved Security process creation findings to $SecurityPowerCfgEventsPath" 'OK'
    }
    else {
        Set-Content -LiteralPath $SecurityPowerCfgEventsPath -Value 'No matching Security 4688 events found.' -Encoding UTF8
        Write-Log 'No matching Security 4688 events found yet.' 'WARN'
    }
}

function Export-SuspiciousScheduledTasks {
    Write-Log 'Inventorying scheduled tasks that may touch power settings...' 'INFO'
    $tasks = Get-ScheduledTask | ForEach-Object {
        $task = $_
        foreach ($action in $task.Actions) {
            [pscustomobject]@{
                TaskName  = $task.TaskName
                TaskPath  = $task.TaskPath
                State     = $task.State
                Execute   = $action.Execute
                Arguments = $action.Arguments
            }
        }
    } | Where-Object {
        $_.Execute -match 'powershell|pwsh|cmd|cscript|wscript|powercfg' -or
        $_.Arguments -match 'powercfg|sleep|monitor-timeout|standby-timeout|VIDEOIDLE|STANDBYIDLE|SUB_SLEEP|SUB_VIDEO'
    } | Sort-Object TaskPath, TaskName

    if ($tasks) {
        $tasks | Format-Table -AutoSize | Out-String | Set-Content -LiteralPath $TaskInventoryPath -Encoding UTF8
        Write-Log "Saved scheduled task inventory to $TaskInventoryPath" 'OK'
    }
    else {
        Set-Content -LiteralPath $TaskInventoryPath -Value 'No suspicious scheduled tasks found.' -Encoding UTF8
        Write-Log 'No suspicious scheduled tasks found.' 'WARN'
    }
}

function Export-SuspiciousServices {
    Write-Log 'Inventorying OEM and power-related services...' 'INFO'
    $services = Get-CimInstance Win32_Service |
        Where-Object {
            $_.Name -match 'Dell|DCU|Optimizer|Power|DPTF|Lenovo|HP|SupportAssist|Energy|Thermal' -or
            $_.DisplayName -match 'Dell|Optimizer|Power|DPTF|Lenovo|HP|SupportAssist|Energy|Thermal'
        } |
        Select-Object Name, DisplayName, State, StartMode, StartName, PathName |
        Sort-Object DisplayName, Name

    if ($services) {
        $services | Format-List | Out-String | Set-Content -LiteralPath $ServiceInventoryPath -Encoding UTF8
        Write-Log "Saved service inventory to $ServiceInventoryPath" 'OK'
    }
    else {
        Set-Content -LiteralPath $ServiceInventoryPath -Value 'No suspicious services found.' -Encoding UTF8
        Write-Log 'No suspicious services found.' 'WARN'
    }
}

function Set-DesiredDesktopPowerSettings {
    param(
        [Parameter(Mandatory)][int]$DisplayMinutes,
        [Parameter(Mandatory)][string]$SleepMinutes
    )

    Write-Log 'Applying desired desktop power settings...' 'INFO'

    $highPerfLine = (Invoke-Native -FilePath 'powercfg.exe' -ArgumentList @('/l')).StdOut -split "`r?`n" |
        Where-Object { $_ -match 'High performance' } |
        Select-Object -First 1

    if (-not $highPerfLine -or $highPerfLine -notmatch '([A-Fa-f0-9\-]{36})') {
        throw 'High Performance power plan not found.'
    }

    $highPerfGuid = $matches[1]
    Invoke-Native -FilePath 'powercfg.exe' -ArgumentList @('/setactive', $highPerfGuid) | Out-Null

    Invoke-Native -FilePath 'powercfg.exe' -ArgumentList @('/change', '-monitor-timeout-ac', $DisplayMinutes) | Out-Null
    Invoke-Native -FilePath 'powercfg.exe' -ArgumentList @('/change', '-monitor-timeout-dc', $DisplayMinutes) | Out-Null

    $sleepValueMinutes = if ($SleepMinutes -eq 'Never') { 0 } else { [int]$SleepMinutes }
    Invoke-Native -FilePath 'powercfg.exe' -ArgumentList @('/change', '-standby-timeout-ac', $sleepValueMinutes) | Out-Null
    Invoke-Native -FilePath 'powercfg.exe' -ArgumentList @('/change', '-standby-timeout-dc', $sleepValueMinutes) | Out-Null

    $displaySeconds = $DisplayMinutes * 60
    $sleepSeconds = $sleepValueMinutes * 60

    Invoke-Native -FilePath 'powercfg.exe' -ArgumentList @('/setacvalueindex', $highPerfGuid, 'SUB_VIDEO', 'VIDEOIDLE', $displaySeconds) | Out-Null
    Invoke-Native -FilePath 'powercfg.exe' -ArgumentList @('/setdcvalueindex', $highPerfGuid, 'SUB_VIDEO', 'VIDEOIDLE', $displaySeconds) | Out-Null
    Invoke-Native -FilePath 'powercfg.exe' -ArgumentList @('/setacvalueindex', $highPerfGuid, 'SUB_SLEEP', 'STANDBYIDLE', $sleepSeconds) | Out-Null
    Invoke-Native -FilePath 'powercfg.exe' -ArgumentList @('/setdcvalueindex', $highPerfGuid, 'SUB_SLEEP', 'STANDBYIDLE', $sleepSeconds) | Out-Null
    Invoke-Native -FilePath 'powercfg.exe' -ArgumentList @('/hibernate', 'off') | Out-Null
    Invoke-Native -FilePath 'powercfg.exe' -ArgumentList @('/S', $highPerfGuid) | Out-Null

    $summary = Get-PowerStateSummary
    Write-Log ("Applied settings. Display AC/DC: {0}/{1} minutes; Sleep AC/DC: {2}/{3}" -f $summary.DisplayACMinutes, $summary.DisplayDCMinutes, $summary.SleepACMinutes, $summary.SleepDCMinutes) 'OK'
}

function Install-SelfHealScheduledTask {
    param(
        [Parameter(Mandatory)][string]$TaskName,
        [Parameter(Mandatory)][int]$DisplayMinutes,
        [Parameter(Mandatory)][string]$SleepMinutes
    )

    $sleepValueMinutes = if ($SleepMinutes -eq 'Never') { 0 } else { [int]$SleepMinutes }
    $displaySeconds = $DisplayMinutes * 60
    $sleepSeconds = $sleepValueMinutes * 60

    $command = @"
powercfg -setactive SCHEME_MIN
powercfg -change -monitor-timeout-ac $DisplayMinutes
powercfg -change -monitor-timeout-dc $DisplayMinutes
powercfg -change -standby-timeout-ac $sleepValueMinutes
powercfg -change -standby-timeout-dc $sleepValueMinutes
for /f "tokens=3" %%G in ('powercfg /getactivescheme ^| findstr /r "[0-9A-Fa-f-][0-9A-Fa-f-]*"') do set GUID=%%G
powercfg -setacvalueindex %GUID% SUB_VIDEO VIDEOIDLE $displaySeconds
powercfg -setdcvalueindex %GUID% SUB_VIDEO VIDEOIDLE $displaySeconds
powercfg -setacvalueindex %GUID% SUB_SLEEP STANDBYIDLE $sleepSeconds
powercfg -setdcvalueindex %GUID% SUB_SLEEP STANDBYIDLE $sleepSeconds
powercfg -hibernate off
powercfg /S %GUID%
"@

    $tempCmd = Join-Path $env:ProgramData 'PowerSettingSelfHeal.cmd'
    Set-Content -LiteralPath $tempCmd -Value $command -Encoding ASCII

    $action = New-ScheduledTaskAction -Execute 'cmd.exe' -Argument "/c `"$tempCmd`""
    $trigger = New-ScheduledTaskTrigger -Once -At ((Get-Date).AddMinutes(1))
    $trigger.Repetition = [Microsoft.Management.Infrastructure.CimInstance]::new(
        'MSFT_TaskRepetitionPattern',
        @{ Interval = 'PT5M'; Duration = 'P1D' },
        'root/Microsoft/Windows/TaskScheduler'
    )
    $principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
    $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable

    Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Force | Out-Null
    Write-Log "Installed self-heal scheduled task '$TaskName' running every 5 minutes." 'OK'
}

function Disable-SelfHealScheduledTask {
    param([Parameter(Mandatory)][string]$TaskName)
    $task = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
    if ($task) {
        Disable-ScheduledTask -TaskName $TaskName | Out-Null
        Write-Log "Disabled scheduled task '$TaskName'." 'OK'
    }
    else {
        Write-Log "Scheduled task '$TaskName' was not found." 'WARN'
    }
}

function Start-LiveWatcher {
    param(
        [Parameter(Mandatory)][int]$DurationMinutes,
        [Parameter(Mandatory)][int]$IntervalSeconds
    )

    Write-Log ("Starting live watcher for {0} minutes with a {1}-second poll interval..." -f $DurationMinutes, $IntervalSeconds) 'INFO'
    $endTime = (Get-Date).AddMinutes($DurationMinutes)
    $changeDetected = $false

    Add-Content -LiteralPath $WatcherPath -Value ("Watcher started: {0:yyyy-MM-dd HH:mm:ss}" -f (Get-Date))

    while ((Get-Date) -lt $endTime) {
        $state = Get-PowerStateSummary
        $line = "{0:yyyy-MM-dd HH:mm:ss} | Scheme={1} | DisplayAC={2}s | DisplayDC={3}s | SleepAC={4}s | SleepDC={5}s" -f (Get-Date), $state.SchemeGuid, $state.DisplayACSeconds, $state.DisplayDCSeconds, $state.SleepACSeconds, $state.SleepDCSeconds
        Add-Content -LiteralPath $WatcherPath -Value $line

        $hitFiveMinutes = @($state.DisplayACSeconds, $state.DisplayDCSeconds, $state.SleepACSeconds, $state.SleepDCSeconds) -contains 300
        if ($hitFiveMinutes) {
            $changeDetected = $true
            Write-Log 'Detected a timeout value set to 300 seconds (5 minutes).' 'WARN'
            Add-Content -LiteralPath $WatcherPath -Value '*** DETECTED 5-MINUTE VALUE ***'
            break
        }

        Start-Sleep -Seconds $IntervalSeconds
    }

    if (-not $changeDetected) {
        Write-Log 'Live watcher completed without detecting a 5-minute value.' 'INFO'
    }

    return $changeDetected
}

function Write-Summary {
    param(
        [Parameter(Mandatory)][pscustomobject]$Before,
        [Parameter(Mandatory)][pscustomobject]$After,
        [Parameter(Mandatory)][bool]$DetectedChange
    )

    $lines = @(
        "Computer: $ComputerName",
        "Script Version: $ScriptVersion",
        ("Started: {0:yyyy-MM-dd HH:mm:ss}" -f $RunTimestamp),
        "Output Directory: $OutputDir",
        '',
        'Before:',
        ("  SchemeGuid: {0}" -f $Before.SchemeGuid),
        ("  Display AC/DC: {0}/{1} minutes" -f $Before.DisplayACMinutes, $Before.DisplayDCMinutes),
        ("  Sleep AC/DC: {0}/{1} minutes" -f $Before.SleepACMinutes, $Before.SleepDCMinutes),
        '',
        'After:',
        ("  SchemeGuid: {0}" -f $After.SchemeGuid),
        ("  Display AC/DC: {0}/{1} minutes" -f $After.DisplayACMinutes, $After.DisplayDCMinutes),
        ("  Sleep AC/DC: {0}/{1} minutes" -f $After.SleepACMinutes, $After.SleepDCMinutes),
        '',
        ("Detected 5-minute change during watch: {0}" -f $DetectedChange),
        '',
        'Important files:',
        "  $LogPath",
        "  $YamlPath",
        "  $BeforePath",
        "  $AfterPath",
        "  $PowerCfgChangesPath",
        "  $KernelPowerEventsPath",
        "  $SecurityPowerCfgEventsPath",
        "  $TaskInventoryPath",
        "  $ServiceInventoryPath",
        "  $WatcherPath"
    )

    Set-Content -LiteralPath $SummaryPath -Value $lines -Encoding UTF8
}

try {
    Ensure-Directory -Path $OutputDir
    New-Item -ItemType File -Path $LogPath -Force | Out-Null

    if (-not (Test-IsAdministrator)) {
        throw 'This script must be run in an elevated PowerShell session as Administrator.'
    }

    Write-Log "Starting Power Setting Forensics $ScriptVersion" 'INFO'
    Write-Log "Output directory: $OutputDir" 'INFO'

    Enable-KernelPowerDiagnosticLogging
    Enable-ProcessCreationAuditing

    if ($ForceDesiredSettingsNow) {
        Set-DesiredDesktopPowerSettings -DisplayMinutes $DesiredDisplayMinutes -SleepMinutes $DesiredSleepMinutes
    }

    if ($InstallSelfHealTask) {
        Install-SelfHealScheduledTask -TaskName $SelfHealTaskName -DisplayMinutes $DesiredDisplayMinutes -SleepMinutes $DesiredSleepMinutes
    }

    if ($DisableSelfHealTask) {
        Disable-SelfHealScheduledTask -TaskName $SelfHealTaskName
    }

    $before = Get-PowerStateSummary
    Write-Log ("Initial state - Display AC/DC: {0}/{1} minutes; Sleep AC/DC: {2}/{3}" -f $before.DisplayACMinutes, $before.DisplayDCMinutes, $before.SleepACMinutes, $before.SleepDCMinutes) 'INFO'
    Export-PowerSnapshot -Path $BeforePath
    Export-SuspiciousScheduledTasks
    Export-SuspiciousServices

    $detected = Start-LiveWatcher -DurationMinutes $WatchMinutes -IntervalSeconds $PollSeconds

    $after = Get-PowerStateSummary
    Export-PowerSnapshot -Path $AfterPath

    $compare = Compare-Object -ReferenceObject (Get-Content -LiteralPath $BeforePath) -DifferenceObject (Get-Content -LiteralPath $AfterPath)
    if ($compare) {
        $compare | Out-String | Set-Content -LiteralPath $PowerCfgChangesPath -Encoding UTF8
        Write-Log "Saved powercfg diff to $PowerCfgChangesPath" 'OK'
    }
    else {
        Set-Content -LiteralPath $PowerCfgChangesPath -Value 'No differences found between before and after powercfg /qh snapshots.' -Encoding UTF8
        Write-Log 'No differences found between before and after powercfg snapshots.' 'INFO'
    }

    Export-RecentKernelPowerEvents
    Export-RecentPowerCfgProcessCreationEvents
    Write-Summary -Before $before -After $after -DetectedChange:$detected

    Write-Log "Summary written to $SummaryPath" 'OK'
    Write-Log 'Power setting forensics collection completed successfully.' 'OK'
}
catch {
    Write-Log $_.Exception.Message 'ERROR'
    throw
}
finally {
    Save-YamlLog
}
