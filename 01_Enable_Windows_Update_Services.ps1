# =====================================================================
# ScriptName: 01_Enable_Windows_Update_Services.ps1
# ScriptVersion: 2.0
# LastUpdated: 2026-04-28
# Purpose: Restore Windows Update services, tasks, policy settings,
#          and Windows 11 classic right-click context menu behavior for all users;
#          verify required services are running, retry startup failures
#          up to 4 total attempts, and force a reboot if critical
#          services still refuse to start.
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

function Set-ClassicRightClickMenuForHive {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$RegistryRoot,

        [Parameter(Mandatory)]
        [string]$DisplayName
    )

    $clsid = '{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}'
    $basePath = "$RegistryRoot\Software\Classes\CLSID\$clsid"
    $subPath  = "$basePath\InprocServer32"

    try {
        if (-not (Test-Path -LiteralPath $basePath)) {
            New-Item -Path $basePath -Force -ErrorAction Stop | Out-Null
            Write-Status "Created classic right-click menu CLSID key for $DisplayName" 'OK'
        }
        else {
            Write-Status "Classic right-click menu CLSID key already exists for $DisplayName" 'INFO'
        }

        if (-not (Test-Path -LiteralPath $subPath)) {
            New-Item -Path $subPath -Force -ErrorAction Stop | Out-Null
            Write-Status "Created classic right-click menu InprocServer32 key for $DisplayName" 'OK'
        }
        else {
            Write-Status "Classic right-click menu InprocServer32 key already exists for $DisplayName" 'INFO'
        }

        Set-Item -Path $subPath -Value '' -ErrorAction Stop
        Write-Status "Enabled Windows 11 classic right-click menu for $DisplayName." 'OK'
        return $true
    }
    catch {
        Write-Status "Failed to enable Windows 11 classic right-click menu for $DisplayName : $($_.Exception.Message)" 'WARN'
        return $false
    }
}

function Invoke-WithLoadedUserHive {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$HiveFile,

        [Parameter(Mandatory)]
        [string]$DisplayName,

        [Parameter(Mandatory)]
        [scriptblock]$Action
    )

    if (-not (Test-Path -LiteralPath $HiveFile)) {
        Write-Status "User hive file not found for $DisplayName : $HiveFile" 'WARN'
        return $false
    }

    $tempHiveName = "TempClassicContextMenu_$([guid]::NewGuid().ToString('N'))"
    $loaded = $false

    try {
        $loadOutput = & reg.exe load "HKU\$tempHiveName" "$HiveFile" 2>&1
        if ($LASTEXITCODE -ne 0) {
            Write-Status "Failed to load hive for $DisplayName : $($loadOutput -join ' ')" 'WARN'
            return $false
        }

        $loaded = $true
        $registryRoot = "Registry::HKEY_USERS\$tempHiveName"
        & $Action $registryRoot $DisplayName | Out-Null
        return $true
    }
    catch {
        Write-Status "Unexpected error while processing hive for $DisplayName : $($_.Exception.Message)" 'WARN'
        return $false
    }
    finally {
        if ($loaded) {
            try {
                [gc]::Collect()
                [gc]::WaitForPendingFinalizers()
                Start-Sleep -Milliseconds 300
                $unloadOutput = & reg.exe unload "HKU\$tempHiveName" 2>&1
                if ($LASTEXITCODE -eq 0) {
                    Write-Status "Unloaded temporary registry hive for $DisplayName" 'OK'
                }
                else {
                    Write-Status "Failed to unload temporary registry hive for $DisplayName : $($unloadOutput -join ' ')" 'WARN'
                }
            }
            catch {
                Write-Status "Unexpected error unloading hive for $DisplayName : $($_.Exception.Message)" 'WARN'
            }
        }
    }
}

function Enable-ClassicWindows11RightClickMenu {
    [CmdletBinding()]
    param()

    Write-Status "Applying Windows 11 classic right-click menu for all users..." 'INFO'

    $processedSids = New-Object 'System.Collections.Generic.HashSet[string]'

    [void](Set-ClassicRightClickMenuForHive -RegistryRoot 'HKCU:' -DisplayName 'current user')

    try {
        $currentSid = [System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value
        [void]$processedSids.Add($currentSid)
    }
    catch {
        Write-Status "Could not determine current user SID: $($_.Exception.Message)" 'WARN'
    }

    try {
        Get-ChildItem -Path 'Registry::HKEY_USERS' -ErrorAction Stop |
            Where-Object {
                $_.PSChildName -match '^S-1-5-21-' -and
                $_.PSChildName -notmatch '_Classes$'
            } |
            ForEach-Object {
                $sid = $_.PSChildName
                [void]$processedSids.Add($sid)
                [void](Set-ClassicRightClickMenuForHive -RegistryRoot "Registry::HKEY_USERS\$sid" -DisplayName "loaded profile $sid")
            }
    }
    catch {
        Write-Status "Failed to enumerate loaded user registry hives: $($_.Exception.Message)" 'WARN'
    }

    try {
        $profiles = Get-CimInstance -ClassName Win32_UserProfile -ErrorAction Stop |
            Where-Object {
                -not $_.Special -and
                $_.SID -match '^S-1-5-21-' -and
                $_.LocalPath -and
                (Test-Path -LiteralPath (Join-Path $_.LocalPath 'NTUSER.DAT'))
            }

        foreach ($profile in $profiles) {
            if ($processedSids.Contains($profile.SID) -or (Test-Path -LiteralPath "Registry::HKEY_USERS\$($profile.SID)")) {
                Write-Status "Profile already loaded or processed; skipping offline load for $($profile.LocalPath)" 'INFO'
                continue
            }

            $ntUserDat = Join-Path $profile.LocalPath 'NTUSER.DAT'
            [void](Invoke-WithLoadedUserHive -HiveFile $ntUserDat -DisplayName "offline profile $($profile.LocalPath)" -Action {
                param($RegistryRoot, $DisplayName)
                Set-ClassicRightClickMenuForHive -RegistryRoot $RegistryRoot -DisplayName $DisplayName
            })
        }
    }
    catch {
        Write-Status "Failed to process offline user profiles: $($_.Exception.Message)" 'WARN'
    }

    $defaultHive = Join-Path $env:SystemDrive 'Users\Default\NTUSER.DAT'
    if (Test-Path -LiteralPath $defaultHive) {
        [void](Invoke-WithLoadedUserHive -HiveFile $defaultHive -DisplayName 'Default User profile for future users' -Action {
            param($RegistryRoot, $DisplayName)
            Set-ClassicRightClickMenuForHive -RegistryRoot $RegistryRoot -DisplayName $DisplayName
        })
    }
    else {
        Write-Status "Default User hive not found at expected path: $defaultHive" 'WARN'
    }

    Write-Status "Completed Windows 11 classic right-click menu application for all available users." 'OK'
}

function Get-ScriptHeaderValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter(Mandatory)]
        [string]$HeaderName
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        return $null
    }

    try {
        $pattern = '^\s*#\s*' + [regex]::Escape($HeaderName) + '\s*:\s*(.+?)\s*$'
        $match = Select-String -LiteralPath $Path -Pattern $pattern -CaseSensitive:$false -ErrorAction Stop | Select-Object -First 1
        if ($match -and $match.Matches.Count -gt 0) {
            return $match.Matches[0].Groups[1].Value.Trim()
        }
    }
    catch {
        Write-Status "Unable to read $HeaderName from $Path : $($_.Exception.Message)" 'WARN'
    }

    return $null
}

function Convert-ToVersionObjectSafe {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [string]$VersionText
    )

    if ([string]::IsNullOrWhiteSpace($VersionText)) {
        return [version]'0.0'
    }

    try {
        return [version]$VersionText.Trim()
    }
    catch {
        $cleanVersion = ($VersionText -replace '[^0-9\.]', '').Trim('.')
        if ([string]::IsNullOrWhiteSpace($cleanVersion)) {
            return [version]'0.0'
        }

        try {
            return [version]$cleanVersion
        }
        catch {
            Write-Status "Unable to parse version [$VersionText]. Treating as 0.0." 'WARN'
            return [version]'0.0'
        }
    }
}

function Ensure-ShareScriptUpdater {
    [CmdletBinding()]
    param(
        [string]$SourcePath = '\\filesvr\labshare\00_Update-Scripts-FromShare.ps1',
        [string]$DestinationFolder = 'C:\Scripts',
        [string]$DestinationFileName = '00_Update-Scripts-FromShare.ps1',
        [string[]]$LegacyFilesToDelete = @(
            '00_Update-Scripts-FromGithub.ps1',
            'Register-Tasks_SYSTEM.ps1'
        )
    )

    $destinationPath = Join-Path $DestinationFolder $DestinationFileName

    try {
        if (-not (Test-Path -LiteralPath $DestinationFolder)) {
            New-Item -Path $DestinationFolder -ItemType Directory -Force | Out-Null
            Write-Status "Created destination folder: $DestinationFolder" 'OK'
        }

        foreach ($legacyFileName in $LegacyFilesToDelete) {
            $legacyPath = Join-Path $DestinationFolder $legacyFileName
            if (Test-Path -LiteralPath $legacyPath) {
                try {
                    Remove-Item -LiteralPath $legacyPath -Force -ErrorAction Stop
                    Write-Status "Removed legacy script: $legacyPath" 'OK'
                }
                catch {
                    Write-Status "Failed to remove legacy script [$legacyPath]: $($_.Exception.Message)" 'WARN'
                }
            }
            else {
                Write-Status "Legacy script not present: $legacyPath" 'INFO'
            }
        }

        if (-not (Test-Path -LiteralPath $SourcePath)) {
            if (Test-Path -LiteralPath $destinationPath) {
                Write-Status "Share source unavailable, but updater already exists locally: $destinationPath" 'WARN'
                return $destinationPath
            }

            throw "Required updater script is missing locally and source is unavailable. Source: $SourcePath ; Local: $destinationPath"
        }

        $sourceVersionText = Get-ScriptHeaderValue -Path $SourcePath -HeaderName 'ScriptVersion'
        $localVersionText  = Get-ScriptHeaderValue -Path $destinationPath -HeaderName 'ScriptVersion'

        $sourceVersion = Convert-ToVersionObjectSafe -VersionText $sourceVersionText
        $localVersion  = Convert-ToVersionObjectSafe -VersionText $localVersionText

        if ([string]::IsNullOrWhiteSpace($sourceVersionText)) {
            Write-Status "Source updater is missing a ScriptVersion header. It will only be copied if the local file is missing." 'WARN'
        }

        if (-not (Test-Path -LiteralPath $destinationPath)) {
            Copy-Item -LiteralPath $SourcePath -Destination $destinationPath -Force -ErrorAction Stop
            Write-Status "Local updater missing. Copied updater from [$SourcePath] to [$destinationPath]. Source version: [$sourceVersionText]" 'OK'
        }
        elseif ($sourceVersion -gt $localVersion) {
            Write-Status "Updater source version [$sourceVersionText] is newer than local version [$localVersionText]. Updating local copy..." 'WARN'
            Copy-Item -LiteralPath $SourcePath -Destination $destinationPath -Force -ErrorAction Stop
            Write-Status "Updated local updater script: $destinationPath" 'OK'
        }
        else {
            Write-Status "Local updater is current. Local version [$localVersionText], source version [$sourceVersionText]." 'OK'
        }

        try {
            Unblock-File -LiteralPath $destinationPath -ErrorAction SilentlyContinue
        }
        catch {
        }

        return $destinationPath
    }
    catch {
        Write-Status "Failed to verify/stage share updater script: $($_.Exception.Message)" 'ERROR'
        return $null
    }
}

function Get-WeeklySundayTriggerSummary {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object[]]$Triggers
    )

    foreach ($trigger in $Triggers) {
        try {
            $days = [string]$trigger.DaysOfWeek
            $startBoundary = [string]$trigger.StartBoundary
            $startTime = $null

            if (-not [string]::IsNullOrWhiteSpace($startBoundary)) {
                $startTime = ([datetime]$startBoundary).ToString('HH:mm')
            }

            if ($days -match 'Sunday' -and $startTime -eq '01:15') {
                return $true
            }
        }
        catch {
        }
    }

    return $false
}

function Test-CheckForUpdatedScriptsTaskCompliant {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$Task,

        [Parameter(Mandatory)]
        [string]$ExpectedScriptPath,

        [Parameter(Mandatory)]
        [string]$ExpectedPowerShellExe
    )

    $issues = New-Object System.Collections.Generic.List[string]

    $action = $Task.Actions | Select-Object -First 1
    if ($null -eq $action) {
        $issues.Add('Task has no action.') | Out-Null
    }
    else {
        $actualExecute = [string]$action.Execute
        $actualArguments = [string]$action.Arguments

        if (($actualExecute -ine $ExpectedPowerShellExe) -and ((Split-Path -Leaf $actualExecute) -ine 'powershell.exe')) {
            $issues.Add("Action executable is not Windows PowerShell. Current: $actualExecute") | Out-Null
        }

        if ($actualArguments -notmatch [regex]::Escape($ExpectedScriptPath)) {
            $issues.Add("Action does not point to expected updater script. Current arguments: $actualArguments") | Out-Null
        }

        if ($actualArguments -notmatch '-NoProfile') {
            $issues.Add('Action is missing -NoProfile.') | Out-Null
        }

        if ($actualArguments -notmatch '-ExecutionPolicy\s+Bypass') {
            $issues.Add('Action is missing -ExecutionPolicy Bypass.') | Out-Null
        }
    }

    if ($Task.Principal.UserId -notmatch 'SYSTEM') {
        $issues.Add("Task principal is not SYSTEM. Current: $($Task.Principal.UserId)") | Out-Null
    }

    if ($Task.Principal.RunLevel -ne 'Highest') {
        $issues.Add("Task run level is not Highest. Current: $($Task.Principal.RunLevel)") | Out-Null
    }

    if (-not (Get-WeeklySundayTriggerSummary -Triggers $Task.Triggers)) {
        $issues.Add('Task trigger is not Sunday at 01:15.') | Out-Null
    }

    return $issues
}

function Ensure-CheckForUpdatedScriptsTask {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ScriptPath,

        [string]$TaskName = '01. Check for Updated Scripts'
    )

    $windowsPowerShellExe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'

    if (-not (Test-Path -LiteralPath $windowsPowerShellExe)) {
        Write-Status "Windows PowerShell executable not found: $windowsPowerShellExe" 'ERROR'
        return $false
    }

    if (-not (Test-Path -LiteralPath $ScriptPath)) {
        Write-Status "Cannot register task because updater script is missing: $ScriptPath" 'ERROR'
        return $false
    }

    try {
        Import-Module ScheduledTasks -ErrorAction Stop | Out-Null
    }
    catch {
        Write-Status "ScheduledTasks module is unavailable: $($_.Exception.Message)" 'ERROR'
        return $false
    }

    $argumentString = "-NoProfile -ExecutionPolicy Bypass -File `"$ScriptPath`""
    $needsRebuild = $true

    try {
        $existingTask = Get-ScheduledTask -TaskName $TaskName -ErrorAction Stop
        $issues = Test-CheckForUpdatedScriptsTaskCompliant -Task $existingTask -ExpectedScriptPath $ScriptPath -ExpectedPowerShellExe $windowsPowerShellExe

        if ($issues.Count -eq 0) {
            Write-Status "Scheduled task '$TaskName' is already configured correctly." 'OK'
            $needsRebuild = $false
        }
        else {
            Write-Status "Scheduled task '$TaskName' needs repair:" 'WARN'
            foreach ($issue in $issues) {
                Write-Status " - $issue" 'WARN'
            }
        }
    }
    catch {
        Write-Status "Scheduled task '$TaskName' does not exist and will be created." 'WARN'
    }

    if (-not $needsRebuild) {
        return $true
    }

    try {
        $action = New-ScheduledTaskAction -Execute $windowsPowerShellExe -Argument $argumentString
        $trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Sunday -WeeksInterval 1 -At ([datetime]::Today.Add([timespan]::Parse('01:15')))
        $principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
        $settings = New-ScheduledTaskSettingsSet `
            -AllowStartIfOnBatteries `
            -DontStopIfGoingOnBatteries `
            -StartWhenAvailable `
            -ExecutionTimeLimit (New-TimeSpan -Hours 12)

        Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Force | Out-Null
        Write-Status "Scheduled task '$TaskName' now points to: $ScriptPath" 'OK'
        Write-Status "Task action: $windowsPowerShellExe $argumentString" 'INFO'
        Write-Status "Task trigger: Sunday at 01:15 as SYSTEM with highest privileges." 'OK'
        return $true
    }
    catch {
        Write-Status "Failed to create/repair scheduled task '$TaskName': $($_.Exception.Message)" 'ERROR'
        return $false
    }
}


if (-not (Test-IsAdmin)) {
    Write-Host ""
    Write-Host "This script must be run as Administrator." -ForegroundColor Red
    exit 1
}

Write-Status "Initializing script..." 'INFO'
Enable-ClassicWindows11RightClickMenu
$shareUpdaterPath = Ensure-ShareScriptUpdater
if ($shareUpdaterPath) {
    [void](Ensure-CheckForUpdatedScriptsTask -ScriptPath $shareUpdaterPath)
}

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
# =====================================================================
# ScriptName: 01_Enable_Windows_Update_Services.ps1
# ScriptVersion: 1.5
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

function Ensure-ShareScriptUpdater {
    param(
        [string]$SourcePath = '\\filesvr\Labscripts\00_Update-Scripts-FromShare.ps1',
        [string]$DestinationFolder = 'C:\Scripts',
        [string]$DestinationFileName = '00_Update-Scripts-FromShare.ps1'
    )

    try {
        if (-not (Test-Path -LiteralPath $SourcePath)) {
            throw "Source updater script not found: $SourcePath"
        }

        if (-not (Test-Path -LiteralPath $DestinationFolder)) {
            New-Item -Path $DestinationFolder -ItemType Directory -Force | Out-Null
            Write-Status "Created destination folder: $DestinationFolder" 'OK'
        }

        $destinationPath = Join-Path $DestinationFolder $DestinationFileName
        Copy-Item -LiteralPath $SourcePath -Destination $destinationPath -Force
        Write-Status "Copied updater script to: $destinationPath" 'OK'

        try {
            Unblock-File -LiteralPath $destinationPath -ErrorAction SilentlyContinue
        }
        catch {
        }

        return $destinationPath
    }
    catch {
        Write-Status "Failed to stage share updater script: $($_.Exception.Message)" 'ERROR'
        return $null
    }
}

function Update-CheckForUpdatedScriptsTask {
    param(
        [Parameter(Mandatory)]
        [string]$ScriptPath,

        [string]$TaskName = '01. Check for Updated Scripts'
    )

    try {
        $task = Get-ScheduledTask -TaskName $TaskName -ErrorAction Stop
    }
    catch {
        Write-Status "Scheduled task not found: $TaskName" 'WARN'
        return $false
    }

    try {
        $existingAction = $task.Actions | Select-Object -First 1
        if ($null -eq $existingAction) {
            throw 'Scheduled task has no executable action to modify.'
        }

        $newArguments = "-NoProfile -ExecutionPolicy Bypass -File `"$ScriptPath`""
        $updatedAction = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument $newArguments

        $registerParams = @{
            TaskName = $task.TaskName
            Action   = $updatedAction
            Settings = $task.Settings
        }

        if ($null -ne $task.Description) {
            $registerParams['Description'] = $task.Description
        }

        if ($null -ne $task.Triggers -and $task.Triggers.Count -gt 0) {
            $registerParams['Trigger'] = $task.Triggers
        }

        if ($null -ne $task.Principal) {
            $registerParams['Principal'] = $task.Principal
        }

        Register-ScheduledTask @registerParams -Force | Out-Null
        Write-Status "Updated scheduled task '$TaskName' to run: $ScriptPath" 'OK'
        return $true
    }
    catch {
        Write-Status "Failed to update scheduled task '$TaskName' : $($_.Exception.Message)" 'ERROR'
        return $false
    }
}

if (-not (Test-IsAdmin)) {
    Write-Host ""
    Write-Host "This script must be run as Administrator." -ForegroundColor Red
    exit 1
}

Write-Status "Initializing script..." 'INFO'
$shareUpdaterPath = Ensure-ShareScriptUpdater
if ($shareUpdaterPath) {
    [void](Update-CheckForUpdatedScriptsTask -ScriptPath $shareUpdaterPath)
}

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
