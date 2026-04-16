# =====================================================================
# ScriptName: 00_Update-Scripts-FromGitHub.ps1
# ScriptVersion: 2.0
# LastUpdated: 2026-04-16
# =====================================================================

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

# ---------------------------
# Configuration
# ---------------------------
$PreferredSourceRoot    = '\\filesvr\Labscripts'
$FallbackSourceRoot     = '\\10.2.3.30\Labscripts'
$ShareScriptFileName    = '00_Update-Scripts-FromShare.ps1'
$LegacyScriptFileName   = '00_Update-Scripts-FromGitHub.ps1'
$LocalScripts           = 'C:\Scripts'
$LogFolder              = 'C:\Logs'
$BackupFolder           = 'C:\Scripts\Backup'
$ScheduledTaskName      = '01. Check for Updated Scripts'
$DefaultTaskStartTime   = '01:15'
$WindowsPowerShellExe   = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'

# ---------------------------
# Runtime State
# ---------------------------
$script:RunStart                 = Get-Date
$script:ComputerName             = $env:COMPUTERNAME
$script:YamlLogPath              = $null
$script:OverallResult            = 'Unknown'
$script:FailureMessage           = $null
$script:ActionHistory            = New-Object System.Collections.Generic.List[object]
$script:SelectedSourceRoot       = $null
$script:SelectedSourceLabel      = $null
$script:PreferredSourceReachable = $false
$script:PreferredSourceUsed      = $false
$script:FallbackSourceUsed       = $false
$script:FallbackReason           = $null
$script:TaskFound                = $false
$script:TaskPath                 = '\\'
$script:TaskActionBefore         = $null
$script:TaskActionAfter          = $null
$script:TaskTriggerSummary       = $null
$script:TaskMigrationStatus      = 'NotStarted'
$script:TaskMigrationMessage     = $null
$script:LocalShareScriptPath     = Join-Path $LocalScripts $ShareScriptFileName
$script:LegacyLocalScriptPath    = Join-Path $LocalScripts $LegacyScriptFileName
$script:DownloadedShareVersion   = $null
$script:DownloadedShareUpdated   = $null
$script:ShareScriptBackupPath    = $null

# ---------------------------
# Helpers
# ---------------------------
function Ensure-Folder {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -Path $Path -ItemType Directory -Force | Out-Null
    }
}

function Initialize-YamlLog {
    Ensure-Folder -Path $LogFolder

    $timestamp = $script:RunStart.ToString('yyyy-MM-dd_HH-mm-ss')
    $baseName = "$($script:ComputerName)-BootstrapShareUpdater-$timestamp"
    $script:YamlLogPath = Join-Path $LogFolder ($baseName + '.yaml')
}

function Write-Status {
    param(
        [Parameter(Mandatory)]
        [string]$Message,

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

    $script:ActionHistory.Add([PSCustomObject]@{
        Time    = $timestamp
        Level   = $Level
        Message = $Message
    }) | Out-Null
}

function ConvertTo-YamlScalar {
    param(
        [AllowNull()]$Value
    )

    if ($null -eq $Value) {
        return 'null'
    }

    if ($Value -is [bool]) {
        return $Value.ToString().ToLowerInvariant()
    }

    if ($Value -is [int] -or $Value -is [long] -or $Value -is [double] -or $Value -is [decimal]) {
        return [string]$Value
    }

    if ($Value -is [datetime]) {
        return "'" + $Value.ToString('yyyy-MM-dd HH:mm:ss') + "'"
    }

    $text = [string]$Value
    $text = $text -replace "`r", ' '
    $text = $text -replace "`n", ' '
    $text = $text -replace "'", "''"
    return "'" + $text + "'"
}

function Write-YamlLog {
    try {
        if ([string]::IsNullOrWhiteSpace($script:YamlLogPath)) {
            Initialize-YamlLog
        }

        $runEnd = Get-Date
        $duration = [math]::Round(($runEnd - $script:RunStart).TotalSeconds, 0)

        $lines = New-Object System.Collections.Generic.List[string]
        $lines.Add("computer_name: $(ConvertTo-YamlScalar $script:ComputerName)") | Out-Null
        $lines.Add("script_name: '00_Update-Scripts-FromGitHub.ps1'") | Out-Null
        $lines.Add("script_version: '2.0'") | Out-Null
        $lines.Add("run_started: $(ConvertTo-YamlScalar $script:RunStart)") | Out-Null
        $lines.Add("run_finished: $(ConvertTo-YamlScalar $runEnd)") | Out-Null
        $lines.Add("duration_seconds: $duration") | Out-Null
        $lines.Add("preferred_source_root: $(ConvertTo-YamlScalar $PreferredSourceRoot)") | Out-Null
        $lines.Add("fallback_source_root: $(ConvertTo-YamlScalar $FallbackSourceRoot)") | Out-Null
        $lines.Add("selected_source_root: $(ConvertTo-YamlScalar $script:SelectedSourceRoot)") | Out-Null
        $lines.Add("selected_source_label: $(ConvertTo-YamlScalar $script:SelectedSourceLabel)") | Out-Null
        $lines.Add("preferred_source_reachable: $(ConvertTo-YamlScalar $script:PreferredSourceReachable)") | Out-Null
        $lines.Add("preferred_source_used: $(ConvertTo-YamlScalar $script:PreferredSourceUsed)") | Out-Null
        $lines.Add("fallback_source_used: $(ConvertTo-YamlScalar $script:FallbackSourceUsed)") | Out-Null
        $lines.Add("fallback_reason: $(ConvertTo-YamlScalar $script:FallbackReason)") | Out-Null
        $lines.Add("local_scripts_path: $(ConvertTo-YamlScalar $LocalScripts)") | Out-Null
        $lines.Add("downloaded_share_script_path: $(ConvertTo-YamlScalar $script:LocalShareScriptPath)") | Out-Null
        $lines.Add("legacy_script_path: $(ConvertTo-YamlScalar $script:LegacyLocalScriptPath)") | Out-Null
        $lines.Add("downloaded_share_version: $(ConvertTo-YamlScalar $script:DownloadedShareVersion)") | Out-Null
        $lines.Add("downloaded_share_last_updated: $(ConvertTo-YamlScalar $script:DownloadedShareUpdated)") | Out-Null
        $lines.Add("downloaded_share_backup_path: $(ConvertTo-YamlScalar $script:ShareScriptBackupPath)") | Out-Null
        $lines.Add("scheduled_task_name: $(ConvertTo-YamlScalar $ScheduledTaskName)") | Out-Null
        $lines.Add("scheduled_task_found: $(ConvertTo-YamlScalar $script:TaskFound)") | Out-Null
        $lines.Add("scheduled_task_path: $(ConvertTo-YamlScalar $script:TaskPath)") | Out-Null
        $lines.Add("scheduled_task_action_before: $(ConvertTo-YamlScalar $script:TaskActionBefore)") | Out-Null
        $lines.Add("scheduled_task_action_after: $(ConvertTo-YamlScalar $script:TaskActionAfter)") | Out-Null
        $lines.Add("scheduled_task_trigger_summary: $(ConvertTo-YamlScalar $script:TaskTriggerSummary)") | Out-Null
        $lines.Add("scheduled_task_migration_status: $(ConvertTo-YamlScalar $script:TaskMigrationStatus)") | Out-Null
        $lines.Add("scheduled_task_migration_message: $(ConvertTo-YamlScalar $script:TaskMigrationMessage)") | Out-Null
        $lines.Add("yaml_log_path: $(ConvertTo-YamlScalar $script:YamlLogPath)") | Out-Null
        $lines.Add("overall_result: $(ConvertTo-YamlScalar $script:OverallResult)") | Out-Null
        $lines.Add("failure_message: $(ConvertTo-YamlScalar $script:FailureMessage)") | Out-Null
        $lines.Add('') | Out-Null
        $lines.Add('actions:') | Out-Null

        if ($script:ActionHistory.Count -gt 0) {
            foreach ($action in $script:ActionHistory) {
                $lines.Add('  -') | Out-Null
                $lines.Add("    time: $(ConvertTo-YamlScalar $action.Time)") | Out-Null
                $lines.Add("    level: $(ConvertTo-YamlScalar $action.Level)") | Out-Null
                $lines.Add("    message: $(ConvertTo-YamlScalar $action.Message)") | Out-Null
            }
        }
        else {
            $lines.Add('  []') | Out-Null
        }

        Set-Content -Path $script:YamlLogPath -Value $lines -Encoding UTF8
    }
    catch {
        Write-Warning "Failed to write YAML log: $($_.Exception.Message)"
    }
}

function Get-FileTextSafe {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        return $null
    }

    try {
        return Get-Content -LiteralPath $Path -Raw -Encoding UTF8
    }
    catch {
        try {
            return Get-Content -LiteralPath $Path -Raw
        }
        catch {
            throw "Failed reading file [$Path] : $($_.Exception.Message)"
        }
    }
}

function Get-ScriptHeaderValue {
    param(
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$Content,

        [Parameter(Mandatory)]
        [string]$HeaderName
    )

    if ([string]::IsNullOrWhiteSpace($Content)) {
        return $null
    }

    $normalized = $Content -replace "^\uFEFF", ''
    $normalized = $normalized -replace "`r`n", "`n"
    $normalized = $normalized -replace "`r", "`n"

    $patternLine = "(?im)^\s*#\s*" + [regex]::Escape($HeaderName) + "\s*:\s*([^\r\n]+?)\s*$"
    $matchLine = [regex]::Match($normalized, $patternLine)
    if ($matchLine.Success) {
        return $matchLine.Groups[1].Value.Trim()
    }

    $patternInline = "(?is)#\s*" + [regex]::Escape($HeaderName) + "\s*:\s*([^#\r\n]+)"
    $matchInline = [regex]::Match($normalized, $patternInline)
    if ($matchInline.Success) {
        return $matchInline.Groups[1].Value.Trim()
    }

    return $null
}

function Backup-File {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        return $null
    }

    Ensure-Folder -Path $BackupFolder

    $baseName   = [System.IO.Path]::GetFileNameWithoutExtension($Path)
    $extension  = [System.IO.Path]::GetExtension($Path)
    $timestamp  = Get-Date -Format 'yyyyMMdd_HHmmss'
    $backupName = "${baseName}_${timestamp}${extension}.bak"
    $backupPath = Join-Path $BackupFolder $backupName

    Copy-Item -LiteralPath $Path -Destination $backupPath -Force
    Write-Status "Backed up [$Path] to [$backupPath]" 'OK'
    return $backupPath
}

function Save-Utf8NoBom {
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter(Mandatory)]
        [string]$Content
    )

    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($Path, $Content, $utf8NoBom)
}

function Test-ShareReachable {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    try {
        return [bool](Test-Path -LiteralPath $Path -ErrorAction Stop)
    }
    catch {
        return $false
    }
}

function Initialize-SourceRoot {
    $script:PreferredSourceReachable = Test-ShareReachable -Path $PreferredSourceRoot

    if ($script:PreferredSourceReachable) {
        $script:SelectedSourceRoot  = $PreferredSourceRoot
        $script:SelectedSourceLabel = 'PreferredUNC'
        $script:PreferredSourceUsed = $true
        $script:FallbackSourceUsed  = $false
        $script:FallbackReason      = $null
        Write-Status "Using preferred source path: $PreferredSourceRoot" 'OK'
        return
    }

    Write-Status "Preferred source path unavailable: $PreferredSourceRoot" 'WARN'

    if (Test-ShareReachable -Path $FallbackSourceRoot) {
        $script:SelectedSourceRoot  = $FallbackSourceRoot
        $script:SelectedSourceLabel = 'FallbackIP'
        $script:PreferredSourceUsed = $false
        $script:FallbackSourceUsed  = $true
        $script:FallbackReason      = "Preferred path [$PreferredSourceRoot] could not be resolved or reached. Using fallback path [$FallbackSourceRoot]."
        Write-Status $script:FallbackReason 'WARN'
        return
    }

    throw "Neither source path is reachable. Preferred [$PreferredSourceRoot], Fallback [$FallbackSourceRoot]."
}

function Get-SourceFileContent {
    param(
        [Parameter(Mandatory)]
        [string]$FileName
    )

    if ([string]::IsNullOrWhiteSpace($script:SelectedSourceRoot)) {
        throw 'Source root has not been initialized.'
    }

    $sourcePath = Join-Path $script:SelectedSourceRoot $FileName
    if (-not (Test-Path -LiteralPath $sourcePath)) {
        throw "Source file not found: $sourcePath"
    }

    $content = Get-FileTextSafe -Path $sourcePath
    if ([string]::IsNullOrWhiteSpace($content)) {
        throw "Source file [$sourcePath] was empty."
    }

    return [PSCustomObject]@{
        Path    = $sourcePath
        Content = $content
        Label   = $script:SelectedSourceLabel
    }
}

function Download-ShareUpdater {
    $sourceFile = Get-SourceFileContent -FileName $ShareScriptFileName
    $shareVersion = Get-ScriptHeaderValue -Content $sourceFile.Content -HeaderName 'ScriptVersion'
    $shareLastUpdated = Get-ScriptHeaderValue -Content $sourceFile.Content -HeaderName 'LastUpdated'

    if ([string]::IsNullOrWhiteSpace($shareVersion)) {
        throw "Source file [$($sourceFile.Path)] is missing or has an unreadable '# ScriptVersion:' header."
    }

    if (Test-Path -LiteralPath $script:LocalShareScriptPath) {
        $script:ShareScriptBackupPath = Backup-File -Path $script:LocalShareScriptPath
    }

    Save-Utf8NoBom -Path $script:LocalShareScriptPath -Content $sourceFile.Content
    $script:DownloadedShareVersion = $shareVersion
    $script:DownloadedShareUpdated = $shareLastUpdated

    Write-Status "Downloaded [$ShareScriptFileName] version [$shareVersion] from [$($sourceFile.Path)] to [$script:LocalShareScriptPath]." 'OK'
}

function Get-TaskTriggerSummary {
    param(
        [AllowNull()]$Triggers
    )

    if ($null -eq $Triggers) {
        return $null
    }

    $parts = foreach ($trigger in @($Triggers)) {
        if ($null -eq $trigger) { continue }

        $type = [string]$trigger.CimClass.CimClassName
        $start = $null
        if ($trigger.PSObject.Properties.Name -contains 'StartBoundary') {
            $start = $trigger.StartBoundary
        }

        $days = $null
        if ($trigger.PSObject.Properties.Name -contains 'DaysOfWeek') {
            $days = ($trigger.DaysOfWeek | ForEach-Object { [string]$_ }) -join ','
        }

        if (-not [string]::IsNullOrWhiteSpace($days)) {
            "$type @ $start [$days]"
        }
        else {
            "$type @ $start"
        }
    }

    return ($parts -join '; ')
}

function Get-DefaultWeeklyTrigger {
    $startAt = [datetime]::Today.Add([timespan]::Parse($DefaultTaskStartTime))
    return New-ScheduledTaskTrigger -Weekly -DaysOfWeek Sunday -WeeksInterval 1 -At $startAt
}

function Get-DefaultPrincipal {
    return New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
}

function Get-DefaultSettings {
    return New-ScheduledTaskSettingsSet `
        -AllowStartIfOnBatteries `
        -DontStopIfGoingOnBatteries `
        -StartWhenAvailable `
        -ExecutionTimeLimit (New-TimeSpan -Hours 12)
}

function Update-ScheduledTaskToShareScript {
    Import-Module ScheduledTasks -ErrorAction Stop | Out-Null

    $newArgument = "-NoProfile -ExecutionPolicy Bypass -File `"$script:LocalShareScriptPath`""
    $newAction = New-ScheduledTaskAction -Execute $WindowsPowerShellExe -Argument $newArgument
    $script:TaskActionAfter = "$WindowsPowerShellExe $newArgument"

    $existingTask = Get-ScheduledTask -TaskName $ScheduledTaskName -ErrorAction SilentlyContinue
    if ($null -ne $existingTask) {
        $script:TaskFound = $true
        $script:TaskPath = $existingTask.TaskPath

        $existingExec = $null
        $existingArgs = $null
        if ($existingTask.Actions.Count -gt 0) {
            $existingExec = $existingTask.Actions[0].Execute
            $existingArgs = $existingTask.Actions[0].Arguments
        }
            if ([string]::IsNullOrWhiteSpace($existingArgs)) {
            $script:TaskActionBefore = $existingExec
        }
        else {
            $script:TaskActionBefore = "$existingExec $existingArgs"
        }
        $script:TaskTriggerSummary = Get-TaskTriggerSummary -Triggers $existingTask.Triggers

        $triggersToUse = @($existingTask.Triggers)
        if (-not $triggersToUse -or $triggersToUse.Count -eq 0) {
            $triggersToUse = @(Get-DefaultWeeklyTrigger)
            Write-Status "Existing task [$ScheduledTaskName] had no usable trigger. Applying default Sunday schedule [$DefaultTaskStartTime]." 'WARN'
        }

        try {
            Set-ScheduledTask -TaskName $existingTask.TaskName -TaskPath $existingTask.TaskPath -Action $newAction -Trigger $triggersToUse -Settings $existingTask.Settings -Principal $existingTask.Principal -ErrorAction Stop | Out-Null
            $script:TaskMigrationStatus = 'UpdatedExistingTask'
            $script:TaskMigrationMessage = 'Existing scheduled task was updated to point to 00_Update-Scripts-FromShare.ps1 while preserving the existing trigger schedule.'
            Write-Status "Updated existing scheduled task [$ScheduledTaskName] to use [$script:LocalShareScriptPath]." 'OK'
            return
        }
        catch {
            Write-Status "Set-ScheduledTask failed for [$ScheduledTaskName]. Falling back to re-register: $($_.Exception.Message)" 'WARN'

            $settingsToUse = $existingTask.Settings
            if ($null -eq $settingsToUse) {
                $settingsToUse = Get-DefaultSettings
            }

            $principalToUse = $existingTask.Principal
            if ($null -eq $principalToUse) {
                $principalToUse = Get-DefaultPrincipal
            }

            Register-ScheduledTask -TaskName $existingTask.TaskName -TaskPath $existingTask.TaskPath -Action $newAction -Trigger $triggersToUse -Settings $settingsToUse -Principal $principalToUse -Force -ErrorAction Stop | Out-Null
            $script:TaskMigrationStatus = 'ReRegisteredExistingTask'
            $script:TaskMigrationMessage = 'Existing scheduled task was re-registered to point to 00_Update-Scripts-FromShare.ps1 while preserving the existing trigger schedule.'
            Write-Status "Re-registered existing scheduled task [$ScheduledTaskName] to use [$script:LocalShareScriptPath]." 'OK'
            return
        }
    }

    $defaultTrigger = Get-DefaultWeeklyTrigger
    $script:TaskFound = $false
    $script:TaskPath = '\\'
    $script:TaskActionBefore = $null
    $script:TaskTriggerSummary = Get-TaskTriggerSummary -Triggers @($defaultTrigger)

    Register-ScheduledTask -TaskName $ScheduledTaskName -TaskPath '\\' -Action $newAction -Trigger $defaultTrigger -Principal (Get-DefaultPrincipal) -Settings (Get-DefaultSettings) -Force -ErrorAction Stop | Out-Null
    $script:TaskMigrationStatus = 'CreatedDefaultTask'
    $script:TaskMigrationMessage = "Scheduled task did not already exist, so a default Sunday task was created for $DefaultTaskStartTime using 00_Update-Scripts-FromShare.ps1."
    Write-Status "Created scheduled task [$ScheduledTaskName] for Sunday at [$DefaultTaskStartTime] using [$script:LocalShareScriptPath]." 'OK'
}

# ---------------------------
# Main
# ---------------------------
Initialize-YamlLog

try {
    Ensure-Folder -Path $LocalScripts
    Ensure-Folder -Path $LogFolder
    Ensure-Folder -Path $BackupFolder

    Write-Status 'Initializing bootstrap migration from GitHub updater to share updater...' 'INFO'
    Write-Status "Preferred source: $PreferredSourceRoot" 'INFO'
    Write-Status "Fallback source: $FallbackSourceRoot" 'INFO'
    Write-Status "Target local share script path: $script:LocalShareScriptPath" 'INFO'
    Write-Status "Target scheduled task: $ScheduledTaskName" 'INFO'

    Initialize-SourceRoot
    Download-ShareUpdater
    Update-ScheduledTaskToShareScript

    $script:OverallResult = 'Succeeded'
    Write-Status 'Bootstrap migration completed successfully.' 'OK'
    Write-YamlLog
    exit 0
}
catch {
    $script:FailureMessage = $_.Exception.Message
    $script:OverallResult = 'Failed'
    if ($script:TaskMigrationStatus -eq 'NotStarted') {
        $script:TaskMigrationStatus = 'Failed'
        $script:TaskMigrationMessage = $_.Exception.Message
    }
    Write-Status "Fatal error: $($_.Exception.Message)" 'ERROR'
    Write-YamlLog
    exit 1
}
