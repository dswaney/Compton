# =====================================================================
# ScriptName: 00_Update-Scripts-FromShare.ps1
# ScriptVersion: 2.2
# LastUpdated: 2026-04-28
# =====================================================================

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

# ---------------------------
# Configuration
# ---------------------------
$PreferredSourceRoot = '\\filesvr\Labscripts'
$FallbackSourceRoot  = '\\10.2.3.30\Labscripts'

$LocalScripts  = 'C:\Scripts'
$LogFolder     = 'C:\Logs'
$BackupFolder  = 'C:\Scripts\Backup'

$ScriptFiles = @(
    '00_Update-Scripts-FromShare.ps1',
    '01_Enable_Windows_Update_Services.ps1',
    '02_Remove_User_Profiles.ps1',
    '03_Weekend_Apps_Update.ps1',
    '04_Update_Edge_Silent.ps1',
    '05_Weekend_HP_Drivers_Update.ps1',
    '06_Weekend_Windows_Updates.ps1',
    '07_Force_Reboot_Install_Updates.ps1',
    '08_System_Repair.ps1',
    '09_Disable_Windows_Update_Services.ps1',
    'Register-Tasks_SYSTEM_v2.0.ps1'
)

$LegacyFilesToDelete = @(
    '00_Update-Scripts-FromGithub.ps1',
    'Register-Tasks_SYSTEM.ps1'
)

# ---------------------------
# Runtime State
# ---------------------------
$script:RunStart             = Get-Date
$script:ComputerName         = $env:COMPUTERNAME
$script:YamlLogPath          = $null
$script:OverallResult        = 'Unknown'
$script:FailureMessage       = $null
$script:ActionHistory        = New-Object System.Collections.Generic.List[object]
$script:FileResults          = New-Object System.Collections.Generic.List[object]
$script:SelectedSourceRoot   = $null
$script:SelectedSourceLabel  = $null
$script:PreferredSourceUsed  = $false
$script:FallbackSourceUsed   = $false
$script:PreferredSourceReachable = $false
$script:FallbackReason       = $null

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
    $baseName = "$($script:ComputerName)-UpdateScriptsFromShare-$timestamp"
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

        $updatedCount = @($script:FileResults | Where-Object { $_.Status -eq 'Updated' }).Count
        $currentCount = @($script:FileResults | Where-Object { $_.Status -eq 'Current' }).Count
        $downloadedMissingCount = @($script:FileResults | Where-Object { $_.Status -eq 'DownloadedMissing' }).Count
        $errorCount = @($script:FileResults | Where-Object { $_.Status -eq 'Error' }).Count

        $lines = New-Object System.Collections.Generic.List[string]

        $lines.Add("computer_name: $(ConvertTo-YamlScalar $script:ComputerName)") | Out-Null
        $lines.Add("script_name: '00_Update-Scripts-FromShare.ps1'") | Out-Null
        $lines.Add("script_version: '2.2'") | Out-Null
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
        $lines.Add("backup_folder: $(ConvertTo-YamlScalar $BackupFolder)") | Out-Null
        $lines.Add("yaml_log_path: $(ConvertTo-YamlScalar $script:YamlLogPath)") | Out-Null
        $lines.Add("overall_result: $(ConvertTo-YamlScalar $script:OverallResult)") | Out-Null
        $lines.Add("failure_message: $(ConvertTo-YamlScalar $script:FailureMessage)") | Out-Null
        $lines.Add('') | Out-Null

        $lines.Add('summary:') | Out-Null
        $lines.Add("  updated: $updatedCount") | Out-Null
        $lines.Add("  current: $currentCount") | Out-Null
        $lines.Add("  downloaded_missing: $downloadedMissingCount") | Out-Null
        $lines.Add("  errors: $errorCount") | Out-Null
        $lines.Add('') | Out-Null

        $lines.Add('file_results:') | Out-Null
        if ($script:FileResults.Count -gt 0) {
            foreach ($item in $script:FileResults) {
                $lines.Add('  -') | Out-Null
                $lines.Add("    file_name: $(ConvertTo-YamlScalar $item.FileName)") | Out-Null
                $lines.Add("    local_path: $(ConvertTo-YamlScalar $item.LocalPath)") | Out-Null
                $lines.Add("    source_path: $(ConvertTo-YamlScalar $item.SourcePath)") | Out-Null
                $lines.Add("    source_label: $(ConvertTo-YamlScalar $item.SourceLabel)") | Out-Null
                $lines.Add("    status: $(ConvertTo-YamlScalar $item.Status)") | Out-Null
                $lines.Add("    local_version: $(ConvertTo-YamlScalar $item.LocalVersion)") | Out-Null
                $lines.Add("    remote_version: $(ConvertTo-YamlScalar $item.RemoteVersion)") | Out-Null
                $lines.Add("    local_last_updated: $(ConvertTo-YamlScalar $item.LocalLastUpdated)") | Out-Null
                $lines.Add("    remote_last_updated: $(ConvertTo-YamlScalar $item.RemoteLastUpdated)") | Out-Null
                $lines.Add("    backup_path: $(ConvertTo-YamlScalar $item.BackupPath)") | Out-Null
                $lines.Add("    message: $(ConvertTo-YamlScalar $item.Message)") | Out-Null
            }
        }
        else {
            $lines.Add('  []') | Out-Null
        }

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

function Convert-ToVersionObject {
    param(
        [Parameter(Mandatory)]
        [string]$VersionText
    )

    try {
        return [version]$VersionText.Trim()
    }
    catch {
        $clean = ($VersionText -replace '[^\d\.]', '').Trim('.')
        if ([string]::IsNullOrWhiteSpace($clean)) {
            return [version]'0.0'
        }

        try {
            return [version]$clean
        }
        catch {
            return [version]'0.0'
        }
    }
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

function Remove-LegacyScriptFiles {
    foreach ($legacyFile in $LegacyFilesToDelete) {
        $legacyPath = Join-Path $LocalScripts $legacyFile

        if (Test-Path -LiteralPath $legacyPath) {
            try {
                Remove-Item -LiteralPath $legacyPath -Force -ErrorAction Stop
                Write-Status "Deleted legacy script: $legacyPath" 'OK'
            }
            catch {
                Write-Status "Failed deleting legacy script [$legacyPath] : $($_.Exception.Message)" 'ERROR'
                $script:FileResults.Add([PSCustomObject]@{
                    FileName          = $legacyFile
                    LocalPath         = $legacyPath
                    SourcePath        = $null
                    SourceLabel       = 'LegacyCleanup'
                    Status            = 'Error'
                    LocalVersion      = $null
                    RemoteVersion     = $null
                    LocalLastUpdated  = $null
                    RemoteLastUpdated = $null
                    BackupPath        = $null
                    Message           = "Failed deleting legacy script: $($_.Exception.Message)"
                }) | Out-Null
            }
        }
        else {
            Write-Status "Legacy script not present, no deletion needed: $legacyPath" 'INFO'
        }
    }
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

# ---------------------------
# Main
# ---------------------------
Initialize-YamlLog

try {
    Ensure-Folder -Path $LocalScripts
    Ensure-Folder -Path $LogFolder

    Write-Status "Initializing script update check from network share..." 'INFO'
    Write-Status "Preferred source: $PreferredSourceRoot" 'INFO'
    Write-Status "Fallback source: $FallbackSourceRoot" 'INFO'
    Write-Status "Local script folder: $LocalScripts" 'INFO'

    Remove-LegacyScriptFiles

    Initialize-SourceRoot

    foreach ($file in $ScriptFiles) {
        $localPath = Join-Path $LocalScripts $file

        Write-Status "Checking file: $file" 'INFO'

        try {
            $sourceFile = Get-SourceFileContent -FileName $file
            $remoteContent = $sourceFile.Content
            $remotePath = $sourceFile.Path
            $remoteLabel = $sourceFile.Label

            $remoteVersionText = Get-ScriptHeaderValue -Content $remoteContent -HeaderName 'ScriptVersion'
            $remoteLastUpdated = Get-ScriptHeaderValue -Content $remoteContent -HeaderName 'LastUpdated'

            if ([string]::IsNullOrWhiteSpace($remoteVersionText)) {
                throw "Source file [$remotePath] is missing or has an unreadable '# ScriptVersion:' header."
            }

            $remoteVersion = Convert-ToVersionObject -VersionText $remoteVersionText

            $localContent = Get-FileTextSafe -Path $localPath

            if ($null -eq $localContent) {
                Write-Status "Local file missing. Copying [$file] version [$remoteVersionText] from [$remotePath]." 'WARN'
                Save-Utf8NoBom -Path $localPath -Content $remoteContent
                Write-Status "Copied new local file: $localPath" 'OK'

                $script:FileResults.Add([PSCustomObject]@{
                    FileName          = $file
                    LocalPath         = $localPath
                    SourcePath        = $remotePath
                    SourceLabel       = $remoteLabel
                    Status            = 'DownloadedMissing'
                    LocalVersion      = $null
                    RemoteVersion     = $remoteVersionText
                    LocalLastUpdated  = $null
                    RemoteLastUpdated = $remoteLastUpdated
                    BackupPath        = $null
                    Message           = "Local file was missing and was copied from [$remotePath]."
                }) | Out-Null

                continue
            }

            $localVersionText = Get-ScriptHeaderValue -Content $localContent -HeaderName 'ScriptVersion'
            $localLastUpdated = Get-ScriptHeaderValue -Content $localContent -HeaderName 'LastUpdated'

            if ([string]::IsNullOrWhiteSpace($localVersionText)) {
                Write-Status "Local file [$file] is missing ScriptVersion header. Treating local version as 0.0." 'WARN'
                $localVersionText = '0.0'
            }

            $localVersion = Convert-ToVersionObject -VersionText $localVersionText

            Write-Status "Local version: [$localVersionText] | Source version: [$remoteVersionText]" 'INFO'

            if ($remoteVersion -gt $localVersion) {
                Write-Status "Source version is newer for [$file]. Updating local copy..." 'INFO'
                $backupPath = Backup-File -Path $localPath
                Save-Utf8NoBom -Path $localPath -Content $remoteContent
                Write-Status "Updated [$file] from version [$localVersionText] to [$remoteVersionText]" 'OK'

                $script:FileResults.Add([PSCustomObject]@{
                    FileName          = $file
                    LocalPath         = $localPath
                    SourcePath        = $remotePath
                    SourceLabel       = $remoteLabel
                    Status            = 'Updated'
                    LocalVersion      = $localVersionText
                    RemoteVersion     = $remoteVersionText
                    LocalLastUpdated  = $localLastUpdated
                    RemoteLastUpdated = $remoteLastUpdated
                    BackupPath        = $backupPath
                    Message           = "Updated local file from $localVersionText to $remoteVersionText using [$remotePath]."
                }) | Out-Null
            }
            else {
                Write-Status "[$file] is current. Local version [$localVersionText], Source version [$remoteVersionText]." 'OK'

                $script:FileResults.Add([PSCustomObject]@{
                    FileName          = $file
                    LocalPath         = $localPath
                    SourcePath        = $remotePath
                    SourceLabel       = $remoteLabel
                    Status            = 'Current'
                    LocalVersion      = $localVersionText
                    RemoteVersion     = $remoteVersionText
                    LocalLastUpdated  = $localLastUpdated
                    RemoteLastUpdated = $remoteLastUpdated
                    BackupPath        = $null
                    Message           = 'Local file is already current.'
                }) | Out-Null
            }
        }
        catch {
            Write-Status "Failed processing [$file] : $($_.Exception.Message)" 'ERROR'

            $script:FileResults.Add([PSCustomObject]@{
                FileName          = $file
                LocalPath         = $localPath
                SourcePath        = if ($script:SelectedSourceRoot) { Join-Path $script:SelectedSourceRoot $file } else { $null }
                SourceLabel       = $script:SelectedSourceLabel
                Status            = 'Error'
                LocalVersion      = $null
                RemoteVersion     = $null
                LocalLastUpdated  = $null
                RemoteLastUpdated = $null
                BackupPath        = $null
                Message           = $_.Exception.Message
            }) | Out-Null
        }
    }

    $errorCount = @($script:FileResults | Where-Object { $_.Status -eq 'Error' }).Count
    if ($errorCount -gt 0) {
        $script:OverallResult = 'CompletedWithErrors'
    }
    else {
        $script:OverallResult = 'Succeeded'
    }

    Write-Status "Update check complete." 'INFO'
    Write-YamlLog

    if ($errorCount -gt 0) {
        exit 2
    }
    else {
        exit 0
    }
}
catch {
    $script:FailureMessage = $_.Exception.Message
    $script:OverallResult = 'Failed'
    Write-Status "Fatal error: $($_.Exception.Message)" 'ERROR'
    Write-YamlLog
    exit 1
}
