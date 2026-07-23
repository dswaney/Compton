# =====================================================================
# ScriptName: 00_Update-Scripts-FromShare.ps1
# ScriptVersion: 3.2
# LastUpdated: 2026-07-23
# Purpose:
#   - Dynamically synchronize PowerShell scripts from the central share.
#   - Update and relaunch itself when a newer/different updater is found.
#   - Automatically deploy newly added .ps1 files.
#   - Self-heal deleted or missing local scripts from the active share.
#   - Verify every restored or updated script using SHA-256.
#   - Run Register-Tasks_SYSTEM.ps1 after synchronization so missing or
#     changed maintenance tasks are reconciled.
# =====================================================================

[CmdletBinding()]
param(
    [switch]$Relaunched
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

# ---------------------------
# Configuration
# ---------------------------
$PreferredSourceRoot = '\\filesvr\Labscripts'
$FallbackSourceRoot  = '\\10.2.3.30\Labscripts'

$LocalScripts = 'C:\Scripts'
$LogFolder    = 'C:\Logs'
$BackupFolder = 'C:\Scripts\Backup'
$LogPath      = Join-Path $LogFolder '00_Update-Scripts-FromShare.log'

$UpdaterFileName      = '00_Update-Scripts-FromShare.ps1'
$RegisterTasksName    = 'Register-Tasks_SYSTEM.ps1'
$RequiredManagedFiles = @(
    'Register-Tasks_SYSTEM.ps1',
    '12. Enable-SystemRestore-And-Create-RestorePoint.ps1'
)
$WindowsPowerShellExe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'

# Files or folders in the share root that should not be treated as managed scripts.
$ExcludedFileNames = @(
    'Register-Tasks_SYSTEM_v2.0.ps1'
)

function Write-Status {
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
        Add-Content -LiteralPath $LogPath -Value $line -Encoding UTF8
    }
    catch {
    }
}

function Ensure-Folder {
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
        New-Item -Path $Path -ItemType Directory -Force | Out-Null
    }
}

function Test-ShareRoot {
    param([Parameter(Mandatory)][string]$Path)

    try {
        return Test-Path -LiteralPath $Path -PathType Container -ErrorAction Stop
    }
    catch {
        return $false
    }
}

function Get-ActiveSourceRoot {
    if (Test-ShareRoot -Path $PreferredSourceRoot) {
        Write-Status "Using preferred source: $PreferredSourceRoot" 'OK'
        return $PreferredSourceRoot
    }

    Write-Status "Preferred source is unavailable: $PreferredSourceRoot" 'WARN'

    if (Test-ShareRoot -Path $FallbackSourceRoot) {
        Write-Status "Using fallback source: $FallbackSourceRoot" 'OK'
        return $FallbackSourceRoot
    }

    throw "Neither script source is available. Preferred: $PreferredSourceRoot | Fallback: $FallbackSourceRoot"
}

function Get-FileSha256 {
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return $null
    }

    return (Get-FileHash -LiteralPath $Path -Algorithm SHA256 -ErrorAction Stop).Hash
}

function Get-ScriptVersionText {
    param([Parameter(Mandatory)][string]$Path)

    try {
        $match = Select-String `
            -LiteralPath $Path `
            -Pattern '^\s*#\s*ScriptVersion\s*:\s*(.+?)\s*$' `
            -ErrorAction Stop |
            Select-Object -First 1

        if ($match) {
            return $match.Matches[0].Groups[1].Value.Trim()
        }
    }
    catch {
    }

    return 'Unknown'
}

function Backup-LocalFile {
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return $null
    }

    Ensure-Folder -Path $BackupFolder

    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $baseName = [IO.Path]::GetFileNameWithoutExtension($Path)
    $extension = [IO.Path]::GetExtension($Path)
    $backupPath = Join-Path $BackupFolder "${baseName}_${stamp}${extension}"

    Copy-Item -LiteralPath $Path -Destination $backupPath -Force
    return $backupPath
}

function Copy-FileAtomically {
    param(
        [Parameter(Mandatory)][string]$SourcePath,
        [Parameter(Mandatory)][string]$DestinationPath
    )

    $destinationDirectory = Split-Path -Path $DestinationPath -Parent
    Ensure-Folder -Path $destinationDirectory

    $temporaryPath = "$DestinationPath.new"
    Copy-Item -LiteralPath $SourcePath -Destination $temporaryPath -Force

    if (Test-Path -LiteralPath $DestinationPath -PathType Leaf) {
        Move-Item -LiteralPath $temporaryPath -Destination $DestinationPath -Force
    }
    else {
        Rename-Item -LiteralPath $temporaryPath -NewName ([IO.Path]::GetFileName($DestinationPath)) -Force
    }
}

function Invoke-TaskReconciliation {
    $registerScript = Join-Path $LocalScripts $RegisterTasksName

    if (-not (Test-Path -LiteralPath $registerScript -PathType Leaf)) {
        Write-Status "Task reconciliation script is not present: $registerScript" 'WARN'
        return $false
    }

    Write-Status "Reconciling scheduled tasks with $RegisterTasksName." 'ACTION'

    $process = Start-Process `
        -FilePath $WindowsPowerShellExe `
        -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$registerScript`"" `
        -Wait `
        -PassThru `
        -WindowStyle Hidden

    if ($process.ExitCode -ne 0) {
        Write-Status "Task reconciliation returned exit code $($process.ExitCode)." 'ERROR'
        return $false
    }

    Write-Status 'Scheduled task reconciliation completed successfully.' 'OK'
    return $true
}

try {
    Ensure-Folder -Path $LocalScripts
    Ensure-Folder -Path $LogFolder
    Ensure-Folder -Path $BackupFolder

    Write-Status 'Starting dynamic script synchronization.' 'INFO'

    $sourceRoot = Get-ActiveSourceRoot
    $remoteScripts = @(
        Get-ChildItem `
            -LiteralPath $sourceRoot `
            -Filter '*.ps1' `
            -File `
            -ErrorAction Stop |
        Where-Object { $_.Name -notin $ExcludedFileNames } |
        Sort-Object {
            if ($_.Name -ieq $UpdaterFileName) { 0 } else { 1 }
        }, Name
    )

    if (-not $remoteScripts) {
        throw "No PowerShell scripts were found in the source root: $sourceRoot"
    }

    Write-Status "Discovered $($remoteScripts.Count) managed script(s) on the share." 'INFO'

    $updatedFiles  = New-Object System.Collections.Generic.List[string]
    $restoredFiles = New-Object System.Collections.Generic.List[string]
    $selfUpdated = $false

    foreach ($remoteFile in $remoteScripts) {
        $localPath = Join-Path $LocalScripts $remoteFile.Name
        $remoteHash = Get-FileSha256 -Path $remoteFile.FullName
        $localHash = Get-FileSha256 -Path $localPath

        if ($remoteHash -eq $localHash -and $null -ne $localHash) {
            Write-Status "$($remoteFile.Name) is current." 'OK'
            continue
        }

        $remoteVersion = Get-ScriptVersionText -Path $remoteFile.FullName
        $localExists = Test-Path -LiteralPath $localPath -PathType Leaf
        $localVersion = if ($localExists) {
            Get-ScriptVersionText -Path $localPath
        }
        else {
            'Missing'
        }

        if (-not $localExists) {
            Write-Status "Missing managed script detected: $localPath" 'WARN'
            Write-Status "Restoring $($remoteFile.Name) from $sourceRoot." 'ACTION'
        }
        else {
            Write-Status "Updating $($remoteFile.Name): local [$localVersion], share [$remoteVersion]." 'ACTION'

            $backupPath = Backup-LocalFile -Path $localPath
            if ($backupPath) {
                Write-Status "Backup created: $backupPath" 'INFO'
            }
        }

        Copy-FileAtomically -SourcePath $remoteFile.FullName -DestinationPath $localPath

        if (-not (Test-Path -LiteralPath $localPath -PathType Leaf)) {
            throw "The destination file is still missing after copying $($remoteFile.Name)."
        }

        $copiedHash = Get-FileSha256 -Path $localPath
        if ([string]::IsNullOrWhiteSpace($copiedHash) -or $copiedHash -ne $remoteHash) {
            throw "SHA-256 verification failed after copying $($remoteFile.Name)."
        }

        if ($localExists) {
            [void]$updatedFiles.Add($remoteFile.Name)
            Write-Status "Updated and hash-verified: $($remoteFile.Name)" 'OK'
        }
        else {
            [void]$restoredFiles.Add($remoteFile.Name)
            Write-Status "Restored and hash-verified: $($remoteFile.Name)" 'OK'
        }

        if ($remoteFile.Name -ieq $UpdaterFileName) {
            $selfUpdated = $true
            break
        }
    }

    # Stop using the old in-memory updater immediately after replacing it.
    if ($selfUpdated -and -not $Relaunched) {
        $newUpdaterPath = Join-Path $LocalScripts $UpdaterFileName
        Write-Status 'The updater changed. Relaunching the new local updater now.' 'ACTION'

        Start-Process `
            -FilePath $WindowsPowerShellExe `
            -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$newUpdaterPath`" -Relaunched" `
            -WindowStyle Hidden | Out-Null

        exit 0
    }

    if ($restoredFiles.Count -gt 0) {
        Write-Status "Self-healed missing files: $($restoredFiles -join ', ')" 'OK'
    }

    if ($updatedFiles.Count -gt 0) {
        Write-Status "Updated changed files: $($updatedFiles -join ', ')" 'OK'
    }

    if ($updatedFiles.Count -eq 0 -and $restoredFiles.Count -eq 0) {
        Write-Status 'All managed scripts are already synchronized.' 'OK'
    }

    $missingRequiredFiles = @(
        foreach ($requiredFile in $RequiredManagedFiles) {
            $requiredPath = Join-Path $LocalScripts $requiredFile
            if (-not (Test-Path -LiteralPath $requiredPath -PathType Leaf)) {
                $requiredFile
            }
        }
    )

    if ($missingRequiredFiles.Count -gt 0) {
        throw "Required managed file(s) are missing after synchronization: $($missingRequiredFiles -join ', ')"
    }

    Write-Status 'All required managed files are present locally.' 'OK'

    $tasksOk = Invoke-TaskReconciliation
    if (-not $tasksOk) {
        exit 2
    }

    Write-Status 'Script synchronization and task reconciliation completed.' 'OK'
    exit 0
}
catch {
    Write-Status "Fatal error: $($_.Exception.Message)" 'ERROR'
    exit 1
}
