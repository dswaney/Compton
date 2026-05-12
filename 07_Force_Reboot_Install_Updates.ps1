# =====================================================================
# ScriptName: 07_Force_Reboot_Install_Updates.ps1
# ScriptVersion: 1.9
# LastUpdated: 2026-04-29
# ChangeLog: Final scheduled-task hardening: mutex single-instance lock, isolated State folder, write-through atomic JSON state saves, state-file backup/repair, and reboot-safe flush before shutdown.
# =====================================================================

[CmdletBinding()]
param(
    [int]$RebootDelaySeconds = 30,
    [string]$LogDirectory = 'C:\Logs',
    [string]$StateDirectory = 'C:\ProgramData\MISMaintenance\State',
    [string]$StateFileName = '07_Force_Reboot_Install_Updates_State.json'
)

$ErrorActionPreference = 'Stop'

$script:RunStart         = Get-Date
$script:ComputerName     = $env:COMPUTERNAME
$script:StateFilePath    = Join-Path $StateDirectory $StateFileName
$script:YamlLogPath      = $null
$script:OverallResult    = 'Unknown'
$script:FailureMessage   = $null
$script:CurrentFlags     = @()
$script:ClearResults     = @()
$script:ActionHistory    = @()
$script:CurrentStage     = 0
$script:RebootIssued     = $false
$script:RebootReason     = $null
$script:MutexName        = 'Global\MIS_07_Force_Reboot_Install_Updates'
$script:Mutex            = $null
$script:MutexAcquired    = $false

function Ensure-Folder {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -Path $Path -ItemType Directory -Force | Out-Null
    }
}

function Initialize-Paths {
    Ensure-Folder -Path $LogDirectory
    Ensure-Folder -Path $StateDirectory

    $timestamp = $script:RunStart.ToString('yyyy-MM-dd_HH-mm-ss')
    $baseName = "$($script:ComputerName)-ForceRebootInstallUpdates-$timestamp"
    $script:YamlLogPath = Join-Path $LogDirectory ($baseName + '.yaml')
}

function Initialize-SingleInstanceLock {
    try {
        $createdNew = $false
        $script:Mutex = New-Object System.Threading.Mutex($false, $script:MutexName, [ref]$createdNew)

        if (-not $script:Mutex.WaitOne(0)) {
            Write-Status 'Another instance of Script 07 is already running. Exiting to prevent state-file corruption.' 'WARN'
            $script:OverallResult = 'SkippedAlreadyRunning'
            Write-YamlLog
            exit 0
        }

        $script:MutexAcquired = $true
        Write-Status 'Single-instance lock acquired.' 'INFO'
    }
    catch {
        Write-Status "Failed to acquire single-instance lock: $($_.Exception.Message)" 'ERROR'
        throw
    }
}

function Release-SingleInstanceLock {
    try {
        if ($script:MutexAcquired -and $null -ne $script:Mutex) {
            [void]$script:Mutex.ReleaseMutex()
            $script:MutexAcquired = $false
            Write-Status 'Single-instance lock released.' 'INFO'
        }
    }
    catch {
        Write-Warning "Failed to release single-instance lock: $($_.Exception.Message)"
    }
    finally {
        if ($null -ne $script:Mutex) {
            $script:Mutex.Dispose()
            $script:Mutex = $null
        }
    }
}

function Confirm-FileWriteToDisk {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        return
    }

    $stream = $null
    try {
        $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::Read)
        $stream.Flush($true)
    }
    catch {
        Write-Status "Unable to force-flush file to disk [$Path]: $($_.Exception.Message)" 'WARN'
    }
    finally {
        if ($null -ne $stream) {
            $stream.Dispose()
        }
    }
}

function Confirm-StateWriteBeforeReboot {
    try {
        Confirm-FileWriteToDisk -Path $script:StateFilePath
        if ($script:YamlLogPath) {
            Confirm-FileWriteToDisk -Path $script:YamlLogPath
        }
        Start-Sleep -Seconds 3
        Write-Status 'State and YAML log writes confirmed before reboot.' 'INFO'
    }
    catch {
        Write-Status "Pre-reboot write confirmation failed: $($_.Exception.Message)" 'WARN'
    }
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

    $script:ActionHistory += [PSCustomObject]@{
        Time    = $timestamp
        Level   = $Level
        Message = $Message
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

function New-FlagRecord {
    param(
        [string]$Name,
        [string]$Type,
        [string]$Path,
        [string]$ValueName,
        [string]$Details
    )

    [PSCustomObject]@{
        Name      = $Name
        Type      = $Type
        Path      = $Path
        ValueName = $ValueName
        Details   = $Details
    }
}

function New-ClearResult {
    param(
        [string]$Name,
        [string]$Path,
        [string]$Action,
        [string]$Status,
        [string]$Message
    )

    [PSCustomObject]@{
        Name    = $Name
        Path    = $Path
        Action  = $Action
        Status  = $Status
        Message = $Message
    }
}

function Test-RegistryKeyExists {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    try {
        Test-Path -LiteralPath $Path
    }
    catch {
        $false
    }
}

function Get-RegistryValueSafe {
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter(Mandatory)]
        [string]$Name
    )

    try {
        $item = Get-ItemProperty -LiteralPath $Path -Name $Name -ErrorAction Stop
        $item.$Name
    }
    catch {
        $null
    }
}

function Get-PendingRebootFlags {
    $flags = @()

    $wuRebootRequired = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired'
    if (Test-RegistryKeyExists -Path $wuRebootRequired) {
        $flags += New-FlagRecord -Name 'WindowsUpdateRebootRequired' -Type 'RegistryKey' -Path $wuRebootRequired -ValueName '' -Details 'Windows Update indicates a reboot is required.'
    }

    $wuPostRebootReporting = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\PostRebootReporting'
    if (Test-RegistryKeyExists -Path $wuPostRebootReporting) {
        $flags += New-FlagRecord -Name 'WindowsUpdatePostRebootReporting' -Type 'RegistryKey' -Path $wuPostRebootReporting -ValueName '' -Details 'Windows Update post-reboot reporting flag is present.'
    }

    $cbsRebootPending = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending'
    if (Test-RegistryKeyExists -Path $cbsRebootPending) {
        $flags += New-FlagRecord -Name 'CBSRebootPending' -Type 'RegistryKey' -Path $cbsRebootPending -ValueName '' -Details 'Component Based Servicing reports RebootPending.'
    }

    $cbsRebootInProgress = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootInProgress'
    if (Test-RegistryKeyExists -Path $cbsRebootInProgress) {
        $flags += New-FlagRecord -Name 'CBSRebootInProgress' -Type 'RegistryKey' -Path $cbsRebootInProgress -ValueName '' -Details 'Component Based Servicing reports RebootInProgress.'
    }

    $cbsPackagesPending = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\PackagesPending'
    if (Test-RegistryKeyExists -Path $cbsPackagesPending) {
        $flags += New-FlagRecord -Name 'CBSPackagesPending' -Type 'RegistryKey' -Path $cbsPackagesPending -ValueName '' -Details 'Component Based Servicing reports PackagesPending.'
    }

    $sessionManagerPath = 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager'

    $pendingFileRenameOperations = Get-RegistryValueSafe -Path $sessionManagerPath -Name 'PendingFileRenameOperations'
    if ($null -ne $pendingFileRenameOperations) {
        $details = if ($pendingFileRenameOperations -is [System.Array]) {
            ($pendingFileRenameOperations | ForEach-Object { [string]$_ }) -join ' | '
        }
        else {
            [string]$pendingFileRenameOperations
        }

        $flags += New-FlagRecord -Name 'PendingFileRenameOperations' -Type 'RegistryValue' -Path $sessionManagerPath -ValueName 'PendingFileRenameOperations' -Details $details
    }

    $pendingFileRenameOperations2 = Get-RegistryValueSafe -Path $sessionManagerPath -Name 'PendingFileRenameOperations2'
    if ($null -ne $pendingFileRenameOperations2) {
        $details = if ($pendingFileRenameOperations2 -is [System.Array]) {
            ($pendingFileRenameOperations2 | ForEach-Object { [string]$_ }) -join ' | '
        }
        else {
            [string]$pendingFileRenameOperations2
        }

        $flags += New-FlagRecord -Name 'PendingFileRenameOperations2' -Type 'RegistryValue' -Path $sessionManagerPath -ValueName 'PendingFileRenameOperations2' -Details $details
    }

    $updatesPath = 'HKLM:\SOFTWARE\Microsoft\Updates'
    $updateExeVolatile = Get-RegistryValueSafe -Path $updatesPath -Name 'UpdateExeVolatile'
    if ($null -ne $updateExeVolatile) {
        try {
            if ([int]$updateExeVolatile -ne 0) {
                $flags += New-FlagRecord -Name 'UpdateExeVolatile' -Type 'RegistryValue' -Path $updatesPath -ValueName 'UpdateExeVolatile' -Details "Value is $updateExeVolatile"
            }
        }
        catch {
            $flags += New-FlagRecord -Name 'UpdateExeVolatile' -Type 'RegistryValue' -Path $updatesPath -ValueName 'UpdateExeVolatile' -Details "Non-integer value detected: $updateExeVolatile"
        }
    }

    $script:CurrentFlags = @($flags)
    return @($flags)
}

function New-DefaultState {
    [PSCustomObject]@{
        Stage     = 0
        FirstSeen = $null
        LastRun   = $null
        LastFlags = @()
    }
}

function ConvertTo-StateObject {
    param(
        [Parameter(Mandatory)]
        $InputObject
    )

    $stage = 0
    try {
        $stage = [int]($InputObject.Stage)
    }
    catch {
        $stage = 0
    }

    if ($stage -lt 0 -or $stage -gt 2) {
        $stage = 0
    }

    [PSCustomObject]@{
        Stage     = $stage
        FirstSeen = $InputObject.FirstSeen
        LastRun   = $InputObject.LastRun
        LastFlags = @($InputObject.LastFlags)
    }
}

function Backup-BadStateFile {
    param(
        [string]$Reason
    )

    try {
        if (Test-Path -LiteralPath $script:StateFilePath) {
            $backupPath = "$($script:StateFilePath).bad-$((Get-Date).ToString('yyyyMMdd-HHmmss'))"
            Copy-Item -LiteralPath $script:StateFilePath -Destination $backupPath -Force -ErrorAction Stop
            Write-Status "Backed up unreadable reboot state file to: $backupPath" 'WARN'
        }
    }
    catch {
        Write-Status "Could not back up unreadable reboot state file. $($_.Exception.Message)" 'WARN'
    }

    Write-Status "State file reset reason: $Reason" 'WARN'
}

function Get-State {
    if (-not (Test-Path -LiteralPath $script:StateFilePath)) {
        return New-DefaultState
    }

    try {
        $raw = Get-Content -LiteralPath $script:StateFilePath -Raw -Encoding UTF8 -ErrorAction Stop

        if ([string]::IsNullOrWhiteSpace($raw)) {
            Backup-BadStateFile -Reason 'State file was empty.'
            return New-DefaultState
        }

        $trimmed = $raw.Trim()

        try {
            $obj = $trimmed | ConvertFrom-Json -ErrorAction Stop
            return (ConvertTo-StateObject -InputObject $obj)
        }
        catch {
            # Recovery path: some systems have produced a state file missing the opening brace,
            # for example: \"Stage\": 1, ... }. Wrap it and retry instead of forcing a false first pass.
            if (($trimmed -notmatch '^\s*\{') -and ($trimmed -match '\"Stage\"\s*:') -and ($trimmed -match '\}\s*$')) {
                try {
                    $repairedJson = '{' + $trimmed
                    $obj = $repairedJson | ConvertFrom-Json -ErrorAction Stop
                    Write-Status 'State file was missing the opening JSON brace. Repaired in memory and continuing staged reboot logic.' 'WARN'
                    Save-State -Stage ([int]$obj.Stage) -FirstSeen $obj.FirstSeen -LastRun $obj.LastRun -LastFlags @($obj.LastFlags)
                    return (ConvertTo-StateObject -InputObject $obj)
                }
                catch {
                    Backup-BadStateFile -Reason "Automatic JSON brace repair failed. $($_.Exception.Message)"
                    return New-DefaultState
                }
            }

            Backup-BadStateFile -Reason "State file unreadable. $($_.Exception.Message)"
            return New-DefaultState
        }
    }
    catch {
        Backup-BadStateFile -Reason "State file could not be read. $($_.Exception.Message)"
        return New-DefaultState
    }
}

function Save-State {
    param(
        [int]$Stage,
        $FirstSeen,
        $LastRun,
        $LastFlags
    )

    Ensure-Folder -Path $StateDirectory

    $state = [PSCustomObject]@{
        Stage     = $Stage
        FirstSeen = if ($null -ne $FirstSeen) { [string]$FirstSeen } else { $null }
        LastRun   = if ($null -ne $LastRun) { [string]$LastRun } else { $null }
        LastFlags = @($LastFlags)
    }

    $json = $state | ConvertTo-Json -Depth 12
    if ([string]::IsNullOrWhiteSpace($json) -or $json.TrimStart()[0] -ne '{') {
        throw 'Generated state JSON did not begin with an opening brace. Refusing to write invalid state file.'
    }

    $unique = [guid]::NewGuid().ToString('N')
    $tempPath = Join-Path $StateDirectory ("$StateFileName.$unique.tmp")

    try {
        $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
        $fileStream = [System.IO.File]::Open($tempPath, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
        try {
            $writer = New-Object System.IO.StreamWriter($fileStream, $utf8NoBom)
            try {
                $writer.Write($json)
                $writer.Flush()
                $fileStream.Flush($true)
            }
            finally {
                $writer.Dispose()
            }
        }
        finally {
            $fileStream.Dispose()
        }

        # Validate the temp file before replacing the active state file.
        $validated = (Get-Content -LiteralPath $tempPath -Raw -Encoding UTF8 -ErrorAction Stop) | ConvertFrom-Json -ErrorAction Stop
        if ($null -eq $validated -or $null -eq $validated.Stage) {
            throw 'Temporary state file validation failed because required Stage property was missing.'
        }

        if (Test-Path -LiteralPath $script:StateFilePath) {
            Copy-Item -LiteralPath $script:StateFilePath -Destination "$($script:StateFilePath).prev" -Force -ErrorAction SilentlyContinue
        }

        Move-Item -LiteralPath $tempPath -Destination $script:StateFilePath -Force -ErrorAction Stop
        Confirm-FileWriteToDisk -Path $script:StateFilePath
        Write-Status "Saved reboot state tracking at stage $Stage." 'INFO'
    }
    catch {
        Write-Status "Failed to save reboot state file safely. Error: $($_.Exception.Message)" 'ERROR'
        try {
            if (Test-Path -LiteralPath $tempPath) {
                Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue
            }
        }
        catch { }
        throw
    }
}

function Reset-State {
    try {
        if (Test-Path -LiteralPath $script:StateFilePath) {
            Remove-Item -LiteralPath $script:StateFilePath -Force -ErrorAction Stop
            Write-Status "Reset reboot state tracking." 'OK'
        }
    }
    catch {
        Write-Status "Failed to remove state file: $($_.Exception.Message)" 'WARN'
    }
}

function Remove-RegistryKeySafe {
    param(
        [string]$Name,
        [string]$Path
    )

    try {
        if (Test-Path -LiteralPath $Path) {
            Remove-Item -LiteralPath $Path -Recurse -Force -ErrorAction Stop
            $script:ClearResults += New-ClearResult -Name $Name -Path $Path -Action 'RemoveKey' -Status 'Succeeded' -Message 'Registry key removed.'
            Write-Status "Removed key for [$Name]: $Path" 'OK'
        }
        else {
            $script:ClearResults += New-ClearResult -Name $Name -Path $Path -Action 'RemoveKey' -Status 'Skipped' -Message 'Registry key not present.'
            Write-Status "Key already absent for [$Name]: $Path" 'INFO'
        }
    }
    catch {
        $script:ClearResults += New-ClearResult -Name $Name -Path $Path -Action 'RemoveKey' -Status 'Failed' -Message $_.Exception.Message
        Write-Status "Failed removing key for [$Name]: $($_.Exception.Message)" 'ERROR'
    }
}

function Remove-RegistryValueSafe {
    param(
        [string]$Name,
        [string]$Path,
        [string]$ValueName
    )

    try {
        if (Test-Path -LiteralPath $Path) {
            $currentValue = Get-RegistryValueSafe -Path $Path -Name $ValueName
            if ($null -ne $currentValue) {
                Remove-ItemProperty -LiteralPath $Path -Name $ValueName -ErrorAction Stop
                $script:ClearResults += New-ClearResult -Name $Name -Path "$Path\$ValueName" -Action 'RemoveValue' -Status 'Succeeded' -Message 'Registry value removed.'
                Write-Status "Removed value for [$Name]: $Path\$ValueName" 'OK'
            }
            else {
                $script:ClearResults += New-ClearResult -Name $Name -Path "$Path\$ValueName" -Action 'RemoveValue' -Status 'Skipped' -Message 'Registry value not present.'
                Write-Status "Value already absent for [$Name]: $Path\$ValueName" 'INFO'
            }
        }
        else {
            $script:ClearResults += New-ClearResult -Name $Name -Path "$Path\$ValueName" -Action 'RemoveValue' -Status 'Skipped' -Message 'Registry path not present.'
            Write-Status "Path absent for [$Name]: $Path" 'INFO'
        }
    }
    catch {
        $script:ClearResults += New-ClearResult -Name $Name -Path "$Path\$ValueName" -Action 'RemoveValue' -Status 'Failed' -Message $_.Exception.Message
        Write-Status "Failed removing value for [$Name]: $($_.Exception.Message)" 'ERROR'
    }
}

function Set-RegistryValueSafe {
    param(
        [string]$Name,
        [string]$Path,
        [string]$ValueName,
        $Value,
        [string]$Type
    )

    try {
        if (-not (Test-Path -LiteralPath $Path)) {
            New-Item -Path $Path -Force | Out-Null
        }

        New-ItemProperty -LiteralPath $Path -Name $ValueName -PropertyType $Type -Value $Value -Force | Out-Null
        $script:ClearResults += New-ClearResult -Name $Name -Path "$Path\$ValueName" -Action 'SetValue' -Status 'Succeeded' -Message "Set to $Value."
        Write-Status "Set [$Name] value: $Path\$ValueName = $Value" 'OK'
    }
    catch {
        $script:ClearResults += New-ClearResult -Name $Name -Path "$Path\$ValueName" -Action 'SetValue' -Status 'Failed' -Message $_.Exception.Message
        Write-Status "Failed setting value for [$Name]: $($_.Exception.Message)" 'ERROR'
    }
}

function Clear-PendingRebootFlags {
    $script:ClearResults = @()
    Write-Status "Attempting to clear persistent reboot flags..." 'WARN'

    Remove-RegistryKeySafe   -Name 'WindowsUpdateRebootRequired'      -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired'
    Remove-RegistryKeySafe   -Name 'WindowsUpdatePostRebootReporting' -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\PostRebootReporting'
    Remove-RegistryKeySafe   -Name 'CBSRebootPending'                 -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending'
    Remove-RegistryKeySafe   -Name 'CBSRebootInProgress'              -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootInProgress'
    Remove-RegistryKeySafe   -Name 'CBSPackagesPending'               -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\PackagesPending'
    Remove-RegistryValueSafe -Name 'PendingFileRenameOperations'      -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -ValueName 'PendingFileRenameOperations'
    Remove-RegistryValueSafe -Name 'PendingFileRenameOperations2'     -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -ValueName 'PendingFileRenameOperations2'
    Set-RegistryValueSafe    -Name 'UpdateExeVolatile'                -Path 'HKLM:\SOFTWARE\Microsoft\Updates' -ValueName 'UpdateExeVolatile' -Value 0 -Type 'DWord'

    return @($script:ClearResults)
}

function Test-AppPatchPendingRenameFlag {
    param(
        [array]$Flags
    )

    foreach ($flag in @($Flags)) {
        if ($flag.Name -in @('PendingFileRenameOperations','PendingFileRenameOperations2') -and $flag.Details -match '(?i)apppatch\\AcPluginDlls') {
            return $true
        }
    }

    return $false
}

function Stop-UpdateLockingProcessesSafe {
    $lockingProcesses = @(
        'TiWorker',
        'TrustedInstaller',
        'wuauclt',
        'UsoClient'
    )

    foreach ($processName in $lockingProcesses) {
        try {
            $processes = @(Get-Process -Name $processName -ErrorAction SilentlyContinue)
            foreach ($process in $processes) {
                try {
                    Write-Status "Stopping possible update-locking process: $($process.Name) PID $($process.Id)" 'WARN'
                    Stop-Process -Id $process.Id -Force -ErrorAction Stop
                }
                catch {
                    Write-Status "Could not stop process $($process.Name) PID $($process.Id): $($_.Exception.Message)" 'WARN'
                }
            }
        }
        catch {
            Write-Status "Process check failed for [$processName]: $($_.Exception.Message)" 'WARN'
        }
    }
}

function Resolve-PersistentRebootFlags {
    param(
        [array]$InitialFlags
    )

    if (Test-AppPatchPendingRenameFlag -Flags $InitialFlags) {
        Write-Status 'Detected AppPatch/AcPluginDlls PendingFileRenameOperations. These can persist after updates and are treated as non-blocking after cleanup.' 'WARN'
    }

    Stop-UpdateLockingProcessesSafe
    [void](Clear-PendingRebootFlags)
    Start-Sleep -Seconds 2

    $remainingFlags = @(Get-PendingRebootFlags)
    if ($remainingFlags.Count -gt 0) {
        Write-Status 'Reboot flags are still present after cleanup. Logging as warning and continuing instead of failing the maintenance run.' 'WARN'
        foreach ($flag in $remainingFlags) {
            Write-Status "Remaining non-blocking flag: $($flag.Name) | $($flag.Path) | $($flag.Details)" 'WARN'
        }
        $script:OverallResult = 'PersistentFlagsClearedOrIgnored'
    }
    else {
        Write-Status 'Persistent reboot flags cleared successfully.' 'OK'
        $script:OverallResult = 'PersistentFlagsCleared'
    }

    Reset-State
    Write-YamlLog
    exit 0
}


function Write-YamlLog {
    try {
        $runEnd = Get-Date
        $duration = [math]::Round(($runEnd - $script:RunStart).TotalSeconds, 0)
        $flags = @($script:CurrentFlags)
        $clear = @($script:ClearResults)
        $actions = @($script:ActionHistory)

        $lines = New-Object System.Collections.Generic.List[string]

        $lines.Add("computer_name: $(ConvertTo-YamlScalar $script:ComputerName)") | Out-Null
        $lines.Add("script_name: '07_Force_Reboot_Install_Updates.ps1'") | Out-Null
        $lines.Add("script_version: '1.9'") | Out-Null
        $lines.Add("run_started: $(ConvertTo-YamlScalar $script:RunStart)") | Out-Null
        $lines.Add("run_finished: $(ConvertTo-YamlScalar $runEnd)") | Out-Null
        $lines.Add("duration_seconds: $duration") | Out-Null
        $lines.Add("stage: $($script:CurrentStage)") | Out-Null
        $lines.Add("reboot_issued: $(ConvertTo-YamlScalar $script:RebootIssued)") | Out-Null
        $lines.Add("reboot_reason: $(ConvertTo-YamlScalar $script:RebootReason)") | Out-Null
        $lines.Add("overall_result: $(ConvertTo-YamlScalar $script:OverallResult)") | Out-Null
        $lines.Add("failure_message: $(ConvertTo-YamlScalar $script:FailureMessage)") | Out-Null
        $lines.Add('') | Out-Null

        $lines.Add('flags_detected:') | Out-Null
        if ($flags.Count -gt 0) {
            foreach ($flag in $flags) {
                $lines.Add('  -') | Out-Null
                $lines.Add("    name: $(ConvertTo-YamlScalar $flag.Name)") | Out-Null
                $lines.Add("    type: $(ConvertTo-YamlScalar $flag.Type)") | Out-Null
                $lines.Add("    path: $(ConvertTo-YamlScalar $flag.Path)") | Out-Null
                $lines.Add("    value_name: $(ConvertTo-YamlScalar $flag.ValueName)") | Out-Null
                $lines.Add("    details: $(ConvertTo-YamlScalar $flag.Details)") | Out-Null
            }
        }
        else {
            $lines.Add('  []') | Out-Null
        }

        $lines.Add('') | Out-Null
        $lines.Add('clear_actions:') | Out-Null
        if ($clear.Count -gt 0) {
            foreach ($item in $clear) {
                $lines.Add('  -') | Out-Null
                $lines.Add("    name: $(ConvertTo-YamlScalar $item.Name)") | Out-Null
                $lines.Add("    path: $(ConvertTo-YamlScalar $item.Path)") | Out-Null
                $lines.Add("    action: $(ConvertTo-YamlScalar $item.Action)") | Out-Null
                $lines.Add("    status: $(ConvertTo-YamlScalar $item.Status)") | Out-Null
                $lines.Add("    message: $(ConvertTo-YamlScalar $item.Message)") | Out-Null
            }
        }
        else {
            $lines.Add('  []') | Out-Null
        }

        $lines.Add('') | Out-Null
        $lines.Add('actions:') | Out-Null
        if ($actions.Count -gt 0) {
            foreach ($action in $actions) {
                $lines.Add('  -') | Out-Null
                $lines.Add("    time: $(ConvertTo-YamlScalar $action.Time)") | Out-Null
                $lines.Add("    level: $(ConvertTo-YamlScalar $action.Level)") | Out-Null
                $lines.Add("    message: $(ConvertTo-YamlScalar $action.Message)") | Out-Null
            }
        }
        else {
            $lines.Add('  []') | Out-Null
        }

        $yamlTempPath = "$($script:YamlLogPath).tmp"
        Set-Content -LiteralPath $yamlTempPath -Value $lines -Encoding UTF8 -Force
        Move-Item -LiteralPath $yamlTempPath -Destination $script:YamlLogPath -Force
        Confirm-FileWriteToDisk -Path $script:YamlLogPath
    }
    catch {
        Write-Warning "Failed to write YAML log: $($_.Exception.Message)"
    }
}

function Invoke-ForcedReboot {
    param(
        [Parameter(Mandatory)]
        [string]$Reason
    )

    $script:RebootIssued = $true
    $script:RebootReason = $Reason

    Write-Status "Issuing forced reboot in $RebootDelaySeconds seconds. Reason: $Reason" 'WARN'
    Write-YamlLog
    Confirm-StateWriteBeforeReboot

    $arguments = @(
        '/r'
        '/f'
        '/t'
        [string]$RebootDelaySeconds
        '/c'
        $Reason
    )

    & shutdown.exe @arguments | Out-Null

    if ($LASTEXITCODE -ne 0) {
        throw "shutdown.exe returned exit code $LASTEXITCODE"
    }

    exit 3010
}

# Main
Initialize-Paths
Initialize-SingleInstanceLock

try {
    if (-not (Test-IsAdministrator)) {
        Write-Error "Please run this script as Administrator."
        exit 1
    }

    Write-Status "Starting reboot flag evaluation..." 'INFO'

    try {
        $state = Get-State
        $flags = @(Get-PendingRebootFlags)

        $stage = 0
        try {
            $stage = [int]$state.Stage
        }
        catch {
            $stage = 0
        }

        $script:CurrentStage = $stage

        switch ($stage) {
            0 {
                Write-Status "First pass detected. Forcing reboot regardless of flag state." 'WARN'
                Save-State -Stage 1 -FirstSeen ((Get-Date).ToString('o')) -LastRun ((Get-Date).ToString('o')) -LastFlags $flags
                $script:OverallResult = 'ForcedInitialReboot'
                Invoke-ForcedReboot -Reason 'Initial forced reboot for update cycle.'
            }

            1 {
                if ($flags.Count -gt 0) {
                    Write-Status "Second pass: reboot flags still detected. Issuing second reboot." 'WARN'
                    foreach ($flag in $flags) {
                        Write-Status "Flag detected: $($flag.Name) | $($flag.Path) | $($flag.Details)" 'WARN'
                    }

                    Save-State -Stage 2 -FirstSeen $state.FirstSeen -LastRun ((Get-Date).ToString('o')) -LastFlags $flags
                    $script:OverallResult = 'SecondPassRebootIssued'
                    Invoke-ForcedReboot -Reason 'Reboot flags remain after initial reboot. Issuing second reboot.'
                }
                else {
                    Write-Status "Second pass: no reboot flags detected. Resetting state and exiting normally." 'OK'
                    Reset-State
                    $script:OverallResult = 'SecondPassNoFlags'
                    Write-YamlLog
                    exit 0
                }
            }

            2 {
                if ($flags.Count -gt 0) {
                    Write-Status "Third pass: reboot flags still present after two reboots. Logging and clearing as non-blocking warning." 'WARN'
                    foreach ($flag in $flags) {
                        Write-Status "Persistent flag: $($flag.Name) | $($flag.Path) | $($flag.Details)" 'WARN'
                    }

                    Resolve-PersistentRebootFlags -InitialFlags $flags
                }
                else {
                    Write-Status "Third pass: no reboot flags detected. Resetting state and exiting normally." 'OK'
                    Reset-State
                    $script:OverallResult = 'ThirdPassNoFlags'
                    Write-YamlLog
                    exit 0
                }
            }

            default {
                Write-Status "Unexpected state value [$stage]. Resetting state and starting over with forced initial reboot." 'WARN'
                Reset-State
                Save-State -Stage 1 -FirstSeen ((Get-Date).ToString('o')) -LastRun ((Get-Date).ToString('o')) -LastFlags $flags
                $script:CurrentStage = 0
                $script:OverallResult = 'ForcedInitialRebootAfterStateReset'
                Invoke-ForcedReboot -Reason 'State reset occurred. Performing initial forced reboot for update cycle.'
            }
        }
    }
    catch {
        $script:FailureMessage = $_.Exception.Message
        $script:OverallResult = 'Failed'
        Write-Status "Script failed: $($_.Exception.Message)" 'ERROR'
        Write-YamlLog
        exit 3
    }
}
finally {
    Release-SingleInstanceLock
}
