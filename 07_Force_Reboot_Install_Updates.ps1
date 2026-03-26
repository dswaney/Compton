# =====================================================================
# ScriptName: 07_Force_Reboot_Install_Updates.ps1
# ScriptVersion: 1.6
# LastUpdated: 2026-03-24
# =====================================================================

[CmdletBinding()]
param(
    [int]$RebootDelaySeconds = 30,
    [string]$LogDirectory = 'C:\Logs',
    [string]$StateDirectory = 'C:\ProgramData\MISMaintenance',
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

function Get-State {
    if (-not (Test-Path -LiteralPath $script:StateFilePath)) {
        return [PSCustomObject]@{
            Stage     = 0
            FirstSeen = $null
            LastRun   = $null
            LastFlags = @()
        }
    }

    try {
        $raw = Get-Content -LiteralPath $script:StateFilePath -Raw -Encoding UTF8
        $obj = $raw | ConvertFrom-Json

        [PSCustomObject]@{
            Stage     = [int]($obj.Stage)
            FirstSeen = $obj.FirstSeen
            LastRun   = $obj.LastRun
            LastFlags = @($obj.LastFlags)
        }
    }
    catch {
        Write-Status "State file unreadable. Resetting state. Error: $($_.Exception.Message)" 'WARN'
        [PSCustomObject]@{
            Stage     = 0
            FirstSeen = $null
            LastRun   = $null
            LastFlags = @()
        }
    }
}

function Save-State {
    param(
        [int]$Stage,
        $FirstSeen,
        $LastRun,
        $LastFlags
    )

    $state = [PSCustomObject]@{
        Stage     = $Stage
        FirstSeen = if ($null -ne $FirstSeen) { [string]$FirstSeen } else { $null }
        LastRun   = if ($null -ne $LastRun) { [string]$LastRun } else { $null }
        LastFlags = @($LastFlags)
    }

    $json = $state | ConvertTo-Json -Depth 8
    Set-Content -LiteralPath $script:StateFilePath -Value $json -Encoding UTF8
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
        $lines.Add("script_version: '1.6'") | Out-Null
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

        Set-Content -Path $script:YamlLogPath -Value $lines -Encoding UTF8
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
                Write-Status "Third pass: reboot flags still present after two reboots. Logging and clearing." 'ERROR'
                foreach ($flag in $flags) {
                    Write-Status "Persistent flag: $($flag.Name) | $($flag.Path) | $($flag.Details)" 'ERROR'
                }

                $script:OverallResult = 'FlagsPersistedAfterTwoReboots'
                [void](Clear-PendingRebootFlags)
                Write-YamlLog
                Reset-State
                exit 2
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
