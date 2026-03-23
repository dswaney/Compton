# ScriptVersion: 1.0
# LastUpdated: 2026-03-23

<#
.SYNOPSIS
    Reboots Windows and verifies after startup whether Windows Update / CBS
    pending reboot flags were cleared.

.DESCRIPTION
    Designed for Deep Freeze or similar environments where updates are staged
    during a thawed window and must complete during reboot/startup.

    This script:
      - DOES NOT clear reboot flags
      - Detects pre-reboot pending state
      - Creates a one-time startup scheduled task
      - Writes a post-boot verification script
      - Reboots the machine
      - After boot, verifies whether Windows Update / CBS reboot indicators cleared
      - Logs all results to disk

.NOTES
    Run while the machine is still THAWED.
    Keep the machine thawed through reboot and until Windows finishes processing.
#>

[CmdletBinding()]
param(
    [int]$DelaySeconds = 20,
    [switch]$ForceReboot = $true,
    [switch]$WaitForStaging = $true,
    [int]$StagingWaitSeconds = 60,
    [int]$PostBootInitialWaitSeconds = 180,
    [string]$BaseFolder = "$env:ProgramData\UpdateRebootVerifier"
)

$ErrorActionPreference = 'Stop'

$LogPath               = Join-Path $BaseFolder 'UpdateRebootVerifier.log'
$PostBootScriptPath    = Join-Path $BaseFolder 'PostBoot-VerifyUpdateFlags.ps1'
$PreBootStatePath      = Join-Path $BaseFolder 'PreBoot-State.json'
$PostBootStatePath     = Join-Path $BaseFolder 'PostBoot-State.json'
$ResultPath            = Join-Path $BaseFolder 'Verification-Result.txt'
$TaskName              = 'Verify-WindowsUpdate-Reboot-Completion'

function Write-Log {
    param(
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

    try {
        if (-not (Test-Path -Path $BaseFolder)) {
            New-Item -Path $BaseFolder -ItemType Directory -Force | Out-Null
        }
        Add-Content -Path $LogPath -Value $line
    }
    catch {
        # Do not break execution if logging fails
    }
}

function Test-IsAdministrator {
    try {
        $identity  = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object Security.Principal.WindowsPrincipal($identity)
        return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    }
    catch {
        return $false
    }
}

function Test-RegKeyExists {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    try {
        return (Test-Path -Path $Path)
    }
    catch {
        return $false
    }
}

function Get-RegValue {
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter(Mandatory)]
        [string]$Name
    )

    try {
        return (Get-ItemProperty -Path $Path -Name $Name -ErrorAction Stop).$Name
    }
    catch {
        return $null
    }
}

function Get-PendingRebootState {
    $result = [ordered]@{
        Timestamp                         = (Get-Date).ToString('o')
        CBServicing_RebootPending         = $false
        CBServicing_RebootInProgress      = $false
        WindowsUpdate_RebootRequired      = $false
        SessionManager_PendingFileRename  = $false
        SessionManager_PendingFileRename2 = $false
        UpdateExeVolatile                 = $false
        ComputerNameChangePending         = $false
        PackagesPending                   = $false
        WUAU_RebootRequired_COM           = $false
        AnyPendingReboot                  = $false
        WindowsUpdatePending              = $false
        GenericPendingOnly                = $false
    }

    $cbsRebootPending     = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending'
    $cbsRebootInProgress  = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootInProgress'
    $wuRebootRequired     = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired'
    $sessionMgr           = 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager'
    $updateExeVolatile    = 'HKLM:\SOFTWARE\Microsoft\Updates'
    $packagesPending      = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\PackagesPending'
    $activeComputerName   = 'HKLM:\SYSTEM\CurrentControlSet\Control\ComputerName\ActiveComputerName'
    $pendingComputerName  = 'HKLM:\SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName'

    $result.CBServicing_RebootPending    = Test-RegKeyExists -Path $cbsRebootPending
    $result.CBServicing_RebootInProgress = Test-RegKeyExists -Path $cbsRebootInProgress
    $result.WindowsUpdate_RebootRequired = Test-RegKeyExists -Path $wuRebootRequired
    $result.PackagesPending              = Test-RegKeyExists -Path $packagesPending

    $pendingRename = Get-RegValue -Path $sessionMgr -Name 'PendingFileRenameOperations'
    if ($null -ne $pendingRename -and $pendingRename.Count -gt 0) {
        $result.SessionManager_PendingFileRename = $true
    }

    $pendingRename2 = Get-RegValue -Path $sessionMgr -Name 'PendingFileRenameOperations2'
    if ($null -ne $pendingRename2 -and $pendingRename2.Count -gt 0) {
        $result.SessionManager_PendingFileRename2 = $true
    }

    $uev = Get-RegValue -Path $updateExeVolatile -Name 'UpdateExeVolatile'
    if ($null -ne $uev) {
        try {
            if ([int]$uev -ne 0) {
                $result.UpdateExeVolatile = $true
            }
        }
        catch {
            if (-not [string]::IsNullOrWhiteSpace([string]$uev)) {
                $result.UpdateExeVolatile = $true
            }
        }
    }

    $activeName  = Get-RegValue -Path $activeComputerName -Name 'ComputerName'
    $pendingName = Get-RegValue -Path $pendingComputerName -Name 'ComputerName'
    if ($activeName -and $pendingName -and $activeName -ne $pendingName) {
        $result.ComputerNameChangePending = $true
    }

    try {
        $sysInfo = New-Object -ComObject Microsoft.Update.SystemInfo
        if ($sysInfo.RebootRequired) {
            $result.WUAU_RebootRequired_COM = $true
        }
    }
    catch {
        # Keep silent in this function
    }

    $result.WindowsUpdatePending =
        $result.WindowsUpdate_RebootRequired -or
        $result.WUAU_RebootRequired_COM -or
        $result.CBServicing_RebootPending -or
        $result.PackagesPending

    $result.AnyPendingReboot =
        $result.WindowsUpdatePending -or
        $result.CBServicing_RebootInProgress -or
        $result.SessionManager_PendingFileRename -or
        $result.SessionManager_PendingFileRename2 -or
        $result.UpdateExeVolatile -or
        $result.ComputerNameChangePending

    $result.GenericPendingOnly =
        $result.AnyPendingReboot -and -not $result.WindowsUpdatePending

    return [PSCustomObject]$result
}

function Wait-ForUpdateStaging {
    param(
        [int]$Seconds = 60
    )

    Write-Log "Waiting up to $Seconds seconds for update staging activity to settle..." 'INFO'

    $serviceNames = @(
        'wuauserv',
        'UsoSvc',
        'BITS',
        'TrustedInstaller'
    )

    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    do {
        $busyServices = @()

        foreach ($svcName in $serviceNames) {
            try {
                $svc = Get-Service -Name $svcName -ErrorAction Stop
                if ($svc.Status -eq 'Running') {
                    $busyServices += $svcName
                }
            }
            catch {
                # Ignore lookup failures
            }
        }

        if ($busyServices.Count -eq 0) {
            Write-Log "Update-related services are not actively running." 'OK'
            return
        }

        Write-Log "Still seeing update-related service activity: $($busyServices -join ', ')" 'INFO'
        Start-Sleep -Seconds 5
    }
    while ($stopwatch.Elapsed.TotalSeconds -lt $Seconds)

    Write-Log "Reached staging wait timeout. Proceeding with reboot." 'WARN'
}

function Write-PostBootVerifierScript {
    $scriptContent = @"
`$ErrorActionPreference = 'Stop'

`$BaseFolder            = '$BaseFolder'
`$LogPath               = '$LogPath'
`$PreBootStatePath      = '$PreBootStatePath'
`$PostBootStatePath     = '$PostBootStatePath'
`$ResultPath            = '$ResultPath'
`$TaskName              = '$TaskName'
`$PostBootInitialWait   = $PostBootInitialWaitSeconds

function Write-Log {
    param(
        [string]`$Message,
        [ValidateSet('INFO','OK','WARN','ERROR')]
        [string]`$Level = 'INFO'
    )

    try {
        if (-not (Test-Path -Path `$BaseFolder)) {
            New-Item -Path `$BaseFolder -ItemType Directory -Force | Out-Null
        }
    }
    catch {}

    `$timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    `$line = "[`$timestamp] [`$(''{0,-5}'' -f `$Level)] `$Message"

    try { Add-Content -Path `$LogPath -Value `$line } catch {}
}

function Test-RegKeyExists {
    param([string]`$Path)
    try { return (Test-Path -Path `$Path) } catch { return `$false }
}

function Get-RegValue {
    param([string]`$Path,[string]`$Name)
    try { return (Get-ItemProperty -Path `$Path -Name `$Name -ErrorAction Stop).`$Name } catch { return `$null }
}

function Get-PendingRebootState {
    `$result = [ordered]@{
        Timestamp                         = (Get-Date).ToString('o')
        CBServicing_RebootPending         = `$false
        CBServicing_RebootInProgress      = `$false
        WindowsUpdate_RebootRequired      = `$false
        SessionManager_PendingFileRename  = `$false
        SessionManager_PendingFileRename2 = `$false
        UpdateExeVolatile                 = `$false
        ComputerNameChangePending         = `$false
        PackagesPending                   = `$false
        WUAU_RebootRequired_COM           = `$false
        AnyPendingReboot                  = `$false
        WindowsUpdatePending              = `$false
        GenericPendingOnly                = `$false
    }

    `$cbsRebootPending     = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending'
    `$cbsRebootInProgress  = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootInProgress'
    `$wuRebootRequired     = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired'
    `$sessionMgr           = 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager'
    `$updateExeVolatile    = 'HKLM:\SOFTWARE\Microsoft\Updates'
    `$packagesPending      = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\PackagesPending'
    `$activeComputerName   = 'HKLM:\SYSTEM\CurrentControlSet\Control\ComputerName\ActiveComputerName'
    `$pendingComputerName  = 'HKLM:\SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName'

    `$result.CBServicing_RebootPending    = Test-RegKeyExists -Path `$cbsRebootPending
    `$result.CBServicing_RebootInProgress = Test-RegKeyExists -Path `$cbsRebootInProgress
    `$result.WindowsUpdate_RebootRequired = Test-RegKeyExists -Path `$wuRebootRequired
    `$result.PackagesPending              = Test-RegKeyExists -Path `$packagesPending

    `$pendingRename = Get-RegValue -Path `$sessionMgr -Name 'PendingFileRenameOperations'
    if (`$null -ne `$pendingRename -and `$pendingRename.Count -gt 0) {
        `$result.SessionManager_PendingFileRename = `$true
    }

    `$pendingRename2 = Get-RegValue -Path `$sessionMgr -Name 'PendingFileRenameOperations2'
    if (`$null -ne `$pendingRename2 -and `$pendingRename2.Count -gt 0) {
        `$result.SessionManager_PendingFileRename2 = `$true
    }

    `$uev = Get-RegValue -Path `$updateExeVolatile -Name 'UpdateExeVolatile'
    if (`$null -ne `$uev) {
        try {
            if ([int]`$uev -ne 0) { `$result.UpdateExeVolatile = `$true }
        }
        catch {
            if (-not [string]::IsNullOrWhiteSpace([string]`$uev)) {
                `$result.UpdateExeVolatile = `$true
            }
        }
    }

    `$activeName  = Get-RegValue -Path `$activeComputerName -Name 'ComputerName'
    `$pendingName = Get-RegValue -Path `$pendingComputerName -Name 'ComputerName'
    if (`$activeName -and `$pendingName -and `$activeName -ne `$pendingName) {
        `$result.ComputerNameChangePending = `$true
    }

    try {
        `$sysInfo = New-Object -ComObject Microsoft.Update.SystemInfo
        if (`$sysInfo.RebootRequired) {
            `$result.WUAU_RebootRequired_COM = `$true
        }
    }
    catch {}

    `$result.WindowsUpdatePending =
        `$result.WindowsUpdate_RebootRequired -or
        `$result.WUAU_RebootRequired_COM -or
        `$result.CBServicing_RebootPending -or
        `$result.PackagesPending

    `$result.AnyPendingReboot =
        `$result.WindowsUpdatePending -or
        `$result.CBServicing_RebootInProgress -or
        `$result.SessionManager_PendingFileRename -or
        `$result.SessionManager_PendingFileRename2 -or
        `$result.UpdateExeVolatile -or
        `$result.ComputerNameChangePending

    `$result.GenericPendingOnly =
        `$result.AnyPendingReboot -and -not `$result.WindowsUpdatePending

    return [PSCustomObject]`$result
}

function Wait-ForServicesToSettle {
    param([int]`$MaxSeconds = 300)

    Write-Log "Post-boot verifier waiting `$PostBootInitialWait second(s) before checking state..." 'INFO'
    Start-Sleep -Seconds `$PostBootInitialWait

    `$serviceNames = @('wuauserv','UsoSvc','BITS','TrustedInstaller')
    `$sw = [System.Diagnostics.Stopwatch]::StartNew()

    do {
        `$busy = @()

        foreach (`$svcName in `$serviceNames) {
            try {
                `$svc = Get-Service -Name `$svcName -ErrorAction Stop
                if (`$svc.Status -eq 'Running') {
                    `$busy += `$svcName
                }
            }
            catch {}
        }

        if (`$busy.Count -eq 0) {
            Write-Log "Post-boot verifier sees no active update-related services." 'OK'
            return
        }

        Write-Log "Post-boot verifier still sees service activity: `$(`$busy -join ', ')" 'INFO'
        Start-Sleep -Seconds 10
    }
    while (`$sw.Elapsed.TotalSeconds -lt `$MaxSeconds)

    Write-Log "Post-boot verifier timeout reached; proceeding with current state evaluation." 'WARN'
}

try {
    Write-Log "Post-boot verification started." 'INFO'
    Wait-ForServicesToSettle -MaxSeconds 300

    `$postState = Get-PendingRebootState
    `$postState | ConvertTo-Json -Depth 5 | Set-Content -Path `$PostBootStatePath -Encoding UTF8

    `$success = -not `$postState.WindowsUpdatePending

    if (`$success) {
        Write-Log "SUCCESS: Windows Update / CBS-specific pending reboot flags are no longer present after reboot." 'OK'
        "SUCCESS" | Set-Content -Path `$ResultPath -Encoding UTF8
    }
    else {
        Write-Log "FAILURE: Windows Update / CBS-specific pending reboot flags are still present after reboot." 'ERROR'
        "FAILURE" | Set-Content -Path `$ResultPath -Encoding UTF8
    }

    if (`$postState.GenericPendingOnly) {
        Write-Log "INFO: Some generic reboot indicators remain, but not Windows Update / CBS-specific ones." 'WARN'
    }
}
catch {
    Write-Log "Post-boot verification error: `$(`$_.Exception.Message)" 'ERROR'
    "ERROR" | Set-Content -Path `$ResultPath -Encoding UTF8
}
finally {
    try {
        Unregister-ScheduledTask -TaskName `$TaskName -Confirm:`$false -ErrorAction SilentlyContinue | Out-Null
        Write-Log "Removed startup verification task." 'INFO'
    }
    catch {}
}
"@

    Set-Content -Path $PostBootScriptPath -Value $scriptContent -Encoding UTF8 -Force
    Write-Log "Wrote post-boot verifier script: $PostBootScriptPath" 'OK'
}

function Register-PostBootTask {
    Write-Log "Registering startup scheduled task: $TaskName" 'INFO'

    try {
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
    }
    catch {}

    $action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$PostBootScriptPath`""
    $trigger = New-ScheduledTaskTrigger -AtStartup
    $principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -RunLevel Highest
    $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -MultipleInstances Ignore

    Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Force | Out-Null
    Write-Log "Startup scheduled task registered successfully." 'OK'
}

function Invoke-ImmediateReboot {
    param(
        [int]$Delay = 20,
        [bool]$Force = $true
    )

    $comment = 'Restarting to allow pending Windows Update and servicing operations to complete while system remains thawed.'

    $arguments = @(
        '/r'
        '/t', $Delay.ToString()
        '/d', 'p:2:17'
        '/c', "`"$comment`""
    )

    if ($Force) {
        $arguments += '/f'
    }

    Write-Log "Issuing reboot command: shutdown.exe $($arguments -join ' ')" 'INFO'

    & "$env:SystemRoot\System32\shutdown.exe" @arguments

    $shutdownExitCode = $LASTEXITCODE
    if ($shutdownExitCode -ne 0) {
        throw "shutdown.exe returned exit code $shutdownExitCode"
    }
}

if (-not (Test-IsAdministrator)) {
    throw "This script must be run as Administrator."
}

if (-not (Test-Path -Path $BaseFolder)) {
    New-Item -Path $BaseFolder -ItemType Directory -Force | Out-Null
}

Write-Log "Initializing script..." 'INFO'
Write-Log "Checking pending reboot and Windows Update servicing state..." 'INFO'

$preState = Get-PendingRebootState
$preState | ConvertTo-Json -Depth 5 | Set-Content -Path $PreBootStatePath -Encoding UTF8

$preState | Format-List | Out-String | ForEach-Object {
    $_.TrimEnd() -split "`r?`n" | ForEach-Object {
        if ($_ -match '\S') {
            Write-Host $_ -ForegroundColor Gray
            try { Add-Content -Path $LogPath -Value $_ } catch {}
        }
    }
}

if ($preState.WindowsUpdatePending) {
    Write-Log "Windows Update / CBS indicates a reboot is pending to complete servicing." 'OK'
}
elseif ($preState.GenericPendingOnly) {
    Write-Log "A reboot is pending, but it is not clearly Windows Update-specific." 'WARN'
}
else {
    Write-Log "No pending reboot indicators were detected." 'WARN'
}

if ($WaitForStaging) {
    Wait-ForUpdateStaging -Seconds $StagingWaitSeconds
}

Write-PostBootVerifierScript
Register-PostBootTask

Write-Log "IMPORTANT: This script does NOT clear Windows Update, CBS, or PendingFileRename reboot flags." 'INFO'
Write-Log "The post-boot verifier will log whether Windows Update / CBS reboot flags cleared after startup." 'INFO'
Write-Log "Rebooting now..." 'INFO'

try {
    Invoke-ImmediateReboot -Delay $DelaySeconds -Force:$ForceReboot
}
catch {
    Write-Log "Failed to issue reboot command: $($_.Exception.Message)" 'ERROR'
    exit 2
}

Write-Log "If you are seeing this message, the reboot command may have been blocked or delayed." 'WARN'
exit 0