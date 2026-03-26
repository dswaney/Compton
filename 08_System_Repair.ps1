# =====================================================================
# ScriptName: 08_System_Repair.ps1
# ScriptVersion: 1.3
# LastUpdated: 2026-03-26
# =====================================================================

[CmdletBinding()]
param(
    [switch]$AutoRepairOnDetection = $true,
    [switch]$AllowWmiRepair = $true,
    [switch]$AllowNetworkReset = $false,
    [switch]$AllowWindowsUpdateReset = $false,
    [switch]$AllowOfflineDiskRepair = $false,
    [switch]$AllowFirewallReset = $false,
    [switch]$AllowIconCacheRebuild = $false,
    [switch]$AggressiveCleanup = $false,
    [switch]$ClearEventLogs = $false,
    [switch]$AutoRebootIfNeeded = $false,
    [int]$AutoRebootDelaySeconds = 60,
    [string]$LogDirectory = 'C:\Logs'
)

$ErrorActionPreference = 'Stop'

$script:RunStart = Get-Date
$script:ComputerName = $env:COMPUTERNAME
$script:TimestampForFile = $script:RunStart.ToString('yyyy-MM-dd_HH-mm-ss')
$script:BaseFileName = "{0}-SystemRepair-{1}" -f $script:ComputerName, $script:TimestampForFile
$script:YamlLogPath = Join-Path $LogDirectory ($script:BaseFileName + '.yaml')

$script:Summary = [ordered]@{
    ComputerName                 = $script:ComputerName
    StartTime                    = $script:RunStart
    EndTime                      = $null
    StepsSucceeded               = 0
    StepsFailed                  = 0
    Warnings                     = 0
    RebootRequired               = $false
    PendingRebootDetected        = $false
    DiskCorruptionSuspected      = $false
    DismCorruptionDetected       = $false
    SfcIntegrityViolations       = $false
    WmiRepositoryInconsistent    = $false
    RepairsAttempted             = New-Object System.Collections.Generic.List[string]
    Notes                        = New-Object System.Collections.Generic.List[string]
}

$script:DetailedResults = New-Object System.Collections.Generic.List[object]

function Ensure-LogDirectory {
    if (-not (Test-Path -LiteralPath $LogDirectory)) {
        New-Item -Path $LogDirectory -ItemType Directory -Force | Out-Null
    }
}

function ConvertTo-YamlScalar {
    param(
        [Parameter(Mandatory)][AllowNull()]$Value
    )

    if ($null -eq $Value) {
        return 'null'
    }

    if ($Value -is [bool]) {
        return $Value.ToString().ToLowerInvariant()
    }

    if ($Value -is [datetime]) {
        return "'" + $Value.ToString('yyyy-MM-dd HH:mm:ss') + "'"
    }

    if ($Value -is [int] -or $Value -is [long] -or $Value -is [double] -or $Value -is [decimal]) {
        return [string]$Value
    }

    $text = [string]$Value
    $text = $text -replace "'", "''"
    return "'" + $text + "'"
}

function Add-DetailedResult {
    param(
        [Parameter(Mandatory)][string]$Step,
        [Parameter(Mandatory)][string]$Status,
        [Parameter(Mandatory)][string]$Message,
        [AllowNull()]$Data = $null
    )

    $script:DetailedResults.Add([PSCustomObject]@{
        Timestamp = Get-Date
        Step      = $Step
        Status    = $Status
        Message   = $Message
        Data      = $Data
    }) | Out-Null
}

function Write-YamlLog {
    try {
        Ensure-LogDirectory

        $lines = New-Object System.Collections.Generic.List[string]

        $lines.Add('run:') | Out-Null
        $lines.Add("  computer_name: $(ConvertTo-YamlScalar $script:ComputerName)") | Out-Null
        $lines.Add("  start_time: $(ConvertTo-YamlScalar $script:Summary.StartTime)") | Out-Null
        $lines.Add("  end_time: $(ConvertTo-YamlScalar $script:Summary.EndTime)") | Out-Null
        $lines.Add("  yaml_log_path: $(ConvertTo-YamlScalar $script:YamlLogPath)") | Out-Null
        $lines.Add('') | Out-Null

        $lines.Add('settings:') | Out-Null
        $lines.Add("  auto_repair_on_detection: $(ConvertTo-YamlScalar $AutoRepairOnDetection)") | Out-Null
        $lines.Add("  allow_wmi_repair: $(ConvertTo-YamlScalar $AllowWmiRepair)") | Out-Null
        $lines.Add("  allow_network_reset: $(ConvertTo-YamlScalar $AllowNetworkReset)") | Out-Null
        $lines.Add("  allow_windows_update_reset: $(ConvertTo-YamlScalar $AllowWindowsUpdateReset)") | Out-Null
        $lines.Add("  allow_offline_disk_repair: $(ConvertTo-YamlScalar $AllowOfflineDiskRepair)") | Out-Null
        $lines.Add("  allow_firewall_reset: $(ConvertTo-YamlScalar $AllowFirewallReset)") | Out-Null
        $lines.Add("  allow_icon_cache_rebuild: $(ConvertTo-YamlScalar $AllowIconCacheRebuild)") | Out-Null
        $lines.Add("  aggressive_cleanup: $(ConvertTo-YamlScalar $AggressiveCleanup)") | Out-Null
        $lines.Add("  clear_event_logs: $(ConvertTo-YamlScalar $ClearEventLogs)") | Out-Null
        $lines.Add("  auto_reboot_if_needed: $(ConvertTo-YamlScalar $AutoRebootIfNeeded)") | Out-Null
        $lines.Add("  auto_reboot_delay_seconds: $(ConvertTo-YamlScalar $AutoRebootDelaySeconds)") | Out-Null
        $lines.Add('') | Out-Null

        $lines.Add('summary:') | Out-Null
        $lines.Add("  steps_succeeded: $(ConvertTo-YamlScalar $script:Summary.StepsSucceeded)") | Out-Null
        $lines.Add("  steps_failed: $(ConvertTo-YamlScalar $script:Summary.StepsFailed)") | Out-Null
        $lines.Add("  warnings: $(ConvertTo-YamlScalar $script:Summary.Warnings)") | Out-Null
        $lines.Add("  reboot_required: $(ConvertTo-YamlScalar $script:Summary.RebootRequired)") | Out-Null
        $lines.Add("  pending_reboot_detected: $(ConvertTo-YamlScalar $script:Summary.PendingRebootDetected)") | Out-Null
        $lines.Add("  disk_corruption_suspected: $(ConvertTo-YamlScalar $script:Summary.DiskCorruptionSuspected)") | Out-Null
        $lines.Add("  dism_corruption_detected: $(ConvertTo-YamlScalar $script:Summary.DismCorruptionDetected)") | Out-Null
        $lines.Add("  sfc_integrity_violations: $(ConvertTo-YamlScalar $script:Summary.SfcIntegrityViolations)") | Out-Null
        $lines.Add("  wmi_repository_inconsistent: $(ConvertTo-YamlScalar $script:Summary.WmiRepositoryInconsistent)") | Out-Null
        $lines.Add('') | Out-Null

        $lines.Add('repairs_attempted:') | Out-Null
        if ($script:Summary.RepairsAttempted.Count -gt 0) {
            foreach ($repair in $script:Summary.RepairsAttempted) {
                $lines.Add("  - $(ConvertTo-YamlScalar $repair)") | Out-Null
            }
        }
        else {
            $lines.Add('  []') | Out-Null
        }
        $lines.Add('') | Out-Null

        $lines.Add('notes:') | Out-Null
        if ($script:Summary.Notes.Count -gt 0) {
            foreach ($note in $script:Summary.Notes) {
                $lines.Add("  - $(ConvertTo-YamlScalar $note)") | Out-Null
            }
        }
        else {
            $lines.Add('  []') | Out-Null
        }
        $lines.Add('') | Out-Null

        $lines.Add('detailed_results:') | Out-Null
        if ($script:DetailedResults.Count -gt 0) {
            foreach ($entry in $script:DetailedResults) {
                $lines.Add('  -') | Out-Null
                $lines.Add("    timestamp: $(ConvertTo-YamlScalar $entry.Timestamp)") | Out-Null
                $lines.Add("    step: $(ConvertTo-YamlScalar $entry.Step)") | Out-Null
                $lines.Add("    status: $(ConvertTo-YamlScalar $entry.Status)") | Out-Null
                $lines.Add("    message: $(ConvertTo-YamlScalar $entry.Message)") | Out-Null

                if ($null -eq $entry.Data) {
                    $lines.Add('    data: null') | Out-Null
                }
                elseif ($entry.Data -is [System.Collections.IDictionary]) {
                    $lines.Add('    data:') | Out-Null
                    foreach ($key in $entry.Data.Keys) {
                        $lines.Add("      $key`: $(ConvertTo-YamlScalar $entry.Data[$key])") | Out-Null
                    }
                }
                else {
                    $lines.Add("    data: $(ConvertTo-YamlScalar $entry.Data)") | Out-Null
                }
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

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','OK','WARN','ERROR')][string]$Level = 'INFO'
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[$timestamp] [$('{0,-5}' -f $Level)] $Message"

    switch ($Level) {
        'INFO'  { Write-Host $line -ForegroundColor Cyan }
        'OK'    { Write-Host $line -ForegroundColor Green }
        'WARN'  { Write-Host $line -ForegroundColor Yellow }
        'ERROR' { Write-Host $line -ForegroundColor Red }
    }
}

function Add-Note {
    param([string]$Message)
    $script:Summary.Notes.Add($Message) | Out-Null
}

function Add-RepairAttempt {
    param([string]$Message)
    $script:Summary.RepairsAttempted.Add($Message) | Out-Null
}

function Complete-Step {
    param([string]$Name)
    $script:Summary.StepsSucceeded++
    Write-Log "$Name completed." 'OK'
    Add-DetailedResult -Step $Name -Status 'Succeeded' -Message "$Name completed successfully."
}

function Fail-Step {
    param(
        [string]$Name,
        [string]$Reason
    )
    $script:Summary.StepsFailed++
    Write-Log "$Name failed: $Reason" 'ERROR'
    Add-Note "$Name failed: $Reason"
    Add-DetailedResult -Step $Name -Status 'Failed' -Message $Reason
}

function Warn-Step {
    param(
        [string]$Name,
        [string]$Reason
    )
    $script:Summary.Warnings++
    Write-Log "$Name warning: $Reason" 'WARN'
    Add-Note "$Name warning: $Reason"
    Add-DetailedResult -Step $Name -Status 'Warning' -Message $Reason
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

function Invoke-Safely {
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][scriptblock]$ScriptBlock,
        [switch]$WarnOnly
    )

    try {
        & $ScriptBlock
        Complete-Step -Name $Name
        return $true
    }
    catch {
        if ($WarnOnly) {
            Warn-Step -Name $Name -Reason $_.Exception.Message
        }
        else {
            Fail-Step -Name $Name -Reason $_.Exception.Message
        }
        return $false
    }
}

function Get-PendingRebootState {
    $result = [ordered]@{
        CBServicing_RebootPending         = $false
        WindowsUpdate_RebootRequired      = $false
        SessionManager_PendingFileRename  = $false
        SessionManager_PendingFileRename2 = $false
        UpdateExeVolatile                 = $false
        PackagesPending                   = $false
        WUAU_RebootRequired_COM           = $false
        AnyPendingReboot                  = $false
    }

    $cbsRebootPending = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending'
    $wuRebootRequired = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired'
    $packagesPending  = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\PackagesPending'
    $sessionMgr       = 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager'
    $updateExe        = 'HKLM:\SOFTWARE\Microsoft\Updates'

    try { $result.CBServicing_RebootPending = Test-Path -LiteralPath $cbsRebootPending } catch {}
    try { $result.WindowsUpdate_RebootRequired = Test-Path -LiteralPath $wuRebootRequired } catch {}
    try { $result.PackagesPending = Test-Path -LiteralPath $packagesPending } catch {}

    try {
        $pendingRename = (Get-ItemProperty -Path $sessionMgr -Name 'PendingFileRenameOperations' -ErrorAction Stop).PendingFileRenameOperations
        if ($null -ne $pendingRename -and $pendingRename.Count -gt 0) {
            $result.SessionManager_PendingFileRename = $true
        }
    }
    catch {}

    try {
        $pendingRename2 = (Get-ItemProperty -Path $sessionMgr -Name 'PendingFileRenameOperations2' -ErrorAction Stop).PendingFileRenameOperations2
        if ($null -ne $pendingRename2 -and $pendingRename2.Count -gt 0) {
            $result.SessionManager_PendingFileRename2 = $true
        }
    }
    catch {}

    try {
        $uev = (Get-ItemProperty -Path $updateExe -Name 'UpdateExeVolatile' -ErrorAction Stop).UpdateExeVolatile
        if ($null -ne $uev -and [int]$uev -ne 0) {
            $result.UpdateExeVolatile = $true
        }
    }
    catch {}

    try {
        $sysInfo = New-Object -ComObject Microsoft.Update.SystemInfo
        if ($sysInfo.RebootRequired) {
            $result.WUAU_RebootRequired_COM = $true
        }
    }
    catch {}

    if ($result.Values -contains $true) {
        $result.AnyPendingReboot = $true
    }

    Add-DetailedResult -Step 'PendingRebootCheckData' -Status 'Info' -Message 'Collected pending reboot state.' -Data $result
    [PSCustomObject]$result
}

function Test-SystemDiskIsSSD {
    try {
        $systemDriveLetter = $env:SystemDrive.TrimEnd(':')
        $partition = Get-Partition -DriveLetter $systemDriveLetter -ErrorAction Stop
        $disk = Get-Disk -Number $partition.DiskNumber -ErrorAction Stop
        return ($disk.MediaType -eq 'SSD')
    }
    catch {
        return $false
    }
}

function Invoke-DismCommand {
    param([string[]]$Arguments)

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = "$env:SystemRoot\System32\dism.exe"
    $psi.Arguments = $Arguments -join ' '
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true

    $proc = New-Object System.Diagnostics.Process
    $proc.StartInfo = $psi
    [void]$proc.Start()

    $stdout = $proc.StandardOutput.ReadToEnd()
    $stderr = $proc.StandardError.ReadToEnd()
    $proc.WaitForExit()

    if (-not [string]::IsNullOrWhiteSpace($stdout)) {
        $stdout -split "`r?`n" | Where-Object { $_.Trim() } | ForEach-Object { Write-Log $_ 'INFO' }
    }

    if (-not [string]::IsNullOrWhiteSpace($stderr)) {
        $stderr -split "`r?`n" | Where-Object { $_.Trim() } | ForEach-Object { Write-Log $_ 'WARN' }
    }

    Add-DetailedResult -Step 'DISM' -Status 'Info' -Message ("Executed DISM: " + ($Arguments -join ' ')) -Data @{
        ExitCode = $proc.ExitCode
    }

    return [PSCustomObject]@{
        ExitCode = $proc.ExitCode
        StdOut   = $stdout
        StdErr   = $stderr
    }
}

function Invoke-SfcCommand {
    param([string]$Arguments)

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = "$env:SystemRoot\System32\sfc.exe"
    $psi.Arguments = $Arguments
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true

    $proc = New-Object System.Diagnostics.Process
    $proc.StartInfo = $psi
    [void]$proc.Start()

    $stdout = $proc.StandardOutput.ReadToEnd()
    $stderr = $proc.StandardError.ReadToEnd()
    $proc.WaitForExit()

    if (-not [string]::IsNullOrWhiteSpace($stdout)) {
        $stdout -split "`r?`n" | Where-Object { $_.Trim() } | ForEach-Object { Write-Log $_ 'INFO' }
    }

    if (-not [string]::IsNullOrWhiteSpace($stderr)) {
        $stderr -split "`r?`n" | Where-Object { $_.Trim() } | ForEach-Object { Write-Log $_ 'WARN' }
    }

    Add-DetailedResult -Step 'SFC' -Status 'Info' -Message ("Executed SFC: " + $Arguments) -Data @{
        ExitCode = $proc.ExitCode
    }

    return [PSCustomObject]@{
        ExitCode = $proc.ExitCode
        StdOut   = $stdout
        StdErr   = $stderr
    }
}


function Get-DiskSpaceInfo {
    param([string]$Path)

    try {
        $driveRoot = Split-Path -Path $Path -Qualifier
        if ([string]::IsNullOrWhiteSpace($driveRoot)) {
            $driveRoot = $env:SystemDrive + '\'
        }

        $drive = [System.IO.DriveInfo]::new($driveRoot)
        return @{
            FreeSpace  = [int64]$drive.AvailableFreeSpace
            TotalSize  = [int64]$drive.TotalSize
            UsedSpace  = [int64]($drive.TotalSize - $drive.AvailableFreeSpace)
        }
    }
    catch {
        return $null
    }
}

function Test-SafeCleanupPath {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return $false
    }

    $normalized = $Path.TrimEnd('\')

    $blockedPaths = @(
        'C:\Windows\System32',
        'C:\Windows\SysWOW64',
        'C:\Program Files',
        'C:\Program Files (x86)',
        'C:\Windows\explorer.exe',
        'C:\Windows\System32\drivers'
    )

    foreach ($blocked in $blockedPaths) {
        if ($normalized -ieq $blocked -or $normalized -like ($blocked + '\*')) {
            return $false
        }
    }

    $allowedPatterns = @(
        'C:\Windows\Temp*',
        'C:\Temp*',
        'C:\SWSetup*',
        'C:\system.sav*',
        'C:\Windows\SoftwareDistribution\Download*',
        'C:\Windows\Prefetch*',
        'C:\Windows\Logs\CBS*',
        'C:\ProgramData\Microsoft\Windows\WER\ReportQueue*',
        "$env:TEMP*",
        "$env:LOCALAPPDATA\Temp*",
        "$env:LOCALAPPDATA\Microsoft\Windows\INetCache*",
        "$env:LOCALAPPDATA\Microsoft\Windows\WebCache*",
        "$env:LOCALAPPDATA\CrashDumps*",
        "$env:LOCALAPPDATA\Microsoft\Windows\DeliveryOptimization\Cache*",
        "$env:LOCALAPPDATA\D3DSCache*",
        "$env:LOCALAPPDATA\NVIDIA\DXCache*",
        "$env:LOCALAPPDATA\NVIDIA\GLCache*"
    ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

    foreach ($pattern in $allowedPatterns) {
        if ($normalized -like $pattern) {
            return $true
        }
    }

    return $false
}

function Remove-FolderContents {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Description,
        [switch]$ContentsOnly
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        return @{
            Success    = $true
            SpaceFreed = [int64]0
            ItemCount  = 0
            Message    = 'Path does not exist'
        }
    }

    if (-not (Test-SafeCleanupPath -Path $Path)) {
        return @{
            Success    = $false
            SpaceFreed = [int64]0
            ItemCount  = 0
            Message    = 'Path blocked for security'
        }
    }

    try {
        $items = if ($ContentsOnly) {
            @(Get-ChildItem -LiteralPath $Path -Force -ErrorAction SilentlyContinue)
        }
        else {
            @(Get-ChildItem -LiteralPath $Path -Recurse -Force -File -ErrorAction SilentlyContinue)
        }

        $itemCount = $items.Count
        [int64]$sizeBefore = 0

        if ($itemCount -gt 0) {
            $files = if ($ContentsOnly) {
                $items | Where-Object { -not $_.PSIsContainer }
            }
            else {
                $items
            }

            if ($files.Count -gt 0) {
                $sizeSum = ($files | Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
                if ($null -ne $sizeSum) {
                    $sizeBefore = [int64]$sizeSum
                }
            }
        }

        if ($itemCount -eq 0) {
            return @{
                Success    = $true
                SpaceFreed = [int64]0
                ItemCount  = 0
                Message    = 'Folder is empty'
            }
        }

        if ($ContentsOnly) {
            Get-ChildItem -LiteralPath $Path -Force -ErrorAction Stop | ForEach-Object {
                try {
                    Remove-Item -LiteralPath $_.FullName -Recurse -Force -ErrorAction Stop
                }
                catch {
                    Write-Log "Could not remove $($_.FullName): $($_.Exception.Message)" 'WARN'
                }
            }
        }
        else {
            Remove-Item -LiteralPath $Path -Recurse -Force -ErrorAction Stop
        }

        return @{
            Success    = $true
            SpaceFreed = $sizeBefore
            ItemCount  = $itemCount
            Message    = 'Successfully cleaned'
        }
    }
    catch {
        return @{
            Success    = $false
            SpaceFreed = [int64]0
            ItemCount  = 0
            Message    = $_.Exception.Message
        }
    }
}

function Invoke-WindowsCleanup {
    param([int]$TimeoutSec = 300)

    try {
        $cleanmgrPath = Join-Path $env:SystemRoot 'System32\cleanmgr.exe'
        if (-not (Test-Path -LiteralPath $cleanmgrPath)) {
            throw 'Windows Disk Cleanup utility not found.'
        }

        Write-Log "Starting Windows Disk Cleanup (timeout: ${TimeoutSec}s)..." 'INFO'

        $job = Start-Job -ScriptBlock {
            Start-Process -FilePath "$env:SystemRoot\System32\cleanmgr.exe" -ArgumentList '/SAGERUN:1','/VERYLOWDISK' -NoNewWindow -Wait -PassThru
        }

        $result = Wait-Job -Job $job -Timeout $TimeoutSec

        if ($null -eq $result -or $job.State -eq 'Running') {
            Stop-Job -Job $job -Force | Out-Null
            Remove-Job -Job $job -Force | Out-Null
            throw "Disk Cleanup timed out after $TimeoutSec seconds."
        }

        $proc = Receive-Job -Job $job
        Remove-Job -Job $job -Force | Out-Null

        return @{
            Success  = $true
            ExitCode = $proc.ExitCode
        }
    }
    catch {
        return @{
            Success = $false
            Error   = $_.Exception.Message
        }
    }
}

function Invoke-TempCleanup {
    [int64]$totalSpaceFreed = 0
    $cleanupResults = New-Object System.Collections.Generic.List[object]
    $initialSpace = Get-DiskSpaceInfo -Path $env:SystemDrive

    Write-Log 'Cleaning temporary files and caches...' 'INFO'

    $windowsCleanup = Invoke-WindowsCleanup -TimeoutSec 300
    if ($windowsCleanup.Success) {
        Write-Log 'Windows Disk Cleanup completed successfully.' 'OK'
        $cleanupResults.Add([PSCustomObject]@{
            Path        = 'cleanmgr.exe'
            Description = 'Windows Disk Cleanup'
            ItemCount   = 0
            SpaceFreed  = [int64]0
            Status      = 'Success'
            Message     = "Exit code $($windowsCleanup.ExitCode)"
        }) | Out-Null
    }
    else {
        Warn-Step -Name 'TempCleanup' -Reason "Windows Disk Cleanup failed: $($windowsCleanup.Error)"
        $cleanupResults.Add([PSCustomObject]@{
            Path        = 'cleanmgr.exe'
            Description = 'Windows Disk Cleanup'
            ItemCount   = 0
            SpaceFreed  = [int64]0
            Status      = 'Warning'
            Message     = $windowsCleanup.Error
        }) | Out-Null
    }

    $cleanupTargets = @(
        @{ Path = 'C:\SWSetup'; Description = 'HP Software Setup'; ContentsOnly = $false },
        @{ Path = 'C:\Temp'; Description = 'System Temp'; ContentsOnly = $false },
        @{ Path = 'C:\system.sav'; Description = 'System Save'; ContentsOnly = $false },
        @{ Path = 'C:\Windows\Temp'; Description = 'Windows Temp'; ContentsOnly = $true },
        @{ Path = $env:TEMP; Description = 'User Temp'; ContentsOnly = $true },
        @{ Path = "$env:LOCALAPPDATA\Temp"; Description = 'Local Temp'; ContentsOnly = $true },
        @{ Path = 'C:\Windows\SoftwareDistribution\Download'; Description = 'Windows Update Cache'; ContentsOnly = $true },
        @{ Path = 'C:\Windows\Prefetch'; Description = 'Windows Prefetch'; ContentsOnly = $true },
        @{ Path = 'C:\Windows\Logs\CBS'; Description = 'CBS Logs'; ContentsOnly = $true },
        @{ Path = "$env:LOCALAPPDATA\Microsoft\Windows\INetCache"; Description = 'Internet Cache'; ContentsOnly = $true },
        @{ Path = "$env:LOCALAPPDATA\Microsoft\Windows\WebCache"; Description = 'Web Cache'; ContentsOnly = $true },
        @{ Path = 'C:\ProgramData\Microsoft\Windows\WER\ReportQueue'; Description = 'Error Report Queue'; ContentsOnly = $true },
        @{ Path = "$env:LOCALAPPDATA\CrashDumps"; Description = 'Crash Dumps'; ContentsOnly = $true },
        @{ Path = "$env:LOCALAPPDATA\Microsoft\Windows\DeliveryOptimization\Cache"; Description = 'Delivery Optimization Cache'; ContentsOnly = $true }
    )

    if ($AggressiveCleanup) {
        $cleanupTargets += @(
            @{ Path = "$env:LOCALAPPDATA\D3DSCache"; Description = 'Direct3D Shader Cache'; ContentsOnly = $true },
            @{ Path = "$env:LOCALAPPDATA\NVIDIA\DXCache"; Description = 'NVIDIA DX Cache'; ContentsOnly = $true },
            @{ Path = "$env:LOCALAPPDATA\NVIDIA\GLCache"; Description = 'NVIDIA GL Cache'; ContentsOnly = $true }
        )
    }

    foreach ($target in $cleanupTargets) {
        Write-Log "Cleaning $($target.Description) at $($target.Path)" 'INFO'
        $result = Remove-FolderContents -Path $target.Path -Description $target.Description -ContentsOnly:$target.ContentsOnly

        if ($result.Success) {
            if ($result.SpaceFreed -gt 0) {
                $totalSpaceFreed += $result.SpaceFreed
                $sizeText = if ($result.SpaceFreed -ge 1GB) {
                    '{0} GB' -f [math]::Round($result.SpaceFreed / 1GB, 2)
                }
                else {
                    '{0} MB' -f [math]::Round($result.SpaceFreed / 1MB, 1)
                }
                Write-Log "Cleaned $($target.Description): $sizeText freed across $($result.ItemCount) item(s)." 'OK'
            }
            else {
                Write-Log "$($target.Description): $($result.Message)" 'INFO'
            }

            $cleanupResults.Add([PSCustomObject]@{
                Path        = $target.Path
                Description = $target.Description
                ItemCount   = $result.ItemCount
                SpaceFreed  = [int64]$result.SpaceFreed
                Status      = 'Success'
                Message     = $result.Message
            }) | Out-Null
        }
        else {
            Warn-Step -Name 'TempCleanup' -Reason "$($target.Description) failed: $($result.Message)"
            $cleanupResults.Add([PSCustomObject]@{
                Path        = $target.Path
                Description = $target.Description
                ItemCount   = 0
                SpaceFreed  = [int64]0
                Status      = 'Failed'
                Message     = $result.Message
            }) | Out-Null
        }
    }

    $finalSpace = Get-DiskSpaceInfo -Path $env:SystemDrive
    $actualFreed = [int64]0
    if ($initialSpace -and $finalSpace) {
        $actualFreed = [int64]($finalSpace.FreeSpace - $initialSpace.FreeSpace)
    }

    Add-DetailedResult -Step 'TempCleanup' -Status 'Info' -Message 'Enhanced temporary file cleanup completed.' -Data @{
        EstimatedSpaceFreedMB = [math]::Round($totalSpaceFreed / 1MB, 2)
        ActualSpaceFreedMB    = [math]::Round($actualFreed / 1MB, 2)
        TargetsProcessed      = $cleanupResults.Count
        ResultsJson           = ($cleanupResults | ForEach-Object {
            [ordered]@{
                Path         = $_.Path
                Description  = $_.Description
                ItemCount    = $_.ItemCount
                SpaceFreedMB = [math]::Round(([double]$_.SpaceFreed) / 1MB, 2)
                Status       = $_.Status
                Message      = $_.Message
            }
        } | ConvertTo-Json -Compress)
    }
}

function Invoke-RepairVolumeScan {

    $systemDrive = $env:SystemDrive.TrimEnd(':')
    Write-Log "Running Repair-Volume scan on $($env:SystemDrive)" 'INFO'
    $result = Repair-Volume -DriveLetter $systemDrive -Scan -ErrorAction Stop

    $resultText = $null
    if ($null -ne $result) {
        $resultText = $result | Out-String
        if ($resultText -match 'corrupt|error|repair') {
            $script:Summary.DiskCorruptionSuspected = $true
            Warn-Step -Name 'RepairVolumeScan' -Reason 'Disk scan output suggests errors or corruption may exist.'
        }
    }

    Add-DetailedResult -Step 'RepairVolumeScan' -Status 'Info' -Message 'Repair-Volume scan completed.' -Data @{
        Output = $resultText
    }
}

function Invoke-RepairVolumeOfflineFix {
    $systemDrive = $env:SystemDrive.TrimEnd(':')
    Write-Log "Running offline disk repair on $($env:SystemDrive)" 'WARN'
    Repair-Volume -DriveLetter $systemDrive -OfflineScanAndFix -ErrorAction Stop | Out-Null
    $script:Summary.RebootRequired = $true
    Add-RepairAttempt 'Repair-Volume -OfflineScanAndFix'
    Add-DetailedResult -Step 'OfflineDiskRepair' -Status 'Info' -Message 'Offline disk repair was started.'
}

function Invoke-DismDetection {
    Write-Log "Running DISM CheckHealth..." 'INFO'
    $check = Invoke-DismCommand -Arguments @('/Online','/Cleanup-Image','/CheckHealth')

    Write-Log "Running DISM ScanHealth..." 'INFO'
    $scan = Invoke-DismCommand -Arguments @('/Online','/Cleanup-Image','/ScanHealth')

    $combined = ($check.StdOut + "`n" + $scan.StdOut + "`n" + $check.StdErr + "`n" + $scan.StdErr)

    if ($check.ExitCode -ne 0 -or $scan.ExitCode -ne 0) {
        $script:Summary.DismCorruptionDetected = $true
        Warn-Step -Name 'DISMDetection' -Reason 'DISM detection returned a non-zero exit code.'
        return
    }

    if ($combined -match 'repairable|corrupt|component store corruption|The component store is repairable') {
        $script:Summary.DismCorruptionDetected = $true
        Warn-Step -Name 'DISMDetection' -Reason 'DISM detected component store corruption.'
    }
}

function Invoke-DismRepair {
    Write-Log "Running DISM RestoreHealth..." 'WARN'
    $repair = Invoke-DismCommand -Arguments @('/Online','/Cleanup-Image','/RestoreHealth')
    if ($repair.ExitCode -ne 0) {
        throw "DISM RestoreHealth exited with code $($repair.ExitCode)"
    }

    Write-Log "Running DISM StartComponentCleanup..." 'INFO'
    $cleanup = Invoke-DismCommand -Arguments @('/Online','/Cleanup-Image','/StartComponentCleanup')
    if ($cleanup.ExitCode -ne 0) {
        throw "DISM StartComponentCleanup exited with code $($cleanup.ExitCode)"
    }

    Add-RepairAttempt 'DISM RestoreHealth + StartComponentCleanup'
}

function Invoke-SfcDetection {
    Write-Log "Running SFC verify-only scan..." 'INFO'
    $result = Invoke-SfcCommand -Arguments '/verifyonly'

    $combined = ($result.StdOut + "`n" + $result.StdErr)

    if ($combined -match 'found integrity violations|Windows Resource Protection found integrity violations') {
        $script:Summary.SfcIntegrityViolations = $true
        Warn-Step -Name 'SFCDetection' -Reason 'SFC detected integrity violations.'
    }
    elseif ($result.ExitCode -notin 0,1) {
        $script:Summary.SfcIntegrityViolations = $true
        Warn-Step -Name 'SFCDetection' -Reason "SFC verify returned unusual exit code $($result.ExitCode)."
    }
}

function Invoke-SfcRepair {
    Write-Log "Running SFC /SCANNOW..." 'WARN'
    $result = Invoke-SfcCommand -Arguments '/scannow'

    if ($result.ExitCode -notin 0,1) {
        throw "SFC /SCANNOW exited with code $($result.ExitCode)"
    }

    Add-RepairAttempt 'SFC /SCANNOW'
}

function Invoke-WmiCheck {
    Write-Log "Checking WMI repository consistency..." 'INFO'
    $output = & "$env:SystemRoot\System32\wbem\winmgmt.exe" /verifyrepository 2>&1
    $text = ($output | Out-String).Trim()

    if ($text) {
        $text -split "`r?`n" | Where-Object { $_.Trim() } | ForEach-Object { Write-Log $_ 'INFO' }
    }

    Add-DetailedResult -Step 'WmiRepositoryCheck' -Status 'Info' -Message 'WMI verify completed.' -Data @{
        Output = $text
    }

    if ($text -match 'inconsistent') {
        $script:Summary.WmiRepositoryInconsistent = $true
        Warn-Step -Name 'WmiRepositoryCheck' -Reason 'WMI repository reported as inconsistent.'
    }
}

function Invoke-WmiRepair {
    Write-Log "Repairing WMI repository..." 'WARN'
    $output = & "$env:SystemRoot\System32\wbem\winmgmt.exe" /salvagerepository 2>&1
    $text = ($output | Out-String).Trim()

    if ($text) {
        $text -split "`r?`n" | Where-Object { $_.Trim() } | ForEach-Object { Write-Log $_ 'INFO' }
    }

    Add-RepairAttempt 'winmgmt /salvagerepository'
    Add-DetailedResult -Step 'WmiRepair' -Status 'Info' -Message 'WMI salvage completed.' -Data @{
        Output = $text
    }
}

function Invoke-NetworkReset {
    Write-Log "Flushing DNS cache..." 'INFO'
    ipconfig /flushdns | Out-Null

    Write-Log "Resetting Winsock..." 'WARN'
    netsh winsock reset | Out-Null

    Write-Log "Resetting TCP/IP stack..." 'WARN'
    netsh int ip reset | Out-Null

    $script:Summary.RebootRequired = $true
    Add-RepairAttempt 'Winsock/TCPIP reset'
    Add-DetailedResult -Step 'NetworkReset' -Status 'Info' -Message 'Network reset completed.'
}

function Invoke-DnsFlushOnly {
    Write-Log "Flushing DNS cache..." 'INFO'
    ipconfig /flushdns | Out-Null
    Add-DetailedResult -Step 'DnsFlush' -Status 'Info' -Message 'DNS cache flushed.'
}

function Invoke-StorageOptimization {
    $systemDrive = $env:SystemDrive.TrimEnd(':')
    $isSsd = Test-SystemDiskIsSSD

    if ($isSsd) {
        Write-Log "SSD detected. Running ReTrim on $($env:SystemDrive)" 'INFO'
        Optimize-Volume -DriveLetter $systemDrive -ReTrim -ErrorAction Stop | Out-Null
    }
    else {
        Write-Log "SSD not detected. Running standard optimization on $($env:SystemDrive)" 'INFO'
        Optimize-Volume -DriveLetter $systemDrive -Analyze -ErrorAction Stop | Out-Null
    }

    Add-DetailedResult -Step 'StorageOptimization' -Status 'Info' -Message 'Storage optimization completed.' -Data @{
        IsSSD = $isSsd
    }
}

function Invoke-IconCacheRebuild {
    Write-Log "Rebuilding icon cache..." 'WARN'

    Get-Process explorer -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2

    $iconPaths = @(
        "$env:LOCALAPPDATA\IconCache.db",
        "$env:LOCALAPPDATA\Microsoft\Windows\Explorer"
    )

    foreach ($path in $iconPaths) {
        if (Test-Path -LiteralPath $path) {
            if ((Get-Item -LiteralPath $path).PSIsContainer) {
                Get-ChildItem -LiteralPath $path -Filter 'iconcache*' -Force -ErrorAction SilentlyContinue | ForEach-Object {
                    Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue
                }
            }
            else {
                Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue
            }
        }
    }

    Start-Process explorer.exe
    Add-RepairAttempt 'Icon cache rebuild'
    Add-DetailedResult -Step 'IconCacheRebuild' -Status 'Info' -Message 'Icon cache rebuild completed.'
}

function Invoke-FirewallReset {
    Write-Log "Resetting Windows Firewall to defaults..." 'WARN'
    netsh advfirewall reset | Out-Null
    Add-RepairAttempt 'Firewall reset'
    Add-DetailedResult -Step 'FirewallReset' -Status 'Info' -Message 'Firewall reset completed.'
}

function Invoke-WindowsUpdateComponentReset {
    Write-Log "Resetting Windows Update components..." 'WARN'

    $services = @('wuauserv','bits','cryptsvc','msiserver')
    foreach ($svc in $services) {
        try { Stop-Service -Name $svc -Force -ErrorAction SilentlyContinue } catch {}
    }

    Start-Sleep -Seconds 2

    $paths = @(
        "$env:WINDIR\SoftwareDistribution",
        "$env:WINDIR\System32\catroot2"
    )

    $renamed = @()

    foreach ($path in $paths) {
        if (Test-Path -LiteralPath $path) {
            $backup = "$path.old_$(Get-Date -Format yyyyMMddHHmmss)"
            Rename-Item -LiteralPath $path -NewName (Split-Path -Leaf $backup) -ErrorAction Stop
            Write-Log "Renamed $path to $backup" 'INFO'
            $renamed += $backup
        }
    }

    foreach ($svc in $services) {
        try { Start-Service -Name $svc -ErrorAction SilentlyContinue } catch {}
    }

    $script:Summary.RebootRequired = $true
    Add-RepairAttempt 'Windows Update component reset'
    Add-DetailedResult -Step 'WindowsUpdateComponentReset' -Status 'Info' -Message 'Windows Update components reset.' -Data @{
        RenamedPaths = ($renamed -join '; ')
    }
}

function Invoke-ScheduledTaskHealthCheck {
    Write-Log "Checking scheduled task health..." 'INFO'

    $badTasks = Get-ScheduledTask -ErrorAction SilentlyContinue | Where-Object {
        $_.State -eq 'Unknown' -or $_.TaskPath -eq $null
    }

    $taskNames = @()

    if ($badTasks) {
        foreach ($task in $badTasks) {
            $name = "$($task.TaskPath)$($task.TaskName)"
            $taskNames += $name
            Warn-Step -Name 'ScheduledTaskCheck' -Reason "Task may need review: $name"
        }
    }

    Add-DetailedResult -Step 'ScheduledTaskHealthCheck' -Status 'Info' -Message 'Scheduled task health check completed.' -Data @{
        SuspectTaskCount = $taskNames.Count
        SuspectTasks     = ($taskNames -join '; ')
    }
}

function Invoke-EventLogSummary {
    Write-Log "Collecting recent system health events..." 'INFO'

    try {
        $recentUnexpected = Get-WinEvent -FilterHashtable @{
            LogName = 'System'
            Id      = 41, 6008
            StartTime = (Get-Date).AddDays(-7)
        } -ErrorAction Stop

        $count = @($recentUnexpected).Count

        Add-DetailedResult -Step 'EventLogSummary' -Status 'Info' -Message 'Collected recent unexpected shutdown events.' -Data @{
            UnexpectedShutdownCount = $count
        }

        if ($recentUnexpected) {
            Warn-Step -Name 'EventLogSummary' -Reason "Recent unexpected shutdown events found: $count"
        }
    }
    catch {
        Warn-Step -Name 'EventLogSummary' -Reason $_.Exception.Message
    }
}

function Invoke-EventLogClear {
    Write-Log "Clearing classic event logs..." 'WARN'
    wevtutil el | ForEach-Object {
        try {
            wevtutil cl $_ 2>$null
        }
        catch {
        }
    }
    Add-RepairAttempt 'Event log clear'
    Add-DetailedResult -Step 'EventLogClear' -Status 'Info' -Message 'Event log clear attempted.'
}

function Invoke-DiskSpaceCheck {
    $drive = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='$env:SystemDrive'"
    if ($drive) {
        $freeGB = [math]::Round($drive.FreeSpace / 1GB, 2)
        $sizeGB = [math]::Round($drive.Size / 1GB, 2)
        Write-Log "System drive free space: $freeGB GB of $sizeGB GB" 'INFO'

        Add-DetailedResult -Step 'DiskSpaceCheck' -Status 'Info' -Message 'Disk space checked.' -Data @{
            FreeGB  = $freeGB
            TotalGB = $sizeGB
        }

        if ($freeGB -lt 10) {
            Warn-Step -Name 'DiskSpaceCheck' -Reason "Low free space on system drive: $freeGB GB"
        }
    }
}

function Invoke-DefenderStatusCheck {
    try {
        $status = Get-MpComputerStatus -ErrorAction Stop
        Write-Log "Defender Antivirus Enabled: $($status.AntivirusEnabled)" 'INFO'
        Write-Log "Defender RealTime Protection Enabled: $($status.RealTimeProtectionEnabled)" 'INFO'

        Add-DetailedResult -Step 'DefenderStatusCheck' -Status 'Info' -Message 'Defender status checked.' -Data @{
            AntivirusEnabled          = $status.AntivirusEnabled
            RealTimeProtectionEnabled = $status.RealTimeProtectionEnabled
        }
    }
    catch {
        Warn-Step -Name 'DefenderStatusCheck' -Reason $_.Exception.Message
    }
}

function Invoke-RebootIfNeeded {
    param(
        [int]$DelaySeconds = 60
    )

    $comment = 'Restarting after automated system repair operations.'

    $args = @(
        '/r'
        '/t', $DelaySeconds.ToString()
        '/d', 'p:2:17'
        '/c', "`"$comment`""
        '/f'
    )

    Write-Log "Issuing reboot command: shutdown.exe $($args -join ' ')" 'WARN'
    & "$env:SystemRoot\System32\shutdown.exe" @args

    if ($LASTEXITCODE -ne 0) {
        throw "shutdown.exe returned exit code $LASTEXITCODE"
    }

    Add-DetailedResult -Step 'AutoReboot' -Status 'Info' -Message 'Automatic reboot command issued.' -Data @{
        DelaySeconds = $DelaySeconds
    }
}

function Show-Summary {
    $script:Summary.EndTime = Get-Date

    Write-Log "---------------- Summary ----------------" 'INFO'
    Write-Log "Computer Name: $($script:Summary.ComputerName)" 'INFO'
    Write-Log "Start Time: $($script:Summary.StartTime)" 'INFO'
    Write-Log "End Time:   $($script:Summary.EndTime)" 'INFO'
    Write-Log "YAML Log:   $($script:YamlLogPath)" 'INFO'
    Write-Log "Succeeded:  $($script:Summary.StepsSucceeded)" 'INFO'
    Write-Log "Failed:     $($script:Summary.StepsFailed)" 'INFO'
    Write-Log "Warnings:   $($script:Summary.Warnings)" 'INFO'
    Write-Log "Pending Reboot Detected: $($script:Summary.PendingRebootDetected)" 'INFO'
    Write-Log "Reboot Required: $($script:Summary.RebootRequired)" 'INFO'
    Write-Log "Disk Corruption Suspected: $($script:Summary.DiskCorruptionSuspected)" 'INFO'
    Write-Log "DISM Corruption Detected: $($script:Summary.DismCorruptionDetected)" 'INFO'
    Write-Log "SFC Integrity Violations: $($script:Summary.SfcIntegrityViolations)" 'INFO'
    Write-Log "WMI Repository Inconsistent: $($script:Summary.WmiRepositoryInconsistent)" 'INFO'

    if ($script:Summary.RepairsAttempted.Count -gt 0) {
        Write-Log "Repairs Attempted:" 'INFO'
        foreach ($repair in $script:Summary.RepairsAttempted) {
            Write-Log " - $repair" 'INFO'
        }
    }

    if ($script:Summary.Notes.Count -gt 0) {
        Write-Log "Notes:" 'INFO'
        foreach ($note in $script:Summary.Notes) {
            Write-Log " - $note" 'INFO'
        }
    }
}

if (-not (Test-IsAdministrator)) {
    Write-Error "This script must be run as Administrator."
    exit 1
}

Ensure-LogDirectory
Write-Log "Initializing automated system health and repair script..." 'INFO'
Write-Log "Detection-first mode is enabled. Repairs run automatically only when needed." 'INFO'
Write-Log "Detailed YAML log will be written to $($script:YamlLogPath)" 'INFO'

Invoke-Safely -Name 'DiskSpaceCheck' -ScriptBlock {
    Invoke-DiskSpaceCheck
} -WarnOnly | Out-Null

Invoke-Safely -Name 'PendingRebootCheck' -ScriptBlock {
    $pending = Get-PendingRebootState
    $script:Summary.PendingRebootDetected = [bool]$pending.AnyPendingReboot

    $pending | Format-List | Out-String | ForEach-Object {
        $_.TrimEnd() -split "`r?`n" | Where-Object { $_.Trim() } | ForEach-Object {
            Write-Log $_ 'INFO'
        }
    }

    if ($pending.AnyPendingReboot) {
        Warn-Step -Name 'PendingRebootCheck' -Reason 'A pending reboot was detected before maintenance began.'
    }
} -WarnOnly | Out-Null

Invoke-Safely -Name 'RepairVolumeScan' -ScriptBlock {
    Invoke-RepairVolumeScan
} -WarnOnly | Out-Null

Invoke-Safely -Name 'DISMDetection' -ScriptBlock {
    Invoke-DismDetection
} -WarnOnly | Out-Null

Invoke-Safely -Name 'SFCDetection' -ScriptBlock {
    Invoke-SfcDetection
} -WarnOnly | Out-Null

Invoke-Safely -Name 'WmiRepositoryCheck' -ScriptBlock {
    Invoke-WmiCheck
} -WarnOnly | Out-Null

Invoke-Safely -Name 'DnsFlush' -ScriptBlock {
    Invoke-DnsFlushOnly
} -WarnOnly | Out-Null

Invoke-Safely -Name 'TempCleanup' -ScriptBlock {
    Invoke-TempCleanup
} -WarnOnly | Out-Null

Invoke-Safely -Name 'StorageOptimization' -ScriptBlock {
    Invoke-StorageOptimization
} -WarnOnly | Out-Null

Invoke-Safely -Name 'ScheduledTaskHealthCheck' -ScriptBlock {
    Invoke-ScheduledTaskHealthCheck
} -WarnOnly | Out-Null

Invoke-Safely -Name 'EventLogSummary' -ScriptBlock {
    Invoke-EventLogSummary
} -WarnOnly | Out-Null

Invoke-Safely -Name 'DefenderStatusCheck' -ScriptBlock {
    Invoke-DefenderStatusCheck
} -WarnOnly | Out-Null

if ($AutoRepairOnDetection) {
    if ($script:Summary.DismCorruptionDetected) {
        Invoke-Safely -Name 'DISMRepair' -ScriptBlock {
            Invoke-DismRepair
        } | Out-Null
    }

    if ($script:Summary.SfcIntegrityViolations) {
        Invoke-Safely -Name 'SFCRepair' -ScriptBlock {
            Invoke-SfcRepair
        } | Out-Null
    }

    if ($script:Summary.WmiRepositoryInconsistent -and $AllowWmiRepair) {
        Invoke-Safely -Name 'WMIRepair' -ScriptBlock {
            Invoke-WmiRepair
        } -WarnOnly | Out-Null
    }

    if ($script:Summary.DiskCorruptionSuspected -and $AllowOfflineDiskRepair) {
        Invoke-Safely -Name 'OfflineDiskRepair' -ScriptBlock {
            Invoke-RepairVolumeOfflineFix
        } -WarnOnly | Out-Null
    }
}

if ($AllowNetworkReset) {
    Invoke-Safely -Name 'NetworkReset' -ScriptBlock {
        Invoke-NetworkReset
    } -WarnOnly | Out-Null
}

if ($AllowIconCacheRebuild) {
    Invoke-Safely -Name 'IconCacheRebuild' -ScriptBlock {
        Invoke-IconCacheRebuild
    } -WarnOnly | Out-Null
}

if ($AllowFirewallReset) {
    Invoke-Safely -Name 'FirewallReset' -ScriptBlock {
        Invoke-FirewallReset
    } -WarnOnly | Out-Null
}

if ($AllowWindowsUpdateReset) {
    Invoke-Safely -Name 'WindowsUpdateComponentReset' -ScriptBlock {
        Invoke-WindowsUpdateComponentReset
    } -WarnOnly | Out-Null
}

if ($ClearEventLogs) {
    Invoke-Safely -Name 'EventLogClear' -ScriptBlock {
        Invoke-EventLogClear
    } -WarnOnly | Out-Null
}

Show-Summary
Write-YamlLog

if ($AutoRebootIfNeeded -and ($script:Summary.RebootRequired -or $script:Summary.PendingRebootDetected)) {
    try {
        Invoke-RebootIfNeeded -DelaySeconds $AutoRebootDelaySeconds
        Write-YamlLog
    }
    catch {
        Fail-Step -Name 'AutoReboot' -Reason $_.Exception.Message
        Show-Summary
        Write-YamlLog
        exit 2
    }
}

if ($script:Summary.StepsFailed -gt 0) {
    Write-YamlLog
    exit 2
}
elseif ($script:Summary.RebootRequired -or $script:Summary.PendingRebootDetected) {
    Write-YamlLog
    exit 3010
}
else {
    Write-YamlLog
    exit 0
}
