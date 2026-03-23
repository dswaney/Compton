# ScriptVersion: 1.0
# LastUpdated: 2026-03-23

[CmdletBinding()]
param(
    [switch]$IncludeUnknown = $true,
    [switch]$IncludePinned = $false,
    [switch]$AttemptMSStore = $false,
    [switch]$UpdateOffice = $true,
    [int]$OfficeWaitMinutes = 30,
    [string]$LogPath = "$env:SystemDrive\Temp\Weekend-Apps-Update.log"
)

$ErrorActionPreference = 'Stop'

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
        $logDir = Split-Path -Path $LogPath -Parent
        if (-not [string]::IsNullOrWhiteSpace($logDir) -and -not (Test-Path -Path $logDir)) {
            New-Item -Path $logDir -ItemType Directory -Force | Out-Null
        }
        Add-Content -Path $LogPath -Value $line -Encoding UTF8
    }
    catch {
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

function Get-WingetPath {
    $cmd = Get-Command winget.exe -ErrorAction SilentlyContinue
    if ($cmd) {
        return $cmd.Source
    }

    $paths = @(
        "$env:ProgramFiles\WindowsApps\Microsoft.DesktopAppInstaller_*_x64__8wekyb3d8bbwe\winget.exe",
        "$env:LOCALAPPDATA\Microsoft\WindowsApps\winget.exe"
    )

    foreach ($pattern in $paths) {
        $match = Get-ChildItem -Path $pattern -ErrorAction SilentlyContinue |
            Sort-Object FullName -Descending |
            Select-Object -First 1
        if ($match) {
            return $match.FullName
        }
    }

    return $null
}

function Test-IsNoiseLine {
    param([string]$Line)

    if ([string]::IsNullOrWhiteSpace($Line)) { return $true }

    $trimmed = $Line.Trim()

    if ($trimmed -match '^[\-/\\|]+$') { return $true }
    if ($trimmed -match 'Γû') { return $true }
    if ($trimmed -match '^\d+%$') { return $true }
    if ($trimmed -match '^\d+(\.\d+)?\s*(KB|MB|GB)\s*/\s*\d+(\.\d+)?\s*(KB|MB|GB)$') { return $true }

    return $false
}

function Get-CleanWingetOutput {
    param([object[]]$RawOutput)

    $clean = @()

    foreach ($line in $RawOutput) {
        if ($null -eq $line) { continue }

        $text = [string]$line
        $text = $text -replace "`r", ''
        $text = $text.TrimEnd()

        if (Test-IsNoiseLine -Line $text) { continue }

        $clean += $text
    }

    return @($clean)
}

function Invoke-Winget {
    param(
        [Parameter(Mandatory)]
        [string[]]$Arguments,
        [switch]$IgnoreExitCode
    )

    $wingetPath = Get-WingetPath
    if (-not $wingetPath) {
        throw "winget.exe was not found on this system."
    }

    $display = $Arguments -join ' '
    Write-Log "Running: winget $display" 'INFO'

    $rawOutput = & $wingetPath @Arguments 2>&1
    $exitCode = $LASTEXITCODE
    $cleanOutput = Get-CleanWingetOutput -RawOutput @($rawOutput)

    foreach ($line in $cleanOutput) {
        Write-Log $line 'INFO'
    }

    if (-not $IgnoreExitCode -and $exitCode -ne 0) {
        throw "winget exited with code $exitCode while running: $display"
    }

    return [PSCustomObject]@{
        ExitCode  = $exitCode
        RawOutput = @($rawOutput)
        Output    = @($cleanOutput)
    }
}

function Initialize-NetworkDefaults {
    try {
        [Net.ServicePointManager]::SecurityProtocol =
            [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls13
    }
    catch {
        try {
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        }
        catch {}
    }

    try {
        [Net.ServicePointManager]::DefaultConnectionLimit = 64
    }
    catch {}
}

function Update-WingetSources {
    Write-Log "Refreshing WinGet sources..." 'INFO'
    $null = Invoke-Winget -Arguments @('source', 'update') -IgnoreExitCode
    Write-Log "WinGet source refresh completed." 'OK'
}

function Get-UpgradeInventory {
    param([string]$Source = 'winget')

    $args = @('upgrade', '--source', $Source)
    if ($IncludeUnknown) { $args += '--include-unknown' }
    if ($IncludePinned)  { $args += '--include-pinned' }

    return (Invoke-Winget -Arguments $args -IgnoreExitCode)
}

function Parse-WingetInventory {
    param([string[]]$Lines)

    $mainPackages = @()
    $explicitPackages = @()
    $mode = 'None'

    foreach ($line in $Lines) {
        $trimmed = $line.Trim()

        if ([string]::IsNullOrWhiteSpace($trimmed)) {
            continue
        }

        if ($trimmed -match 'require explicit targeting for upgrade' -or
            $trimmed -match 'need to be explicitly upgraded') {
            $mode = 'Explicit'
            continue
        }

        if ($trimmed -match '^Name\s+Id\s+Version\s+Available(\s+Source)?$') {
            if ($mode -eq 'None') {
                $mode = 'Main'
            }
            continue
        }

        if ($trimmed -match '^\d+\s+upgrades available\.?$') {
            continue
        }

        if ($trimmed -match '^Installing dependencies:$') { continue }
        if ($trimmed -match '^This package requires the following dependencies:$') { continue }
        if ($trimmed -match '^- Packages$') { continue }
        if ($trimmed -match '^[A-Za-z0-9._+-]+\.[A-Za-z0-9.+-]+$') { continue }

        if ($trimmed -match '^\(\d+/\d+\)\s+Found ') { continue }
        if ($trimmed -match '^Found .+ Version .+$') { continue }
        if ($trimmed -match '^This application is licensed to you by its owner\.$') { continue }
        if ($trimmed -match '^Microsoft is not responsible for, nor does it grant any licenses to, third-party packages\.$') { continue }
        if ($trimmed -match '^Downloading ') { continue }
        if ($trimmed -match '^Successfully verified installer hash$') { continue }
        if ($trimmed -match '^Starting package install\.\.\.$') { continue }
        if ($trimmed -match '^Successfully installed$') { continue }
        if ($trimmed -match '^No installed package found matching input criteria\.$') { continue }
        if ($trimmed -match '^A newer version was found, but the install technology is different from the current version installed\..+$') { continue }

        $pattern = '^(?<Name>.+?)\s{2,}(?<Id>[A-Za-z0-9][A-Za-z0-9._-]+)\s{2,}(?<Version>\S+)\s{2,}(?<Available>\S+)(?:\s{2,}(?<Source>\S+))?$'
        if ($trimmed -match $pattern) {
            $obj = [PSCustomObject]@{
                Name      = $matches.Name.Trim()
                Id        = $matches.Id.Trim()
                Version   = $matches.Version.Trim()
                Available = $matches.Available.Trim()
                Source    = if ($matches.Source) { $matches.Source.Trim() } else { '' }
            }

            if ($mode -eq 'Explicit') {
                $explicitPackages += $obj
            }
            elseif ($mode -eq 'Main') {
                $mainPackages += $obj
            }

            continue
        }
    }

    return [PSCustomObject]@{
        Main     = @($mainPackages)
        Explicit = @($explicitPackages)
    }
}

function Invoke-WingetUpgradeAll {
    param([string]$Source = 'winget')

    $args = @(
        'upgrade', '--all',
        '--source', $Source,
        '--silent',
        '--accept-source-agreements',
        '--accept-package-agreements',
        '--disable-interactivity'
    )

    if ($IncludeUnknown) { $args += '--include-unknown' }
    if ($IncludePinned)  { $args += '--include-pinned' }

    return (Invoke-Winget -Arguments $args -IgnoreExitCode)
}

function Invoke-TargetedUpgrade {
    param(
        [Parameter(Mandatory)]$Package,
        [switch]$UseNameFallback
    )

    $args = @(
        'upgrade',
        '--id', $Package.Id,
        '--exact',
        '--silent',
        '--accept-source-agreements',
        '--accept-package-agreements',
        '--disable-interactivity'
    )

    if ($Package.Version -eq 'Unknown' -or $IncludeUnknown) {
        $args += '--include-unknown'
    }

    if (-not [string]::IsNullOrWhiteSpace($Package.Source)) {
        $args += @('--source', $Package.Source)
    }

    $result = Invoke-Winget -Arguments $args -IgnoreExitCode

    if ($result.ExitCode -eq 0) {
        return $result
    }

    $text = ($result.Output -join "`n")
    if ($text -match 'install technology is different from the current version installed') {
        Write-Log "Package $($Package.Id) cannot be upgraded in-place because Winget reports the installed technology differs from the new package. Manual uninstall/reinstall is required." 'WARN'
        return $result
    }

    if ($UseNameFallback) {
        Write-Log "ID-targeted upgrade failed for $($Package.Id). Trying name fallback for $($Package.Name)." 'WARN'

        $fallbackArgs = @(
            'upgrade',
            '--name', $Package.Name,
            '--exact',
            '--silent',
            '--accept-source-agreements',
            '--accept-package-agreements',
            '--disable-interactivity'
        )

        if ($Package.Version -eq 'Unknown' -or $IncludeUnknown) {
            $fallbackArgs += '--include-unknown'
        }

        if (-not [string]::IsNullOrWhiteSpace($Package.Source)) {
            $fallbackArgs += @('--source', $Package.Source)
        }

        return (Invoke-Winget -Arguments $fallbackArgs -IgnoreExitCode)
    }

    return $result
}

function Invoke-PackageSet {
    param(
        [object[]]$Packages,
        [string]$Label,
        [switch]$UseNameFallback
    )

    $failures = 0

    if (-not $Packages -or $Packages.Count -eq 0) {
        Write-Log "No packages found for $Label." 'INFO'
        return 0
    }

    foreach ($pkg in $Packages) {
        Write-Log "$($Label): $($pkg.Name) [$($pkg.Id)] $($pkg.Version) -> $($pkg.Available)" 'WARN'
        $result = Invoke-TargetedUpgrade -Package $pkg -UseNameFallback:$UseNameFallback

        if ($result.ExitCode -eq 0) {
            Write-Log "Targeted upgrade completed for $($pkg.Id)." 'OK'
        }
        else {
            Write-Log "Targeted upgrade failed for $($pkg.Id) with exit code $($result.ExitCode)." 'WARN'
            $failures++
        }
    }

    return $failures
}

function Update-OfficeClickToRun {
    param([int]$WaitMinutes = 30)

    $officePath = 'C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeC2RClient.exe'
    if (-not (Test-Path -Path $officePath)) {
        Write-Log "Office Click-to-Run client not found. Skipping Office update." 'INFO'
        return
    }

    Write-Log "Starting Office Click-to-Run update..." 'INFO'

    $proc = Start-Process -FilePath $officePath `
                          -ArgumentList '/update', 'USER', 'displaylevel=False', 'forceappshutdown=True' `
                          -PassThru `
                          -WindowStyle Hidden

    $completed = $proc.WaitForExit($WaitMinutes * 60 * 1000)

    if (-not $completed) {
        Write-Log "Office update process did not finish within $WaitMinutes minute(s)." 'WARN'
        return
    }

    Write-Log "Office Click-to-Run exited with code $($proc.ExitCode)." 'INFO'
}

function Test-RebootRequired {
    try {
        $sysInfo = New-Object -ComObject Microsoft.Update.SystemInfo
        return [bool]$sysInfo.RebootRequired
    }
    catch {
        return $false
    }
}

function Remove-PackagesById {
    param(
        [object[]]$Packages,
        [string[]]$IdsToRemove
    )

    if (-not $Packages) { return @() }
    if (-not $IdsToRemove -or $IdsToRemove.Count -eq 0) { return @($Packages) }

    $idSet = @{}
    foreach ($id in $IdsToRemove) {
        $idSet[$id.ToLowerInvariant()] = $true
    }

    $filtered = @()
    foreach ($pkg in $Packages) {
        if (-not $idSet.ContainsKey($pkg.Id.ToLowerInvariant())) {
            $filtered += $pkg
        }
    }

    return @($filtered)
}

if (-not (Test-IsAdministrator)) {
    Write-Error "This script must be run as Administrator."
    exit 1
}

Initialize-NetworkDefaults
Write-Log "Initializing application update script..." 'INFO'

try {
    Update-WingetSources

    Write-Log "Collecting pre-upgrade inventory..." 'INFO'
    $preInventoryResult = Get-UpgradeInventory -Source 'winget'
    $preParsed = Parse-WingetInventory -Lines $preInventoryResult.Output
    $preMain = $preParsed.Main
    $preExplicit = $preParsed.Explicit

    Write-Log "Main inventory found $($preMain.Count) upgradeable package(s)." $(if ($preMain.Count -gt 0) { 'OK' } else { 'INFO' })
    Write-Log "Explicit-target inventory found $($preExplicit.Count) package(s)." $(if ($preExplicit.Count -gt 0) { 'WARN' } else { 'INFO' })

    if ($preExplicit.Count -gt 0) {
        foreach ($pkg in $preExplicit) {
            Write-Log "Pre-scan explicit target: $($pkg.Name) [$($pkg.Id)] $($pkg.Version) -> $($pkg.Available)" 'WARN'
        }
    }

    $bulkResult = Invoke-WingetUpgradeAll -Source 'winget'

    # Always process explicit-target packages from the inventory scan.
    $explicitFailures = Invoke-PackageSet -Packages $preExplicit -Label 'Explicit target required' -UseNameFallback

    if ($AttemptMSStore) {
        Write-Log "Attempting second pass against msstore source..." 'INFO'
        $null = Invoke-WingetUpgradeAll -Source 'msstore'
    }

    if ($UpdateOffice) {
        Update-OfficeClickToRun -WaitMinutes $OfficeWaitMinutes
    }

    Write-Log "Collecting post-upgrade inventory..." 'INFO'
    $postInventoryResult = Get-UpgradeInventory -Source 'winget'
    $postParsed = Parse-WingetInventory -Lines $postInventoryResult.Output
    $remainingMain = $postParsed.Main
    $remainingExplicit = $postParsed.Explicit

    # Do not keep retrying packages already handled in explicit-target pass.
    $remainingMain = Remove-PackagesById -Packages $remainingMain -IdsToRemove ($preExplicit.Id)

    $retryFailures = 0
    $retryFailures += Invoke-PackageSet -Packages $remainingMain -Label 'Retry remaining package' -UseNameFallback
    $retryFailures += Invoke-PackageSet -Packages $remainingExplicit -Label 'Retry explicit-target package' -UseNameFallback

    Write-Log "Collecting final inventory..." 'INFO'
    $finalInventoryResult = Get-UpgradeInventory -Source 'winget'
    $finalParsed = Parse-WingetInventory -Lines $finalInventoryResult.Output
    $finalMain = Remove-PackagesById -Packages $finalParsed.Main -IdsToRemove ($preExplicit.Id)
    $finalExplicit = $finalParsed.Explicit

    if ($finalMain.Count -gt 0) {
        Write-Log "Final remaining main packages: $($finalMain.Count)" 'WARN'
        foreach ($pkg in $finalMain) {
            Write-Log "Still remaining: $($pkg.Name) [$($pkg.Id)] $($pkg.Version) -> $($pkg.Available)" 'WARN'
        }
    }
    else {
        Write-Log "Final remaining main packages: 0" 'OK'
    }

    if ($finalExplicit.Count -gt 0) {
        Write-Log "Final remaining explicit-target packages: $($finalExplicit.Count)" 'WARN'
        foreach ($pkg in $finalExplicit) {
            Write-Log "Still explicit-target remaining: $($pkg.Name) [$($pkg.Id)] $($pkg.Version) -> $($pkg.Available)" 'WARN'
        }
    }
    else {
        Write-Log "Final remaining explicit-target packages: 0" 'OK'
    }

    if (Test-RebootRequired) {
        Write-Log "A reboot is required after application updates." 'WARN'
        exit 3010
    }

    if ($explicitFailures -eq 0 -and $retryFailures -eq 0 -and $finalMain.Count -eq 0 -and $finalExplicit.Count -eq 0) {
        Write-Log "Application update script completed successfully." 'OK'
        exit 0
    }
    else {
        Write-Log "Application update script completed with remaining packages or non-zero package results." 'WARN'
        exit 2
    }
}
catch {
    Write-Log "Script failed: $($_.Exception.Message)" 'ERROR'
    exit 3
}