# ScriptVersion: 2.0
# LastUpdated: 2026-04-24 18:10

[CmdletBinding()]
param(
    [switch]$IncludeUnknown = $true,
    [switch]$IncludePinned = $false,
    [switch]$AttemptMSStore = $false,
    [switch]$UpdateOffice = $true,
    [switch]$EnableClassicContextMenu = $true,
    [int]$OfficeWaitMinutes = 30,
    [string]$LogPath = "$env:SystemDrive\Temp\Weekend-Apps-Update.log",
    [switch]$PowerShellGetBootstrapDone
)

$ErrorActionPreference = 'Stop'

# Writes color-coded status output to the console and appends the same message to the log file.
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

# Confirms the script is running elevated because package updates and system changes require admin rights.
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

# Finds winget.exe from the PATH first, then falls back to the common App Installer locations.
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

# Filters out progress bars, separators, and other noisy winget output that does not help with logging or parsing.
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

# Normalizes raw winget output into cleaner line-based text so it can be logged and parsed consistently.
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

# Runs winget with the supplied arguments, captures output, logs it, and returns a structured result object.
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

# Enables modern TLS defaults and raises the connection limit to reduce network-related package source issues.
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

# Applies the registry setting that restores the classic Windows 10 style right-click context menu in Windows 11.
function Set-ClassicContextMenuForHive {
    param(
        [Parameter(Mandatory)]
        [string]$RootKey
    )

    try {
        $basePath = Join-Path $RootKey 'Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}'
        $inprocPath = Join-Path $basePath 'InprocServer32'

        if (-not (Test-Path $basePath)) {
            New-Item -Path $basePath -Force | Out-Null
        }

        if (-not (Test-Path $inprocPath)) {
            New-Item -Path $inprocPath -Force | Out-Null
        }

        New-ItemProperty -Path $inprocPath -Name '(default)' -Value '' -PropertyType String -Force -ErrorAction SilentlyContinue | Out-Null
        Set-ItemProperty -Path $inprocPath -Name '(default)' -Value '' -ErrorAction SilentlyContinue

        Write-Log "Enabled classic context menu for hive: $RootKey" 'OK'
    }
    catch {
        Write-Log "Failed to update hive $RootKey : $($_.Exception.Message)" 'ERROR'
    }
}

function Enable-ClassicContextMenuAllUsers {
    Write-Log 'Applying classic Windows 10-style context menu for all users...' 'INFO'

    # Apply to the current user profile.
    Set-ClassicContextMenuForHive -RootKey 'Registry::HKEY_CURRENT_USER'

    # Apply to all currently loaded user profiles.
    $userSids = Get-ChildItem Registry::HKEY_USERS |
        Where-Object {
            $_.PSChildName -match '^S-1-5-21-' -and
            $_.PSChildName -notmatch '_Classes$'
        } |
        Select-Object -ExpandProperty PSChildName

    foreach ($sid in $userSids) {
        Set-ClassicContextMenuForHive -RootKey "Registry::HKEY_USERS\$sid"
    }

    # Apply to the Default User profile so future users inherit the classic context menu.
    $defaultHiveName = 'HKU\DefaultTemp'
    $defaultHivePsPath = 'Registry::HKEY_USERS\DefaultTemp'
    $defaultUserNtUserDat = 'C:\Users\Default\NTUSER.DAT'

    if (Test-Path $defaultUserNtUserDat) {
        $hiveLoaded = $false

        try {
            $loadResult = & reg.exe load $defaultHiveName $defaultUserNtUserDat 2>&1
            if ($LASTEXITCODE -ne 0) {
                throw "Failed to load Default User hive: $($loadResult -join ' ')"
            }

            $hiveLoaded = $true
            Start-Sleep -Milliseconds 750

            Set-ClassicContextMenuForHive -RootKey $defaultHivePsPath

            # Release handles before unloading the hive.
            Start-Sleep -Milliseconds 750
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
            Start-Sleep -Milliseconds 750
        }
        catch {
            Write-Log "Failed to update Default User profile: $($_.Exception.Message)" 'ERROR'
        }
        finally {
            if ($hiveLoaded) {
                $unloaded = $false
                $unloadResult = $null

                foreach ($attempt in 1..5) {
                    $unloadResult = & reg.exe unload $defaultHiveName 2>&1
                    if ($LASTEXITCODE -eq 0) {
                        $unloaded = $true
                        Write-Log 'Applied classic context menu to Default User profile.' 'OK'
                        break
                    }

                    Start-Sleep -Seconds 1
                    [System.GC]::Collect()
                    [System.GC]::WaitForPendingFinalizers()
                }

                if (-not $unloaded) {
                    Write-Log "Classic context menu was written to Default User profile, but unloading the hive failed after multiple attempts. Last response: $($unloadResult -join ' ')" 'WARN'
                }
            }
        }
    }
    else {
        Write-Log 'Default User NTUSER.DAT not found; future new users were not updated.' 'WARN'
    }

    Write-Log 'Classic context menu registry changes applied. Users may need to sign out and back in.' 'INFO'
}

# Refreshes winget package sources before inventory collection and upgrades are attempted.
function Update-WingetSources {
    Write-Log "Refreshing WinGet sources..." 'INFO'
    $null = Invoke-Winget -Arguments @('source', 'update') -IgnoreExitCode
    Write-Log "WinGet source refresh completed." 'OK'
}

# Retrieves the current list of upgradeable packages from a specific winget source.
function Get-UpgradeInventory {
    param([string]$Source = 'winget')

    $args = @('upgrade', '--source', $Source)
    if ($IncludeUnknown) { $args += '--include-unknown' }
    if ($IncludePinned)  { $args += '--include-pinned' }

    return (Invoke-Winget -Arguments $args -IgnoreExitCode)
}

# Parses winget inventory text into normal upgrades and packages that require explicit targeting.
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

# Performs a bulk upgrade pass against the requested winget source using non-interactive switches.
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

# Upgrades one package by ID first, then optionally retries by package name if the ID-based attempt fails.
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

# Loops through a package collection and performs targeted upgrade attempts while tracking failures.
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

# Starts a Microsoft Office Click-to-Run update and waits up to the configured timeout for completion.
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


# Ensures NuGet exists so PowerShell Gallery module lookups and installs work reliably.
function Ensure-NuGetPackageProvider {
    [CmdletBinding()]
    param()

    try {
        $provider = Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue |
            Sort-Object Version -Descending |
            Select-Object -First 1

        $minimum = [version]'2.8.5.201'
        if ($provider -and ([version]$provider.Version -ge $minimum)) {
            Write-Log "NuGet package provider is current enough: $($provider.Version)" 'OK'
            return $true
        }

        Write-Log "NuGet package provider is missing or outdated. Installing minimum version $minimum..." 'WARN'
        Install-PackageProvider -Name NuGet -MinimumVersion $minimum -Force -ErrorAction Stop | Out-Null
        Write-Log "NuGet package provider installed/updated successfully." 'OK'
        return $true
    }
    catch {
        Write-Log "Failed to install or validate NuGet package provider: $($_.Exception.Message)" 'WARN'
        return $false
    }
}

# Ensures PSGallery is available and trusted enough for unattended module updates.
function Ensure-PSGalleryRepository {
    [CmdletBinding()]
    param()

    try {
        $repo = Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue
        if (-not $repo) {
            Write-Log "PSGallery repository is not registered. Registering default PowerShell repository..." 'WARN'
            Register-PSRepository -Default -ErrorAction Stop
            $repo = Get-PSRepository -Name PSGallery -ErrorAction Stop
        }

        if ($repo.InstallationPolicy -ne 'Trusted') {
            Write-Log "Setting PSGallery installation policy to Trusted for unattended module updates..." 'INFO'
            Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction Stop
        }

        Write-Log "PSGallery repository is available." 'OK'
        return $true
    }
    catch {
        Write-Log "Failed to validate PSGallery repository: $($_.Exception.Message)" 'WARN'
        return $false
    }
}

# Returns the highest locally installed version of a PowerShell module.
function Get-HighestInstalledModuleVersion {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Name
    )

    try {
        $installed = Get-Module -Name $Name -ListAvailable -ErrorAction SilentlyContinue |
            Sort-Object Version -Descending |
            Select-Object -First 1

        if ($installed) {
            return [version]$installed.Version
        }
    }
    catch {
        Write-Log "Could not determine installed version for module ${Name}: $($_.Exception.Message)" 'WARN'
    }

    return $null
}

# Checks whether a command supports a parameter on the current PowerShellGet/PowerShell version.
function Test-CommandParameterAvailable {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$CommandName,

        [Parameter(Mandatory)]
        [string]$ParameterName
    )

    try {
        $command = Get-Command -Name $CommandName -ErrorAction Stop
        return $command.Parameters.ContainsKey($ParameterName)
    }
    catch {
        return $false
    }
}

# Checks the PowerShell Gallery and installs/updates a module only when the online version is newer.
function Update-ModuleIfNewerOnline {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Name,

        [ValidateSet('CurrentUser','AllUsers')]
        [string]$Scope = 'AllUsers',

        [switch]$AcceptLicense
    )

    Write-Log "Checking PowerShell Gallery module: $Name" 'INFO'

    $installedVersion = Get-HighestInstalledModuleVersion -Name $Name
    if ($installedVersion) {
        Write-Log "$Name installed version: $installedVersion" 'INFO'
    }
    else {
        Write-Log "$Name is not currently installed." 'WARN'
    }

    try {
        $galleryModule = Find-Module -Name $Name -Repository PSGallery -ErrorAction Stop
        $onlineVersion = [version]$galleryModule.Version
        Write-Log "$Name latest PSGallery version: $onlineVersion" 'INFO'
    }
    catch {
        Write-Log "Unable to check PSGallery for $Name. Skipping module update. Error: $($_.Exception.Message)" 'WARN'
        return $false
    }

    if ($installedVersion -and $installedVersion -ge $onlineVersion) {
        Write-Log "$Name is already current. No update needed." 'OK'
        try { Import-Module $Name -Force -ErrorAction SilentlyContinue } catch {}
        return $true
    }

    if ($installedVersion) {
        Write-Log "Updating $Name from $installedVersion to $onlineVersion..." 'WARN'
    }
    else {
        Write-Log "Installing $Name version $onlineVersion..." 'WARN'
    }

    $installParams = @{
        Name         = $Name
        Repository   = 'PSGallery'
        Scope        = $Scope
        Force        = $true
        AllowClobber = $true
        ErrorAction  = 'Stop'
    }

    $installSupportsAcceptLicense = Test-CommandParameterAvailable -CommandName 'Install-Module' -ParameterName 'AcceptLicense'

    if ($AcceptLicense) {
        if ($installSupportsAcceptLicense) {
            $installParams['AcceptLicense'] = $true
            Write-Log "Install-Module supports -AcceptLicense. License acceptance parameter will be used for $Name." 'INFO'
        }
        else {
            Write-Log "Install-Module on this system does not support -AcceptLicense. Continuing without that parameter for $Name." 'WARN'
        }
    }

    try {
        Install-Module @installParams
    }
    catch {
        $firstError = $_.Exception.Message

        if ($firstError -match "parameter.*AcceptLicense|AcceptLicense") {
            Write-Log "Install-Module rejected -AcceptLicense for ${Name}; retrying without -AcceptLicense." 'WARN'
            if ($installParams.ContainsKey('AcceptLicense')) {
                $installParams.Remove('AcceptLicense')
            }

            try {
                Install-Module @installParams
            }
            catch {
                Write-Log "Failed to install/update ${Name} after retry without -AcceptLicense: $($_.Exception.Message)" 'WARN'
                return $false
            }
        }
        else {
            Write-Log "Failed to install/update ${Name}: $firstError" 'WARN'
            return $false
        }
    }

    $newVersion = Get-HighestInstalledModuleVersion -Name $Name
    if ($newVersion) {
        Write-Log "$Name installed/updated successfully. Active installed version: $newVersion" 'OK'
    }
    else {
        Write-Log "$Name install command completed, but the installed version could not be verified." 'WARN'
    }

    try {
        Import-Module $Name -Force -ErrorAction Stop
        Write-Log "$Name imported successfully." 'OK'
    }
    catch {
        Write-Log "$Name was installed/updated but could not be imported in the current session: $($_.Exception.Message)" 'WARN'
    }

    return $true
}



# Restarts this same script in a fresh elevated PowerShell session after PowerShellGet/PackageManagement are updated.
# This is needed because Windows PowerShell often keeps the inbox PowerShellGet loaded for the life of the current process.
function Restart-ScriptAfterPowerShellGetBootstrap {
    [CmdletBinding()]
    param()

    if ([string]::IsNullOrWhiteSpace($PSCommandPath) -or -not (Test-Path -LiteralPath $PSCommandPath)) {
        Write-Log 'Cannot restart script automatically because PSCommandPath is unavailable. Run the script again from a new elevated PowerShell session.' 'WARN'
        return $false
    }

    $argList = New-Object System.Collections.Generic.List[string]
    $argList.Add('-NoProfile')
    $argList.Add('-ExecutionPolicy')
    $argList.Add('Bypass')
    $argList.Add('-File')
    $argList.Add(('"{0}"' -f $PSCommandPath))

    if ($IncludeUnknown) { $argList.Add('-IncludeUnknown') }
    if ($IncludePinned) { $argList.Add('-IncludePinned') }
    if ($AttemptMSStore) { $argList.Add('-AttemptMSStore') }
    if ($UpdateOffice) { $argList.Add('-UpdateOffice') }
    if ($EnableClassicContextMenu) { $argList.Add('-EnableClassicContextMenu') }

    $argList.Add('-OfficeWaitMinutes')
    $argList.Add([string]$OfficeWaitMinutes)
    $argList.Add('-LogPath')
    $argList.Add(('"{0}"' -f $LogPath))
    $argList.Add('-PowerShellGetBootstrapDone')

    try {
        Write-Log 'Starting a fresh PowerShell session so the newer PowerShellGet/PackageManagement modules are used...' 'WARN'
        $process = Start-Process -FilePath powershell.exe -ArgumentList ($argList -join ' ') -Wait -PassThru -WindowStyle Normal
        Write-Log "Fresh PowerShell session completed with exit code $($process.ExitCode)." 'OK'
        return $true
    }
    catch {
        Write-Log "Failed to restart script after PowerShellGet bootstrap: $($_.Exception.Message)" 'WARN'
        return $false
    }
}

# Force-updates the PowerShell module installer stack used by Windows PowerShell.
# This installs/updates NuGet, PackageManagement, and PowerShellGet, then requests a fresh PowerShell process when required.
function Invoke-PowerShellGetBootstrap {
    [CmdletBinding()]
    param(
        [version]$MinimumPowerShellGetVersion = [version]'2.2.5',
        [version]$MinimumPackageManagementVersion = [version]'1.4.8.1'
    )

    Write-Log 'Forcing PowerShellGet/PackageManagement bootstrap before HPCMSL update...' 'WARN'

    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Write-Log 'TLS 1.2 enabled for PowerShell Gallery access.' 'OK'
    }
    catch {
        Write-Log "Could not force TLS 1.2: $($_.Exception.Message)" 'WARN'
    }

    try {
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope AllUsers -ErrorAction Stop | Out-Null
        Write-Log 'NuGet package provider installed/validated for AllUsers.' 'OK'
    }
    catch {
        Write-Log "NuGet package provider install/validation failed: $($_.Exception.Message)" 'WARN'
    }

    try {
        Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue
        Write-Log 'PSGallery set to Trusted.' 'OK'
    }
    catch {
        Write-Log "Could not set PSGallery to Trusted: $($_.Exception.Message)" 'WARN'
    }

    $moduleInstallParams = @{
        Repository   = 'PSGallery'
        Scope        = 'AllUsers'
        Force        = $true
        AllowClobber = $true
        ErrorAction  = 'Stop'
    }

    try {
        Write-Log 'Installing/updating PackageManagement from PSGallery...' 'WARN'
        Install-Module @moduleInstallParams -Name 'PackageManagement' -MinimumVersion $MinimumPackageManagementVersion
        Write-Log 'PackageManagement install/update completed.' 'OK'
    }
    catch {
        Write-Log "PackageManagement force update failed: $($_.Exception.Message)" 'WARN'
    }

    try {
        Write-Log 'Installing/updating PowerShellGet from PSGallery...' 'WARN'
        Install-Module @moduleInstallParams -Name 'PowerShellGet' -MinimumVersion $MinimumPowerShellGetVersion
        Write-Log 'PowerShellGet install/update completed.' 'OK'
    }
    catch {
        Write-Log "PowerShellGet force update failed: $($_.Exception.Message)" 'WARN'
        return $false
    }

    $highestPowerShellGet = Get-Module -Name PowerShellGet -ListAvailable -ErrorAction SilentlyContinue |
        Sort-Object Version -Descending |
        Select-Object -First 1

    $highestPackageManagement = Get-Module -Name PackageManagement -ListAvailable -ErrorAction SilentlyContinue |
        Sort-Object Version -Descending |
        Select-Object -First 1

    if ($highestPowerShellGet) {
        Write-Log "Highest installed PowerShellGet after bootstrap: $($highestPowerShellGet.Version) at $($highestPowerShellGet.ModuleBase)" 'INFO'
    }
    if ($highestPackageManagement) {
        Write-Log "Highest installed PackageManagement after bootstrap: $($highestPackageManagement.Version) at $($highestPackageManagement.ModuleBase)" 'INFO'
    }

    try {
        Remove-Module PowerShellGet -Force -ErrorAction SilentlyContinue
        Remove-Module PackageManagement -Force -ErrorAction SilentlyContinue
        Import-Module PackageManagement -MinimumVersion $MinimumPackageManagementVersion -Force -ErrorAction Stop
        Import-Module PowerShellGet -MinimumVersion $MinimumPowerShellGetVersion -Force -ErrorAction Stop

        $loadedPowerShellGet = (Get-Module PowerShellGet | Sort-Object Version -Descending | Select-Object -First 1).Version
        $loadedPackageManagement = (Get-Module PackageManagement | Sort-Object Version -Descending | Select-Object -First 1).Version

        Write-Log "Loaded PackageManagement version: $loadedPackageManagement" 'INFO'
        Write-Log "Loaded PowerShellGet version: $loadedPowerShellGet" 'INFO'

        if ([version]$loadedPowerShellGet -ge $MinimumPowerShellGetVersion) {
            Write-Log 'Current session is now using a modern PowerShellGet module.' 'OK'
            return $true
        }
    }
    catch {
        Write-Log "New PowerShellGet/PackageManagement could not be loaded into this current process: $($_.Exception.Message)" 'WARN'
    }

    return $false
}

# Ensures the current session is using a PowerShellGet version new enough to handle PowerShellGetFormatVersion 2.0 modules.
# Newer HPCMSL dependencies require this.
function Ensure-ModernPowerShellGetForGalleryFormat2 {
    [CmdletBinding()]
    param(
        [version]$MinimumPowerShellGetVersion = [version]'2.2.5'
    )

    Write-Log 'Validating active PowerShellGet support for PowerShellGetFormatVersion 2.0 modules...' 'INFO'

    $loadedVersion = $null
    try {
        $loaded = Get-Module -Name PowerShellGet | Sort-Object Version -Descending | Select-Object -First 1
        if ($loaded) {
            $loadedVersion = [version]$loaded.Version
            Write-Log "Currently loaded PowerShellGet version: $loadedVersion" 'INFO'
        }
    }
    catch {}

    $highestInstalled = Get-Module -Name PowerShellGet -ListAvailable -ErrorAction SilentlyContinue |
        Sort-Object Version -Descending |
        Select-Object -First 1

    if ($highestInstalled) {
        Write-Log "Highest installed PowerShellGet version: $($highestInstalled.Version)" 'INFO'
    }
    else {
        Write-Log 'PowerShellGet is not detected in ListAvailable output.' 'WARN'
    }

    if ($loadedVersion -and $loadedVersion -ge $MinimumPowerShellGetVersion) {
        Write-Log 'Loaded PowerShellGet version is modern enough.' 'OK'
        return $true
    }

    if ($highestInstalled -and ([version]$highestInstalled.Version -ge $MinimumPowerShellGetVersion)) {
        try {
            Remove-Module PowerShellGet -Force -ErrorAction SilentlyContinue
            Remove-Module PackageManagement -Force -ErrorAction SilentlyContinue
            Import-Module PackageManagement -Force -ErrorAction SilentlyContinue
            Import-Module PowerShellGet -MinimumVersion $MinimumPowerShellGetVersion -Force -ErrorAction Stop
            $loadedVersion = [version](Get-Module PowerShellGet | Sort-Object Version -Descending | Select-Object -First 1).Version
            Write-Log "Imported newer installed PowerShellGet version: $loadedVersion" 'OK'
            return $true
        }
        catch {
            Write-Log "A newer PowerShellGet exists but could not be loaded in this process: $($_.Exception.Message)" 'WARN'
        }
    }

    if ($PowerShellGetBootstrapDone) {
        Write-Log 'PowerShellGet bootstrap already ran in this execution chain, but the modern module still is not active.' 'WARN'
        return $false
    }

    $bootstrapReady = Invoke-PowerShellGetBootstrap -MinimumPowerShellGetVersion $MinimumPowerShellGetVersion

    if ($bootstrapReady) {
        return $true
    }

    $restarted = Restart-ScriptAfterPowerShellGetBootstrap
    if ($restarted) {
        Write-Log 'Current script instance will stop because the fresh PowerShell process handled the remaining work.' 'WARN'
        exit 0
    }

    return $false
}


# Installs/updates HPCMSL from a completely fresh Windows PowerShell process.
# This avoids the common issue where the current session keeps the inbox PowerShellGet loaded,
# which cannot install PowerShellGetFormatVersion 2.0 HP dependency modules.
function Invoke-HPCMSLUpdateInFreshPowerShell {
    [CmdletBinding()]
    param(
        [ValidateSet('CurrentUser','AllUsers')]
        [string]$Scope = 'AllUsers',

        [version]$MinimumPowerShellGetVersion = [version]'2.2.5',
        [version]$MinimumPackageManagementVersion = [version]'1.4.8.1'
    )

    Write-Log 'Starting isolated HPCMSL update using PowerShellGet/PSResourceGet when available...' 'WARN'

    $helperPath = Join-Path $env:TEMP ('HPCMSL_Update_Helper_{0}.ps1' -f ([guid]::NewGuid().ToString('N')))
    $helperLog  = Join-Path $env:TEMP ('HPCMSL_Update_Helper_{0}.log' -f ([guid]::NewGuid().ToString('N')))

    $helperScript = @"
`$ErrorActionPreference = 'Stop'
function HelperLog([string]`$Message) {
    `$line = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), `$Message
    Write-Host `$line
    Add-Content -Path '$helperLog' -Value `$line -Encoding UTF8
}

function Get-HighestModuleVersion([string]`$Name) {
    `$m = Get-Module -Name `$Name -ListAvailable -ErrorAction SilentlyContinue | Sort-Object Version -Descending | Select-Object -First 1
    if (`$m) { return [version]`$m.Version }
    return `$null
}

function Get-OnlineModuleVersion([string]`$Name) {
    try {
        `$found = Find-Module -Name `$Name -Repository PSGallery -ErrorAction Stop
        return [version]`$found.Version
    }
    catch {
        HelperLog ('Find-Module failed for {0}: {1}' -f `$Name, `$_.Exception.Message)
        return `$null
    }
}

function Install-PowerShellGalleryModuleManual {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=`$true)][string]`$Name,
        [string]`$Version,
        [string]`$InstallRoot,
        [hashtable]`$Visited
    )

    if (-not `$Visited) { `$Visited = @{} }
    `$visitKey = ('{0}|{1}' -f `$Name.ToLowerInvariant(), `$Version)
    if (`$Visited.ContainsKey(`$visitKey)) { return }
    `$Visited[`$visitKey] = `$true

    if ([string]::IsNullOrWhiteSpace(`$InstallRoot)) {
        throw 'InstallRoot is blank.'
    }

    New-Item -Path `$InstallRoot -ItemType Directory -Force -ErrorAction Stop | Out-Null

    `$tempRoot = Join-Path `$env:TEMP ('PSGalleryManual_{0}_{1}' -f `$Name, ([guid]::NewGuid().ToString('N')))
    New-Item -Path `$tempRoot -ItemType Directory -Force -ErrorAction Stop | Out-Null

    try {
        if ([string]::IsNullOrWhiteSpace(`$Version)) {
            `$packageUrl = 'https://www.powershellgallery.com/api/v2/package/{0}' -f `$Name
        }
        else {
            `$packageUrl = 'https://www.powershellgallery.com/api/v2/package/{0}/{1}' -f `$Name, `$Version
        }

        `$nupkgPath = Join-Path `$tempRoot ('{0}.nupkg' -f `$Name)
        HelperLog ('Manual PSGallery download: {0}' -f `$packageUrl)
        Invoke-WebRequest -Uri `$packageUrl -OutFile `$nupkgPath -UseBasicParsing -ErrorAction Stop

        `$extractPath = Join-Path `$tempRoot 'extract'
        New-Item -Path `$extractPath -ItemType Directory -Force -ErrorAction Stop | Out-Null
        Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
        [System.IO.Compression.ZipFile]::ExtractToDirectory(`$nupkgPath, `$extractPath)

        `$nuspec = Get-ChildItem -Path `$extractPath -Filter '*.nuspec' -File -ErrorAction Stop | Select-Object -First 1
        if (-not `$nuspec) { throw ('No nuspec found in package {0}.' -f `$Name) }

        [xml]`$nuspecXml = Get-Content -LiteralPath `$nuspec.FullName -Raw
        `$metadata = `$nuspecXml.package.metadata
        `$actualName = [string]`$metadata.id
        `$actualVersion = [string]`$metadata.version
        if ([string]::IsNullOrWhiteSpace(`$actualName)) { `$actualName = `$Name }
        if ([string]::IsNullOrWhiteSpace(`$actualVersion)) { throw ('No version found in nuspec for {0}.' -f `$Name) }

        # Install dependencies first. This bypasses old PowerShellGet limitations with PowerShellGetFormatVersion 2.0 modules.
        `$dependencyNodes = @()
        if (`$metadata.dependencies) {
            foreach (`$group in @(`$metadata.dependencies.group)) {
                if (`$group.dependency) { `$dependencyNodes += @(`$group.dependency) }
            }
            if (`$metadata.dependencies.dependency) { `$dependencyNodes += @(`$metadata.dependencies.dependency) }
        }

        foreach (`$dep in `$dependencyNodes) {
            `$depName = [string]`$dep.id
            if ([string]::IsNullOrWhiteSpace(`$depName)) { continue }
            if (`$depName -like 'Microsoft.PowerShell.*') { continue }

            `$depInstalled = Get-HighestModuleVersion -Name `$depName
            if (`$depInstalled) {
                HelperLog ('Dependency already present: {0} {1}' -f `$depName, `$depInstalled)
            }
            else {
                HelperLog ('Installing dependency manually: {0}' -f `$depName)
                Install-PowerShellGalleryModuleManual -Name `$depName -InstallRoot `$InstallRoot -Visited `$Visited
            }
        }

        `$moduleBase = Join-Path `$InstallRoot `$actualName
        `$versionPath = Join-Path `$moduleBase `$actualVersion

        if (Test-Path -LiteralPath `$versionPath) {
            HelperLog ('Module already exists at {0}' -f `$versionPath)
            return
        }

        New-Item -Path `$moduleBase -ItemType Directory -Force -ErrorAction Stop | Out-Null
        New-Item -Path `$versionPath -ItemType Directory -Force -ErrorAction Stop | Out-Null

        HelperLog ('Installing {0} {1} to {2}' -f `$actualName, `$actualVersion, `$versionPath)

        `$skipNames = @('_rels','package','[Content_Types].xml')
        Get-ChildItem -LiteralPath `$extractPath -Force | Where-Object { `$skipNames -notcontains `$_.Name -and `$_.Name -notlike '*.nuspec' } | ForEach-Object {
            Copy-Item -LiteralPath `$_.FullName -Destination `$versionPath -Recurse -Force -ErrorAction Stop
        }
    }
    finally {
        Remove-Item -LiteralPath `$tempRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}

try {
    HelperLog 'Enabling TLS 1.2 for PSGallery.'
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    `$installed = Get-HighestModuleVersion -Name 'HPCMSL'
    `$online = Get-OnlineModuleVersion -Name 'HPCMSL'
    if (-not `$online) { throw 'Could not determine latest HPCMSL version from PSGallery.' }

    if (`$installed) {
        HelperLog ('Installed HPCMSL version: {0}' -f `$installed)
    } else {
        HelperLog 'HPCMSL is not currently installed.'
    }
    HelperLog ('Latest PSGallery HPCMSL version: {0}' -f `$online)

    if (`$installed -and `$installed -ge `$online) {
        HelperLog 'HPCMSL is already current. No update needed.'
        exit 0
    }

    # First try PSResourceGet because it handles newer Gallery package metadata better than old Windows PowerShellGet.
    `$psResourceWorked = `$false
    try {
        HelperLog 'Installing/updating Microsoft.PowerShell.PSResourceGet.'
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope AllUsers -ErrorAction SilentlyContinue | Out-Null
        Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue
        Install-Module -Name Microsoft.PowerShell.PSResourceGet -Repository PSGallery -Scope AllUsers -Force -AllowClobber -SkipPublisherCheck -ErrorAction Stop
        Import-Module Microsoft.PowerShell.PSResourceGet -Force -ErrorAction Stop

        `$installPSResource = Get-Command Install-PSResource -ErrorAction Stop
        HelperLog ('Using Install-PSResource from {0}' -f `$installPSResource.Source)
        Set-PSResourceRepository -Name PSGallery -Trusted -ErrorAction SilentlyContinue
        Install-PSResource -Name HPCMSL -Repository PSGallery -Scope AllUsers -TrustRepository -Reinstall -Quiet -AcceptLicense -ErrorAction Stop
        `$psResourceWorked = `$true
    }
    catch {
        HelperLog ('PSResourceGet install path failed; falling back to manual NuGet extraction. Error: {0}' -f `$_.Exception.Message)
    }

    `$newInstalled = Get-HighestModuleVersion -Name 'HPCMSL'

    if (-not `$psResourceWorked -or -not `$newInstalled -or `$newInstalled -lt `$online) {
        HelperLog 'Starting manual PSGallery NuGet extraction fallback for HPCMSL and HP dependency modules.'
        if ('$Scope' -eq 'AllUsers') {
            `$installRoot = Join-Path `$env:ProgramFiles 'WindowsPowerShell\Modules'
        }
        else {
            `$installRoot = Join-Path `$HOME 'Documents\WindowsPowerShell\Modules'
        }
        Install-PowerShellGalleryModuleManual -Name 'HPCMSL' -Version ([string]`$online) -InstallRoot `$installRoot -Visited @{}
    }

    `$finalInstalled = Get-HighestModuleVersion -Name 'HPCMSL'
    if (-not `$finalInstalled) { throw 'HPCMSL install completed but no installed HPCMSL module was found.' }

    HelperLog ('HPCMSL version after install/update: {0}' -f `$finalInstalled)
    if (`$finalInstalled -lt `$online) {
        throw ('HPCMSL did not update to the latest available version. Installed: {0}; Online: {1}' -f `$finalInstalled, `$online)
    }

    Import-Module HPCMSL -Force -ErrorAction Stop
    HelperLog 'HPCMSL imported successfully in helper process.'
    exit 0
}
catch {
    HelperLog ('ERROR: {0}' -f `$_.Exception.Message)
    exit 1
}
"@

    try {
        Set-Content -Path $helperPath -Value $helperScript -Encoding UTF8 -Force
        Write-Log "Helper script created at $helperPath" 'INFO'

        $args = @('-NoProfile','-ExecutionPolicy','Bypass','-File',('"{0}"' -f $helperPath))
        $process = Start-Process -FilePath powershell.exe -ArgumentList ($args -join ' ') -Wait -PassThru -WindowStyle Hidden

        if (Test-Path -LiteralPath $helperLog) {
            Get-Content -LiteralPath $helperLog -ErrorAction SilentlyContinue | ForEach-Object {
                if (-not [string]::IsNullOrWhiteSpace($_)) {
                    Write-Log "HPCMSL helper: $_" 'INFO'
                }
            }
        }

        if ($process.ExitCode -ne 0) {
            Write-Log "Isolated HPCMSL helper failed with exit code $($process.ExitCode)." 'WARN'
            return $false
        }

        $finalVersion = Get-HighestInstalledModuleVersion -Name 'HPCMSL'
        if ($finalVersion) {
            Write-Log "HPCMSL final installed version after isolated update: $finalVersion" 'OK'
        }
        else {
            Write-Log 'HPCMSL helper completed, but this session could not verify the installed HPCMSL version.' 'WARN'
            return $false
        }

        return $true
    }
    catch {
        Write-Log "Failed to run isolated HPCMSL update helper: $($_.Exception.Message)" 'WARN'
        return $false
    }
    finally {
        Remove-Item -LiteralPath $helperPath -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $helperLog -Force -ErrorAction SilentlyContinue
    }
}


# Integrates the former PS_Update.ps1 and PS_Update_Afterwards.txt logic, but avoids unnecessary reinstalls.
function Update-PowerShellSupportModules {
    [CmdletBinding()]
    param()

    Write-Log "Checking PowerShell support modules before application updates..." 'INFO'

    Initialize-NetworkDefaults

    $nugetOk = Ensure-NuGetPackageProvider
    if (-not $nugetOk) {
        Write-Log "NuGet provider validation failed. Continuing with app updates, but PowerShell module checks may be limited." 'WARN'
    }

    $repoOk = Ensure-PSGalleryRepository
    if (-not $repoOk) {
        Write-Log "PSGallery validation failed. Skipping PowerShell module update checks." 'WARN'
        return
    }

    # Force-refresh the PowerShell module installer stack before HPCMSL.
    # This avoids PowerShellGetFormatVersion 2.0 dependency failures with newer HP modules.
    $modernPowerShellGetReady = Ensure-ModernPowerShellGetForGalleryFormat2

    if (-not $modernPowerShellGetReady) {
        Write-Log 'Skipping HPCMSL update because the active PowerShellGet version cannot install PowerShellGetFormatVersion 2.0 dependencies in this session.' 'WARN'
        Write-Log 'The script can be run again after reboot or after opening a new elevated PowerShell session.' 'WARN'
        return
    }

    # Former PS_Update_Afterwards.txt logic, now version-aware.
    # HPCMSL 1.8.6+ uses HP dependency modules with PowerShellGetFormatVersion 2.0.
    # Run this step in a fresh PowerShell process so the inbox PowerShellGet is not reused.
    $hpcmslUpdated = Invoke-HPCMSLUpdateInFreshPowerShell -Scope AllUsers
    if (-not $hpcmslUpdated) {
        Write-Log 'HPCMSL update did not complete successfully. Continuing with other application updates.' 'WARN'
    }
}

# Checks whether Windows Update reports that a reboot is pending after software maintenance.
function Test-RebootRequired {
    try {
        $sysInfo = New-Object -ComObject Microsoft.Update.SystemInfo
        return [bool]$sysInfo.RebootRequired
    }
    catch {
        return $false
    }
}

# Removes a defined set of package IDs from a package list so they are not retried unnecessarily.
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

# Main execution starts here: validate elevation, prepare networking, apply optional desktop tweak, then process app updates.
if (-not (Test-IsAdministrator)) {
    Write-Error "This script must be run as Administrator."
    exit 1
}

Initialize-NetworkDefaults
Write-Log "Initializing application update script..." 'INFO'
Write-Log "Script version: 2.0 | Last updated: 2026-04-24 18:10" 'INFO'
Update-PowerShellSupportModules

try {
    if ($EnableClassicContextMenu) {
        Write-Log "Applying Option 1: Enable classic Windows 10 style right-click context menu..." 'INFO'
        Enable-ClassicContextMenuAllUsers
        Write-Log "Classic context menu change will fully apply after Explorer restarts or the user signs out and back in." 'INFO'
    }
    else {
        Write-Log "Option 1 skipped because EnableClassicContextMenu was set to false." 'INFO'
    }

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
    exit 1
}
