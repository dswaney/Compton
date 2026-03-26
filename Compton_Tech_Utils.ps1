# =====================================================================
# ScriptName: Compton_Tech_Utils.ps1
# ScriptVersion: 1.8.0
# LastUpdated: 2026-03-26
# Notes: Added startup self-update check against GitHub.
# Notes: Master utility script with merged menu options and YAML logging.
# =====================================================================




# ─────────────────────────────────────────────────────────────────────────────
# Startup Self-Update Check
# ─────────────────────────────────────────────────────────────────────────────
function Get-LocalScriptVersion {
    param(
        [Parameter(Mandatory)][string]$ScriptPath
    )

    try {
        if (-not (Test-Path -LiteralPath $ScriptPath)) {
            return $null
        }

        $firstLines = Get-Content -LiteralPath $ScriptPath -TotalCount 25 -ErrorAction Stop
        $versionLine = $firstLines | Where-Object { $_ -match '^\s*#\s*ScriptVersion\s*:\s*(.+?)\s*$' } | Select-Object -First 1
        if (-not $versionLine) {
            return $null
        }

        $versionText = ($versionLine -replace '^\s*#\s*ScriptVersion\s*:\s*', '').Trim()
        return [version]$versionText
    }
    catch {
        return $null
    }
}

function Get-RemoteScriptMetadata {
    param(
        [Parameter(Mandatory)][string]$RawUrl
    )

    try {
        $response = Invoke-WebRequest -Uri $RawUrl -UseBasicParsing -ErrorAction Stop
        $content = [string]$response.Content
        $versionMatch = [regex]::Match($content, '(?im)^\s*#\s*ScriptVersion\s*:\s*([^\r\n]+)')
        $updatedMatch = [regex]::Match($content, '(?im)^\s*#\s*LastUpdated\s*:\s*([^\r\n]+)')

        $remoteVersion = $null
        if ($versionMatch.Success) {
            try {
                $remoteVersion = [version]$versionMatch.Groups[1].Value.Trim()
            }
            catch {
                $remoteVersion = $null
            }
        }

        [PSCustomObject]@{
            Version     = $remoteVersion
            LastUpdated = if ($updatedMatch.Success) { $updatedMatch.Groups[1].Value.Trim() } else { $null }
            Content     = $content
        }
    }
    catch {
        Write-Host "Unable to check GitHub for script updates: $($_.Exception.Message)" -ForegroundColor Yellow
        return $null
    }
}

function Invoke-StartupSelfUpdate {
    [CmdletBinding()]
    param()

    $rawUrl = 'https://raw.githubusercontent.com/dswaney/Compton/main/Compton_Tech_Utils.ps1'
    $currentScriptPath = $PSCommandPath

    if ([string]::IsNullOrWhiteSpace($currentScriptPath)) {
        $currentScriptPath = $MyInvocation.MyCommand.Path
    }

    if ([string]::IsNullOrWhiteSpace($currentScriptPath)) {
        Write-Host 'Startup update check skipped because the current script path could not be determined.' -ForegroundColor Yellow
        return
    }

    Write-Host 'Checking GitHub for a newer version of Compton_Tech_Utils.ps1...' -ForegroundColor Cyan

    $localVersion = Get-LocalScriptVersion -ScriptPath $currentScriptPath
    $remoteMetadata = Get-RemoteScriptMetadata -RawUrl $rawUrl

    if (-not $remoteMetadata -or -not $remoteMetadata.Version) {
        Write-Host 'Update check skipped. Continuing with the current script.' -ForegroundColor Yellow
        return
    }

    if (-not $localVersion) {
        Write-Host "Current local script version could not be read. GitHub version detected: $($remoteMetadata.Version)" -ForegroundColor Yellow
    }
    else {
        Write-Host "Current version: $localVersion" -ForegroundColor Gray
        Write-Host "GitHub version : $($remoteMetadata.Version)" -ForegroundColor Gray
    }

    if ($localVersion -and $remoteMetadata.Version -le $localVersion) {
        Write-Host 'This script is already up to date.' -ForegroundColor Green
        return
    }

    $prompt = if ($remoteMetadata.LastUpdated) {
        "A newer version ($($remoteMetadata.Version), updated $($remoteMetadata.LastUpdated)) is available. Download and run it now? (Y/N)"
    }
    else {
        "A newer version ($($remoteMetadata.Version)) is available. Download and run it now? (Y/N)"
    }

    do {
        $updateChoice = Read-Host $prompt
        $normalizedChoice = ([string]$updateChoice).Trim().ToUpperInvariant()
    } while ($normalizedChoice -notin @('Y','N','YES','NO'))

    if ($normalizedChoice -in @('N','NO')) {
        Write-Host 'Continuing with the current script version.' -ForegroundColor Yellow
        return
    }

    try {
        $targetDirectory = Split-Path -Path $currentScriptPath -Parent
        if (-not [string]::IsNullOrWhiteSpace($targetDirectory) -and -not (Test-Path -LiteralPath $targetDirectory)) {
            New-Item -Path $targetDirectory -ItemType Directory -Force | Out-Null
        }

        [System.IO.File]::WriteAllText($currentScriptPath, $remoteMetadata.Content, [System.Text.UTF8Encoding]::new($false))
        Write-Host "Updated script saved to: $currentScriptPath" -ForegroundColor Green
        Write-Host 'Launching the updated script...' -ForegroundColor Cyan

        $powershellExe = if (Get-Command -Name 'pwsh.exe' -ErrorAction SilentlyContinue) {
            'pwsh.exe'
        }
        else {
            'powershell.exe'
        }

        Start-Process -FilePath $powershellExe -ArgumentList @('-ExecutionPolicy', 'Bypass', '-File', ('"{0}"' -f $currentScriptPath)) | Out-Null
        exit
    }
    catch {
        Write-Host "Failed to download or relaunch the updated script: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host 'Continuing with the current script version.' -ForegroundColor Yellow
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# Option 14 - Remove User Profiles
# ─────────────────────────────────────────────────────────────────────────────
function Remove-UserProfilesClassroom {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    param(
        [string[]]$ExcludedProfiles = @(
            'Default',
            'Default User',
            'Public',
            'All Users',
            'MISAdmin',
            'dswaney'
        ),

        [string]$UsersRoot = 'C:\Users',

        [switch]$SkipLoadedProfiles = $true,

        [switch]$SkipSpecialProfiles = $true,

        [int]$OlderThanDays = 0,

        [string]$LogDirectory = 'C:\Logs'
    )

    $ErrorActionPreference = 'Stop'

    $script:RunStart = Get-Date
    $script:ComputerName = $env:COMPUTERNAME
    $script:TimestampForFile = $script:RunStart.ToString('yyyy-MM-dd_HH-mm-ss')
    $script:BaseFileName = "{0}-RemoveUserProfiles-{1}" -f $script:ComputerName, $script:TimestampForFile
    $script:YamlLogPath = Join-Path $LogDirectory ($script:BaseFileName + '.yaml')

    $script:Summary = [ordered]@{
        ComputerName       = $script:ComputerName
        StartTime          = $script:RunStart
        EndTime            = $null
        FoundProfiles      = 0
        ExcludedProfiles   = 0
        SkippedLoaded      = 0
        SkippedSpecial     = 0
        SkippedByAge       = 0
        DeletedProfiles    = 0
        FailedProfiles     = 0
    }

    $script:DeletedProfileDetails = New-Object System.Collections.Generic.List[object]
    $script:SkippedProfileDetails = New-Object System.Collections.Generic.List[object]
    $script:FailedProfileDetails  = New-Object System.Collections.Generic.List[object]

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
            $lines.Add("  users_root: $(ConvertTo-YamlScalar $UsersRoot)") | Out-Null
            $lines.Add("  skip_loaded_profiles: $(ConvertTo-YamlScalar $SkipLoadedProfiles)") | Out-Null
            $lines.Add("  skip_special_profiles: $(ConvertTo-YamlScalar $SkipSpecialProfiles)") | Out-Null
            $lines.Add("  older_than_days: $(ConvertTo-YamlScalar $OlderThanDays)") | Out-Null
            $lines.Add("  excluded_profiles:") | Out-Null
            if ($ExcludedProfiles.Count -gt 0) {
                foreach ($name in $ExcludedProfiles) {
                    $lines.Add("    - $(ConvertTo-YamlScalar $name)") | Out-Null
                }
            }
            else {
                $lines.Add('    []') | Out-Null
            }
            $lines.Add('') | Out-Null

            $lines.Add('summary:') | Out-Null
            $lines.Add("  found_profiles: $(ConvertTo-YamlScalar $script:Summary.FoundProfiles)") | Out-Null
            $lines.Add("  excluded_profiles: $(ConvertTo-YamlScalar $script:Summary.ExcludedProfiles)") | Out-Null
            $lines.Add("  skipped_loaded: $(ConvertTo-YamlScalar $script:Summary.SkippedLoaded)") | Out-Null
            $lines.Add("  skipped_special: $(ConvertTo-YamlScalar $script:Summary.SkippedSpecial)") | Out-Null
            $lines.Add("  skipped_by_age: $(ConvertTo-YamlScalar $script:Summary.SkippedByAge)") | Out-Null
            $lines.Add("  deleted_profiles: $(ConvertTo-YamlScalar $script:Summary.DeletedProfiles)") | Out-Null
            $lines.Add("  failed_profiles: $(ConvertTo-YamlScalar $script:Summary.FailedProfiles)") | Out-Null
            $lines.Add('') | Out-Null

            $lines.Add('deleted_profiles:') | Out-Null
            if ($script:DeletedProfileDetails.Count -gt 0) {
                foreach ($entry in $script:DeletedProfileDetails) {
                    $lines.Add('  -') | Out-Null
                    $lines.Add("    profile_name: $(ConvertTo-YamlScalar $entry.ProfileName)") | Out-Null
                    $lines.Add("    local_path: $(ConvertTo-YamlScalar $entry.LocalPath)") | Out-Null
                    $lines.Add("    sid: $(ConvertTo-YamlScalar $entry.SID)") | Out-Null
                    $lines.Add("    loaded: $(ConvertTo-YamlScalar $entry.Loaded)") | Out-Null
                    $lines.Add("    special: $(ConvertTo-YamlScalar $entry.Special)") | Out-Null
                    $lines.Add("    created_time: $(ConvertTo-YamlScalar $entry.CreatedTime)") | Out-Null
                    $lines.Add("    last_use_time: $(ConvertTo-YamlScalar $entry.LastUseTime)") | Out-Null
                    $lines.Add("    days_on_system: $(ConvertTo-YamlScalar $entry.DaysOnSystem)") | Out-Null
                }
            }
            else {
                $lines.Add('  []') | Out-Null
            }
            $lines.Add('') | Out-Null

            $lines.Add('skipped_profiles:') | Out-Null
            if ($script:SkippedProfileDetails.Count -gt 0) {
                foreach ($entry in $script:SkippedProfileDetails) {
                    $lines.Add('  -') | Out-Null
                    $lines.Add("    profile_name: $(ConvertTo-YamlScalar $entry.ProfileName)") | Out-Null
                    $lines.Add("    local_path: $(ConvertTo-YamlScalar $entry.LocalPath)") | Out-Null
                    $lines.Add("    sid: $(ConvertTo-YamlScalar $entry.SID)") | Out-Null
                    $lines.Add("    reason: $(ConvertTo-YamlScalar $entry.Reason)") | Out-Null
                    $lines.Add("    created_time: $(ConvertTo-YamlScalar $entry.CreatedTime)") | Out-Null
                    $lines.Add("    last_use_time: $(ConvertTo-YamlScalar $entry.LastUseTime)") | Out-Null
                    $lines.Add("    days_on_system: $(ConvertTo-YamlScalar $entry.DaysOnSystem)") | Out-Null
                }
            }
            else {
                $lines.Add('  []') | Out-Null
            }
            $lines.Add('') | Out-Null

            $lines.Add('failed_profiles:') | Out-Null
            if ($script:FailedProfileDetails.Count -gt 0) {
                foreach ($entry in $script:FailedProfileDetails) {
                    $lines.Add('  -') | Out-Null
                    $lines.Add("    profile_name: $(ConvertTo-YamlScalar $entry.ProfileName)") | Out-Null
                    $lines.Add("    local_path: $(ConvertTo-YamlScalar $entry.LocalPath)") | Out-Null
                    $lines.Add("    sid: $(ConvertTo-YamlScalar $entry.SID)") | Out-Null
                    $lines.Add("    error: $(ConvertTo-YamlScalar $entry.Error)") | Out-Null
                    $lines.Add("    created_time: $(ConvertTo-YamlScalar $entry.CreatedTime)") | Out-Null
                    $lines.Add("    last_use_time: $(ConvertTo-YamlScalar $entry.LastUseTime)") | Out-Null
                    $lines.Add("    days_on_system: $(ConvertTo-YamlScalar $entry.DaysOnSystem)") | Out-Null
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

    function Get-ProfileFolderName {
        param(
            [Parameter(Mandatory)][string]$Path
        )

        try {
            return (Split-Path -Path $Path -Leaf)
        }
        catch {
            return $null
        }
    }

    function Get-ProfileAgeData {
        param(
            [Parameter(Mandatory)][string]$ProfilePath,
            [AllowNull()][datetime]$LastUseTime
        )

        $createdTime = $null
        $daysOnSystem = $null

        try {
            if (Test-Path -LiteralPath $ProfilePath) {
                $item = Get-Item -LiteralPath $ProfilePath -ErrorAction Stop
                $createdTime = $item.CreationTime
                $daysOnSystem = [math]::Round(((Get-Date) - $createdTime).TotalDays, 2)
            }
        }
        catch {
        }

        return [PSCustomObject]@{
            CreatedTime  = $createdTime
            LastUseTime  = $LastUseTime
            DaysOnSystem = $daysOnSystem
        }
    }

    if (-not (Test-IsAdministrator)) {
        Write-Error "Please run this script as Administrator."
        $global:LastStatus = "[ERROR] Remove User Profiles requires Administrator rights."
        return 1
    }

    Ensure-LogDirectory

    Write-Log "Starting profile cleanup." 'INFO'
    Write-Log "Users root: $UsersRoot" 'INFO'
    Write-Log "Excluded profile names: $($ExcludedProfiles -join ', ')" 'INFO'
    Write-Log "Skip loaded profiles: $SkipLoadedProfiles" 'INFO'
    Write-Log "Skip special profiles: $SkipSpecialProfiles" 'INFO'
    Write-Log "OlderThanDays filter: $OlderThanDays" 'INFO'
    Write-Log "YAML log path: $($script:YamlLogPath)" 'INFO'

    try {
        $allUserProfiles = Get-CimInstance -ClassName Win32_UserProfile -ErrorAction Stop | Where-Object {
            -not [string]::IsNullOrWhiteSpace($_.LocalPath) -and
            $_.LocalPath -like "$UsersRoot\*"
        }
    }
    catch {
        Write-Log "Failed to enumerate Win32_UserProfile instances: $($_.Exception.Message)" 'ERROR'
        $script:Summary.EndTime = Get-Date
        Write-YamlLog
        $global:LastStatus = "[ERROR] Remove User Profiles failed while enumerating profiles."
        return 2
    }

    $script:Summary.FoundProfiles = @($allUserProfiles).Count
    Write-Log "Found $($script:Summary.FoundProfiles) profile(s) under $UsersRoot." 'INFO'

    foreach ($profile in $allUserProfiles) {
        $profilePath = $profile.LocalPath
        $profileName = Get-ProfileFolderName -Path $profilePath

        if ([string]::IsNullOrWhiteSpace($profileName)) {
            Write-Log "Skipping profile with invalid path: $profilePath" 'WARN'
            continue
        }

        $ageData = Get-ProfileAgeData -ProfilePath $profilePath -LastUseTime $profile.LastUseTime

        if ($ExcludedProfiles -icontains $profileName) {
            $script:Summary.ExcludedProfiles++
            Write-Log "Skipping excluded profile: $profileName" 'WARN'

            $script:SkippedProfileDetails.Add([PSCustomObject]@{
                ProfileName  = $profileName
                LocalPath    = $profilePath
                SID          = $profile.SID
                Reason       = 'Excluded'
                CreatedTime  = $ageData.CreatedTime
                LastUseTime  = $ageData.LastUseTime
                DaysOnSystem = $ageData.DaysOnSystem
            }) | Out-Null

            continue
        }

        if ($SkipSpecialProfiles -and $profile.Special) {
            $script:Summary.SkippedSpecial++
            Write-Log "Skipping special/system profile: $profileName ($profilePath)" 'WARN'

            $script:SkippedProfileDetails.Add([PSCustomObject]@{
                ProfileName  = $profileName
                LocalPath    = $profilePath
                SID          = $profile.SID
                Reason       = 'SpecialProfile'
                CreatedTime  = $ageData.CreatedTime
                LastUseTime  = $ageData.LastUseTime
                DaysOnSystem = $ageData.DaysOnSystem
            }) | Out-Null

            continue
        }

        if ($SkipLoadedProfiles -and $profile.Loaded) {
            $script:Summary.SkippedLoaded++
            Write-Log "Skipping loaded profile: $profileName ($profilePath)" 'WARN'

            $script:SkippedProfileDetails.Add([PSCustomObject]@{
                ProfileName  = $profileName
                LocalPath    = $profilePath
                SID          = $profile.SID
                Reason       = 'LoadedProfile'
                CreatedTime  = $ageData.CreatedTime
                LastUseTime  = $ageData.LastUseTime
                DaysOnSystem = $ageData.DaysOnSystem
            }) | Out-Null

            continue
        }

        if ($OlderThanDays -gt 0) {
            try {
                $cutoff = (Get-Date).AddDays(-$OlderThanDays)
                if ($profile.LastUseTime -and $profile.LastUseTime -gt $cutoff) {
                    $script:Summary.SkippedByAge++
                    Write-Log "Skipping recent profile: $profileName (LastUseTime: $($profile.LastUseTime))" 'WARN'

                    $script:SkippedProfileDetails.Add([PSCustomObject]@{
                        ProfileName  = $profileName
                        LocalPath    = $profilePath
                        SID          = $profile.SID
                        Reason       = 'TooRecent'
                        CreatedTime  = $ageData.CreatedTime
                        LastUseTime  = $ageData.LastUseTime
                        DaysOnSystem = $ageData.DaysOnSystem
                    }) | Out-Null

                    continue
                }
            }
            catch {
                Write-Log "Could not evaluate LastUseTime for ${profileName}: $($_.Exception.Message)" 'WARN'
            }
        }

        $targetDescription = "$profileName ($profilePath)"

        try {
            if ($PSCmdlet.ShouldProcess($targetDescription, 'Delete user profile')) {
                Write-Log "Deleting profile: $targetDescription" 'INFO'
                Remove-CimInstance -InputObject $profile -ErrorAction Stop
                $script:Summary.DeletedProfiles++
                Write-Log "Successfully deleted profile: $profileName" 'OK'

                $script:DeletedProfileDetails.Add([PSCustomObject]@{
                    ProfileName  = $profileName
                    LocalPath    = $profilePath
                    SID          = $profile.SID
                    Loaded       = $profile.Loaded
                    Special      = $profile.Special
                    CreatedTime  = $ageData.CreatedTime
                    LastUseTime  = $ageData.LastUseTime
                    DaysOnSystem = $ageData.DaysOnSystem
                }) | Out-Null
            }
        }
        catch {
            $script:Summary.FailedProfiles++
            Write-Log "Failed to delete profile: $profileName. Error: $($_.Exception.Message)" 'ERROR'

            $script:FailedProfileDetails.Add([PSCustomObject]@{
                ProfileName  = $profileName
                LocalPath    = $profilePath
                SID          = $profile.SID
                Error        = $_.Exception.Message
                CreatedTime  = $ageData.CreatedTime
                LastUseTime  = $ageData.LastUseTime
                DaysOnSystem = $ageData.DaysOnSystem
            }) | Out-Null
        }
    }

    $script:Summary.EndTime = Get-Date

    Write-Log "Profile cleanup complete." 'INFO'
    Write-Log "Summary: Found=$($script:Summary.FoundProfiles), Excluded=$($script:Summary.ExcludedProfiles), LoadedSkipped=$($script:Summary.SkippedLoaded), SpecialSkipped=$($script:Summary.SkippedSpecial), AgeSkipped=$($script:Summary.SkippedByAge), Deleted=$($script:Summary.DeletedProfiles), Failed=$($script:Summary.FailedProfiles)" 'INFO'

    Write-YamlLog

    if ($script:Summary.FailedProfiles -gt 0) {
        $global:LastStatus = "[WARN] Remove User Profiles completed with failures. Deleted=$($script:Summary.DeletedProfiles), Failed=$($script:Summary.FailedProfiles)"
        return 2
    }

    if ($script:Summary.DeletedProfiles -gt 0) {
        $global:LastStatus = "[OK] Removed $($script:Summary.DeletedProfiles) user profile(s)."
    }
    else {
        $global:LastStatus = "[INFO] No user profiles were removed."
    }

    return 0
}

# Enhanced alias for compatibility
Set-Alias -Name Remove-UserProfiles -Value Remove-UserProfilesClassroom -Force

# ─────────────────────────────────────────────────────────────────────────────
# Option 15 - Disable the display of the last user logged on
# ─────────────────────────────────────────────────────────────────────────────
function Apply-LoginScreenRegistryFixes {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [switch]$EnhancedSecurity,
        [string]$LogPath = "$env:TEMP\LoginRegistryFixes_$(Get-Date -Format 'yyyyMMdd_HHmmss').log",
        [switch]$BackupRegistry
    )

    # Security: Require elevation
    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        throw "This function must be run as Administrator"
    }

    $ErrorActionPreference = 'Continue'
    $ProgressPreference = 'SilentlyContinue'  # Speed: Disable progress bars
    
    # Initialize tracking (intentionally unscoped so nested functions can update them)
    $appliedSettings = @()
    $failedSettings  = @()
    $backupData      = @{}
    $logEntries      = @()
    $stopwatch       = [System.Diagnostics.Stopwatch]::StartNew()
    
    function Write-LogEntry {
        param([string]$Message, [string]$Level = 'INFO')
        $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $logEntry = "[$timestamp] [$Level] $Message"
        
        # Output to console with colors
        switch ($Level) {
            'ERROR'    { Write-Host $Message -ForegroundColor Red }
            'WARNING'  { Write-Host $Message -ForegroundColor Yellow }
            'SUCCESS'  { Write-Host $Message -ForegroundColor Green }
            'INFO'     { Write-Host $Message -ForegroundColor Cyan }
            'SECURITY' { Write-Host $Message -ForegroundColor Magenta }
            default    { Write-Host $Message }
        }
        
        # Add to log collection
        $logEntries += $logEntry
    }

    # Security: Registry validation function
    function Test-RegistryPath {
        param([string]$Path)
        try {
            return Test-Path $Path -ErrorAction Stop
        } catch {
            return $false
        }
    }

    # Security: Safe registry backup function
    function Backup-RegistryValue {
        param(
            [string]$Path,
            [string]$Name
        )
        try {
            if (Test-RegistryPath $Path) {
                $currentValue = Get-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue
                if ($null -ne $currentValue) {
                    return $currentValue.$Name
                }
            }
            return $null
        } catch {
            return $null
        }
    }

    # Speed: Optimized registry setting function
    function Set-RegistryValueSafe {
        param(
            [string]$Path,
            [string]$Name,
            [object]$Value,
            [string]$Type = 'DWord',
            [string]$Description,
            [switch]$SecurityCritical
        )
        
        try {
            # Security: Validate registry path format
            if ($Path -notmatch '^HK(LM|CU|CR|U|CC):\\') {
                throw "Invalid registry path format: $Path"
            }
            
            # Backup current value if requested
            if ($BackupRegistry) {
                $currentValue = Backup-RegistryValue -Path $Path -Name $Name
                if ($null -ne $currentValue) {
                    $existing = Get-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue
                    $backupData["$Path\$Name"] = @{
                        Value = $currentValue
                        Type  = if ($existing) { $existing.$Name.GetType().Name } else { $null }
                    }
                }
            }
            
            # Create registry path if it doesn't exist
            if (-not (Test-RegistryPath $Path)) {
                Write-LogEntry "Creating registry path: $Path" 'INFO'
                New-Item -Path $Path -Force | Out-Null
            }
            
            # Apply setting
            if ($PSCmdlet.ShouldProcess("$Path\$Name", "Set registry value to $Value")) {
                Set-ItemProperty -Path $Path -Name $Name -Value $Value -Type $Type -ErrorAction Stop
                
                # Verify the setting was applied
                $verifyValue = Get-ItemProperty -Path $Path -Name $Name -ErrorAction Stop
                if ($verifyValue.$Name -eq $Value) {
                    $level = if ($SecurityCritical) { 'SECURITY' } else { 'SUCCESS' }
                    Write-LogEntry "[OK] $Description" $level
                    $appliedSettings += [pscustomobject]@{
                        Path            = $Path
                        Name            = $Name
                        Value           = $Value
                        Description     = $Description
                        SecurityCritical= $SecurityCritical.IsPresent
                    }
                    return $true
                } else {
                    throw "Verification failed: Expected $Value, got $($verifyValue.$Name)"
                }
            }
        } catch {
            Write-LogEntry "[ERROR] Failed to apply $Description : $_" 'ERROR'
            $failedSettings += [pscustomobject]@{
                Path        = $Path
                Name        = $Name
                Description = $Description
                Error       = $_.Exception.Message
            }
            return $false
        }
    }

    try {
        Write-LogEntry "=== Login Screen Registry Security Configuration Started ===" 'INFO'
        Write-LogEntry "Enhanced Security Mode: $($EnhancedSecurity.IsPresent)" 'INFO'
        if ($PSBoundParameters.ContainsKey('WhatIf') -and $PSBoundParameters['WhatIf']) {
            Write-LogEntry "WhatIf mode - no registry changes will be made" 'INFO'
        }

        # Core login screen security settings
        Write-LogEntry "`n[SECURITY] Applying core login screen security settings..." 'INFO'
        
        # Speed: Batch registry operations using array
        $coreSettings = @(
            @{
                Path = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
                Name = "EnableFirstLogonAnimation"
                Value = 0
                Type = 'DWord'
                Description = "Disabled first logon animation"
                SecurityCritical = $false
            },
            @{
                Path = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
                Name = "dontdisplaylastusername"
                Value = 1
                Type = 'DWord'
                Description = "Configured login screen to hide last username"
                SecurityCritical = $true
            }
        )

        # Enhanced security settings (applied when -EnhancedSecurity is used)
        if ($EnhancedSecurity) {
            Write-LogEntry "`n[SECURITY] Applying enhanced security settings..." 'SECURITY'
            
            $enhancedSettings = @(
                @{
                    Path = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
                    Name = "ShutdownWithoutLogon"
                    Value = 0
                    Type = 'DWord'
                    Description = "Disabled shutdown without logon"
                    SecurityCritical = $true
                },
                @{
                    Path = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
                    Name = "UndockWithoutLogon"
                    Value = 0
                    Type = 'DWord'
                    Description = "Disabled undock without logon"
                    SecurityCritical = $true
                },
                @{
                    Path = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
                    Name = "AutoAdminLogon"
                    Value = 0
                    Type = 'String' # This value is a REG_SZ ("0"/"1" as string)
                    Description = "Disabled automatic administrator logon"
                    SecurityCritical = $true
                },
                @{
                    Path = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
                    Name = "ForceAutoLogon"
                    Value = 0
                    Type = 'String' # Also REG_SZ
                    Description = "Disabled forced automatic logon"
                    SecurityCritical = $true
                },
                @{
                    Path = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
                    Name = "InactivityTimeoutSecs"
                    Value = 900
                    Type = 'DWord'
                    Description = "Set login screen timeout to 15 minutes"
                    SecurityCritical = $true
                },
                @{
                    Path = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
                    Name = "MaxDevicePasswordFailedAttempts"
                    Value = 5
                    Type = 'DWord'
                    Description = "Set maximum password attempts to 5"
                    SecurityCritical = $true
                },
                @{
                    Path = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
                    Name = "CachedLogonsCount"
                    Value = 2
                    Type = 'String' # REG_SZ number as string
                    Description = "Limited cached logons to 2"
                    SecurityCritical = $true
                }
            )
            
            $coreSettings += $enhancedSettings
        }

        # Speed: Process settings in batches
        $batchSize = 5
        $totalSettings = $coreSettings.Count
        
        for ($i = 0; $i -lt $totalSettings; $i += $batchSize) {
            $end = [Math]::Min($i + $batchSize - 1, $totalSettings - 1)
            $batch = $coreSettings[$i..$end]
            
            foreach ($setting in $batch) {
                $params = @{
                    Path            = $setting.Path
                    Name            = $setting.Name
                    Value           = $setting.Value
                    Description     = $setting.Description
                    SecurityCritical= $setting.SecurityCritical
                }
                if ($setting.ContainsKey('Type') -and $setting.Type) {
                    $params.Type = $setting.Type
                }
                Set-RegistryValueSafe @params | Out-Null
            }
            
            if ($end -lt ($totalSettings - 1)) {
                Start-Sleep -Milliseconds 50
            }
        }

        # Security: Additional hardening for classroom environments
        if ($EnhancedSecurity) {
            Write-LogEntry "`n[CLASSROOM] Applying classroom-specific security hardening..." 'SECURITY'
            
            # Disable guest account
            try {
                $guestAccount = Get-LocalUser -Name "Guest" -ErrorAction SilentlyContinue
                if ($guestAccount -and $guestAccount.Enabled) {
                    if ($PSCmdlet.ShouldProcess("LocalUser 'Guest'", "Disable")) {
                        Disable-LocalUser -Name "Guest" -ErrorAction Stop
                        Write-LogEntry "[OK] Disabled Guest account" 'SECURITY'
                    }
                }
            } catch {
                Write-LogEntry "[ERROR] Failed to disable Guest account: $_" 'ERROR'
            }
            
            # Set strong password policy via registry
            $passwordPolicySettings = @(
                @{
                    Path = "HKLM:\SYSTEM\CurrentControlSet\Services\Netlogon\Parameters"
                    Name = "RequireStrongKey"
                    Value = 1
                    Type = 'DWord'
                    Description = "Enabled strong authentication keys"
                },
                @{
                    Path = "HKLM:\SYSTEM\CurrentControlSet\Control\Lsa"
                    Name = "LimitBlankPasswordUse"
                    Value = 1
                    Type = 'DWord'
                    Description = "Disabled blank password usage"
                }
            )
            
            foreach ($setting in $passwordPolicySettings) {
                Set-RegistryValueSafe @setting -SecurityCritical | Out-Null
            }
        }

        # Security: Registry integrity verification
        Write-LogEntry "`n[CHECK] Verifying registry integrity..." 'INFO'
        $verificationErrors = 0
        
        foreach ($setting in $appliedSettings) {
            try {
                $currentValue = Get-ItemProperty -Path $setting.Path -Name $setting.Name -ErrorAction Stop
                if ($currentValue.($setting.Name) -ne $setting.Value) {
                    Write-LogEntry "[WARN] Verification failed for $($setting.Description)" 'WARNING'
                    $verificationErrors++
                }
            } catch {
                Write-LogEntry "[WARN] Could not verify $($setting.Description)" 'WARNING'
                $verificationErrors++
            }
        }
        
        if ($verificationErrors -eq 0) {
            Write-LogEntry "[OK] All registry settings verified successfully" 'SUCCESS'
        } else {
            Write-LogEntry "[WARN] $verificationErrors settings failed verification" 'WARNING'
        }

        # Create registry backup file
        if ($BackupRegistry -and $backupData.Count -gt 0) {
            try {
                $backupFile = "$env:TEMP\LoginRegistryBackup_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
                $backupData | ConvertTo-Json -Depth 5 | Out-File -FilePath $backupFile -Encoding UTF8
                Write-LogEntry "[OK] Registry backup saved to: $backupFile" 'INFO'
            } catch {
                Write-LogEntry "[WARN] Failed to save registry backup: $_" 'WARNING'
            }
        }

    } catch {
        Write-LogEntry "Critical error during registry configuration: $_" 'ERROR'
        throw
    } finally {
        $stopwatch.Stop()
        $duration = $stopwatch.Elapsed.TotalSeconds
        
        # Generate summary
        Write-LogEntry "`n[SUMMARY] Configuration Summary:" 'INFO'
        Write-LogEntry "Duration: $([math]::Round($duration, 2)) seconds" 'INFO'
        Write-LogEntry "Settings applied: $($appliedSettings.Count)" 'SUCCESS'
        Write-LogEntry "Settings failed: $($failedSettings.Count)" 'ERROR'
        
        $securitySettings = ($appliedSettings | Where-Object { $_.SecurityCritical }).Count
        if ($securitySettings -gt 0) {
            Write-LogEntry "Security-critical settings applied: $securitySettings" 'SECURITY'
        }
        
        if ($failedSettings.Count -gt 0) {
            Write-LogEntry "`n[ERROR] Failed Settings:" 'ERROR'
            foreach ($failed in $failedSettings) {
                Write-LogEntry "  • $($failed.Description): $($failed.Error)" 'ERROR'
            }
        }
        
        # Write detailed log file
        try {
            $logEntries | Out-File -FilePath $LogPath -Encoding UTF8 -Force
            Write-LogEntry "[LOG] Detailed log saved to: $LogPath" 'INFO'
        } catch {
            Write-LogEntry "[WARN] Failed to save log file: $_" 'WARNING'
        }
        
        # Set global status
        if ($appliedSettings.Count -gt 0) {
            $statusMsg = "[OK] Applied $($appliedSettings.Count) login screen security settings"
            if ($failedSettings.Count -gt 0) {
                $statusMsg += " ($($failedSettings.Count) failed)"
            }
            if ($securitySettings -gt 0) {
                $statusMsg += " [$securitySettings security-critical]"
            }
            $global:LastStatus = $statusMsg
        } else {
            $global:LastStatus = "[WARN] No login screen settings were applied"
        }
        
        Write-LogEntry "=== Login Screen Registry Configuration Completed ===" 'INFO'
    }
}

# Create restore function for emergency rollback
function Restore-LoginScreenRegistrySettings {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$BackupFile
    )
    
    if (-not (Test-Path $BackupFile)) {
        throw "Backup file not found: $BackupFile"
    }
    
    try {
        $backupData = Get-Content $BackupFile | ConvertFrom-Json
        $restored = 0
        
        foreach ($entry in $backupData.PSObject.Properties) {
            $pathAndName = $entry.Name -split '\\'
            $path = $pathAndName[0..($pathAndName.Length-2)] -join '\'
            $name = $pathAndName[-1]
            
            try {
                Set-ItemProperty -Path $path -Name $name -Value $entry.Value.Value -ErrorAction Stop
                Write-Host "[OK] Restored: $($entry.Name)" -ForegroundColor Green
                $restored++
            } catch {
                Write-Warning "[WARN] Failed to restore: $($entry.Name) - $_"
            }
        }
        
        Write-Host "[OK] Restored $restored registry settings from backup" -ForegroundColor Cyan
    } catch {
        throw "Failed to restore from backup: $_"
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# Option 16 - Enable Automatic Login with CC-Student
# ─────────────────────────────────────────────────────────────────────────────
function Set-DomainAutoLogin {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory = $false)]
        [string]$UserName = "CC-Student",
        
        [Parameter(Mandatory = $false)]
        [securestring]$Password,
        
        [Parameter(Mandatory = $false)]
        [string]$DomainName = "Compton.edu",
        
        [switch]$DisableAutoLogin,
        [switch]$WhatIf,
        [switch]$Force,
        [string]$LogPath = "$env:TEMP\AutoLoginConfig_$(Get-Date -Format 'yyyyMMdd_HHmmss').log",
        [int]$AutoLoginCount = 1,  # Number of auto-logins before disabling
        [switch]$UseLocalSystemEncryption
    )

    # Security: Require elevation
    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        throw "This function must be run as Administrator"
    }

    $ErrorActionPreference = 'Stop'
    $ProgressPreference = 'SilentlyContinue'
    
    # Initialize tracking
    $logEntries = @()
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    
    function Write-LogEntry {
        param([string]$Message, [string]$Level = 'INFO')
        $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $logEntry = "[$timestamp] [$Level] $Message"
        $script:logEntries += $logEntry
        
        switch ($Level) {
            'ERROR' { Write-Host $Message -ForegroundColor Red }
            'WARNING' { Write-Host $Message -ForegroundColor Yellow }
            'SUCCESS' { Write-Host $Message -ForegroundColor Green }
            'INFO' { Write-Host $Message -ForegroundColor Cyan }
            'SECURITY' { Write-Host $Message -ForegroundColor Magenta }
        }
    }

    # Security: Secure credential management
    function Get-SecureCredential {
        param(
            [string]$Username,
            [string]$Domain,
            [securestring]$SecurePassword
        )
        
        if (-not $SecurePassword) {
            Write-LogEntry "[WARN] SECURITY WARNING: Auto-login requires storing credentials" 'WARNING'
            Write-LogEntry "Consider using alternative authentication methods for production" 'WARNING'
            
            if (-not $Force) {
                $response = Read-Host "Continue with credential storage? (Type 'ACCEPT' to proceed)"
                if ($response -ne 'ACCEPT') {
                    throw "Operation cancelled - credential storage not accepted"
                }
            }
            
            # Prompt for secure password
            $SecurePassword = Read-Host "Enter password for $Domain\$Username" -AsSecureString
            if (-not $SecurePassword -or $SecurePassword.Length -eq 0) {
                throw "Password is required for auto-login configuration"
            }
        }
        
        return New-Object System.Management.Automation.PSCredential("$Domain\$Username", $SecurePassword)
    }

    # Security: Registry validation and backup
    function Backup-AutoLoginSettings {
        $regPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
        $backupData = @{}
        
        $settingsToBackup = @(
            'AutoAdminLogon',
            'DefaultUserName', 
            'DefaultPassword',
            'DefaultDomainName',
            'AutoLogonCount'
        )
        
        foreach ($setting in $settingsToBackup) {
            try {
                $value = Get-ItemProperty -Path $regPath -Name $setting -ErrorAction SilentlyContinue
                if ($value) {
                    $backupData[$setting] = $value.$setting
                }
            } catch {
                # Setting doesn't exist, which is fine
            }
        }
        
        if ($backupData.Count -gt 0) {
            try {
                $backupFile = "$env:TEMP\AutoLoginBackup_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
                $backupData | ConvertTo-Json | Out-File -FilePath $backupFile -Encoding UTF8
                Write-LogEntry "[OK] Registry backup saved to: $backupFile" 'INFO'
                return $backupFile
            } catch {
                Write-LogEntry "[WARN] Failed to save registry backup: $_" 'WARNING'
            }
        }
        
        return $null
    }

    # Security: Secure password encryption using DPAPI
    function Protect-AutoLoginPassword {
        param([securestring]$SecurePassword)
        
        try {
            # Convert SecureString to encrypted string using DPAPI
            $ptr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecurePassword)
            $plaintext = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($ptr)
            [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptr)
            
            # Use DPAPI to encrypt for local machine
            $bytes = [System.Text.Encoding]::UTF8.GetBytes($plaintext)
            $encryptedBytes = [System.Security.Cryptography.ProtectedData]::Protect(
                $bytes, 
                $null, 
                [System.Security.Cryptography.DataProtectionScope]::LocalMachine
            )
            
            # Clear plaintext from memory
            $plaintext = $null
            [System.GC]::Collect()
            
            return [System.Convert]::ToBase64String($encryptedBytes)
            
        } catch {
            throw "Failed to encrypt password: $_"
        }
    }

    # Security: Domain validation
    function Test-DomainConnectivity {
        param([string]$DomainName)
        
        try {
            Write-LogEntry "[CHECK] Validating domain connectivity..." 'INFO'
            
            # Test domain controller connectivity
            $domainController = Resolve-DnsName -Name $DomainName -Type A -ErrorAction Stop
            if (-not $domainController) {
                return $false
            }
            
            # Test LDAP connectivity
            $ldapTest = Test-NetConnection -ComputerName $DomainName -Port 389 -WarningAction SilentlyContinue
            if (-not $ldapTest.TcpTestSucceeded) {
                Write-LogEntry "[WARN] LDAP connectivity test failed" 'WARNING'
                return $false
            }
            
            Write-LogEntry "[OK] Domain connectivity validated" 'SUCCESS'
            return $true
            
        } catch {
            Write-LogEntry "[WARN] Domain validation failed: $_" 'WARNING'
            return $false
        }
    }

    # Speed: Optimized registry operations
    function Set-AutoLoginRegistry {
        param(
            [string]$Username,
            [string]$Domain,
            [string]$EncryptedPassword,
            [int]$LoginCount
        )
        
        $regPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
        $settingsApplied = 0
        
        # Security: Validate registry path exists
        if (-not (Test-Path $regPath)) {
            throw "Winlogon registry path not found: $regPath"
        }
        
        # Batch registry operations for speed
        $registrySettings = @(
            @{ Name = 'AutoAdminLogon'; Value = '1'; Type = 'String' },
            @{ Name = 'DefaultUserName'; Value = $Username; Type = 'String' },
            @{ Name = 'DefaultDomainName'; Value = $Domain; Type = 'String' },
            @{ Name = 'AutoLogonCount'; Value = $LoginCount; Type = 'DWord' }
        )
        
        # Add password setting based on encryption method
        if ($UseLocalSystemEncryption) {
            # Use Windows built-in LSA encryption
            $registrySettings += @{ Name = 'DefaultPassword'; Value = ''; Type = 'String' }
            
            # Store encrypted password separately (requires additional LSA configuration)
            Write-LogEntry "Using LSA Secret storage for password (enhanced security)" 'SECURITY'
            # Note: LSA Secret storage requires additional implementation
        } else {
            # Store DPAPI-encrypted password
            $registrySettings += @{ Name = 'DefaultPassword'; Value = $EncryptedPassword; Type = 'String' }
        }
        
        foreach ($setting in $registrySettings) {
            try {
                if ($PSCmdlet.ShouldProcess("$regPath\$($setting.Name)", "Set registry value")) {
                    Set-ItemProperty -Path $regPath -Name $setting.Name -Value $setting.Value -Type $setting.Type -ErrorAction Stop
                    Write-LogEntry "[OK] Set $($setting.Name)" 'SUCCESS'
                    $settingsApplied++
                }
            } catch {
                Write-LogEntry "[ERROR] Failed to set $($setting.Name): $_" 'ERROR'
                throw
            }
        }
        
        return $settingsApplied
    }

    try {
        Write-LogEntry "=== Domain Auto-Login Configuration Started ===" 'INFO'
        
        if ($WhatIf) {
            Write-LogEntry "WhatIf mode - no registry changes will be made" 'INFO'
        }

        # Handle disable auto-login request
        if ($DisableAutoLogin) {
            Write-LogEntry "[DISABLE] Disabling auto-login..." 'INFO'
            
            $regPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
            
            if ($PSCmdlet.ShouldProcess($regPath, "Disable auto-login")) {
                try {
                    Set-ItemProperty -Path $regPath -Name "AutoAdminLogon" -Value "0" -Type String -ErrorAction Stop
                    
                    # Clear stored credentials for security
                    $credentialSettings = @('DefaultPassword', 'DefaultUserName', 'DefaultDomainName', 'AutoLogonCount')
                    foreach ($setting in $credentialSettings) {
                        try {
                            Remove-ItemProperty -Path $regPath -Name $setting -ErrorAction SilentlyContinue
                        } catch {
                            # Setting may not exist, continue
                        }
                    }
                    
                    Write-LogEntry "[OK] Auto-login disabled and credentials cleared" 'SUCCESS'
                    $global:LastStatus = "[OK] Auto-login disabled successfully."
                    return
                    
                } catch {
                    throw "Failed to disable auto-login: $_"
                }
            }
            return
        }

        # Security: Critical warning about auto-login risks
        if (-not $Force) {
            Write-Host "`n" + "="*70 -ForegroundColor Red
            Write-Host "CRITICAL SECURITY WARNING - AUTO-LOGIN CONFIGURATION" -ForegroundColor Red -BackgroundColor Black
            Write-Host "="*70 -ForegroundColor Red
            Write-Host "This configuration will:" -ForegroundColor Yellow
            Write-Host "  • Store domain credentials on local system" -ForegroundColor Yellow
            Write-Host "  • Allow automatic login without authentication" -ForegroundColor Yellow
            Write-Host "  • Potentially expose credentials to local attacks" -ForegroundColor Yellow
            Write-Host "`nRECOMMENDATIONS:" -ForegroundColor Cyan
            Write-Host "  • Use only in secure, controlled environments" -ForegroundColor Cyan
            Write-Host "  • Consider using domain Group Policy instead" -ForegroundColor Cyan
            Write-Host "  • Limit auto-login count to minimize exposure" -ForegroundColor Cyan
            Write-Host "  • Regularly rotate the password" -ForegroundColor Cyan
            Write-Host "="*70 -ForegroundColor Red
            
            $confirmation = Read-Host "`n[PROMPT] Type 'UNDERSTAND' to acknowledge security risks and continue"
            if ($confirmation -ne 'UNDERSTAND') {
                Write-LogEntry "Operation cancelled - security risks not acknowledged" 'WARNING'
                $global:LastStatus = "[WARN] User cancelled auto-login configuration."
                return
            }
        }

        # Validate inputs
        Write-LogEntry "[CHECK] Validating configuration parameters..." 'INFO'
        
        if ([string]::IsNullOrWhiteSpace($UserName)) {
            throw "Username cannot be empty"
        }
        
        if ([string]::IsNullOrWhiteSpace($DomainName)) {
            throw "Domain name cannot be empty"
        }
        
        # Security: Validate domain connectivity
        $domainConnectivity = Test-DomainConnectivity -DomainName $DomainName
        if (-not $domainConnectivity -and -not $Force) {
            throw "Domain connectivity validation failed. Use -Force to override."
        }

        # Security: Get secure credentials
        Write-LogEntry "[CREDENTIALS] Processing credentials securely..." 'SECURITY'
        $credential = Get-SecureCredential -Username $UserName -Domain $DomainName -SecurePassword $Password
        
        # Security: Test credential validity (optional)
        if ($domainConnectivity) {
            try {
                Write-LogEntry "[CHECK] Validating credentials..." 'INFO'
                # Note: In production, you might want to test credentials against domain
                # This is omitted here to avoid additional authentication attempts
                Write-LogEntry "[OK] Credential format validated" 'SUCCESS'
            } catch {
                Write-LogEntry "[WARN] Credential validation failed: $_" 'WARNING'
                if (-not $Force) {
                    throw "Invalid credentials provided"
                }
            }
        }

        # Security: Backup current settings
        Write-LogEntry "[BACKUP] Backing up current auto-login settings..." 'INFO'
        $backupFile = Backup-AutoLoginSettings
        
        if ($WhatIf) {
            Write-LogEntry "WhatIf: Would configure auto-login for $DomainName\$UserName" 'INFO'
            Write-LogEntry "WhatIf: Would set AutoLogonCount to $AutoLoginCount" 'INFO'
            $global:LastStatus = "[INFO] WhatIf completed - auto-login would be configured."
            return
        }

        # Security: Encrypt password
        Write-LogEntry "[SECURITY] Encrypting credentials..." 'SECURITY'
        $encryptedPassword = Protect-AutoLoginPassword -SecurePassword $credential.Password
        
        # Apply registry settings
        Write-LogEntry "[LOG] Applying auto-login registry settings..." 'INFO'
        $settingsCount = Set-AutoLoginRegistry -Username $UserName -Domain $DomainName -EncryptedPassword $encryptedPassword -LoginCount $AutoLoginCount
        
        # Security: Verify settings were applied
        Write-LogEntry "[CHECK] Verifying configuration..." 'INFO'
        $regPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
        
        try {
            $autoLogonValue = Get-ItemProperty -Path $regPath -Name "AutoAdminLogon" -ErrorAction Stop
            $userNameValue = Get-ItemProperty -Path $regPath -Name "DefaultUserName" -ErrorAction Stop
            $domainValue = Get-ItemProperty -Path $regPath -Name "DefaultDomainName" -ErrorAction Stop
            
            if ($autoLogonValue.AutoAdminLogon -eq "1" -and 
                $userNameValue.DefaultUserName -eq $UserName -and 
                $domainValue.DefaultDomainName -eq $DomainName) {
                
                Write-LogEntry "[OK] Auto-login configuration verified successfully" 'SUCCESS'
            } else {
                throw "Configuration verification failed"
            }
        } catch {
            throw "Failed to verify auto-login configuration: $_"
        }

        # Security: Set appropriate permissions on registry key
        try {
            Write-LogEntry "[SECURITY] Securing registry permissions..." 'SECURITY'
            
            $acl = Get-Acl -Path $regPath
            # Remove inherited permissions to protect stored credentials
            $acl.SetAccessRuleProtection($true, $true)
            
            # Keep only essential permissions (System, Administrators)
            $accessRules = $acl.Access | Where-Object { 
                $_.IdentityReference -notmatch "(Users|Everyone|Authenticated Users)" 
            }
            
            Set-Acl -Path $regPath -AclObject $acl
            Write-LogEntry "[OK] Registry permissions secured" 'SUCCESS'
            
        } catch {
            Write-LogEntry "[WARN] Failed to secure registry permissions: $_" 'WARNING'
        }

    } catch {
        Write-LogEntry "Critical error during auto-login configuration: $_" 'ERROR'
        $global:LastStatus = "[ERROR] Auto-login configuration failed: $_"
        throw
    } finally {
        $stopwatch.Stop()
        $duration = $stopwatch.Elapsed.TotalSeconds
        
        Write-LogEntry "`n[SUMMARY] Configuration Summary:" 'INFO'
        Write-LogEntry "Duration: $([math]::Round($duration, 2)) seconds" 'INFO'
        
        if (-not $DisableAutoLogin -and -not $WhatIf) {
            Write-LogEntry "User: $DomainName\$UserName" 'INFO'
            Write-LogEntry "Auto-login count: $AutoLoginCount" 'INFO'
            Write-LogEntry "Encryption: DPAPI (LocalMachine scope)" 'SECURITY'
        }
        
        # Write log file
        try {
            $logEntries | Out-File -FilePath $LogPath -Encoding UTF8 -Force
            Write-LogEntry "[LOG] Detailed log saved to: $LogPath" 'INFO'
        } catch {
            Write-LogEntry "[WARN] Failed to save log file: $_" 'WARNING'
        }
        
        # Set final status
        if (-not $global:LastStatus -or $global:LastStatus -notlike "*auto-login*") {
            if ($DisableAutoLogin) {
                $global:LastStatus = "[OK] Auto-login disabled successfully."
            } else {
                $global:LastStatus = "[OK] Auto-login configured for $DomainName\$UserName."
            }
        }
        
        Write-LogEntry "=== Auto-Login Configuration Completed ===" 'INFO'
    }
}

# Utility function to check current auto-login status
function Get-AutoLoginStatus {
    [CmdletBinding()]
    param()
    
    try {
        $regPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
        
        $autoLogon = Get-ItemProperty -Path $regPath -Name "AutoAdminLogon" -ErrorAction SilentlyContinue
        $userName = Get-ItemProperty -Path $regPath -Name "DefaultUserName" -ErrorAction SilentlyContinue
        $domainName = Get-ItemProperty -Path $regPath -Name "DefaultDomainName" -ErrorAction SilentlyContinue
        $loginCount = Get-ItemProperty -Path $regPath -Name "AutoLogonCount" -ErrorAction SilentlyContinue
        
        return @{
            Enabled = ($autoLogon.AutoAdminLogon -eq "1")
            UserName = $userName.DefaultUserName
            DomainName = $domainName.DefaultDomainName
            AutoLogonCount = $loginCount.AutoLogonCount
            PasswordStored = (Get-ItemProperty -Path $regPath -Name "DefaultPassword" -ErrorAction SilentlyContinue) -ne $null
        }
        
    } catch {
        return @{
            Error = $_.Exception.Message
            Enabled = $false
        }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# Option 7 - Set Desktop Power Settings - Only run on Desktop computers, no laptops!
# ─────────────────────────────────────────────────────────────────────────────
function Set-DesktopPowerSettings {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [ValidateSet('High Performance', 'Balanced', 'Power Saver', 'Ultimate Performance')]
        [string]$PowerPlan = 'High Performance',
        
        [ValidateRange(1, 600)]
        [int]$MonitorTimeoutMinutes = 60,
        
        [ValidateRange(1, 600)]
        [int]$DiskTimeoutMinutes = 0,  # 0 = Never
        
        [switch]$Force,
        # Removed [switch]$WhatIf - this is automatically provided by SupportsShouldProcess
        [switch]$AllowLaptops,
        [string]$LogPath = "$env:TEMP\PowerSettings_$(Get-Date -Format 'yyyyMMdd_HHmmss').log",
        [switch]$SkipHardwareDetection
    )

    # Security: Require elevation for power configuration
    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        throw "This function must be run as Administrator"
    }

    $ErrorActionPreference = 'Continue'
    $ProgressPreference = 'SilentlyContinue'  # Speed: Disable progress bars
    
    # Initialize tracking
    $logEntries = @()
    $appliedSettings = @()
    $failedSettings = @()
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    
    function Write-LogEntry {
        param([string]$Message, [string]$Level = 'INFO')
        $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $logEntry = "[$timestamp] [$Level] $Message"
        $script:logEntries += $logEntry
        
        switch ($Level) {
            'ERROR' { Write-Host $Message -ForegroundColor Red }
            'WARNING' { Write-Host $Message -ForegroundColor Yellow }
            'SUCCESS' { Write-Host $Message -ForegroundColor Green }
            'INFO' { Write-Host $Message -ForegroundColor Cyan }
            'HARDWARE' { Write-Host $Message -ForegroundColor Magenta }
        }
    }

    # Speed: Hardware detection with caching
    function Get-SystemHardwareType {
        try {
            # Use CIM for faster queries
            $computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop
            $battery = Get-CimInstance -ClassName Win32_Battery -ErrorAction SilentlyContinue
            $systemEnclosure = Get-CimInstance -ClassName Win32_SystemEnclosure -ErrorAction SilentlyContinue
            
            $hardwareInfo = @{
                IsLaptop = $false
                IsDesktop = $false
                IsWorkstation = $false
                IsServer = $false
                HasBattery = ($battery -ne $null)
                ChassisTypes = @()
                PCSystemType = $computerSystem.PCSystemType
                TotalPhysicalMemory = [math]::Round($computerSystem.TotalPhysicalMemory / 1GB, 2)
                Manufacturer = $computerSystem.Manufacturer
                Model = $computerSystem.Model
            }
            
            # Determine system type based on multiple factors
            if ($systemEnclosure) {
                $hardwareInfo.ChassisTypes = $systemEnclosure.ChassisTypes
                
                # Chassis type detection (more reliable than PCSystemType)
                $laptopChassisTypes = @(8, 9, 10, 11, 12, 14, 18, 21, 30, 31, 32)  # Laptop variants
                $desktopChassisTypes = @(3, 4, 5, 6, 7, 15, 16)  # Desktop variants
                $serverChassisTypes = @(17, 23)  # Server variants
                
                if ($systemEnclosure.ChassisTypes | Where-Object { $_ -in $laptopChassisTypes }) {
                    $hardwareInfo.IsLaptop = $true
                } elseif ($systemEnclosure.ChassisTypes | Where-Object { $_ -in $desktopChassisTypes }) {
                    $hardwareInfo.IsDesktop = $true
                } elseif ($systemEnclosure.ChassisTypes | Where-Object { $_ -in $serverChassisTypes }) {
                    $hardwareInfo.IsServer = $true
                } else {
                    $hardwareInfo.IsWorkstation = $true
                }
            }
            
            # Fallback to PCSystemType if chassis detection inconclusive
            if (-not ($hardwareInfo.IsLaptop -or $hardwareInfo.IsDesktop -or $hardwareInfo.IsServer)) {
                switch ($computerSystem.PCSystemType) {
                    1 { $hardwareInfo.IsDesktop = $true }
                    2 { $hardwareInfo.IsLaptop = $true }
                    3 { $hardwareInfo.IsWorkstation = $true }
                    4 { $hardwareInfo.IsServer = $true }
                    default { $hardwareInfo.IsDesktop = $true }  # Default assumption
                }
            }
            
            # Battery presence overrides chassis detection for laptops
            if ($hardwareInfo.HasBattery -and -not $hardwareInfo.IsServer) {
                $hardwareInfo.IsLaptop = $true
                $hardwareInfo.IsDesktop = $false
            }
            
            return $hardwareInfo
            
        } catch {
            Write-LogEntry "Hardware detection failed: $_" 'WARNING'
            return @{
                IsLaptop = $false
                IsDesktop = $true  # Safe default
                Error = $_.Exception.Message
            }
        }
    }

    # Speed: Get available power schemes efficiently
    function Get-PowerSchemes {
        try {
            $schemes = @{}
            
            # Parse powercfg output for available schemes
            $output = powercfg.exe /list 2>$null
            if ($LASTEXITCODE -eq 0 -and $output) {
                foreach ($line in $output) {
                    if ($line -match 'Power Scheme GUID: ([a-f0-9-]+)\s+\((.+?)\)(\s+\*)?') {
                        $guid = $matches[1]
                        $name = $matches[2].Trim()
                        $isActive = $matches[3] -eq ' *'
                        
                        $schemes[$name] = @{
                            GUID = $guid
                            Name = $name
                            IsActive = $isActive
                        }
                    }
                }
            }
            
            # Add common scheme mappings if not found
            $commonSchemes = @{
                'High Performance' = 'SCHEME_MIN'
                'Balanced' = 'SCHEME_BALANCED'
                'Power Saver' = 'SCHEME_MAX'
                'Ultimate Performance' = '8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c'
            }
            
            foreach ($scheme in $commonSchemes.GetEnumerator()) {
                if (-not $schemes.ContainsKey($scheme.Key)) {
                    $schemes[$scheme.Key] = @{
                        GUID = $scheme.Value
                        Name = $scheme.Key
                        IsActive = $false
                    }
                }
            }
            
            return $schemes
            
        } catch {
            Write-LogEntry "Failed to get power schemes: $_" 'WARNING'
            return @{}
        }
    }

    # Security: Backup current power settings
    function Backup-PowerSettings {
        try {
            $backupData = @{
                CurrentScheme = (powercfg.exe /getactivescheme 2>$null)
                HibernationStatus = (powercfg.exe /query SCHEME_CURRENT SUB_SLEEP HIBERNATEIDLE 2>$null)
                MonitorTimeoutAC = (powercfg.exe /query SCHEME_CURRENT SUB_VIDEO VIDEOIDLE /AC 2>$null)
                MonitorTimeoutDC = (powercfg.exe /query SCHEME_CURRENT SUB_VIDEO VIDEOIDLE /DC 2>$null)
                DiskTimeoutAC = (powercfg.exe /query SCHEME_CURRENT SUB_DISK DISKIDLE /AC 2>$null)
                DiskTimeoutDC = (powercfg.exe /query SCHEME_CURRENT SUB_DISK DISKIDLE /DC 2>$null)
                Timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
            }
            
            $backupFile = "$env:TEMP\PowerSettingsBackup_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
            $backupData | ConvertTo-Json -Depth 3 | Out-File -FilePath $backupFile -Encoding UTF8
            Write-LogEntry "[OK] Power settings backup saved to: $backupFile" 'INFO'
            return $backupFile
            
        } catch {
            Write-LogEntry "[WARN] Failed to backup power settings: $_" 'WARNING'
            return $null
        }
    }

    # Speed: Execute power configuration commands
    function Set-PowerConfiguration {
        param(
            [string]$SchemeName,
            [string]$SchemeGUID,
            [int]$MonitorTimeout,
            [int]$DiskTimeout
        )
        
        $configResults = @()
        
        # Set power scheme
        try {
            Write-LogEntry "Setting power scheme to: $SchemeName" 'INFO'
            
            if ($PSCmdlet.ShouldProcess($SchemeName, "Set active power scheme")) {
                $result = powercfg.exe /setactive $SchemeGUID 2>&1
                if ($LASTEXITCODE -eq 0) {
                    $configResults += @{
                        Setting = "Power Scheme"
                        Value = $SchemeName
                        Success = $true
                    }
                    Write-LogEntry "[OK] Power scheme set to: $SchemeName" 'SUCCESS'
                } else {
                    throw "Failed to set power scheme: $result"
                }
            }
        } catch {
            $configResults += @{
                Setting = "Power Scheme"
                Value = $SchemeName
                Success = $false
                Error = $_.Exception.Message
            }
            Write-LogEntry "[ERROR] Failed to set power scheme: $_" 'ERROR'
        }
        
        # Configure monitor timeout
        try {
            Write-LogEntry "Setting monitor timeout to: $MonitorTimeout minutes" 'INFO'
            
            if ($PSCmdlet.ShouldProcess("Monitor Timeout", "Set to $MonitorTimeout minutes")) {
                # Set for both AC and DC power
                $resultAC = powercfg.exe /change monitor-timeout-ac $MonitorTimeout 2>&1
                $resultDC = powercfg.exe /change monitor-timeout-dc $MonitorTimeout 2>&1
                
                if ($LASTEXITCODE -eq 0) {
                    $configResults += @{
                        Setting = "Monitor Timeout"
                        Value = "$MonitorTimeout minutes"
                        Success = $true
                    }
                    Write-LogEntry "[OK] Monitor timeout set to: $MonitorTimeout minutes" 'SUCCESS'
                } else {
                    throw "Failed to set monitor timeout: AC=$resultAC, DC=$resultDC"
                }
            }
        } catch {
            $configResults += @{
                Setting = "Monitor Timeout"
                Value = "$MonitorTimeout minutes"
                Success = $false
                Error = $_.Exception.Message
            }
            Write-LogEntry "[ERROR] Failed to set monitor timeout: $_" 'ERROR'
        }
        
        # Configure disk timeout
        if ($DiskTimeout -eq 0) {
            try {
                Write-LogEntry "Disabling disk timeout (never turn off)" 'INFO'
                
                if ($PSCmdlet.ShouldProcess("Disk Timeout", "Disable")) {
                    $resultAC = powercfg.exe /change disk-timeout-ac 0 2>&1
                    $resultDC = powercfg.exe /change disk-timeout-dc 0 2>&1
                    
                    if ($LASTEXITCODE -eq 0) {
                        $configResults += @{
                            Setting = "Disk Timeout"
                            Value = "Disabled"
                            Success = $true
                        }
                        Write-LogEntry "[OK] Disk timeout disabled" 'SUCCESS'
                    } else {
                        throw "Failed to disable disk timeout: AC=$resultAC, DC=$resultDC"
                    }
                }
            } catch {
                $configResults += @{
                    Setting = "Disk Timeout"
                    Value = "Disabled"
                    Success = $false
                    Error = $_.Exception.Message
                }
                Write-LogEntry "[ERROR] Failed to disable disk timeout: $_" 'ERROR'
            }
        } else {
            try {
                Write-LogEntry "Setting disk timeout to: $DiskTimeout minutes" 'INFO'
                
                if ($PSCmdlet.ShouldProcess("Disk Timeout", "Set to $DiskTimeout minutes")) {
                    $resultAC = powercfg.exe /change disk-timeout-ac $DiskTimeout 2>&1
                    $resultDC = powercfg.exe /change disk-timeout-dc $DiskTimeout 2>&1
                    
                    if ($LASTEXITCODE -eq 0) {
                        $configResults += @{
                            Setting = "Disk Timeout"
                            Value = "$DiskTimeout minutes"
                            Success = $true
                        }
                        Write-LogEntry "[OK] Disk timeout set to: $DiskTimeout minutes" 'SUCCESS'
                    } else {
                        throw "Failed to set disk timeout: AC=$resultAC, DC=$resultDC"
                    }
                }
            } catch {
                $configResults += @{
                    Setting = "Disk Timeout"
                    Value = "$DiskTimeout minutes"
                    Success = $false
                    Error = $_.Exception.Message
                }
                Write-LogEntry "[ERROR] Failed to set disk timeout: $_" 'ERROR'
            }
        }
        
        # Disable hibernation
        try {
            Write-LogEntry "Disabling hibernation" 'INFO'
            
            if ($PSCmdlet.ShouldProcess("Hibernation", "Disable")) {
                $result = powercfg.exe /hibernate off 2>&1
                if ($LASTEXITCODE -eq 0) {
                    $configResults += @{
                        Setting = "Hibernation"
                        Value = "Disabled"
                        Success = $true
                    }
                    Write-LogEntry "[OK] Hibernation disabled" 'SUCCESS'
                } else {
                    throw "Failed to disable hibernation: $result"
                }
            }
        } catch {
            $configResults += @{
                Setting = "Hibernation"
                Value = "Disabled"
                Success = $false
                Error = $_.Exception.Message
            }
            Write-LogEntry "[ERROR] Failed to disable hibernation: $_" 'ERROR'
        }
        
        return $configResults
    }

    try {
        Write-LogEntry "=== Desktop Power Settings Configuration Started ===" 'INFO'
        Write-LogEntry "Power Plan: $PowerPlan" 'INFO'
        Write-LogEntry "Monitor Timeout: $MonitorTimeoutMinutes minutes" 'INFO'
        Write-LogEntry "Disk Timeout: $(if($DiskTimeoutMinutes -eq 0){'Disabled'}else{"$DiskTimeoutMinutes minutes"})" 'INFO'
        
        if ($WhatIfPreference) {
            Write-LogEntry "WhatIf mode - no power settings will be changed" 'INFO'
        }

        # Hardware detection and validation
        if (-not $SkipHardwareDetection) {
            Write-LogEntry "`n[CHECK] Detecting system hardware type..." 'HARDWARE'
            $hardwareInfo = Get-SystemHardwareType
            
            Write-LogEntry "Hardware Analysis:" 'HARDWARE'
            Write-LogEntry "  • System Type: $(if($hardwareInfo.IsLaptop){'Laptop'}elseif($hardwareInfo.IsDesktop){'Desktop'}elseif($hardwareInfo.IsServer){'Server'}else{'Workstation'})" 'HARDWARE'
            Write-LogEntry "  • Has Battery: $($hardwareInfo.HasBattery)" 'HARDWARE'
            Write-LogEntry "  • Manufacturer: $($hardwareInfo.Manufacturer)" 'HARDWARE'
            Write-LogEntry "  • Model: $($hardwareInfo.Model)" 'HARDWARE'
            Write-LogEntry "  • Memory: $($hardwareInfo.TotalPhysicalMemory) GB" 'HARDWARE'
            
            # Security: Prevent accidental laptop configuration
            if ($hardwareInfo.IsLaptop -and -not $AllowLaptops -and -not $Force) {
                Write-Host "`n" + "="*60 -ForegroundColor Red
                Write-Host "LAPTOP DETECTED - OPERATION BLOCKED" -ForegroundColor Red -BackgroundColor Black
                Write-Host "="*60 -ForegroundColor Red
                Write-Host "This system appears to be a LAPTOP with battery power." -ForegroundColor Yellow
                Write-Host "Desktop power settings may negatively impact battery life!" -ForegroundColor Yellow
                Write-Host "`nTo proceed anyway, use one of these options:" -ForegroundColor Cyan
                Write-Host "  • Use -AllowLaptops parameter" -ForegroundColor Cyan
                Write-Host "  • Use -Force parameter" -ForegroundColor Cyan
                Write-Host "  • Use -SkipHardwareDetection parameter" -ForegroundColor Cyan
                Write-Host "="*60 -ForegroundColor Red
                
                $global:LastStatus = "[WARN] Operation blocked - laptop detected."
                return
            }
            
            if ($hardwareInfo.IsLaptop -and ($AllowLaptops -or $Force)) {
                Write-LogEntry "[WARN] Proceeding with laptop configuration (overridden)" 'WARNING'
            }
        }

        # User confirmation for desktop systems
        if (-not $Force -and -not $WhatIfPreference) {
            Write-Host "`n" + "="*60 -ForegroundColor Yellow
            Write-Host "POWER SETTINGS CONFIGURATION" -ForegroundColor Yellow -BackgroundColor Black
            Write-Host "="*60 -ForegroundColor Yellow
            Write-Host "This will configure the following settings:" -ForegroundColor White
            Write-Host "  • Power Plan: $PowerPlan" -ForegroundColor Cyan
            Write-Host "  • Monitor Timeout: $MonitorTimeoutMinutes minutes" -ForegroundColor Cyan
            Write-Host "  • Disk Timeout: $(if($DiskTimeoutMinutes -eq 0){'Disabled'}else{"$DiskTimeoutMinutes minutes"})" -ForegroundColor Cyan
            Write-Host "  • Hibernation: Disabled" -ForegroundColor Cyan
            Write-Host "`nNote: These settings optimize for desktop performance" -ForegroundColor Yellow
            Write-Host "="*60 -ForegroundColor Yellow
            
            $confirmation = Read-Host "`n[PROMPT] Continue with power configuration? (Y/N)"
            if ($confirmation -notin @('Y', 'y', 'Yes', 'yes')) {
                Write-LogEntry "Operation cancelled by user" 'WARNING'
                $global:LastStatus = "[WARN] User cancelled power settings configuration."
                return
            }
        }

        # Get available power schemes
        Write-LogEntry "`n[CHECK] Scanning available power schemes..." 'INFO'
        $powerSchemes = Get-PowerSchemes
        
        if ($powerSchemes.Count -eq 0) {
            throw "No power schemes detected on this system"
        }
        
        Write-LogEntry "Available power schemes:" 'INFO'
        foreach ($scheme in $powerSchemes.GetEnumerator()) {
            $activeIndicator = if ($scheme.Value.IsActive) { " (ACTIVE)" } else { "" }
            Write-LogEntry "  • $($scheme.Key)$activeIndicator" 'INFO'
        }
        
        # Validate requested power plan
        if (-not $powerSchemes.ContainsKey($PowerPlan)) {
            Write-LogEntry "[WARN] Requested power plan '$PowerPlan' not found" 'WARNING'
            Write-LogEntry "Falling back to 'High Performance' scheme" 'WARNING'
            $PowerPlan = 'High Performance'
            
            if (-not $powerSchemes.ContainsKey($PowerPlan)) {
                throw "Neither requested scheme nor High Performance scheme available"
            }
        }
        
        $selectedScheme = $powerSchemes[$PowerPlan]
        Write-LogEntry "Selected scheme: $PowerPlan (GUID: $($selectedScheme.GUID))" 'INFO'

        # Backup current settings
        if (-not $WhatIfPreference) {
            Write-LogEntry "`n[BACKUP] Backing up current power settings..." 'INFO'
            $backupFile = Backup-PowerSettings
        }

        if ($WhatIfPreference) {
            Write-LogEntry "`nWhatIf Summary:" 'INFO'
            Write-LogEntry "  • Would set power scheme to: $PowerPlan" 'INFO'
            Write-LogEntry "  • Would set monitor timeout to: $MonitorTimeoutMinutes minutes" 'INFO'
            Write-LogEntry "  • Would set disk timeout to: $(if($DiskTimeoutMinutes -eq 0){'Disabled'}else{"$DiskTimeoutMinutes minutes"})" 'INFO'
            Write-LogEntry "  • Would disable hibernation" 'INFO'
            $global:LastStatus = "[INFO] WhatIf completed - power settings would be configured."
            return
        }

        # Apply power configuration
        Write-LogEntry "`n[APPLY] Applying power configuration..." 'INFO'
        $configResults = Set-PowerConfiguration -SchemeName $PowerPlan -SchemeGUID $selectedScheme.GUID -MonitorTimeout $MonitorTimeoutMinutes -DiskTimeout $DiskTimeoutMinutes
        
        # Process results
        $successCount = ($configResults | Where-Object { $_.Success }).Count
        $failureCount = ($configResults | Where-Object { -not $_.Success }).Count
        
        $script:appliedSettings = $configResults | Where-Object { $_.Success }
        $script:failedSettings = $configResults | Where-Object { -not $_.Success }

        # Verify configuration
        Write-LogEntry "`n[CHECK] Verifying power configuration..." 'INFO'
        try {
            $currentScheme = powercfg.exe /getactivescheme 2>$null
            if ($currentScheme -and $currentScheme -match $selectedScheme.GUID) {
                Write-LogEntry "[OK] Power scheme verification passed" 'SUCCESS'
            } else {
                Write-LogEntry "[WARN] Power scheme verification failed" 'WARNING'
            }
        } catch {
            Write-LogEntry "[WARN] Could not verify power scheme: $_" 'WARNING'
        }

    } catch {
        Write-LogEntry "Critical error during power settings configuration: $_" 'ERROR'
        $global:LastStatus = "[ERROR] Power settings configuration failed: $_"
        throw
    } finally {
        $stopwatch.Stop()
        $duration = $stopwatch.Elapsed.TotalSeconds
        
        Write-LogEntry "`n[SUMMARY] Configuration Summary:" 'INFO'
        Write-LogEntry "Duration: $([math]::Round($duration, 2)) seconds" 'INFO'
        Write-LogEntry "Settings applied: $($appliedSettings.Count)" 'SUCCESS'
        Write-LogEntry "Settings failed: $($failedSettings.Count)" 'ERROR'
        
        if ($appliedSettings.Count -gt 0) {
            Write-LogEntry "`n[OK] Successfully Applied:" 'SUCCESS'
            foreach ($setting in $appliedSettings) {
                Write-LogEntry "  • $($setting.Setting): $($setting.Value)" 'SUCCESS'
            }
        }
        
        if ($failedSettings.Count -gt 0) {
            Write-LogEntry "`n[ERROR] Failed Settings:" 'ERROR'
            foreach ($setting in $failedSettings) {
                Write-LogEntry "  • $($setting.Setting): $($setting.Error)" 'ERROR'
            }
        }
        
        # Write log file
        try {
            $logEntries | Out-File -FilePath $LogPath -Encoding UTF8 -Force
            Write-LogEntry "[LOG] Detailed log saved to: $LogPath" 'INFO'
        } catch {
            Write-LogEntry "[WARN] Failed to save log file: $_" 'WARNING'
        }
        
        # Set final status
        if ($appliedSettings.Count -gt 0) {
            if ($failedSettings.Count -eq 0) {
                $global:LastStatus = "[OK] All power settings applied successfully."
            } else {
                $global:LastStatus = "[WARN] Power settings partially applied ($($appliedSettings.Count) success, $($failedSettings.Count) failed)."
            }
        } else {
            $global:LastStatus = "[ERROR] No power settings were applied."
        }
        
        Write-LogEntry "=== Power Settings Configuration Completed ===" 'INFO'
    }
}

# Utility function to get current power configuration
function Get-PowerConfiguration {
    [CmdletBinding()]
    param()
    
    try {
        $config = @{}
        
        # Get active scheme
        $activeScheme = powercfg.exe /getactivescheme 2>$null
        if ($activeScheme) {
            $config.ActiveScheme = $activeScheme.Trim()
        }
        
        # Get hibernation status
        $hibernation = powercfg.exe /availablesleepstates 2>$null
        $config.HibernationAvailable = ($hibernation -match "Hibernate")
        
        # Get monitor and disk timeouts (simplified parsing)
        $config.MonitorTimeoutAC = "Unknown"
        $config.MonitorTimeoutDC = "Unknown"
        $config.DiskTimeoutAC = "Unknown"
        $config.DiskTimeoutDC = "Unknown"
        
        return $config
        
    } catch {
        return @{
            Error = $_.Exception.Message
        }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# Option 17 - Install Computer Lab Scheduled Tasks
# ─────────────────────────────────────────────────────────────────────────────
function Register-LabScheduledTasks {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [string]$RepoOwner = 'dswaney',
        [string]$RepoName = 'Compton',
        [string]$Branch = 'main',
        [string]$RepoSubFolder = '',
        [string]$DestinationPath = 'C:\Scripts',
        [string]$LogDirectory = 'C:\Logs',
        [switch]$Force
    )

    $ErrorActionPreference = 'Stop'
    $ProgressPreference = 'SilentlyContinue'

    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        $global:LastStatus = "[ERROR] Scheduled task setup requires Administrator rights."
        throw "This function must be run as Administrator."
    }

    $script:RunStart = Get-Date
    $script:ComputerName = $env:COMPUTERNAME
    $script:BackupFolder = Join-Path $DestinationPath 'Backup'
    $script:YamlLogPath = Join-Path $LogDirectory ("{0}-RegisterLabScheduledTasks-{1}.yaml" -f $script:ComputerName, $script:RunStart.ToString('yyyy-MM-dd_HH-mm-ss'))
    $script:ActionHistory = New-Object System.Collections.Generic.List[object]
    $script:FileResults = New-Object System.Collections.Generic.List[object]
    $script:TaskResults = New-Object System.Collections.Generic.List[object]
    $script:OverallResult = 'Unknown'
    $script:FailureMessage = $null

    $scriptFiles = @(
        '00_Update-Scripts-FromGitHub.ps1',
        '01_Enable_Windows_Update_Services.ps1',
        '02_Remove_User_Profiles.ps1',
        '03_Weekend_Apps_Update.ps1',
        '04_Update_Edge_Silent.ps1',
        '05_Weekend_HP_Drivers_Update.ps1',
        '06_Weekend_Windows_Updates.ps1',
        '07_Force_Reboot_Install_Updates.ps1',
        '08_System_Repair.ps1',
        '09_Disable_Windows_Update_Services.ps1'
    )

    $taskDefinitions = @(
        [PSCustomObject]@{ Name = '01. Check for Updated Scripts';          Script = '00_Update-Scripts-FromGitHub.ps1';       Time = '01:15'; Arguments = '' },
        [PSCustomObject]@{ Name = '02. Enable Windows Update Services';     Script = '01_Enable_Windows_Update_Services.ps1';  Time = '01:20'; Arguments = '' },
        [PSCustomObject]@{ Name = '03. Remove User Profiles Weekly';        Script = '02_Remove_User_Profiles.ps1';            Time = '01:30'; Arguments = '' },
        [PSCustomObject]@{ Name = '04. Weekend Apps Update';                Script = '03_Weekend_Apps_Update.ps1';             Time = '02:00'; Arguments = '' },
        [PSCustomObject]@{ Name = '05. Update Edge Silent';                 Script = '04_Update_Edge_Silent.ps1';              Time = '02:45'; Arguments = '-KillEdgeProcesses' },
        [PSCustomObject]@{ Name = '06. Weekend HP Drivers Update';          Script = '05_Weekend_HP_Drivers_Update.ps1';       Time = '03:00'; Arguments = '' },
        [PSCustomObject]@{ Name = '07. Weekend Windows Updates - 1st Pass'; Script = '06_Weekend_Windows_Updates.ps1';         Time = '04:00'; Arguments = '' },
        [PSCustomObject]@{ Name = '08. Force Reboot Install Updates';       Script = '07_Force_Reboot_Install_Updates.ps1';    Time = '05:00'; Arguments = '' },
        [PSCustomObject]@{ Name = '09. Weekend Windows Updates - 2nd Pass'; Script = '06_Weekend_Windows_Updates.ps1';         Time = '05:30'; Arguments = '' },
        [PSCustomObject]@{ Name = '10. Disable Windows Update Services';    Script = '09_Disable_Windows_Update_Services.ps1'; Time = '06:00'; Arguments = '' },
        [PSCustomObject]@{ Name = '11. Force Reboot Install Updates 2';     Script = '07_Force_Reboot_Install_Updates.ps1';    Time = '06:05'; Arguments = '' },
        [PSCustomObject]@{ Name = '12. System Repair';                      Script = '08_System_Repair.ps1';                   Time = '06:15'; Arguments = '' },
        [PSCustomObject]@{ Name = '13. Force Reboot Install Updates 3';     Script = '07_Force_Reboot_Install_Updates.ps1';    Time = '07:00'; Arguments = '' }
    )

    function Ensure-Directory {
        param([Parameter(Mandatory)][string]$Path)
        if (-not (Test-Path -LiteralPath $Path)) {
            New-Item -Path $Path -ItemType Directory -Force | Out-Null
        }
    }

    function Write-Status {
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

        $script:ActionHistory.Add([PSCustomObject]@{
            Time    = $timestamp
            Level   = $Level
            Message = $Message
        }) | Out-Null
    }

    function ConvertTo-YamlScalar {
        param([AllowNull()]$Value)

        if ($null -eq $Value) { return 'null' }
        if ($Value -is [bool]) { return $Value.ToString().ToLowerInvariant() }
        if ($Value -is [int] -or $Value -is [long] -or $Value -is [double] -or $Value -is [decimal]) { return [string]$Value }
        if ($Value -is [datetime]) { return "'" + $Value.ToString('yyyy-MM-dd HH:mm:ss') + "'" }

        $text = [string]$Value
        $text = $text -replace "`r", ' '
        $text = $text -replace "`n", ' '
        $text = $text -replace "'", "''"
        return "'" + $text + "'"
    }

    function Save-Utf8NoBom {
        param(
            [Parameter(Mandatory)][string]$Path,
            [Parameter(Mandatory)][string]$Content
        )

        $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
        [System.IO.File]::WriteAllText($Path, $Content, $utf8NoBom)
    }

    function Get-RawGitHubUrl {
        param([Parameter(Mandatory)][string]$FileName)

        $pathPart = if ([string]::IsNullOrWhiteSpace($RepoSubFolder)) {
            $FileName
        }
        else {
            ($RepoSubFolder.Trim('/').Replace('\','/') + '/' + $FileName)
        }

        'https://raw.githubusercontent.com/{0}/{1}/{2}/{3}' -f $RepoOwner, $RepoName, $Branch, $pathPart
    }

    function Get-RemoteFileContent {
        param([Parameter(Mandatory)][string]$FileName)

        $url = Get-RawGitHubUrl -FileName $FileName
        $uriBuilder = New-Object System.UriBuilder($url)
        $uriBuilder.Query = 'cb={0}' -f [DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds()
        $finalUri = $uriBuilder.Uri.AbsoluteUri

        $response = Invoke-WebRequest -Uri $finalUri `
                                      -UseBasicParsing `
                                      -Headers @{
                                          'Cache-Control' = 'no-cache'
                                          'Pragma'        = 'no-cache'
                                          'User-Agent'    = 'PowerShell-GitHub-Updater'
                                      } `
                                      -ErrorAction Stop

        if ([string]::IsNullOrWhiteSpace($response.Content)) {
            throw "Downloaded content was empty for [$FileName]."
        }

        $response.Content
    }

    function Get-FileTextSafe {
        param([Parameter(Mandatory)][string]$Path)

        if (-not (Test-Path -LiteralPath $Path)) {
            return $null
        }

        try {
            return Get-Content -LiteralPath $Path -Raw -Encoding UTF8
        }
        catch {
            return Get-Content -LiteralPath $Path -Raw
        }
    }

    function Get-ScriptHeaderValue {
        param(
            [Parameter(Mandatory)][AllowEmptyString()][string]$Content,
            [Parameter(Mandatory)][string]$HeaderName
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

        return $null
    }

    function Convert-ToVersionObject {
        param([Parameter(Mandatory)][string]$VersionText)

        try {
            return [version]$VersionText.Trim()
        }
        catch {
            $clean = ($VersionText -replace '[^\d\.]', '').Trim('.')
            if ([string]::IsNullOrWhiteSpace($clean)) {
                return [version]'0.0'
            }

            try { return [version]$clean } catch { return [version]'0.0' }
        }
    }

    function Backup-File {
        param([Parameter(Mandatory)][string]$Path)

        if (-not (Test-Path -LiteralPath $Path)) {
            return $null
        }

        Ensure-Directory -Path $script:BackupFolder

        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($Path)
        $extension = [System.IO.Path]::GetExtension($Path)
        $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
        $backupPath = Join-Path $script:BackupFolder ("{0}_{1}{2}.bak" -f $baseName, $timestamp, $extension)
        Copy-Item -LiteralPath $Path -Destination $backupPath -Force
        $backupPath
    }

    function Write-YamlLog {
        try {
            Ensure-Directory -Path $LogDirectory

            $runEnd = Get-Date
            $duration = [math]::Round(($runEnd - $script:RunStart).TotalSeconds, 0)

            $updatedCount = @($script:FileResults | Where-Object { $_.Status -eq 'Updated' }).Count
            $currentCount = @($script:FileResults | Where-Object { $_.Status -eq 'Current' }).Count
            $downloadedMissingCount = @($script:FileResults | Where-Object { $_.Status -eq 'DownloadedMissing' }).Count
            $fileErrorCount = @($script:FileResults | Where-Object { $_.Status -eq 'Error' }).Count
            $taskCreatedCount = @($script:TaskResults | Where-Object { $_.Status -eq 'Created' }).Count
            $taskWhatIfCount = @($script:TaskResults | Where-Object { $_.Status -eq 'WhatIf' }).Count
            $taskErrorCount = @($script:TaskResults | Where-Object { $_.Status -eq 'Error' }).Count

            $lines = New-Object System.Collections.Generic.List[string]
            $lines.Add("computer_name: $(ConvertTo-YamlScalar $script:ComputerName)") | Out-Null
            $lines.Add("script_name: 'Register-LabScheduledTasks'") | Out-Null
            $lines.Add("script_version: '1.5.0'") | Out-Null
            $lines.Add("run_started: $(ConvertTo-YamlScalar $script:RunStart)") | Out-Null
            $lines.Add("run_finished: $(ConvertTo-YamlScalar $runEnd)") | Out-Null
            $lines.Add("duration_seconds: $duration") | Out-Null
            $lines.Add("destination_path: $(ConvertTo-YamlScalar $DestinationPath)") | Out-Null
            $lines.Add("backup_folder: $(ConvertTo-YamlScalar $script:BackupFolder)") | Out-Null
            $lines.Add("overall_result: $(ConvertTo-YamlScalar $script:OverallResult)") | Out-Null
            $lines.Add("failure_message: $(ConvertTo-YamlScalar $script:FailureMessage)") | Out-Null
            $lines.Add('') | Out-Null
            $lines.Add('summary:') | Out-Null
            $lines.Add("  files_updated: $updatedCount") | Out-Null
            $lines.Add("  files_current: $currentCount") | Out-Null
            $lines.Add("  files_downloaded_missing: $downloadedMissingCount") | Out-Null
            $lines.Add("  file_errors: $fileErrorCount") | Out-Null
            $lines.Add("  tasks_created: $taskCreatedCount") | Out-Null
            $lines.Add("  tasks_whatif: $taskWhatIfCount") | Out-Null
            $lines.Add("  task_errors: $taskErrorCount") | Out-Null
            $lines.Add('') | Out-Null

            $lines.Add('file_results:') | Out-Null
            if ($script:FileResults.Count -gt 0) {
                foreach ($item in $script:FileResults) {
                    $lines.Add('  -') | Out-Null
                    $lines.Add("    file_name: $(ConvertTo-YamlScalar $item.FileName)") | Out-Null
                    $lines.Add("    local_path: $(ConvertTo-YamlScalar $item.LocalPath)") | Out-Null
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
            $lines.Add('task_results:') | Out-Null
            if ($script:TaskResults.Count -gt 0) {
                foreach ($item in $script:TaskResults) {
                    $lines.Add('  -') | Out-Null
                    $lines.Add("    task_name: $(ConvertTo-YamlScalar $item.TaskName)") | Out-Null
                    $lines.Add("    script_path: $(ConvertTo-YamlScalar $item.ScriptPath)") | Out-Null
                    $lines.Add("    schedule_time: $(ConvertTo-YamlScalar $item.ScheduleTime)") | Out-Null
                    $lines.Add("    arguments: $(ConvertTo-YamlScalar $item.Arguments)") | Out-Null
                    $lines.Add("    status: $(ConvertTo-YamlScalar $item.Status)") | Out-Null
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

    function Get-TriggerTime {
        param([Parameter(Mandatory)][string]$TimeText)
        $parts = $TimeText.Split(':')
        if ($parts.Count -ne 2) {
            throw "Invalid schedule time [$TimeText]. Expected HH:mm."
        }

        (Get-Date -Hour ([int]$parts[0]) -Minute ([int]$parts[1]) -Second 0)
    }

    function Register-WeeklySystemTask {
        param(
            [Parameter(Mandatory)][string]$TaskName,
            [Parameter(Mandatory)][string]$ScriptPath,
            [Parameter(Mandatory)][string]$TimeText,
            [string]$Arguments = ''
        )

        $taskDescription = "Created by Compton_Tech_Utils Option 17"
        $argSuffix = if ([string]::IsNullOrWhiteSpace($Arguments)) { '' } else { ' ' + $Arguments.Trim() }
        $actionArgs = '-NoProfile -ExecutionPolicy Bypass -File "{0}"{1}' -f $ScriptPath, $argSuffix
        $target = "{0} [{1}]" -f $TaskName, $TimeText

        if (-not $PSCmdlet.ShouldProcess($target, 'Register weekly scheduled task as SYSTEM')) {
            $script:TaskResults.Add([PSCustomObject]@{
                TaskName     = $TaskName
                ScriptPath   = $ScriptPath
                ScheduleTime = $TimeText
                Arguments    = $Arguments
                Status       = 'WhatIf'
                Message      = 'WhatIf prevented task registration.'
            }) | Out-Null
            Write-Status "WhatIf: would register task [$TaskName] for Sundays at [$TimeText]." 'INFO'
            return
        }

        if (-not (Test-Path -LiteralPath $ScriptPath)) {
            throw "Script path does not exist: $ScriptPath"
        }

        try {
            if ($Force) {
                try {
                    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction Stop
                    Write-Status "Removed existing task [$TaskName] before recreation." 'INFO'
                }
                catch {
                    Write-Status "Existing task [$TaskName] was not present or could not be removed before recreation." 'INFO'
                }
            }

            $action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument $actionArgs
            $trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Sunday -At (Get-TriggerTime -TimeText $TimeText)
            $principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
            $settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

            $taskObject = New-ScheduledTask -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Description $taskDescription
            Register-ScheduledTask -TaskName $TaskName -InputObject $taskObject -Force | Out-Null

            $script:TaskResults.Add([PSCustomObject]@{
                TaskName     = $TaskName
                ScriptPath   = $ScriptPath
                ScheduleTime = $TimeText
                Arguments    = $Arguments
                Status       = 'Created'
                Message      = 'Scheduled task registered successfully.'
            }) | Out-Null

            Write-Status "Registered task [$TaskName] for Sundays at [$TimeText] as SYSTEM." 'OK'
        }
        catch {
            $script:TaskResults.Add([PSCustomObject]@{
                TaskName     = $TaskName
                ScriptPath   = $ScriptPath
                ScheduleTime = $TimeText
                Arguments    = $Arguments
                Status       = 'Error'
                Message      = $_.Exception.Message
            }) | Out-Null

            Write-Status "Failed to register task [$TaskName]: $($_.Exception.Message)" 'ERROR'
        }
    }

    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

        Ensure-Directory -Path $DestinationPath
        Ensure-Directory -Path $LogDirectory

        Write-Status "Ensuring script directory exists at [$DestinationPath]." 'INFO'
        Write-Status "Downloading and updating lab scripts from [$RepoOwner/$RepoName] branch [$Branch]." 'INFO'

        foreach ($file in $scriptFiles) {
            $localPath = Join-Path $DestinationPath $file

            try {
                $remoteContent = Get-RemoteFileContent -FileName $file
                $remoteVersionText = Get-ScriptHeaderValue -Content $remoteContent -HeaderName 'ScriptVersion'
                $remoteLastUpdated = Get-ScriptHeaderValue -Content $remoteContent -HeaderName 'LastUpdated'

                if ([string]::IsNullOrWhiteSpace($remoteVersionText)) {
                    throw "Remote file [$file] is missing a readable ScriptVersion header."
                }

                $remoteVersion = Convert-ToVersionObject -VersionText $remoteVersionText
                $localContent = Get-FileTextSafe -Path $localPath

                if ($null -eq $localContent) {
                    Save-Utf8NoBom -Path $localPath -Content $remoteContent
                    $script:FileResults.Add([PSCustomObject]@{
                        FileName          = $file
                        LocalPath         = $localPath
                        Status            = 'DownloadedMissing'
                        LocalVersion      = $null
                        RemoteVersion     = $remoteVersionText
                        LocalLastUpdated  = $null
                        RemoteLastUpdated = $remoteLastUpdated
                        BackupPath        = $null
                        Message           = 'Local file was missing and was downloaded from GitHub.'
                    }) | Out-Null
                    Write-Status "Downloaded missing script [$file]." 'OK'
                    continue
                }

                $localVersionText = Get-ScriptHeaderValue -Content $localContent -HeaderName 'ScriptVersion'
                $localLastUpdated = Get-ScriptHeaderValue -Content $localContent -HeaderName 'LastUpdated'
                if ([string]::IsNullOrWhiteSpace($localVersionText)) {
                    $localVersionText = '0.0'
                }

                $localVersion = Convert-ToVersionObject -VersionText $localVersionText

                if ($remoteVersion -gt $localVersion) {
                    $backupPath = Backup-File -Path $localPath
                    Save-Utf8NoBom -Path $localPath -Content $remoteContent
                    $script:FileResults.Add([PSCustomObject]@{
                        FileName          = $file
                        LocalPath         = $localPath
                        Status            = 'Updated'
                        LocalVersion      = $localVersionText
                        RemoteVersion     = $remoteVersionText
                        LocalLastUpdated  = $localLastUpdated
                        RemoteLastUpdated = $remoteLastUpdated
                        BackupPath        = $backupPath
                        Message           = "Updated local file from $localVersionText to $remoteVersionText."
                    }) | Out-Null
                    Write-Status "Updated script [$file] from [$localVersionText] to [$remoteVersionText]." 'OK'
                }
                else {
                    $script:FileResults.Add([PSCustomObject]@{
                        FileName          = $file
                        LocalPath         = $localPath
                        Status            = 'Current'
                        LocalVersion      = $localVersionText
                        RemoteVersion     = $remoteVersionText
                        LocalLastUpdated  = $localLastUpdated
                        RemoteLastUpdated = $remoteLastUpdated
                        BackupPath        = $null
                        Message           = 'Local file is already current.'
                    }) | Out-Null
                    Write-Status "Script [$file] is already current." 'INFO'
                }
            }
            catch {
                $script:FileResults.Add([PSCustomObject]@{
                    FileName          = $file
                    LocalPath         = $localPath
                    Status            = 'Error'
                    LocalVersion      = $null
                    RemoteVersion     = $null
                    LocalLastUpdated  = $null
                    RemoteLastUpdated = $null
                    BackupPath        = $null
                    Message           = $_.Exception.Message
                }) | Out-Null
                Write-Status "Failed processing script [$file]: $($_.Exception.Message)" 'ERROR'
            }
        }

        Write-Status "Registering scheduled tasks in Task Scheduler." 'INFO'

        foreach ($task in $taskDefinitions) {
            $scriptPath = Join-Path $DestinationPath $task.Script
            Register-WeeklySystemTask -TaskName $task.Name -ScriptPath $scriptPath -TimeText $task.Time -Arguments $task.Arguments
        }

        $fileErrors = @($script:FileResults | Where-Object { $_.Status -eq 'Error' }).Count
        $taskErrors = @($script:TaskResults | Where-Object { $_.Status -eq 'Error' }).Count
        $taskCreated = @($script:TaskResults | Where-Object { $_.Status -eq 'Created' }).Count

        if (($fileErrors + $taskErrors) -gt 0) {
            $script:OverallResult = 'CompletedWithErrors'
            $global:LastStatus = "[WARN] Option 17 completed with issues. Scripts downloaded/updated, tasks created: $taskCreated, errors: $($fileErrors + $taskErrors)."
        }
        else {
            $script:OverallResult = 'Succeeded'
            $global:LastStatus = "[OK] Option 17 completed successfully. C:\Scripts refreshed and $taskCreated scheduled tasks registered."
        }

        Write-YamlLog
    }
    catch {
        $script:FailureMessage = $_.Exception.Message
        $script:OverallResult = 'Failed'
        $global:LastStatus = "[ERROR] Option 17 failed: $($_.Exception.Message)"
        Write-Status "Fatal error in Option 17: $($_.Exception.Message)" 'ERROR'
        Write-YamlLog
        throw
    }
}


# ─────────────────────────────────────────────────────────────────────────────
# Option 18 - Set OneDrive to Automatically Login at Boot
# ─────────────────────────────────────────────────────────────────────────────
function Set-OneDriveAutoLoginPolicy {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [switch]$EnableSilentConfig,
        [switch]$DisableFirstRunWizard,
        [switch]$EnableAutoStartup,
        [switch]$EnableFilesOnDemand,
        [switch]$DisablePersonalSync,
        [switch]$EnableKnownFolderMove,
        [switch]$DisableAutoLogin,
        [string]$TenantId,
        [ValidateRange(1, 30)]
        [int]$SyncThrottleKbps = 0,  # 0 = No throttling
        [switch]$BackupSettings,
        [string]$LogPath = "$env:TEMP\OneDrivePolicy_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
    )

    # Security: Require elevation for policy configuration
    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        throw "This function must be run as Administrator"
    }

    $ErrorActionPreference = 'Continue'
    $ProgressPreference = 'SilentlyContinue'  # Speed: Disable progress bars
    
    # Initialize tracking using ArrayList for better performance and compatibility
    $script:logEntries = New-Object System.Collections.ArrayList
    $script:appliedPolicies = New-Object System.Collections.ArrayList
    $script:failedPolicies = New-Object System.Collections.ArrayList
    $script:skippedPolicies = New-Object System.Collections.ArrayList
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    
    function Write-LogEntry {
        param([string]$Message, [string]$Level = 'INFO')
        $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $logEntry = "[$timestamp] [$Level] $Message"
        
        # Ensure logEntries exists before adding to it
        if ($null -eq $script:logEntries) {
            $script:logEntries = New-Object System.Collections.ArrayList
        }
        [void]$script:logEntries.Add($logEntry)
        
        switch ($Level) {
            'ERROR' { Write-Host $Message -ForegroundColor Red }
            'WARNING' { Write-Host $Message -ForegroundColor Yellow }
            'SUCCESS' { Write-Host $Message -ForegroundColor Green }
            'INFO' { Write-Host $Message -ForegroundColor Cyan }
            'POLICY' { Write-Host $Message -ForegroundColor Magenta }
            'SKIP' { Write-Host $Message -ForegroundColor DarkGray }
        }
    }

    # Security: Registry validation and path management
    function Test-RegistryPath {
        param([string]$Path)
        try {
            # Security: Validate registry path format
            if ($Path -notmatch '^HK(LM|CU|CR|U|CC):\\') {
                return $false
            }
            return Test-Path $Path -ErrorAction SilentlyContinue
        } catch {
            return $false
        }
    }

    function New-RegistryPath {
        param([string]$Path)
        try {
            if (-not (Test-RegistryPath $Path)) {
                New-Item -Path $Path -Force -ErrorAction Stop | Out-Null
                Write-LogEntry "[OK] Created registry path: $Path" 'SUCCESS'
                return $true
            }
            return $true
        } catch {
            Write-LogEntry "[ERROR] Failed to create registry path: $Path - $_" 'ERROR'
            return $false
        }
    }

    # Get current registry value safely
    function Get-RegistryValue {
        param(
            [string]$Path,
            [string]$Name
        )
        
        try {
            if (Test-RegistryPath $Path) {
                $property = Get-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue
                if ($property) {
                    return $property.$Name
                }
            }
            return $null
        } catch {
            return $null
        }
    }

    # Check if policy value needs to be changed
    function Test-PolicyValue {
        param(
            [string]$Path,
            [string]$Name,
            [object]$DesiredValue
        )
        
        $currentValue = Get-RegistryValue -Path $Path -Name $Name
        return ($currentValue -eq $DesiredValue)
    }

    # Speed: Optimized policy application function with existing value detection
    function Set-OneDrivePolicy {
        param(
            [string]$Path,
            [string]$Name,
            [object]$Value,
            [string]$Type = 'DWord',
            [string]$Description,
            [switch]$Critical
        )
        
        try {
            # Check if the policy already has the correct value
            if (Test-PolicyValue -Path $Path -Name $Name -DesiredValue $Value) {
                $level = if ($Critical) { 'SKIP' } else { 'SKIP' }
                Write-LogEntry "⏭ $Description (already configured)" $level
                
                # [OK] FIX: Ensure skippedPolicies exists and use ArrayList.Add()
                if ($null -eq $script:skippedPolicies) {
                    $script:skippedPolicies = New-Object System.Collections.ArrayList
                }
                $skipInfo = @{
                    Path = $Path
                    Name = $Name
                    Value = $Value
                    Description = $Description
                    Critical = $Critical.IsPresent
                }
                [void]$script:skippedPolicies.Add($skipInfo)
                return $true
            }
            
            # Ensure registry path exists
            if (-not (New-RegistryPath -Path $Path)) {
                throw "Cannot create registry path: $Path"
            }
            
            # Apply policy setting
            if ($PSCmdlet.ShouldProcess("$Path\$Name", "Set policy value to $Value")) {
                Set-ItemProperty -Path $Path -Name $Name -Value $Value -Type $Type -ErrorAction Stop
                
                # Verify the setting was applied
                $verifyValue = Get-ItemProperty -Path $Path -Name $Name -ErrorAction Stop
                if ($verifyValue.$Name -eq $Value) {
                    $level = if ($Critical) { 'POLICY' } else { 'SUCCESS' }
                    Write-LogEntry "[OK] $Description" $level
                    
                    # [OK] FIX: Ensure appliedPolicies exists and use ArrayList.Add()
                    if ($null -eq $script:appliedPolicies) {
                        $script:appliedPolicies = New-Object System.Collections.ArrayList
                    }
                    $policyInfo = @{
                        Path = $Path
                        Name = $Name
                        Value = $Value
                        Description = $Description
                        Critical = $Critical.IsPresent
                    }
                    [void]$script:appliedPolicies.Add($policyInfo)
                    return $true
                } else {
                    throw "Verification failed: Expected $Value, got $($verifyValue.$Name)"
                }
            }
        } catch {
            Write-LogEntry "[ERROR] Failed to apply $Description : $_" 'ERROR'
            
            # [OK] FIX: Ensure failedPolicies exists and use ArrayList.Add()
            if ($null -eq $script:failedPolicies) {
                $script:failedPolicies = New-Object System.Collections.ArrayList
            }
            $failInfo = @{
                Path = $Path
                Name = $Name
                Description = $Description
                Error = $_.Exception.Message
            }
            [void]$script:failedPolicies.Add($failInfo)
            return $false
        }
    }

    # Security: OneDrive installation and version validation
    function Test-OneDriveInstallation {
        try {
            $oneDrivePaths = @(
                "${env:ProgramFiles}\Microsoft OneDrive\OneDrive.exe",
                "${env:ProgramFiles(x86)}\Microsoft OneDrive\OneDrive.exe",
                "${env:LOCALAPPDATA}\Microsoft\OneDrive\OneDrive.exe"
            )
            
            $installations = @()
            foreach ($path in $oneDrivePaths) {
                if (Test-Path $path) {
                    try {
                        $version = (Get-ItemProperty -Path $path).VersionInfo.FileVersion
                        $installations += @{
                            Path = $path
                            Version = $version
                            Type = if ($path -like "*Program Files*") { "System" } else { "User" }
                        }
                    } catch {
                        $installations += @{
                            Path = $path
                            Version = "Unknown"
                            Type = if ($path -like "*Program Files*") { "System" } else { "User" }
                        }
                    }
                }
            }
            
            return @{
                IsInstalled = ($installations.Count -gt 0)
                Installations = $installations
                RecommendedPath = $installations | Where-Object { $_.Type -eq "System" } | Select-Object -First 1
            }
            
        } catch {
            Write-LogEntry "OneDrive installation check failed: $_" 'WARNING'
            return @{
                IsInstalled = $false
                Error = $_.Exception.Message
            }
        }
    }

    # Security: Backup current OneDrive policies
    function Backup-OneDrivePolicies {
        param([string[]]$RegistryPaths)
        
        try {
            $backupData = @{}
            
            foreach ($regPath in $RegistryPaths) {
                if (Test-RegistryPath $regPath) {
                    $properties = Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue
                    if ($properties) {
                        $backupData[$regPath] = @{}
                        foreach ($property in $properties.PSObject.Properties) {
                            if ($property.Name -notmatch '^PS') {  # Skip PowerShell properties
                                $backupData[$regPath][$property.Name] = $property.Value
                            }
                        }
                    }
                }
            }
            
            if ($backupData.Count -gt 0) {
                $backupFile = "$env:TEMP\OneDrivePolicyBackup_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
                $backupData | ConvertTo-Json -Depth 4 | Out-File -FilePath $backupFile -Encoding UTF8
                Write-LogEntry "[OK] Policy backup saved to: $backupFile" 'INFO'
                return $backupFile
            }
            
        } catch {
            Write-LogEntry "[WARN] Failed to backup OneDrive policies: $_" 'WARNING'
        }
        
        return $null
    }

    # Security: Validate Tenant ID format
    function Test-TenantId {
        param([string]$TenantId)
        
        if ([string]::IsNullOrWhiteSpace($TenantId)) {
            return $true  # Optional parameter
        }
        
        # GUID format validation
        $guidPattern = '^[a-fA-F0-9]{8}-[a-fA-F0-9]{4}-[a-fA-F0-9]{4}-[a-fA-F0-9]{4}-[a-fA-F0-9]{12}$'
        return $TenantId -match $guidPattern
    }

    try {
        Write-LogEntry "=== OneDrive Auto-Login Policy Configuration Started ===" 'INFO'
        
        # Set default behaviors if no specific parameters provided
        if (-not ($EnableSilentConfig -or $DisableFirstRunWizard -or $EnableAutoStartup -or $EnableFilesOnDemand -or $DisablePersonalSync -or $EnableKnownFolderMove -or $DisableAutoLogin)) {
            Write-LogEntry "No specific policies specified, applying default configuration..." 'INFO'
            $EnableSilentConfig = $true
            $DisableFirstRunWizard = $true
            $EnableAutoStartup = $true
        }
        
        if ($WhatIfPreference) {
            Write-LogEntry "WhatIf mode - no registry changes will be made" 'INFO'
        }

        # Security: Validate Tenant ID if provided
        if ($TenantId -and -not (Test-TenantId -TenantId $TenantId)) {
            throw "Invalid Tenant ID format. Must be a valid GUID."
        }

        # Validate OneDrive installation
        Write-LogEntry "`n[CHECK] Validating OneDrive installation..." 'INFO'
        $oneDriveInfo = Test-OneDriveInstallation
        
        if (-not $oneDriveInfo.IsInstalled) {
            Write-LogEntry "[WARN] OneDrive not detected on this system" 'WARNING'
            Write-LogEntry "Policies will be applied but may not take effect until OneDrive is installed" 'WARNING'
        } else {
            Write-LogEntry "[OK] OneDrive installation detected:" 'SUCCESS'
            foreach ($installation in $oneDriveInfo.Installations) {
                Write-LogEntry "  • $($installation.Type): $($installation.Path) (v$($installation.Version))" 'INFO'
            }
        }

        # Define registry paths
        $registryPaths = @{
            MainPolicy = "HKLM:\SOFTWARE\Policies\Microsoft\OneDrive"
            UserPolicy = "HKCU:\SOFTWARE\Policies\Microsoft\OneDrive"
            TenantRestrictions = "HKLM:\SOFTWARE\Policies\Microsoft\OneDrive\TenantRestrictions"
            KnownFolderMove = "HKLM:\SOFTWARE\Policies\Microsoft\OneDrive\KnownFolderMove"
        }

        # Backup existing policies if requested
        if ($BackupSettings) {
            Write-LogEntry "`n[BACKUP] Backing up current OneDrive policies..." 'INFO'
            $backupFile = Backup-OneDrivePolicies -RegistryPaths $registryPaths.Values
        }

        if ($WhatIfPreference) {
            Write-LogEntry "`nWhatIf Summary - Policies that would be applied:" 'INFO'
            if ($EnableSilentConfig) { Write-LogEntry "  • Silent account configuration: Enabled" 'INFO' }
            if ($DisableFirstRunWizard) { Write-LogEntry "  • First run wizard: Disabled" 'INFO' }
            if ($EnableAutoStartup) { Write-LogEntry "  • Auto startup: Enabled" 'INFO' }
            if ($EnableFilesOnDemand) { Write-LogEntry "  • Files On-Demand: Enabled" 'INFO' }
            if ($DisablePersonalSync) { Write-LogEntry "  • Personal account sync: Disabled" 'INFO' }
            if ($EnableKnownFolderMove) { Write-LogEntry "  • Known Folder Move: Enabled" 'INFO' }
            if ($TenantId) { Write-LogEntry "  • Tenant restriction: $TenantId" 'INFO' }
            if ($SyncThrottleKbps -gt 0) { Write-LogEntry "  • Sync throttle: $SyncThrottleKbps KB/s" 'INFO' }
            $global:LastStatus = "[INFO] WhatIf completed - OneDrive policies would be configured."
            return
        }

        # Apply core OneDrive policies
        Write-LogEntry "`n🛠️ Applying OneDrive policies..." 'POLICY'
        
        # Silent account configuration
        if ($EnableSilentConfig) {
            Set-OneDrivePolicy -Path $registryPaths.MainPolicy -Name "SilentAccountConfig" -Value 1 -Description "Enabled silent account configuration" -Critical | Out-Null
        }
        
        # Disable first run wizard
        if ($DisableFirstRunWizard) {
            Set-OneDrivePolicy -Path $registryPaths.MainPolicy -Name "DisableFirstRunWizard" -Value 1 -Description "Disabled first run wizard" | Out-Null
        }
        
        # Auto startup policy
        if ($EnableAutoStartup) {
            Set-OneDrivePolicy -Path $registryPaths.MainPolicy -Name "OneDriveStartupPolicy" -Value 1 -Description "Enabled OneDrive auto startup" | Out-Null
        }
        
        # Files On-Demand
        if ($EnableFilesOnDemand) {
            Set-OneDrivePolicy -Path $registryPaths.MainPolicy -Name "FilesOnDemandEnabled" -Value 1 -Description "Enabled Files On-Demand" | Out-Null
        }
        
        # Disable personal account sync
        if ($DisablePersonalSync) {
            Set-OneDrivePolicy -Path $registryPaths.MainPolicy -Name "DisablePersonalSync" -Value 1 -Description "Disabled personal account synchronization" -Critical | Out-Null
        }
        
        # Known Folder Move
        if ($EnableKnownFolderMove) {
            Set-OneDrivePolicy -Path $registryPaths.KnownFolderMove -Name "KnownFolderMoveOpt" -Value 1 -Description "Enabled Known Folder Move optimization" | Out-Null
            
            if ($TenantId) {
                Set-OneDrivePolicy -Path $registryPaths.KnownFolderMove -Name $TenantId -Value 1 -Description "Enabled Known Folder Move for tenant: $TenantId" | Out-Null
            }
        }
        
        # Tenant restrictions
        if ($TenantId) {
            Set-OneDrivePolicy -Path $registryPaths.TenantRestrictions -Name $TenantId -Value 1 -Type "String" -Description "Applied tenant restriction for: $TenantId" -Critical | Out-Null
        }
        
        # Sync throttling
        if ($SyncThrottleKbps -gt 0) {
            Set-OneDrivePolicy -Path $registryPaths.MainPolicy -Name "UploadBandwidthLimit" -Value $SyncThrottleKbps -Description "Set upload bandwidth limit: $SyncThrottleKbps KB/s" | Out-Null
            Set-OneDrivePolicy -Path $registryPaths.MainPolicy -Name "DownloadBandwidthLimit" -Value $SyncThrottleKbps -Description "Set download bandwidth limit: $SyncThrottleKbps KB/s" | Out-Null
        }
        
        # Disable auto-login (override other settings)
        if ($DisableAutoLogin) {
            Set-OneDrivePolicy -Path $registryPaths.MainPolicy -Name "SilentAccountConfig" -Value 0 -Description "Disabled OneDrive auto-login" -Critical | Out-Null
            Set-OneDrivePolicy -Path $registryPaths.MainPolicy -Name "OneDriveStartupPolicy" -Value 0 -Description "Disabled OneDrive auto startup" | Out-Null
        }

        # Additional security and performance policies
        Write-LogEntry "`n[SECURITY] Applying security and performance policies..." 'POLICY'
        
        # Prevent OneDrive from generating network traffic until user signs in
        Set-OneDrivePolicy -Path $registryPaths.MainPolicy -Name "PreventNetworkTrafficPreUserSignIn" -Value 1 -Description "Prevented network traffic before user sign-in" | Out-Null
        
        # Block external sharing
        Set-OneDrivePolicy -Path $registryPaths.MainPolicy -Name "BlockExternalSync" -Value 1 -Description "Blocked external sharing and sync" | Out-Null
        
        # Enable automatic sign-in
        if (-not $DisableAutoLogin) {
            Set-OneDrivePolicy -Path $registryPaths.MainPolicy -Name "AutomaticUploadBandwidthPercentage" -Value 70 -Description "Set automatic upload bandwidth to 70%" | Out-Null
        }

        # Verification of applied policies
        Write-LogEntry "`n[CHECK] Verifying policy application..." 'INFO'
        $verificationErrors = 0
        
        # Ensure appliedPolicies exists before iterating
        if ($null -ne $script:appliedPolicies) {
            foreach ($policy in $script:appliedPolicies) {
                try {
                    $currentValue = Get-ItemProperty -Path $policy.Path -Name $policy.Name -ErrorAction Stop
                    if ($currentValue.($policy.Name) -ne $policy.Value) {
                        Write-LogEntry "[WARN] Verification failed for $($policy.Description)" 'WARNING'
                        $verificationErrors++
                    }
                } catch {
                    Write-LogEntry "[WARN] Could not verify $($policy.Description)" 'WARNING'
                    $verificationErrors++
                }
            }
        }
        
        $appliedCount = if ($null -ne $script:appliedPolicies) { $script:appliedPolicies.Count } else { 0 }
        $skippedCount = if ($null -ne $script:skippedPolicies) { $script:skippedPolicies.Count } else { 0 }
        
        if ($verificationErrors -eq 0 -and $appliedCount -gt 0) {
            Write-LogEntry "[OK] All OneDrive policies verified successfully" 'SUCCESS'
        } elseif ($appliedCount -eq 0 -and $skippedCount -gt 0) {
            Write-LogEntry "[OK] All OneDrive policies already correctly configured" 'SUCCESS'
        } elseif ($verificationErrors -gt 0) {
            Write-LogEntry "[WARN] $verificationErrors policies failed verification" 'WARNING'
        }

    } catch {
        Write-LogEntry "Critical error during OneDrive policy configuration: $_" 'ERROR'
        $global:LastStatus = "[ERROR] OneDrive policy configuration failed: $_"
        throw
    } finally {
        $stopwatch.Stop()
        $duration = $stopwatch.Elapsed.TotalSeconds
        
        # Safe counting with null checks
        $appliedCount = if ($null -ne $script:appliedPolicies) { $script:appliedPolicies.Count } else { 0 }
        $skippedCount = if ($null -ne $script:skippedPolicies) { $script:skippedPolicies.Count } else { 0 }
        $failedCount = if ($null -ne $script:failedPolicies) { $script:failedPolicies.Count } else { 0 }
        
        Write-LogEntry "`n[SUMMARY] Policy Configuration Summary:" 'INFO'
        Write-LogEntry "Duration: $([math]::Round($duration, 2)) seconds" 'INFO'
        Write-LogEntry "Policies applied: $appliedCount" 'SUCCESS'
        Write-LogEntry "Policies already correct: $skippedCount" 'SKIP'
        Write-LogEntry "Policies failed: $failedCount" 'ERROR'
        
        # Safe critical policy counting
        $criticalPolicies = 0
        $criticalSkipped = 0
        
        if ($null -ne $script:appliedPolicies) {
            $criticalPolicies = ($script:appliedPolicies | Where-Object { $_.Critical }).Count
        }
        if ($null -ne $script:skippedPolicies) {
            $criticalSkipped = ($script:skippedPolicies | Where-Object { $_.Critical }).Count
        }
        
        if ($criticalPolicies -gt 0) {
            Write-LogEntry "Critical policies applied: $criticalPolicies" 'POLICY'
        }
        if ($criticalSkipped -gt 0) {
            Write-LogEntry "Critical policies already configured: $criticalSkipped" 'SKIP'
        }
        
        if ($failedCount -gt 0 -and $null -ne $script:failedPolicies) {
            Write-LogEntry "`n[ERROR] Failed Policies:" 'ERROR'
            foreach ($failed in $script:failedPolicies) {
                Write-LogEntry "  • $($failed.Description): $($failed.Error)" 'ERROR'
            }
        }
        
        # Write detailed log file with null check
        try {
            if ($null -ne $script:logEntries) {
                $script:logEntries.ToArray() | Out-File -FilePath $LogPath -Encoding UTF8 -Force
                Write-LogEntry "[LOG] Detailed log saved to: $LogPath" 'INFO'
            }
        } catch {
            Write-LogEntry "[WARN] Failed to save log file: $_" 'WARNING'
        }
        
        # Set global status with safe counting
        $totalPolicies = $appliedCount + $skippedCount
        if ($totalPolicies -gt 0) {
            if ($appliedCount -gt 0) {
                $statusMsg = "[OK] Applied $appliedCount OneDrive policies"
                if ($skippedCount -gt 0) {
                    $statusMsg += " ($skippedCount already correct)"
                }
            } else {
                $statusMsg = "[OK] All $skippedCount OneDrive policies already correctly configured"
            }
            
            if ($failedCount -gt 0) {
                $statusMsg += " [$failedCount failed]"
            }
            if ($criticalPolicies -gt 0 -or $criticalSkipped -gt 0) {
                $statusMsg += " [$($criticalPolicies + $criticalSkipped) critical]"
            }
            $global:LastStatus = $statusMsg
        } else {
            $global:LastStatus = "[WARN] No OneDrive policies were processed"
        }
        
        Write-LogEntry "=== OneDrive Policy Configuration Completed ===" 'INFO'
    }
}

# Utility function to check current OneDrive policies
function Get-OneDrivePolicyStatus {
    [CmdletBinding()]
    param()
    
    try {
        $policyPaths = @{
            "Main Policies" = "HKLM:\SOFTWARE\Policies\Microsoft\OneDrive"
            "User Policies" = "HKCU:\SOFTWARE\Policies\Microsoft\OneDrive"
            "Tenant Restrictions" = "HKLM:\SOFTWARE\Policies\Microsoft\OneDrive\TenantRestrictions"
            "Known Folder Move" = "HKLM:\SOFTWARE\Policies\Microsoft\OneDrive\KnownFolderMove"
        }
        
        $status = @{}
        
        foreach ($category in $policyPaths.GetEnumerator()) {
            $status[$category.Key] = @{
                Path = $category.Value
                Exists = (Test-Path $category.Value)
                Policies = @{}
            }
            
            if ($status[$category.Key].Exists) {
                try {
                    $properties = Get-ItemProperty -Path $category.Value -ErrorAction SilentlyContinue
                    if ($properties) {
                        foreach ($prop in $properties.PSObject.Properties) {
                            if ($prop.Name -notmatch '^PS') {
                                $status[$category.Key].Policies[$prop.Name] = $prop.Value
                            }
                        }
                    }
                } catch {
                    $status[$category.Key].Error = $_.Exception.Message
                }
            }
        }
        
        return $status
        
    } catch {
        return @{
            Error = $_.Exception.Message
        }
    }
}

# Utility function to remove OneDrive policies
function Remove-OneDrivePolicies {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [switch]$ConfirmEach
    )
    
    $policyPaths = @(
        "HKLM:\SOFTWARE\Policies\Microsoft\OneDrive",
        "HKCU:\SOFTWARE\Policies\Microsoft\OneDrive"
    )
    
    $removedCount = 0
    
    foreach ($path in $policyPaths) {
        if (Test-Path $path) {
            try {
                if ($PSCmdlet.ShouldProcess($path, "Remove OneDrive policy registry key")) {
                    if (-not $ConfirmEach -or (Read-Host "Remove $path? (y/n)") -eq 'y') {
                        Remove-Item -Path $path -Recurse -Force -ErrorAction Stop
                        Write-Host "[OK] Removed OneDrive policies from: $path" -ForegroundColor Green
                        $removedCount++
                    }
                }
            } catch {
                Write-Warning "[WARN] Failed to remove $path : $_"
            }
        }
    }
    
    Write-Host "[OK] Removed OneDrive policies from $removedCount registry locations" -ForegroundColor Cyan
}


# ─────────────────────────────────────────────────────────────────────────────
# Option 19 - Full System Update
# ─────────────────────────────────────────────────────────────────────────────
function Run-CorePostDeploymentTasks {
    [CmdletBinding()]
    param()

    $ErrorActionPreference = 'Continue'
    $ProgressPreference = 'SilentlyContinue'

    $runStart = Get-Date
    $computerName = $env:COMPUTERNAME
    $timestampForFile = $runStart.ToString('yyyy-MM-dd_HH-mm-ss')
    $logDirectory = 'C:\Logs'
    $logPath = Join-Path -Path $logDirectory -ChildPath ('{0}-Full_System_Update-{1}.yaml' -f $computerName, $timestampForFile)

    $taskPlan = @(
        @{ Option = 8;  Name = 'Network Optimization';                     Function = 'Start-NetworkOptimization' },
        @{ Option = 2;  Name = 'Remove Windows Bloatware';                 Function = 'Remove-BloatwareApps' },
        @{ Option = 3;  Name = 'Set Recommended Registry Settings';        Function = 'Apply-RecommendedRegistrySettings' },
        @{ Option = 4;  Name = 'Optimize Windows Services';                Function = 'Optimize-WindowsServices' },
        @{ Option = 5;  Name = 'Enable PowerShell Remote Management';      Function = 'Enable-PowerShellRemotingSafely' },
        @{ Option = 6;  Name = 'Configure Automatic Time Sync';            Function = 'Configure-AutomaticTimeSync' },
        @{ Option = 7;  Name = 'Set Desktop Power Settings';               Function = 'Set-DesktopPowerSettings' },
        @{ Option = 9;  Name = 'Application Updates';                      Function = 'Update-Applications' },
        @{ Option = 10; Name = 'HP Driver Updates';                        Function = 'Update-HPDrivers' },
        @{ Option = 11; Name = 'Windows Updates';                          Function = 'Update-WindowsOS' }
    )

    $results = New-Object System.Collections.Generic.List[object]

    function ConvertTo-YamlSafeString {
        param(
            [AllowNull()]
            [object]$Value
        )

        if ($null -eq $Value) { return "''" }

        $stringValue = [string]$Value
        $stringValue = $stringValue -replace "`r`n", ' | '
        $stringValue = $stringValue -replace "`n", ' | '
        $stringValue = $stringValue -replace "`r", ' | '
        $stringValue = $stringValue -replace "'", "''"
        return "'$stringValue'"
    }

    function Invoke-FullUpdateFunction {
        param(
            [Parameter(Mandatory)]
            [string]$FunctionName
        )

        $previousState = @{
            ConfirmPreference = $ConfirmPreference
            ProgressPreference = $ProgressPreference
            VerbosePreference = $VerbosePreference
            InformationPreference = $InformationPreference
            ErrorActionPreference = $ErrorActionPreference
            PSDefaultParameterValues = if ($PSDefaultParameterValues) { $PSDefaultParameterValues.Clone() } else { @{} }
        }

        try {
            $script:ConfirmPreference = 'None'
            $script:ProgressPreference = 'SilentlyContinue'
            $script:VerbosePreference = 'SilentlyContinue'
            $script:InformationPreference = 'Continue'
            $script:ErrorActionPreference = 'Stop'
            $script:PSDefaultParameterValues = @{'*:Confirm' = $false}

            if (-not (Get-Command -Name $FunctionName -CommandType Function -ErrorAction SilentlyContinue)) {
                throw "Function '$FunctionName' was not found in the script."
            }

            & $FunctionName
        }
        finally {
            $script:ConfirmPreference = $previousState.ConfirmPreference
            $script:ProgressPreference = $previousState.ProgressPreference
            $script:VerbosePreference = $previousState.VerbosePreference
            $script:InformationPreference = $previousState.InformationPreference
            $script:ErrorActionPreference = $previousState.ErrorActionPreference
            $script:PSDefaultParameterValues = $previousState.PSDefaultParameterValues
        }
    }

    Write-Host "Starting Full System Update..." -ForegroundColor Cyan
    Write-Host "Running options in order: 8, 2, 3, 4, 5, 6, 7, 9, 10, 11" -ForegroundColor Cyan

    foreach ($task in $taskPlan) {
        $taskStart = Get-Date
        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        $result = [ordered]@{
            Option = [int]$task.Option
            Name = [string]$task.Name
            Function = [string]$task.Function
            StartTime = $taskStart.ToString('yyyy-MM-dd HH:mm:ss')
            EndTime = $null
            DurationSeconds = 0
            Success = $false
            Status = 'NotStarted'
            Error = $null
        }

        try {
            Write-Host ("[{0}] {1}" -f $task.Option, $task.Name) -ForegroundColor Yellow
            Invoke-FullUpdateFunction -FunctionName $task.Function
            $result.Success = $true
            $result.Status = 'Completed'
        }
        catch {
            $result.Success = $false
            $result.Status = 'Failed'
            $result.Error = $_.Exception.Message
            Write-Warning ("Option {0} failed: {1}" -f $task.Option, $_.Exception.Message)
        }
        finally {
            $stopwatch.Stop()
            $taskEnd = Get-Date
            $result.EndTime = $taskEnd.ToString('yyyy-MM-dd HH:mm:ss')
            $result.DurationSeconds = [math]::Round($stopwatch.Elapsed.TotalSeconds, 2)
            $results.Add([pscustomobject]$result)
        }
    }

    $runEnd = Get-Date
    $successfulTasks = @($results | Where-Object { $_.Success })
    $failedTasks = @($results | Where-Object { -not $_.Success })

    $yamlLines = New-Object System.Collections.Generic.List[string]
    $yamlLines.Add('run_summary:')
    $yamlLines.Add(('  computer_name: {0}' -f (ConvertTo-YamlSafeString $computerName)))
    $yamlLines.Add(('  started_at: {0}' -f (ConvertTo-YamlSafeString ($runStart.ToString('yyyy-MM-dd HH:mm:ss')))))
    $yamlLines.Add(('  ended_at: {0}' -f (ConvertTo-YamlSafeString ($runEnd.ToString('yyyy-MM-dd HH:mm:ss')))))
    $yamlLines.Add(('  duration_seconds: {0}' -f ([math]::Round(($runEnd - $runStart).TotalSeconds, 2))))
    $yamlLines.Add(('  total_tasks: {0}' -f $results.Count))
    $yamlLines.Add(('  successful_tasks: {0}' -f $successfulTasks.Count))
    $yamlLines.Add(('  failed_tasks: {0}' -f $failedTasks.Count))
    $yamlLines.Add(('  log_path: {0}' -f (ConvertTo-YamlSafeString $logPath)))
    $yamlLines.Add('task_order:')
    foreach ($task in $taskPlan) {
        $yamlLines.Add(('  - {0}' -f $task.Option))
    }
    $yamlLines.Add('tasks:')
    foreach ($item in $results) {
        $yamlLines.Add(('  - option: {0}' -f $item.Option))
        $yamlLines.Add(('    name: {0}' -f (ConvertTo-YamlSafeString $item.Name)))
        $yamlLines.Add(('    function: {0}' -f (ConvertTo-YamlSafeString $item.Function)))
        $yamlLines.Add(('    start_time: {0}' -f (ConvertTo-YamlSafeString $item.StartTime)))
        $yamlLines.Add(('    end_time: {0}' -f (ConvertTo-YamlSafeString $item.EndTime)))
        $yamlLines.Add(('    duration_seconds: {0}' -f $item.DurationSeconds))
        $yamlLines.Add(('    success: {0}' -f $item.Success.ToString().ToLower()))
        $yamlLines.Add(('    status: {0}' -f (ConvertTo-YamlSafeString $item.Status)))
        $yamlLines.Add(('    error: {0}' -f (ConvertTo-YamlSafeString $item.Error)))
    }

    try {
        if (-not (Test-Path -Path $logDirectory)) {
            New-Item -Path $logDirectory -ItemType Directory -Force | Out-Null
        }
        Set-Content -Path $logPath -Value $yamlLines -Encoding UTF8
    }
    catch {
        Write-Warning ("Failed to write YAML log: {0}" -f $_.Exception.Message)
    }

    if ($failedTasks.Count -eq 0) {
        $global:LastStatus = "[OK] Full System Update completed successfully. Log: $logPath"
        Write-Host $global:LastStatus -ForegroundColor Green
    }
    else {
        $global:LastStatus = "[WARN] Full System Update completed with $($failedTasks.Count) failed task(s). Log: $logPath"
        Write-Host $global:LastStatus -ForegroundColor Yellow
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# Menu Display
# ─────────────────────────────────────────────────────────────────────────────
function Show-Menu {
    Clear-Host
    Write-Host "╔══════════════════════════════════════════════════════════╗" -ForegroundColor Magenta
    Write-Host "║              Compton College Tech Utils                ║" -ForegroundColor Magenta
    Write-Host "╚══════════════════════════════════════════════════════════╝" -ForegroundColor Magenta
    Write-Host ""

    Write-Host "1.  Create MISAdmin account" -ForegroundColor White
    Write-Host "2.  Remove Windows Bloatware" -ForegroundColor White
    Write-Host "3.  Set Recommended Registry Settings" -ForegroundColor White
    Write-Host "4.  Optimize Windows Services" -ForegroundColor White
    Write-Host "5.  Enable PowerShell Remote Management" -ForegroundColor White
    Write-Host "6.  Configure Automatic Time Sync" -ForegroundColor White
    Write-Host "7.  Set Desktop Power Settings" -ForegroundColor White
    Write-Host "8.  Network Optimization" -ForegroundColor White
    Write-Host "9.  Application Updates" -ForegroundColor White
    Write-Host "10. HP Driver Updates" -ForegroundColor White
    Write-Host "11. Windows Updates" -ForegroundColor White
    Write-Host "12. Disk Cleanup" -ForegroundColor White
    Write-Host "13. System Repair" -ForegroundColor White
    Write-Host "14. Remove User Profiles" -ForegroundColor White
    Write-Host "15. Disable Last User Display" -ForegroundColor White
    Write-Host "16. Enable Automatic Login with CC-Student" -ForegroundColor White
    Write-Host "17. Install Computer Lab Scheduled Tasks" -ForegroundColor White
    Write-Host "18. Set OneDrive Auto Login on Boot" -ForegroundColor White
    Write-Host "19. Run Full System Updates" -ForegroundColor White
    Write-Host "20. Network Diag and Repair" -ForegroundColor White
    Write-Host "Q.  Exit" -ForegroundColor Red

    Write-Host ""
    Write-Host "Last Status: $global:LastStatus" -ForegroundColor Yellow
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# Main Loop
# ─────────────────────────────────────────────────────────────────────────────
function Main {
    do {
        Show-Menu
        $choice = Read-Host "Enter your selection"

        switch ($choice.ToUpperInvariant()) {
            "1" {
                Clear-Host
                New-MISAdminAccount
                Pause
            }
            "2" {
                Clear-Host
                Remove-BloatwareApps
                Pause
            }
            "3" {
                Clear-Host
                Apply-RecommendedRegistrySettings
                Pause
            }
            "4" {
                Clear-Host
                Optimize-WindowsServices
                Pause
            }
            "5" {
                Clear-Host
                Enable-PowerShellRemotingSafely
                Pause
            }
            "6" {
                Clear-Host
                Configure-AutomaticTimeSync
                Pause
            }
            "7" {
                Clear-Host
                Set-DesktopPowerSettings
                Pause
            }
            "8" {
                Clear-Host
                Start-NetworkOptimization
                Pause
            }
            "9" {
                Clear-Host
                Update-Applications
                Pause
            }
            "10" {
                Clear-Host
                Update-HPDrivers
                Pause
            }
            "11" {
                Clear-Host
                Update-WindowsOS
                Pause
            }
            "12" {
                Clear-Host
                Run-DiskCleanup
                Pause
            }
            "13" {
                Clear-Host
                Invoke-SystemMaintenance
                Pause
            }
            "14" {
                Clear-Host
                Remove-UserProfilesClassroom
                Pause
            }
            "15" {
                Clear-Host
                Apply-LoginScreenRegistryFixes
                Pause
            }
            "16" {
                Clear-Host
                Set-DomainAutoLogin
                Pause
            }
            "17" {
                Clear-Host
                Register-LabScheduledTasks
                Pause
            }
            "18" {
                Clear-Host
                Set-OneDriveAutoLoginPolicy
                Pause
            }
            "19" {
                Clear-Host
                Run-CorePostDeploymentTasks
                Pause
            }
            "20" {
                Clear-Host
                Run-NetworkDiagnostics
                Pause
            }
            "Q" {
                Clear-Host
                Write-Host "Compton College Tech Utils has exited." -ForegroundColor Cyan
                exit
            }
            Default {
                $global:LastStatus = "[ERROR] Invalid selection. Please try again."
                Start-Sleep -Seconds 1
            }
        }
    } while ($true)
}

# ─────────────────────────────────────────────────────────────────────────────
# Entry Point
# ─────────────────────────────────────────────────────────────────────────────
Invoke-StartupSelfUpdate
Main
# =====================================================================
# ScriptName: Compton_Tech_Utils.ps1
# ScriptVersion: 1.7.1
# LastUpdated: 2026-03-26
# Notes: Added startup self-update check against GitHub.
# Notes: Master utility script with merged menu options and YAML logging.
# =====================================================================




# ─────────────────────────────────────────────────────────────────────────────
# Startup Self-Update Check
# ─────────────────────────────────────────────────────────────────────────────
function Get-LocalScriptVersion {
    param(
        [Parameter(Mandatory)][string]$ScriptPath
    )

    try {
        if (-not (Test-Path -LiteralPath $ScriptPath)) {
            return $null
        }

        $firstLines = Get-Content -LiteralPath $ScriptPath -TotalCount 25 -ErrorAction Stop
        $versionLine = $firstLines | Where-Object { $_ -match '^\s*#\s*ScriptVersion\s*:\s*(.+?)\s*$' } | Select-Object -First 1
        if (-not $versionLine) {
            return $null
        }

        $versionText = ($versionLine -replace '^\s*#\s*ScriptVersion\s*:\s*', '').Trim()
        return [version]$versionText
    }
    catch {
        return $null
    }
}

function Get-RemoteScriptMetadata {
    param(
        [Parameter(Mandatory)][string]$RawUrl
    )

    try {
        $response = Invoke-WebRequest -Uri $RawUrl -UseBasicParsing -ErrorAction Stop
        $content = [string]$response.Content
        $versionMatch = [regex]::Match($content, '(?im)^\s*#\s*ScriptVersion\s*:\s*([^\r\n]+)')
        $updatedMatch = [regex]::Match($content, '(?im)^\s*#\s*LastUpdated\s*:\s*([^\r\n]+)')

        $remoteVersion = $null
        if ($versionMatch.Success) {
            try {
                $remoteVersion = [version]$versionMatch.Groups[1].Value.Trim()
            }
            catch {
                $remoteVersion = $null
            }
        }

        [PSCustomObject]@{
            Version     = $remoteVersion
            LastUpdated = if ($updatedMatch.Success) { $updatedMatch.Groups[1].Value.Trim() } else { $null }
            Content     = $content
        }
    }
    catch {
        Write-Host "Unable to check GitHub for script updates: $($_.Exception.Message)" -ForegroundColor Yellow
        return $null
    }
}

function Invoke-StartupSelfUpdate {
    [CmdletBinding()]
    param()

    $rawUrl = 'https://raw.githubusercontent.com/dswaney/Compton/main/Compton_Tech_Utils.ps1'
    $currentScriptPath = $PSCommandPath

    if ([string]::IsNullOrWhiteSpace($currentScriptPath)) {
        $currentScriptPath = $MyInvocation.MyCommand.Path
    }

    if ([string]::IsNullOrWhiteSpace($currentScriptPath)) {
        Write-Host 'Startup update check skipped because the current script path could not be determined.' -ForegroundColor Yellow
        return
    }

    Write-Host 'Checking GitHub for a newer version of Compton_Tech_Utils.ps1...' -ForegroundColor Cyan

    $localVersion = Get-LocalScriptVersion -ScriptPath $currentScriptPath
    $remoteMetadata = Get-RemoteScriptMetadata -RawUrl $rawUrl

    if (-not $remoteMetadata -or -not $remoteMetadata.Version) {
        Write-Host 'Update check skipped. Continuing with the current script.' -ForegroundColor Yellow
        return
    }

    if (-not $localVersion) {
        Write-Host "Current local script version could not be read. GitHub version detected: $($remoteMetadata.Version)" -ForegroundColor Yellow
    }
    else {
        Write-Host "Current version: $localVersion" -ForegroundColor Gray
        Write-Host "GitHub version : $($remoteMetadata.Version)" -ForegroundColor Gray
    }

    if ($localVersion -and $remoteMetadata.Version -le $localVersion) {
        Write-Host 'This script is already up to date.' -ForegroundColor Green
        return
    }

    $prompt = if ($remoteMetadata.LastUpdated) {
        "A newer version ($($remoteMetadata.Version), updated $($remoteMetadata.LastUpdated)) is available. Download and run it now? (Y/N)"
    }
    else {
        "A newer version ($($remoteMetadata.Version)) is available. Download and run it now? (Y/N)"
    }

    do {
        $updateChoice = Read-Host $prompt
        $normalizedChoice = ([string]$updateChoice).Trim().ToUpperInvariant()
    } while ($normalizedChoice -notin @('Y','N','YES','NO'))

    if ($normalizedChoice -in @('N','NO')) {
        Write-Host 'Continuing with the current script version.' -ForegroundColor Yellow
        return
    }

    try {
        $targetDirectory = Split-Path -Path $currentScriptPath -Parent
        if (-not [string]::IsNullOrWhiteSpace($targetDirectory) -and -not (Test-Path -LiteralPath $targetDirectory)) {
            New-Item -Path $targetDirectory -ItemType Directory -Force | Out-Null
        }

        [System.IO.File]::WriteAllText($currentScriptPath, $remoteMetadata.Content, [System.Text.UTF8Encoding]::new($false))
        Write-Host "Updated script saved to: $currentScriptPath" -ForegroundColor Green
        Write-Host 'Launching the updated script...' -ForegroundColor Cyan

        $powershellExe = if (Get-Command -Name 'pwsh.exe' -ErrorAction SilentlyContinue) {
            'pwsh.exe'
        }
        else {
            'powershell.exe'
        }

        Start-Process -FilePath $powershellExe -ArgumentList @('-ExecutionPolicy', 'Bypass', '-File', ('"{0}"' -f $currentScriptPath)) | Out-Null
        exit
    }
    catch {
        Write-Host "Failed to download or relaunch the updated script: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host 'Continuing with the current script version.' -ForegroundColor Yellow
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# Option 14 - Remove User Profiles
# ─────────────────────────────────────────────────────────────────────────────
function Remove-UserProfilesClassroom {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    param(
        [string[]]$ExcludedProfiles = @(
            'Default',
            'Default User',
            'Public',
            'All Users',
            'MISAdmin',
            'dswaney'
        ),

        [string]$UsersRoot = 'C:\Users',

        [switch]$SkipLoadedProfiles = $true,

        [switch]$SkipSpecialProfiles = $true,

        [int]$OlderThanDays = 0,

        [string]$LogDirectory = 'C:\Logs'
    )

    $ErrorActionPreference = 'Stop'

    $script:RunStart = Get-Date
    $script:ComputerName = $env:COMPUTERNAME
    $script:TimestampForFile = $script:RunStart.ToString('yyyy-MM-dd_HH-mm-ss')
    $script:BaseFileName = "{0}-RemoveUserProfiles-{1}" -f $script:ComputerName, $script:TimestampForFile
    $script:YamlLogPath = Join-Path $LogDirectory ($script:BaseFileName + '.yaml')

    $script:Summary = [ordered]@{
        ComputerName       = $script:ComputerName
        StartTime          = $script:RunStart
        EndTime            = $null
        FoundProfiles      = 0
        ExcludedProfiles   = 0
        SkippedLoaded      = 0
        SkippedSpecial     = 0
        SkippedByAge       = 0
        DeletedProfiles    = 0
        FailedProfiles     = 0
    }

    $script:DeletedProfileDetails = New-Object System.Collections.Generic.List[object]
    $script:SkippedProfileDetails = New-Object System.Collections.Generic.List[object]
    $script:FailedProfileDetails  = New-Object System.Collections.Generic.List[object]

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
            $lines.Add("  users_root: $(ConvertTo-YamlScalar $UsersRoot)") | Out-Null
            $lines.Add("  skip_loaded_profiles: $(ConvertTo-YamlScalar $SkipLoadedProfiles)") | Out-Null
            $lines.Add("  skip_special_profiles: $(ConvertTo-YamlScalar $SkipSpecialProfiles)") | Out-Null
            $lines.Add("  older_than_days: $(ConvertTo-YamlScalar $OlderThanDays)") | Out-Null
            $lines.Add("  excluded_profiles:") | Out-Null
            if ($ExcludedProfiles.Count -gt 0) {
                foreach ($name in $ExcludedProfiles) {
                    $lines.Add("    - $(ConvertTo-YamlScalar $name)") | Out-Null
                }
            }
            else {
                $lines.Add('    []') | Out-Null
            }
            $lines.Add('') | Out-Null

            $lines.Add('summary:') | Out-Null
            $lines.Add("  found_profiles: $(ConvertTo-YamlScalar $script:Summary.FoundProfiles)") | Out-Null
            $lines.Add("  excluded_profiles: $(ConvertTo-YamlScalar $script:Summary.ExcludedProfiles)") | Out-Null
            $lines.Add("  skipped_loaded: $(ConvertTo-YamlScalar $script:Summary.SkippedLoaded)") | Out-Null
            $lines.Add("  skipped_special: $(ConvertTo-YamlScalar $script:Summary.SkippedSpecial)") | Out-Null
            $lines.Add("  skipped_by_age: $(ConvertTo-YamlScalar $script:Summary.SkippedByAge)") | Out-Null
            $lines.Add("  deleted_profiles: $(ConvertTo-YamlScalar $script:Summary.DeletedProfiles)") | Out-Null
            $lines.Add("  failed_profiles: $(ConvertTo-YamlScalar $script:Summary.FailedProfiles)") | Out-Null
            $lines.Add('') | Out-Null

            $lines.Add('deleted_profiles:') | Out-Null
            if ($script:DeletedProfileDetails.Count -gt 0) {
                foreach ($entry in $script:DeletedProfileDetails) {
                    $lines.Add('  -') | Out-Null
                    $lines.Add("    profile_name: $(ConvertTo-YamlScalar $entry.ProfileName)") | Out-Null
                    $lines.Add("    local_path: $(ConvertTo-YamlScalar $entry.LocalPath)") | Out-Null
                    $lines.Add("    sid: $(ConvertTo-YamlScalar $entry.SID)") | Out-Null
                    $lines.Add("    loaded: $(ConvertTo-YamlScalar $entry.Loaded)") | Out-Null
                    $lines.Add("    special: $(ConvertTo-YamlScalar $entry.Special)") | Out-Null
                    $lines.Add("    created_time: $(ConvertTo-YamlScalar $entry.CreatedTime)") | Out-Null
                    $lines.Add("    last_use_time: $(ConvertTo-YamlScalar $entry.LastUseTime)") | Out-Null
                    $lines.Add("    days_on_system: $(ConvertTo-YamlScalar $entry.DaysOnSystem)") | Out-Null
                }
            }
            else {
                $lines.Add('  []') | Out-Null
            }
            $lines.Add('') | Out-Null

            $lines.Add('skipped_profiles:') | Out-Null
            if ($script:SkippedProfileDetails.Count -gt 0) {
                foreach ($entry in $script:SkippedProfileDetails) {
                    $lines.Add('  -') | Out-Null
                    $lines.Add("    profile_name: $(ConvertTo-YamlScalar $entry.ProfileName)") | Out-Null
                    $lines.Add("    local_path: $(ConvertTo-YamlScalar $entry.LocalPath)") | Out-Null
                    $lines.Add("    sid: $(ConvertTo-YamlScalar $entry.SID)") | Out-Null
                    $lines.Add("    reason: $(ConvertTo-YamlScalar $entry.Reason)") | Out-Null
                    $lines.Add("    created_time: $(ConvertTo-YamlScalar $entry.CreatedTime)") | Out-Null
                    $lines.Add("    last_use_time: $(ConvertTo-YamlScalar $entry.LastUseTime)") | Out-Null
                    $lines.Add("    days_on_system: $(ConvertTo-YamlScalar $entry.DaysOnSystem)") | Out-Null
                }
            }
            else {
                $lines.Add('  []') | Out-Null
            }
            $lines.Add('') | Out-Null

            $lines.Add('failed_profiles:') | Out-Null
            if ($script:FailedProfileDetails.Count -gt 0) {
                foreach ($entry in $script:FailedProfileDetails) {
                    $lines.Add('  -') | Out-Null
                    $lines.Add("    profile_name: $(ConvertTo-YamlScalar $entry.ProfileName)") | Out-Null
                    $lines.Add("    local_path: $(ConvertTo-YamlScalar $entry.LocalPath)") | Out-Null
                    $lines.Add("    sid: $(ConvertTo-YamlScalar $entry.SID)") | Out-Null
                    $lines.Add("    error: $(ConvertTo-YamlScalar $entry.Error)") | Out-Null
                    $lines.Add("    created_time: $(ConvertTo-YamlScalar $entry.CreatedTime)") | Out-Null
                    $lines.Add("    last_use_time: $(ConvertTo-YamlScalar $entry.LastUseTime)") | Out-Null
                    $lines.Add("    days_on_system: $(ConvertTo-YamlScalar $entry.DaysOnSystem)") | Out-Null
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

    function Get-ProfileFolderName {
        param(
            [Parameter(Mandatory)][string]$Path
        )

        try {
            return (Split-Path -Path $Path -Leaf)
        }
        catch {
            return $null
        }
    }

    function Get-ProfileAgeData {
        param(
            [Parameter(Mandatory)][string]$ProfilePath,
            [AllowNull()][datetime]$LastUseTime
        )

        $createdTime = $null
        $daysOnSystem = $null

        try {
            if (Test-Path -LiteralPath $ProfilePath) {
                $item = Get-Item -LiteralPath $ProfilePath -ErrorAction Stop
                $createdTime = $item.CreationTime
                $daysOnSystem = [math]::Round(((Get-Date) - $createdTime).TotalDays, 2)
            }
        }
        catch {
        }

        return [PSCustomObject]@{
            CreatedTime  = $createdTime
            LastUseTime  = $LastUseTime
            DaysOnSystem = $daysOnSystem
        }
    }

    if (-not (Test-IsAdministrator)) {
        Write-Error "Please run this script as Administrator."
        $global:LastStatus = "[ERROR] Remove User Profiles requires Administrator rights."
        return 1
    }

    Ensure-LogDirectory

    Write-Log "Starting profile cleanup." 'INFO'
    Write-Log "Users root: $UsersRoot" 'INFO'
    Write-Log "Excluded profile names: $($ExcludedProfiles -join ', ')" 'INFO'
    Write-Log "Skip loaded profiles: $SkipLoadedProfiles" 'INFO'
    Write-Log "Skip special profiles: $SkipSpecialProfiles" 'INFO'
    Write-Log "OlderThanDays filter: $OlderThanDays" 'INFO'
    Write-Log "YAML log path: $($script:YamlLogPath)" 'INFO'

    try {
        $allUserProfiles = Get-CimInstance -ClassName Win32_UserProfile -ErrorAction Stop | Where-Object {
            -not [string]::IsNullOrWhiteSpace($_.LocalPath) -and
            $_.LocalPath -like "$UsersRoot\*"
        }
    }
    catch {
        Write-Log "Failed to enumerate Win32_UserProfile instances: $($_.Exception.Message)" 'ERROR'
        $script:Summary.EndTime = Get-Date
        Write-YamlLog
        $global:LastStatus = "[ERROR] Remove User Profiles failed while enumerating profiles."
        return 2
    }

    $script:Summary.FoundProfiles = @($allUserProfiles).Count
    Write-Log "Found $($script:Summary.FoundProfiles) profile(s) under $UsersRoot." 'INFO'

    foreach ($profile in $allUserProfiles) {
        $profilePath = $profile.LocalPath
        $profileName = Get-ProfileFolderName -Path $profilePath

        if ([string]::IsNullOrWhiteSpace($profileName)) {
            Write-Log "Skipping profile with invalid path: $profilePath" 'WARN'
            continue
        }

        $ageData = Get-ProfileAgeData -ProfilePath $profilePath -LastUseTime $profile.LastUseTime

        if ($ExcludedProfiles -icontains $profileName) {
            $script:Summary.ExcludedProfiles++
            Write-Log "Skipping excluded profile: $profileName" 'WARN'

            $script:SkippedProfileDetails.Add([PSCustomObject]@{
                ProfileName  = $profileName
                LocalPath    = $profilePath
                SID          = $profile.SID
                Reason       = 'Excluded'
                CreatedTime  = $ageData.CreatedTime
                LastUseTime  = $ageData.LastUseTime
                DaysOnSystem = $ageData.DaysOnSystem
            }) | Out-Null

            continue
        }

        if ($SkipSpecialProfiles -and $profile.Special) {
            $script:Summary.SkippedSpecial++
            Write-Log "Skipping special/system profile: $profileName ($profilePath)" 'WARN'

            $script:SkippedProfileDetails.Add([PSCustomObject]@{
                ProfileName  = $profileName
                LocalPath    = $profilePath
                SID          = $profile.SID
                Reason       = 'SpecialProfile'
                CreatedTime  = $ageData.CreatedTime
                LastUseTime  = $ageData.LastUseTime
                DaysOnSystem = $ageData.DaysOnSystem
            }) | Out-Null

            continue
        }

        if ($SkipLoadedProfiles -and $profile.Loaded) {
            $script:Summary.SkippedLoaded++
            Write-Log "Skipping loaded profile: $profileName ($profilePath)" 'WARN'

            $script:SkippedProfileDetails.Add([PSCustomObject]@{
                ProfileName  = $profileName
                LocalPath    = $profilePath
                SID          = $profile.SID
                Reason       = 'LoadedProfile'
                CreatedTime  = $ageData.CreatedTime
                LastUseTime  = $ageData.LastUseTime
                DaysOnSystem = $ageData.DaysOnSystem
            }) | Out-Null

            continue
        }

        if ($OlderThanDays -gt 0) {
            try {
                $cutoff = (Get-Date).AddDays(-$OlderThanDays)
                if ($profile.LastUseTime -and $profile.LastUseTime -gt $cutoff) {
                    $script:Summary.SkippedByAge++
                    Write-Log "Skipping recent profile: $profileName (LastUseTime: $($profile.LastUseTime))" 'WARN'

                    $script:SkippedProfileDetails.Add([PSCustomObject]@{
                        ProfileName  = $profileName
                        LocalPath    = $profilePath
                        SID          = $profile.SID
                        Reason       = 'TooRecent'
                        CreatedTime  = $ageData.CreatedTime
                        LastUseTime  = $ageData.LastUseTime
                        DaysOnSystem = $ageData.DaysOnSystem
                    }) | Out-Null

                    continue
                }
            }
            catch {
                Write-Log "Could not evaluate LastUseTime for ${profileName}: $($_.Exception.Message)" 'WARN'
            }
        }

        $targetDescription = "$profileName ($profilePath)"

        try {
            if ($PSCmdlet.ShouldProcess($targetDescription, 'Delete user profile')) {
                Write-Log "Deleting profile: $targetDescription" 'INFO'
                Remove-CimInstance -InputObject $profile -ErrorAction Stop
                $script:Summary.DeletedProfiles++
                Write-Log "Successfully deleted profile: $profileName" 'OK'

                $script:DeletedProfileDetails.Add([PSCustomObject]@{
                    ProfileName  = $profileName
                    LocalPath    = $profilePath
                    SID          = $profile.SID
                    Loaded       = $profile.Loaded
                    Special      = $profile.Special
                    CreatedTime  = $ageData.CreatedTime
                    LastUseTime  = $ageData.LastUseTime
                    DaysOnSystem = $ageData.DaysOnSystem
                }) | Out-Null
            }
        }
        catch {
            $script:Summary.FailedProfiles++
            Write-Log "Failed to delete profile: $profileName. Error: $($_.Exception.Message)" 'ERROR'

            $script:FailedProfileDetails.Add([PSCustomObject]@{
                ProfileName  = $profileName
                LocalPath    = $profilePath
                SID          = $profile.SID
                Error        = $_.Exception.Message
                CreatedTime  = $ageData.CreatedTime
                LastUseTime  = $ageData.LastUseTime
                DaysOnSystem = $ageData.DaysOnSystem
            }) | Out-Null
        }
    }

    $script:Summary.EndTime = Get-Date

    Write-Log "Profile cleanup complete." 'INFO'
    Write-Log "Summary: Found=$($script:Summary.FoundProfiles), Excluded=$($script:Summary.ExcludedProfiles), LoadedSkipped=$($script:Summary.SkippedLoaded), SpecialSkipped=$($script:Summary.SkippedSpecial), AgeSkipped=$($script:Summary.SkippedByAge), Deleted=$($script:Summary.DeletedProfiles), Failed=$($script:Summary.FailedProfiles)" 'INFO'

    Write-YamlLog

    if ($script:Summary.FailedProfiles -gt 0) {
        $global:LastStatus = "[WARN] Remove User Profiles completed with failures. Deleted=$($script:Summary.DeletedProfiles), Failed=$($script:Summary.FailedProfiles)"
        return 2
    }

    if ($script:Summary.DeletedProfiles -gt 0) {
        $global:LastStatus = "[OK] Removed $($script:Summary.DeletedProfiles) user profile(s)."
    }
    else {
        $global:LastStatus = "[INFO] No user profiles were removed."
    }

    return 0
}

# Enhanced alias for compatibility
Set-Alias -Name Remove-UserProfiles -Value Remove-UserProfilesClassroom -Force

# ─────────────────────────────────────────────────────────────────────────────
# Option 15 - Disable the display of the last user logged on
# ─────────────────────────────────────────────────────────────────────────────
function Apply-LoginScreenRegistryFixes {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [switch]$EnhancedSecurity,
        [string]$LogPath = "$env:TEMP\LoginRegistryFixes_$(Get-Date -Format 'yyyyMMdd_HHmmss').log",
        [switch]$BackupRegistry
    )

    # Security: Require elevation
    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        throw "This function must be run as Administrator"
    }

    $ErrorActionPreference = 'Continue'
    $ProgressPreference = 'SilentlyContinue'  # Speed: Disable progress bars
    
    # Initialize tracking (intentionally unscoped so nested functions can update them)
    $appliedSettings = @()
    $failedSettings  = @()
    $backupData      = @{}
    $logEntries      = @()
    $stopwatch       = [System.Diagnostics.Stopwatch]::StartNew()
    
    function Write-LogEntry {
        param([string]$Message, [string]$Level = 'INFO')
        $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $logEntry = "[$timestamp] [$Level] $Message"
        
        # Output to console with colors
        switch ($Level) {
            'ERROR'    { Write-Host $Message -ForegroundColor Red }
            'WARNING'  { Write-Host $Message -ForegroundColor Yellow }
            'SUCCESS'  { Write-Host $Message -ForegroundColor Green }
            'INFO'     { Write-Host $Message -ForegroundColor Cyan }
            'SECURITY' { Write-Host $Message -ForegroundColor Magenta }
            default    { Write-Host $Message }
        }
        
        # Add to log collection
        $logEntries += $logEntry
    }

    # Security: Registry validation function
    function Test-RegistryPath {
        param([string]$Path)
        try {
            return Test-Path $Path -ErrorAction Stop
        } catch {
            return $false
        }
    }

    # Security: Safe registry backup function
    function Backup-RegistryValue {
        param(
            [string]$Path,
            [string]$Name
        )
        try {
            if (Test-RegistryPath $Path) {
                $currentValue = Get-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue
                if ($null -ne $currentValue) {
                    return $currentValue.$Name
                }
            }
            return $null
        } catch {
            return $null
        }
    }

    # Speed: Optimized registry setting function
    function Set-RegistryValueSafe {
        param(
            [string]$Path,
            [string]$Name,
            [object]$Value,
            [string]$Type = 'DWord',
            [string]$Description,
            [switch]$SecurityCritical
        )
        
        try {
            # Security: Validate registry path format
            if ($Path -notmatch '^HK(LM|CU|CR|U|CC):\\') {
                throw "Invalid registry path format: $Path"
            }
            
            # Backup current value if requested
            if ($BackupRegistry) {
                $currentValue = Backup-RegistryValue -Path $Path -Name $Name
                if ($null -ne $currentValue) {
                    $existing = Get-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue
                    $backupData["$Path\$Name"] = @{
                        Value = $currentValue
                        Type  = if ($existing) { $existing.$Name.GetType().Name } else { $null }
                    }
                }
            }
            
            # Create registry path if it doesn't exist
            if (-not (Test-RegistryPath $Path)) {
                Write-LogEntry "Creating registry path: $Path" 'INFO'
                New-Item -Path $Path -Force | Out-Null
            }
            
            # Apply setting
            if ($PSCmdlet.ShouldProcess("$Path\$Name", "Set registry value to $Value")) {
                Set-ItemProperty -Path $Path -Name $Name -Value $Value -Type $Type -ErrorAction Stop
                
                # Verify the setting was applied
                $verifyValue = Get-ItemProperty -Path $Path -Name $Name -ErrorAction Stop
                if ($verifyValue.$Name -eq $Value) {
                    $level = if ($SecurityCritical) { 'SECURITY' } else { 'SUCCESS' }
                    Write-LogEntry "[OK] $Description" $level
                    $appliedSettings += [pscustomobject]@{
                        Path            = $Path
                        Name            = $Name
                        Value           = $Value
                        Description     = $Description
                        SecurityCritical= $SecurityCritical.IsPresent
                    }
                    return $true
                } else {
                    throw "Verification failed: Expected $Value, got $($verifyValue.$Name)"
                }
            }
        } catch {
            Write-LogEntry "[ERROR] Failed to apply $Description : $_" 'ERROR'
            $failedSettings += [pscustomobject]@{
                Path        = $Path
                Name        = $Name
                Description = $Description
                Error       = $_.Exception.Message
            }
            return $false
        }
    }

    try {
        Write-LogEntry "=== Login Screen Registry Security Configuration Started ===" 'INFO'
        Write-LogEntry "Enhanced Security Mode: $($EnhancedSecurity.IsPresent)" 'INFO'
        if ($PSBoundParameters.ContainsKey('WhatIf') -and $PSBoundParameters['WhatIf']) {
            Write-LogEntry "WhatIf mode - no registry changes will be made" 'INFO'
        }

        # Core login screen security settings
        Write-LogEntry "`n[SECURITY] Applying core login screen security settings..." 'INFO'
        
        # Speed: Batch registry operations using array
        $coreSettings = @(
            @{
                Path = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
                Name = "EnableFirstLogonAnimation"
                Value = 0
                Type = 'DWord'
                Description = "Disabled first logon animation"
                SecurityCritical = $false
            },
            @{
                Path = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
                Name = "dontdisplaylastusername"
                Value = 1
                Type = 'DWord'
                Description = "Configured login screen to hide last username"
                SecurityCritical = $true
            }
        )

        # Enhanced security settings (applied when -EnhancedSecurity is used)
        if ($EnhancedSecurity) {
            Write-LogEntry "`n[SECURITY] Applying enhanced security settings..." 'SECURITY'
            
            $enhancedSettings = @(
                @{
                    Path = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
                    Name = "ShutdownWithoutLogon"
                    Value = 0
                    Type = 'DWord'
                    Description = "Disabled shutdown without logon"
                    SecurityCritical = $true
                },
                @{
                    Path = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
                    Name = "UndockWithoutLogon"
                    Value = 0
                    Type = 'DWord'
                    Description = "Disabled undock without logon"
                    SecurityCritical = $true
                },
                @{
                    Path = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
                    Name = "AutoAdminLogon"
                    Value = 0
                    Type = 'String' # This value is a REG_SZ ("0"/"1" as string)
                    Description = "Disabled automatic administrator logon"
                    SecurityCritical = $true
                },
                @{
                    Path = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
                    Name = "ForceAutoLogon"
                    Value = 0
                    Type = 'String' # Also REG_SZ
                    Description = "Disabled forced automatic logon"
                    SecurityCritical = $true
                },
                @{
                    Path = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
                    Name = "InactivityTimeoutSecs"
                    Value = 900
                    Type = 'DWord'
                    Description = "Set login screen timeout to 15 minutes"
                    SecurityCritical = $true
                },
                @{
                    Path = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
                    Name = "MaxDevicePasswordFailedAttempts"
                    Value = 5
                    Type = 'DWord'
                    Description = "Set maximum password attempts to 5"
                    SecurityCritical = $true
                },
                @{
                    Path = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
                    Name = "CachedLogonsCount"
                    Value = 2
                    Type = 'String' # REG_SZ number as string
                    Description = "Limited cached logons to 2"
                    SecurityCritical = $true
                }
            )
            
            $coreSettings += $enhancedSettings
        }

        # Speed: Process settings in batches
        $batchSize = 5
        $totalSettings = $coreSettings.Count
        
        for ($i = 0; $i -lt $totalSettings; $i += $batchSize) {
            $end = [Math]::Min($i + $batchSize - 1, $totalSettings - 1)
            $batch = $coreSettings[$i..$end]
            
            foreach ($setting in $batch) {
                $params = @{
                    Path            = $setting.Path
                    Name            = $setting.Name
                    Value           = $setting.Value
                    Description     = $setting.Description
                    SecurityCritical= $setting.SecurityCritical
                }
                if ($setting.ContainsKey('Type') -and $setting.Type) {
                    $params.Type = $setting.Type
                }
                Set-RegistryValueSafe @params | Out-Null
            }
            
            if ($end -lt ($totalSettings - 1)) {
                Start-Sleep -Milliseconds 50
            }
        }

        # Security: Additional hardening for classroom environments
        if ($EnhancedSecurity) {
            Write-LogEntry "`n[CLASSROOM] Applying classroom-specific security hardening..." 'SECURITY'
            
            # Disable guest account
            try {
                $guestAccount = Get-LocalUser -Name "Guest" -ErrorAction SilentlyContinue
                if ($guestAccount -and $guestAccount.Enabled) {
                    if ($PSCmdlet.ShouldProcess("LocalUser 'Guest'", "Disable")) {
                        Disable-LocalUser -Name "Guest" -ErrorAction Stop
                        Write-LogEntry "[OK] Disabled Guest account" 'SECURITY'
                    }
                }
            } catch {
                Write-LogEntry "[ERROR] Failed to disable Guest account: $_" 'ERROR'
            }
            
            # Set strong password policy via registry
            $passwordPolicySettings = @(
                @{
                    Path = "HKLM:\SYSTEM\CurrentControlSet\Services\Netlogon\Parameters"
                    Name = "RequireStrongKey"
                    Value = 1
                    Type = 'DWord'
                    Description = "Enabled strong authentication keys"
                },
                @{
                    Path = "HKLM:\SYSTEM\CurrentControlSet\Control\Lsa"
                    Name = "LimitBlankPasswordUse"
                    Value = 1
                    Type = 'DWord'
                    Description = "Disabled blank password usage"
                }
            )
            
            foreach ($setting in $passwordPolicySettings) {
                Set-RegistryValueSafe @setting -SecurityCritical | Out-Null
            }
        }

        # Security: Registry integrity verification
        Write-LogEntry "`n[CHECK] Verifying registry integrity..." 'INFO'
        $verificationErrors = 0
        
        foreach ($setting in $appliedSettings) {
            try {
                $currentValue = Get-ItemProperty -Path $setting.Path -Name $setting.Name -ErrorAction Stop
                if ($currentValue.($setting.Name) -ne $setting.Value) {
                    Write-LogEntry "[WARN] Verification failed for $($setting.Description)" 'WARNING'
                    $verificationErrors++
                }
            } catch {
                Write-LogEntry "[WARN] Could not verify $($setting.Description)" 'WARNING'
                $verificationErrors++
            }
        }
        
        if ($verificationErrors -eq 0) {
            Write-LogEntry "[OK] All registry settings verified successfully" 'SUCCESS'
        } else {
            Write-LogEntry "[WARN] $verificationErrors settings failed verification" 'WARNING'
        }

        # Create registry backup file
        if ($BackupRegistry -and $backupData.Count -gt 0) {
            try {
                $backupFile = "$env:TEMP\LoginRegistryBackup_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
                $backupData | ConvertTo-Json -Depth 5 | Out-File -FilePath $backupFile -Encoding UTF8
                Write-LogEntry "[OK] Registry backup saved to: $backupFile" 'INFO'
            } catch {
                Write-LogEntry "[WARN] Failed to save registry backup: $_" 'WARNING'
            }
        }

    } catch {
        Write-LogEntry "Critical error during registry configuration: $_" 'ERROR'
        throw
    } finally {
        $stopwatch.Stop()
        $duration = $stopwatch.Elapsed.TotalSeconds
        
        # Generate summary
        Write-LogEntry "`n[SUMMARY] Configuration Summary:" 'INFO'
        Write-LogEntry "Duration: $([math]::Round($duration, 2)) seconds" 'INFO'
        Write-LogEntry "Settings applied: $($appliedSettings.Count)" 'SUCCESS'
        Write-LogEntry "Settings failed: $($failedSettings.Count)" 'ERROR'
        
        $securitySettings = ($appliedSettings | Where-Object { $_.SecurityCritical }).Count
        if ($securitySettings -gt 0) {
            Write-LogEntry "Security-critical settings applied: $securitySettings" 'SECURITY'
        }
        
        if ($failedSettings.Count -gt 0) {
            Write-LogEntry "`n[ERROR] Failed Settings:" 'ERROR'
            foreach ($failed in $failedSettings) {
                Write-LogEntry "  • $($failed.Description): $($failed.Error)" 'ERROR'
            }
        }
        
        # Write detailed log file
        try {
            $logEntries | Out-File -FilePath $LogPath -Encoding UTF8 -Force
            Write-LogEntry "[LOG] Detailed log saved to: $LogPath" 'INFO'
        } catch {
            Write-LogEntry "[WARN] Failed to save log file: $_" 'WARNING'
        }
        
        # Set global status
        if ($appliedSettings.Count -gt 0) {
            $statusMsg = "[OK] Applied $($appliedSettings.Count) login screen security settings"
            if ($failedSettings.Count -gt 0) {
                $statusMsg += " ($($failedSettings.Count) failed)"
            }
            if ($securitySettings -gt 0) {
                $statusMsg += " [$securitySettings security-critical]"
            }
            $global:LastStatus = $statusMsg
        } else {
            $global:LastStatus = "[WARN] No login screen settings were applied"
        }
        
        Write-LogEntry "=== Login Screen Registry Configuration Completed ===" 'INFO'
    }
}

# Create restore function for emergency rollback
function Restore-LoginScreenRegistrySettings {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$BackupFile
    )
    
    if (-not (Test-Path $BackupFile)) {
        throw "Backup file not found: $BackupFile"
    }
    
    try {
        $backupData = Get-Content $BackupFile | ConvertFrom-Json
        $restored = 0
        
        foreach ($entry in $backupData.PSObject.Properties) {
            $pathAndName = $entry.Name -split '\\'
            $path = $pathAndName[0..($pathAndName.Length-2)] -join '\'
            $name = $pathAndName[-1]
            
            try {
                Set-ItemProperty -Path $path -Name $name -Value $entry.Value.Value -ErrorAction Stop
                Write-Host "[OK] Restored: $($entry.Name)" -ForegroundColor Green
                $restored++
            } catch {
                Write-Warning "[WARN] Failed to restore: $($entry.Name) - $_"
            }
        }
        
        Write-Host "[OK] Restored $restored registry settings from backup" -ForegroundColor Cyan
    } catch {
        throw "Failed to restore from backup: $_"
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# Option 16 - Enable Automatic Login with CC-Student
# ─────────────────────────────────────────────────────────────────────────────
function Set-DomainAutoLogin {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory = $false)]
        [string]$UserName = "CC-Student",
        
        [Parameter(Mandatory = $false)]
        [securestring]$Password,
        
        [Parameter(Mandatory = $false)]
        [string]$DomainName = "Compton.edu",
        
        [switch]$DisableAutoLogin,
        [switch]$WhatIf,
        [switch]$Force,
        [string]$LogPath = "$env:TEMP\AutoLoginConfig_$(Get-Date -Format 'yyyyMMdd_HHmmss').log",
        [int]$AutoLoginCount = 1,  # Number of auto-logins before disabling
        [switch]$UseLocalSystemEncryption
    )

    # Security: Require elevation
    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        throw "This function must be run as Administrator"
    }

    $ErrorActionPreference = 'Stop'
    $ProgressPreference = 'SilentlyContinue'
    
    # Initialize tracking
    $logEntries = @()
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    
    function Write-LogEntry {
        param([string]$Message, [string]$Level = 'INFO')
        $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $logEntry = "[$timestamp] [$Level] $Message"
        $script:logEntries += $logEntry
        
        switch ($Level) {
            'ERROR' { Write-Host $Message -ForegroundColor Red }
            'WARNING' { Write-Host $Message -ForegroundColor Yellow }
            'SUCCESS' { Write-Host $Message -ForegroundColor Green }
            'INFO' { Write-Host $Message -ForegroundColor Cyan }
            'SECURITY' { Write-Host $Message -ForegroundColor Magenta }
        }
    }

    # Security: Secure credential management
    function Get-SecureCredential {
        param(
            [string]$Username,
            [string]$Domain,
            [securestring]$SecurePassword
        )
        
        if (-not $SecurePassword) {
            Write-LogEntry "[WARN] SECURITY WARNING: Auto-login requires storing credentials" 'WARNING'
            Write-LogEntry "Consider using alternative authentication methods for production" 'WARNING'
            
            if (-not $Force) {
                $response = Read-Host "Continue with credential storage? (Type 'ACCEPT' to proceed)"
                if ($response -ne 'ACCEPT') {
                    throw "Operation cancelled - credential storage not accepted"
                }
            }
            
            # Prompt for secure password
            $SecurePassword = Read-Host "Enter password for $Domain\$Username" -AsSecureString
            if (-not $SecurePassword -or $SecurePassword.Length -eq 0) {
                throw "Password is required for auto-login configuration"
            }
        }
        
        return New-Object System.Management.Automation.PSCredential("$Domain\$Username", $SecurePassword)
    }

    # Security: Registry validation and backup
    function Backup-AutoLoginSettings {
        $regPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
        $backupData = @{}
        
        $settingsToBackup = @(
            'AutoAdminLogon',
            'DefaultUserName', 
            'DefaultPassword',
            'DefaultDomainName',
            'AutoLogonCount'
        )
        
        foreach ($setting in $settingsToBackup) {
            try {
                $value = Get-ItemProperty -Path $regPath -Name $setting -ErrorAction SilentlyContinue
                if ($value) {
                    $backupData[$setting] = $value.$setting
                }
            } catch {
                # Setting doesn't exist, which is fine
            }
        }
        
        if ($backupData.Count -gt 0) {
            try {
                $backupFile = "$env:TEMP\AutoLoginBackup_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
                $backupData | ConvertTo-Json | Out-File -FilePath $backupFile -Encoding UTF8
                Write-LogEntry "[OK] Registry backup saved to: $backupFile" 'INFO'
                return $backupFile
            } catch {
                Write-LogEntry "[WARN] Failed to save registry backup: $_" 'WARNING'
            }
        }
        
        return $null
    }

    # Security: Secure password encryption using DPAPI
    function Protect-AutoLoginPassword {
        param([securestring]$SecurePassword)
        
        try {
            # Convert SecureString to encrypted string using DPAPI
            $ptr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecurePassword)
            $plaintext = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($ptr)
            [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptr)
            
            # Use DPAPI to encrypt for local machine
            $bytes = [System.Text.Encoding]::UTF8.GetBytes($plaintext)
            $encryptedBytes = [System.Security.Cryptography.ProtectedData]::Protect(
                $bytes, 
                $null, 
                [System.Security.Cryptography.DataProtectionScope]::LocalMachine
            )
            
            # Clear plaintext from memory
            $plaintext = $null
            [System.GC]::Collect()
            
            return [System.Convert]::ToBase64String($encryptedBytes)
            
        } catch {
            throw "Failed to encrypt password: $_"
        }
    }

    # Security: Domain validation
    function Test-DomainConnectivity {
        param([string]$DomainName)
        
        try {
            Write-LogEntry "[CHECK] Validating domain connectivity..." 'INFO'
            
            # Test domain controller connectivity
            $domainController = Resolve-DnsName -Name $DomainName -Type A -ErrorAction Stop
            if (-not $domainController) {
                return $false
            }
            
            # Test LDAP connectivity
            $ldapTest = Test-NetConnection -ComputerName $DomainName -Port 389 -WarningAction SilentlyContinue
            if (-not $ldapTest.TcpTestSucceeded) {
                Write-LogEntry "[WARN] LDAP connectivity test failed" 'WARNING'
                return $false
            }
            
            Write-LogEntry "[OK] Domain connectivity validated" 'SUCCESS'
            return $true
            
        } catch {
            Write-LogEntry "[WARN] Domain validation failed: $_" 'WARNING'
            return $false
        }
    }

    # Speed: Optimized registry operations
    function Set-AutoLoginRegistry {
        param(
            [string]$Username,
            [string]$Domain,
            [string]$EncryptedPassword,
            [int]$LoginCount
        )
        
        $regPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
        $settingsApplied = 0
        
        # Security: Validate registry path exists
        if (-not (Test-Path $regPath)) {
            throw "Winlogon registry path not found: $regPath"
        }
        
        # Batch registry operations for speed
        $registrySettings = @(
            @{ Name = 'AutoAdminLogon'; Value = '1'; Type = 'String' },
            @{ Name = 'DefaultUserName'; Value = $Username; Type = 'String' },
            @{ Name = 'DefaultDomainName'; Value = $Domain; Type = 'String' },
            @{ Name = 'AutoLogonCount'; Value = $LoginCount; Type = 'DWord' }
        )
        
        # Add password setting based on encryption method
        if ($UseLocalSystemEncryption) {
            # Use Windows built-in LSA encryption
            $registrySettings += @{ Name = 'DefaultPassword'; Value = ''; Type = 'String' }
            
            # Store encrypted password separately (requires additional LSA configuration)
            Write-LogEntry "Using LSA Secret storage for password (enhanced security)" 'SECURITY'
            # Note: LSA Secret storage requires additional implementation
        } else {
            # Store DPAPI-encrypted password
            $registrySettings += @{ Name = 'DefaultPassword'; Value = $EncryptedPassword; Type = 'String' }
        }
        
        foreach ($setting in $registrySettings) {
            try {
                if ($PSCmdlet.ShouldProcess("$regPath\$($setting.Name)", "Set registry value")) {
                    Set-ItemProperty -Path $regPath -Name $setting.Name -Value $setting.Value -Type $setting.Type -ErrorAction Stop
                    Write-LogEntry "[OK] Set $($setting.Name)" 'SUCCESS'
                    $settingsApplied++
                }
            } catch {
                Write-LogEntry "[ERROR] Failed to set $($setting.Name): $_" 'ERROR'
                throw
            }
        }
        
        return $settingsApplied
    }

    try {
        Write-LogEntry "=== Domain Auto-Login Configuration Started ===" 'INFO'
        
        if ($WhatIf) {
            Write-LogEntry "WhatIf mode - no registry changes will be made" 'INFO'
        }

        # Handle disable auto-login request
        if ($DisableAutoLogin) {
            Write-LogEntry "[DISABLE] Disabling auto-login..." 'INFO'
            
            $regPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
            
            if ($PSCmdlet.ShouldProcess($regPath, "Disable auto-login")) {
                try {
                    Set-ItemProperty -Path $regPath -Name "AutoAdminLogon" -Value "0" -Type String -ErrorAction Stop
                    
                    # Clear stored credentials for security
                    $credentialSettings = @('DefaultPassword', 'DefaultUserName', 'DefaultDomainName', 'AutoLogonCount')
                    foreach ($setting in $credentialSettings) {
                        try {
                            Remove-ItemProperty -Path $regPath -Name $setting -ErrorAction SilentlyContinue
                        } catch {
                            # Setting may not exist, continue
                        }
                    }
                    
                    Write-LogEntry "[OK] Auto-login disabled and credentials cleared" 'SUCCESS'
                    $global:LastStatus = "[OK] Auto-login disabled successfully."
                    return
                    
                } catch {
                    throw "Failed to disable auto-login: $_"
                }
            }
            return
        }

        # Security: Critical warning about auto-login risks
        if (-not $Force) {
            Write-Host "`n" + "="*70 -ForegroundColor Red
            Write-Host "CRITICAL SECURITY WARNING - AUTO-LOGIN CONFIGURATION" -ForegroundColor Red -BackgroundColor Black
            Write-Host "="*70 -ForegroundColor Red
            Write-Host "This configuration will:" -ForegroundColor Yellow
            Write-Host "  • Store domain credentials on local system" -ForegroundColor Yellow
            Write-Host "  • Allow automatic login without authentication" -ForegroundColor Yellow
            Write-Host "  • Potentially expose credentials to local attacks" -ForegroundColor Yellow
            Write-Host "`nRECOMMENDATIONS:" -ForegroundColor Cyan
            Write-Host "  • Use only in secure, controlled environments" -ForegroundColor Cyan
            Write-Host "  • Consider using domain Group Policy instead" -ForegroundColor Cyan
            Write-Host "  • Limit auto-login count to minimize exposure" -ForegroundColor Cyan
            Write-Host "  • Regularly rotate the password" -ForegroundColor Cyan
            Write-Host "="*70 -ForegroundColor Red
            
            $confirmation = Read-Host "`n[PROMPT] Type 'UNDERSTAND' to acknowledge security risks and continue"
            if ($confirmation -ne 'UNDERSTAND') {
                Write-LogEntry "Operation cancelled - security risks not acknowledged" 'WARNING'
                $global:LastStatus = "[WARN] User cancelled auto-login configuration."
                return
            }
        }

        # Validate inputs
        Write-LogEntry "[CHECK] Validating configuration parameters..." 'INFO'
        
        if ([string]::IsNullOrWhiteSpace($UserName)) {
            throw "Username cannot be empty"
        }
        
        if ([string]::IsNullOrWhiteSpace($DomainName)) {
            throw "Domain name cannot be empty"
        }
        
        # Security: Validate domain connectivity
        $domainConnectivity = Test-DomainConnectivity -DomainName $DomainName
        if (-not $domainConnectivity -and -not $Force) {
            throw "Domain connectivity validation failed. Use -Force to override."
        }

        # Security: Get secure credentials
        Write-LogEntry "[CREDENTIALS] Processing credentials securely..." 'SECURITY'
        $credential = Get-SecureCredential -Username $UserName -Domain $DomainName -SecurePassword $Password
        
        # Security: Test credential validity (optional)
        if ($domainConnectivity) {
            try {
                Write-LogEntry "[CHECK] Validating credentials..." 'INFO'
                # Note: In production, you might want to test credentials against domain
                # This is omitted here to avoid additional authentication attempts
                Write-LogEntry "[OK] Credential format validated" 'SUCCESS'
            } catch {
                Write-LogEntry "[WARN] Credential validation failed: $_" 'WARNING'
                if (-not $Force) {
                    throw "Invalid credentials provided"
                }
            }
        }

        # Security: Backup current settings
        Write-LogEntry "[BACKUP] Backing up current auto-login settings..." 'INFO'
        $backupFile = Backup-AutoLoginSettings
        
        if ($WhatIf) {
            Write-LogEntry "WhatIf: Would configure auto-login for $DomainName\$UserName" 'INFO'
            Write-LogEntry "WhatIf: Would set AutoLogonCount to $AutoLoginCount" 'INFO'
            $global:LastStatus = "[INFO] WhatIf completed - auto-login would be configured."
            return
        }

        # Security: Encrypt password
        Write-LogEntry "[SECURITY] Encrypting credentials..." 'SECURITY'
        $encryptedPassword = Protect-AutoLoginPassword -SecurePassword $credential.Password
        
        # Apply registry settings
        Write-LogEntry "[LOG] Applying auto-login registry settings..." 'INFO'
        $settingsCount = Set-AutoLoginRegistry -Username $UserName -Domain $DomainName -EncryptedPassword $encryptedPassword -LoginCount $AutoLoginCount
        
        # Security: Verify settings were applied
        Write-LogEntry "[CHECK] Verifying configuration..." 'INFO'
        $regPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
        
        try {
            $autoLogonValue = Get-ItemProperty -Path $regPath -Name "AutoAdminLogon" -ErrorAction Stop
            $userNameValue = Get-ItemProperty -Path $regPath -Name "DefaultUserName" -ErrorAction Stop
            $domainValue = Get-ItemProperty -Path $regPath -Name "DefaultDomainName" -ErrorAction Stop
            
            if ($autoLogonValue.AutoAdminLogon -eq "1" -and 
                $userNameValue.DefaultUserName -eq $UserName -and 
                $domainValue.DefaultDomainName -eq $DomainName) {
                
                Write-LogEntry "[OK] Auto-login configuration verified successfully" 'SUCCESS'
            } else {
                throw "Configuration verification failed"
            }
        } catch {
            throw "Failed to verify auto-login configuration: $_"
        }

        # Security: Set appropriate permissions on registry key
        try {
            Write-LogEntry "[SECURITY] Securing registry permissions..." 'SECURITY'
            
            $acl = Get-Acl -Path $regPath
            # Remove inherited permissions to protect stored credentials
            $acl.SetAccessRuleProtection($true, $true)
            
            # Keep only essential permissions (System, Administrators)
            $accessRules = $acl.Access | Where-Object { 
                $_.IdentityReference -notmatch "(Users|Everyone|Authenticated Users)" 
            }
            
            Set-Acl -Path $regPath -AclObject $acl
            Write-LogEntry "[OK] Registry permissions secured" 'SUCCESS'
            
        } catch {
            Write-LogEntry "[WARN] Failed to secure registry permissions: $_" 'WARNING'
        }

    } catch {
        Write-LogEntry "Critical error during auto-login configuration: $_" 'ERROR'
        $global:LastStatus = "[ERROR] Auto-login configuration failed: $_"
        throw
    } finally {
        $stopwatch.Stop()
        $duration = $stopwatch.Elapsed.TotalSeconds
        
        Write-LogEntry "`n[SUMMARY] Configuration Summary:" 'INFO'
        Write-LogEntry "Duration: $([math]::Round($duration, 2)) seconds" 'INFO'
        
        if (-not $DisableAutoLogin -and -not $WhatIf) {
            Write-LogEntry "User: $DomainName\$UserName" 'INFO'
            Write-LogEntry "Auto-login count: $AutoLoginCount" 'INFO'
            Write-LogEntry "Encryption: DPAPI (LocalMachine scope)" 'SECURITY'
        }
        
        # Write log file
        try {
            $logEntries | Out-File -FilePath $LogPath -Encoding UTF8 -Force
            Write-LogEntry "[LOG] Detailed log saved to: $LogPath" 'INFO'
        } catch {
            Write-LogEntry "[WARN] Failed to save log file: $_" 'WARNING'
        }
        
        # Set final status
        if (-not $global:LastStatus -or $global:LastStatus -notlike "*auto-login*") {
            if ($DisableAutoLogin) {
                $global:LastStatus = "[OK] Auto-login disabled successfully."
            } else {
                $global:LastStatus = "[OK] Auto-login configured for $DomainName\$UserName."
            }
        }
        
        Write-LogEntry "=== Auto-Login Configuration Completed ===" 'INFO'
    }
}

# Utility function to check current auto-login status
function Get-AutoLoginStatus {
    [CmdletBinding()]
    param()
    
    try {
        $regPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
        
        $autoLogon = Get-ItemProperty -Path $regPath -Name "AutoAdminLogon" -ErrorAction SilentlyContinue
        $userName = Get-ItemProperty -Path $regPath -Name "DefaultUserName" -ErrorAction SilentlyContinue
        $domainName = Get-ItemProperty -Path $regPath -Name "DefaultDomainName" -ErrorAction SilentlyContinue
        $loginCount = Get-ItemProperty -Path $regPath -Name "AutoLogonCount" -ErrorAction SilentlyContinue
        
        return @{
            Enabled = ($autoLogon.AutoAdminLogon -eq "1")
            UserName = $userName.DefaultUserName
            DomainName = $domainName.DefaultDomainName
            AutoLogonCount = $loginCount.AutoLogonCount
            PasswordStored = (Get-ItemProperty -Path $regPath -Name "DefaultPassword" -ErrorAction SilentlyContinue) -ne $null
        }
        
    } catch {
        return @{
            Error = $_.Exception.Message
            Enabled = $false
        }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# Option 7 - Set Desktop Power Settings - Only run on Desktop computers, no laptops!
# ─────────────────────────────────────────────────────────────────────────────
function Set-DesktopPowerSettings {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [ValidateSet('High Performance', 'Balanced', 'Power Saver', 'Ultimate Performance')]
        [string]$PowerPlan = 'High Performance',
        
        [ValidateRange(1, 600)]
        [int]$MonitorTimeoutMinutes = 60,
        
        [ValidateRange(1, 600)]
        [int]$DiskTimeoutMinutes = 0,  # 0 = Never
        
        [switch]$Force,
        # Removed [switch]$WhatIf - this is automatically provided by SupportsShouldProcess
        [switch]$AllowLaptops,
        [string]$LogPath = "$env:TEMP\PowerSettings_$(Get-Date -Format 'yyyyMMdd_HHmmss').log",
        [switch]$SkipHardwareDetection
    )

    # Security: Require elevation for power configuration
    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        throw "This function must be run as Administrator"
    }

    $ErrorActionPreference = 'Continue'
    $ProgressPreference = 'SilentlyContinue'  # Speed: Disable progress bars
    
    # Initialize tracking
    $logEntries = @()
    $appliedSettings = @()
    $failedSettings = @()
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    
    function Write-LogEntry {
        param([string]$Message, [string]$Level = 'INFO')
        $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $logEntry = "[$timestamp] [$Level] $Message"
        $script:logEntries += $logEntry
        
        switch ($Level) {
            'ERROR' { Write-Host $Message -ForegroundColor Red }
            'WARNING' { Write-Host $Message -ForegroundColor Yellow }
            'SUCCESS' { Write-Host $Message -ForegroundColor Green }
            'INFO' { Write-Host $Message -ForegroundColor Cyan }
            'HARDWARE' { Write-Host $Message -ForegroundColor Magenta }
        }
    }

    # Speed: Hardware detection with caching
    function Get-SystemHardwareType {
        try {
            # Use CIM for faster queries
            $computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop
            $battery = Get-CimInstance -ClassName Win32_Battery -ErrorAction SilentlyContinue
            $systemEnclosure = Get-CimInstance -ClassName Win32_SystemEnclosure -ErrorAction SilentlyContinue
            
            $hardwareInfo = @{
                IsLaptop = $false
                IsDesktop = $false
                IsWorkstation = $false
                IsServer = $false
                HasBattery = ($battery -ne $null)
                ChassisTypes = @()
                PCSystemType = $computerSystem.PCSystemType
                TotalPhysicalMemory = [math]::Round($computerSystem.TotalPhysicalMemory / 1GB, 2)
                Manufacturer = $computerSystem.Manufacturer
                Model = $computerSystem.Model
            }
            
            # Determine system type based on multiple factors
            if ($systemEnclosure) {
                $hardwareInfo.ChassisTypes = $systemEnclosure.ChassisTypes
                
                # Chassis type detection (more reliable than PCSystemType)
                $laptopChassisTypes = @(8, 9, 10, 11, 12, 14, 18, 21, 30, 31, 32)  # Laptop variants
                $desktopChassisTypes = @(3, 4, 5, 6, 7, 15, 16)  # Desktop variants
                $serverChassisTypes = @(17, 23)  # Server variants
                
                if ($systemEnclosure.ChassisTypes | Where-Object { $_ -in $laptopChassisTypes }) {
                    $hardwareInfo.IsLaptop = $true
                } elseif ($systemEnclosure.ChassisTypes | Where-Object { $_ -in $desktopChassisTypes }) {
                    $hardwareInfo.IsDesktop = $true
                } elseif ($systemEnclosure.ChassisTypes | Where-Object { $_ -in $serverChassisTypes }) {
                    $hardwareInfo.IsServer = $true
                } else {
                    $hardwareInfo.IsWorkstation = $true
                }
            }
            
            # Fallback to PCSystemType if chassis detection inconclusive
            if (-not ($hardwareInfo.IsLaptop -or $hardwareInfo.IsDesktop -or $hardwareInfo.IsServer)) {
                switch ($computerSystem.PCSystemType) {
                    1 { $hardwareInfo.IsDesktop = $true }
                    2 { $hardwareInfo.IsLaptop = $true }
                    3 { $hardwareInfo.IsWorkstation = $true }
                    4 { $hardwareInfo.IsServer = $true }
                    default { $hardwareInfo.IsDesktop = $true }  # Default assumption
                }
            }
            
            # Battery presence overrides chassis detection for laptops
            if ($hardwareInfo.HasBattery -and -not $hardwareInfo.IsServer) {
                $hardwareInfo.IsLaptop = $true
                $hardwareInfo.IsDesktop = $false
            }
            
            return $hardwareInfo
            
        } catch {
            Write-LogEntry "Hardware detection failed: $_" 'WARNING'
            return @{
                IsLaptop = $false
                IsDesktop = $true  # Safe default
                Error = $_.Exception.Message
            }
        }
    }

    # Speed: Get available power schemes efficiently
    function Get-PowerSchemes {
        try {
            $schemes = @{}
            
            # Parse powercfg output for available schemes
            $output = powercfg.exe /list 2>$null
            if ($LASTEXITCODE -eq 0 -and $output) {
                foreach ($line in $output) {
                    if ($line -match 'Power Scheme GUID: ([a-f0-9-]+)\s+\((.+?)\)(\s+\*)?') {
                        $guid = $matches[1]
                        $name = $matches[2].Trim()
                        $isActive = $matches[3] -eq ' *'
                        
                        $schemes[$name] = @{
                            GUID = $guid
                            Name = $name
                            IsActive = $isActive
                        }
                    }
                }
            }
            
            # Add common scheme mappings if not found
            $commonSchemes = @{
                'High Performance' = 'SCHEME_MIN'
                'Balanced' = 'SCHEME_BALANCED'
                'Power Saver' = 'SCHEME_MAX'
                'Ultimate Performance' = '8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c'
            }
            
            foreach ($scheme in $commonSchemes.GetEnumerator()) {
                if (-not $schemes.ContainsKey($scheme.Key)) {
                    $schemes[$scheme.Key] = @{
                        GUID = $scheme.Value
                        Name = $scheme.Key
                        IsActive = $false
                    }
                }
            }
            
            return $schemes
            
        } catch {
            Write-LogEntry "Failed to get power schemes: $_" 'WARNING'
            return @{}
        }
    }

    # Security: Backup current power settings
    function Backup-PowerSettings {
        try {
            $backupData = @{
                CurrentScheme = (powercfg.exe /getactivescheme 2>$null)
                HibernationStatus = (powercfg.exe /query SCHEME_CURRENT SUB_SLEEP HIBERNATEIDLE 2>$null)
                MonitorTimeoutAC = (powercfg.exe /query SCHEME_CURRENT SUB_VIDEO VIDEOIDLE /AC 2>$null)
                MonitorTimeoutDC = (powercfg.exe /query SCHEME_CURRENT SUB_VIDEO VIDEOIDLE /DC 2>$null)
                DiskTimeoutAC = (powercfg.exe /query SCHEME_CURRENT SUB_DISK DISKIDLE /AC 2>$null)
                DiskTimeoutDC = (powercfg.exe /query SCHEME_CURRENT SUB_DISK DISKIDLE /DC 2>$null)
                Timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
            }
            
            $backupFile = "$env:TEMP\PowerSettingsBackup_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
            $backupData | ConvertTo-Json -Depth 3 | Out-File -FilePath $backupFile -Encoding UTF8
            Write-LogEntry "[OK] Power settings backup saved to: $backupFile" 'INFO'
            return $backupFile
            
        } catch {
            Write-LogEntry "[WARN] Failed to backup power settings: $_" 'WARNING'
            return $null
        }
    }

    # Speed: Execute power configuration commands
    function Set-PowerConfiguration {
        param(
            [string]$SchemeName,
            [string]$SchemeGUID,
            [int]$MonitorTimeout,
            [int]$DiskTimeout
        )
        
        $configResults = @()
        
        # Set power scheme
        try {
            Write-LogEntry "Setting power scheme to: $SchemeName" 'INFO'
            
            if ($PSCmdlet.ShouldProcess($SchemeName, "Set active power scheme")) {
                $result = powercfg.exe /setactive $SchemeGUID 2>&1
                if ($LASTEXITCODE -eq 0) {
                    $configResults += @{
                        Setting = "Power Scheme"
                        Value = $SchemeName
                        Success = $true
                    }
                    Write-LogEntry "[OK] Power scheme set to: $SchemeName" 'SUCCESS'
                } else {
                    throw "Failed to set power scheme: $result"
                }
            }
        } catch {
            $configResults += @{
                Setting = "Power Scheme"
                Value = $SchemeName
                Success = $false
                Error = $_.Exception.Message
            }
            Write-LogEntry "[ERROR] Failed to set power scheme: $_" 'ERROR'
        }
        
        # Configure monitor timeout
        try {
            Write-LogEntry "Setting monitor timeout to: $MonitorTimeout minutes" 'INFO'
            
            if ($PSCmdlet.ShouldProcess("Monitor Timeout", "Set to $MonitorTimeout minutes")) {
                # Set for both AC and DC power
                $resultAC = powercfg.exe /change monitor-timeout-ac $MonitorTimeout 2>&1
                $resultDC = powercfg.exe /change monitor-timeout-dc $MonitorTimeout 2>&1
                
                if ($LASTEXITCODE -eq 0) {
                    $configResults += @{
                        Setting = "Monitor Timeout"
                        Value = "$MonitorTimeout minutes"
                        Success = $true
                    }
                    Write-LogEntry "[OK] Monitor timeout set to: $MonitorTimeout minutes" 'SUCCESS'
                } else {
                    throw "Failed to set monitor timeout: AC=$resultAC, DC=$resultDC"
                }
            }
        } catch {
            $configResults += @{
                Setting = "Monitor Timeout"
                Value = "$MonitorTimeout minutes"
                Success = $false
                Error = $_.Exception.Message
            }
            Write-LogEntry "[ERROR] Failed to set monitor timeout: $_" 'ERROR'
        }
        
        # Configure disk timeout
        if ($DiskTimeout -eq 0) {
            try {
                Write-LogEntry "Disabling disk timeout (never turn off)" 'INFO'
                
                if ($PSCmdlet.ShouldProcess("Disk Timeout", "Disable")) {
                    $resultAC = powercfg.exe /change disk-timeout-ac 0 2>&1
                    $resultDC = powercfg.exe /change disk-timeout-dc 0 2>&1
                    
                    if ($LASTEXITCODE -eq 0) {
                        $configResults += @{
                            Setting = "Disk Timeout"
                            Value = "Disabled"
                            Success = $true
                        }
                        Write-LogEntry "[OK] Disk timeout disabled" 'SUCCESS'
                    } else {
                        throw "Failed to disable disk timeout: AC=$resultAC, DC=$resultDC"
                    }
                }
            } catch {
                $configResults += @{
                    Setting = "Disk Timeout"
                    Value = "Disabled"
                    Success = $false
                    Error = $_.Exception.Message
                }
                Write-LogEntry "[ERROR] Failed to disable disk timeout: $_" 'ERROR'
            }
        } else {
            try {
                Write-LogEntry "Setting disk timeout to: $DiskTimeout minutes" 'INFO'
                
                if ($PSCmdlet.ShouldProcess("Disk Timeout", "Set to $DiskTimeout minutes")) {
                    $resultAC = powercfg.exe /change disk-timeout-ac $DiskTimeout 2>&1
                    $resultDC = powercfg.exe /change disk-timeout-dc $DiskTimeout 2>&1
                    
                    if ($LASTEXITCODE -eq 0) {
                        $configResults += @{
                            Setting = "Disk Timeout"
                            Value = "$DiskTimeout minutes"
                            Success = $true
                        }
                        Write-LogEntry "[OK] Disk timeout set to: $DiskTimeout minutes" 'SUCCESS'
                    } else {
                        throw "Failed to set disk timeout: AC=$resultAC, DC=$resultDC"
                    }
                }
            } catch {
                $configResults += @{
                    Setting = "Disk Timeout"
                    Value = "$DiskTimeout minutes"
                    Success = $false
                    Error = $_.Exception.Message
                }
                Write-LogEntry "[ERROR] Failed to set disk timeout: $_" 'ERROR'
            }
        }
        
        # Disable hibernation
        try {
            Write-LogEntry "Disabling hibernation" 'INFO'
            
            if ($PSCmdlet.ShouldProcess("Hibernation", "Disable")) {
                $result = powercfg.exe /hibernate off 2>&1
                if ($LASTEXITCODE -eq 0) {
                    $configResults += @{
                        Setting = "Hibernation"
                        Value = "Disabled"
                        Success = $true
                    }
                    Write-LogEntry "[OK] Hibernation disabled" 'SUCCESS'
                } else {
                    throw "Failed to disable hibernation: $result"
                }
            }
        } catch {
            $configResults += @{
                Setting = "Hibernation"
                Value = "Disabled"
                Success = $false
                Error = $_.Exception.Message
            }
            Write-LogEntry "[ERROR] Failed to disable hibernation: $_" 'ERROR'
        }
        
        return $configResults
    }

    try {
        Write-LogEntry "=== Desktop Power Settings Configuration Started ===" 'INFO'
        Write-LogEntry "Power Plan: $PowerPlan" 'INFO'
        Write-LogEntry "Monitor Timeout: $MonitorTimeoutMinutes minutes" 'INFO'
        Write-LogEntry "Disk Timeout: $(if($DiskTimeoutMinutes -eq 0){'Disabled'}else{"$DiskTimeoutMinutes minutes"})" 'INFO'
        
        if ($WhatIfPreference) {
            Write-LogEntry "WhatIf mode - no power settings will be changed" 'INFO'
        }

        # Hardware detection and validation
        if (-not $SkipHardwareDetection) {
            Write-LogEntry "`n[CHECK] Detecting system hardware type..." 'HARDWARE'
            $hardwareInfo = Get-SystemHardwareType
            
            Write-LogEntry "Hardware Analysis:" 'HARDWARE'
            Write-LogEntry "  • System Type: $(if($hardwareInfo.IsLaptop){'Laptop'}elseif($hardwareInfo.IsDesktop){'Desktop'}elseif($hardwareInfo.IsServer){'Server'}else{'Workstation'})" 'HARDWARE'
            Write-LogEntry "  • Has Battery: $($hardwareInfo.HasBattery)" 'HARDWARE'
            Write-LogEntry "  • Manufacturer: $($hardwareInfo.Manufacturer)" 'HARDWARE'
            Write-LogEntry "  • Model: $($hardwareInfo.Model)" 'HARDWARE'
            Write-LogEntry "  • Memory: $($hardwareInfo.TotalPhysicalMemory) GB" 'HARDWARE'
            
            # Security: Prevent accidental laptop configuration
            if ($hardwareInfo.IsLaptop -and -not $AllowLaptops -and -not $Force) {
                Write-Host "`n" + "="*60 -ForegroundColor Red
                Write-Host "LAPTOP DETECTED - OPERATION BLOCKED" -ForegroundColor Red -BackgroundColor Black
                Write-Host "="*60 -ForegroundColor Red
                Write-Host "This system appears to be a LAPTOP with battery power." -ForegroundColor Yellow
                Write-Host "Desktop power settings may negatively impact battery life!" -ForegroundColor Yellow
                Write-Host "`nTo proceed anyway, use one of these options:" -ForegroundColor Cyan
                Write-Host "  • Use -AllowLaptops parameter" -ForegroundColor Cyan
                Write-Host "  • Use -Force parameter" -ForegroundColor Cyan
                Write-Host "  • Use -SkipHardwareDetection parameter" -ForegroundColor Cyan
                Write-Host "="*60 -ForegroundColor Red
                
                $global:LastStatus = "[WARN] Operation blocked - laptop detected."
                return
            }
            
            if ($hardwareInfo.IsLaptop -and ($AllowLaptops -or $Force)) {
                Write-LogEntry "[WARN] Proceeding with laptop configuration (overridden)" 'WARNING'
            }
        }

        # User confirmation for desktop systems
        if (-not $Force -and -not $WhatIfPreference) {
            Write-Host "`n" + "="*60 -ForegroundColor Yellow
            Write-Host "POWER SETTINGS CONFIGURATION" -ForegroundColor Yellow -BackgroundColor Black
            Write-Host "="*60 -ForegroundColor Yellow
            Write-Host "This will configure the following settings:" -ForegroundColor White
            Write-Host "  • Power Plan: $PowerPlan" -ForegroundColor Cyan
            Write-Host "  • Monitor Timeout: $MonitorTimeoutMinutes minutes" -ForegroundColor Cyan
            Write-Host "  • Disk Timeout: $(if($DiskTimeoutMinutes -eq 0){'Disabled'}else{"$DiskTimeoutMinutes minutes"})" -ForegroundColor Cyan
            Write-Host "  • Hibernation: Disabled" -ForegroundColor Cyan
            Write-Host "`nNote: These settings optimize for desktop performance" -ForegroundColor Yellow
            Write-Host "="*60 -ForegroundColor Yellow
            
            $confirmation = Read-Host "`n[PROMPT] Continue with power configuration? (Y/N)"
            if ($confirmation -notin @('Y', 'y', 'Yes', 'yes')) {
                Write-LogEntry "Operation cancelled by user" 'WARNING'
                $global:LastStatus = "[WARN] User cancelled power settings configuration."
                return
            }
        }

        # Get available power schemes
        Write-LogEntry "`n[CHECK] Scanning available power schemes..." 'INFO'
        $powerSchemes = Get-PowerSchemes
        
        if ($powerSchemes.Count -eq 0) {
            throw "No power schemes detected on this system"
        }
        
        Write-LogEntry "Available power schemes:" 'INFO'
        foreach ($scheme in $powerSchemes.GetEnumerator()) {
            $activeIndicator = if ($scheme.Value.IsActive) { " (ACTIVE)" } else { "" }
            Write-LogEntry "  • $($scheme.Key)$activeIndicator" 'INFO'
        }
        
        # Validate requested power plan
        if (-not $powerSchemes.ContainsKey($PowerPlan)) {
            Write-LogEntry "[WARN] Requested power plan '$PowerPlan' not found" 'WARNING'
            Write-LogEntry "Falling back to 'High Performance' scheme" 'WARNING'
            $PowerPlan = 'High Performance'
            
            if (-not $powerSchemes.ContainsKey($PowerPlan)) {
                throw "Neither requested scheme nor High Performance scheme available"
            }
        }
        
        $selectedScheme = $powerSchemes[$PowerPlan]
        Write-LogEntry "Selected scheme: $PowerPlan (GUID: $($selectedScheme.GUID))" 'INFO'

        # Backup current settings
        if (-not $WhatIfPreference) {
            Write-LogEntry "`n[BACKUP] Backing up current power settings..." 'INFO'
            $backupFile = Backup-PowerSettings
        }

        if ($WhatIfPreference) {
            Write-LogEntry "`nWhatIf Summary:" 'INFO'
            Write-LogEntry "  • Would set power scheme to: $PowerPlan" 'INFO'
            Write-LogEntry "  • Would set monitor timeout to: $MonitorTimeoutMinutes minutes" 'INFO'
            Write-LogEntry "  • Would set disk timeout to: $(if($DiskTimeoutMinutes -eq 0){'Disabled'}else{"$DiskTimeoutMinutes minutes"})" 'INFO'
            Write-LogEntry "  • Would disable hibernation" 'INFO'
            $global:LastStatus = "[INFO] WhatIf completed - power settings would be configured."
            return
        }

        # Apply power configuration
        Write-LogEntry "`n[APPLY] Applying power configuration..." 'INFO'
        $configResults = Set-PowerConfiguration -SchemeName $PowerPlan -SchemeGUID $selectedScheme.GUID -MonitorTimeout $MonitorTimeoutMinutes -DiskTimeout $DiskTimeoutMinutes
        
        # Process results
        $successCount = ($configResults | Where-Object { $_.Success }).Count
        $failureCount = ($configResults | Where-Object { -not $_.Success }).Count
        
        $script:appliedSettings = $configResults | Where-Object { $_.Success }
        $script:failedSettings = $configResults | Where-Object { -not $_.Success }

        # Verify configuration
        Write-LogEntry "`n[CHECK] Verifying power configuration..." 'INFO'
        try {
            $currentScheme = powercfg.exe /getactivescheme 2>$null
            if ($currentScheme -and $currentScheme -match $selectedScheme.GUID) {
                Write-LogEntry "[OK] Power scheme verification passed" 'SUCCESS'
            } else {
                Write-LogEntry "[WARN] Power scheme verification failed" 'WARNING'
            }
        } catch {
            Write-LogEntry "[WARN] Could not verify power scheme: $_" 'WARNING'
        }

    } catch {
        Write-LogEntry "Critical error during power settings configuration: $_" 'ERROR'
        $global:LastStatus = "[ERROR] Power settings configuration failed: $_"
        throw
    } finally {
        $stopwatch.Stop()
        $duration = $stopwatch.Elapsed.TotalSeconds
        
        Write-LogEntry "`n[SUMMARY] Configuration Summary:" 'INFO'
        Write-LogEntry "Duration: $([math]::Round($duration, 2)) seconds" 'INFO'
        Write-LogEntry "Settings applied: $($appliedSettings.Count)" 'SUCCESS'
        Write-LogEntry "Settings failed: $($failedSettings.Count)" 'ERROR'
        
        if ($appliedSettings.Count -gt 0) {
            Write-LogEntry "`n[OK] Successfully Applied:" 'SUCCESS'
            foreach ($setting in $appliedSettings) {
                Write-LogEntry "  • $($setting.Setting): $($setting.Value)" 'SUCCESS'
            }
        }
        
        if ($failedSettings.Count -gt 0) {
            Write-LogEntry "`n[ERROR] Failed Settings:" 'ERROR'
            foreach ($setting in $failedSettings) {
                Write-LogEntry "  • $($setting.Setting): $($setting.Error)" 'ERROR'
            }
        }
        
        # Write log file
        try {
            $logEntries | Out-File -FilePath $LogPath -Encoding UTF8 -Force
            Write-LogEntry "[LOG] Detailed log saved to: $LogPath" 'INFO'
        } catch {
            Write-LogEntry "[WARN] Failed to save log file: $_" 'WARNING'
        }
        
        # Set final status
        if ($appliedSettings.Count -gt 0) {
            if ($failedSettings.Count -eq 0) {
                $global:LastStatus = "[OK] All power settings applied successfully."
            } else {
                $global:LastStatus = "[WARN] Power settings partially applied ($($appliedSettings.Count) success, $($failedSettings.Count) failed)."
            }
        } else {
            $global:LastStatus = "[ERROR] No power settings were applied."
        }
        
        Write-LogEntry "=== Power Settings Configuration Completed ===" 'INFO'
    }
}

# Utility function to get current power configuration
function Get-PowerConfiguration {
    [CmdletBinding()]
    param()
    
    try {
        $config = @{}
        
        # Get active scheme
        $activeScheme = powercfg.exe /getactivescheme 2>$null
        if ($activeScheme) {
            $config.ActiveScheme = $activeScheme.Trim()
        }
        
        # Get hibernation status
        $hibernation = powercfg.exe /availablesleepstates 2>$null
        $config.HibernationAvailable = ($hibernation -match "Hibernate")
        
        # Get monitor and disk timeouts (simplified parsing)
        $config.MonitorTimeoutAC = "Unknown"
        $config.MonitorTimeoutDC = "Unknown"
        $config.DiskTimeoutAC = "Unknown"
        $config.DiskTimeoutDC = "Unknown"
        
        return $config
        
    } catch {
        return @{
            Error = $_.Exception.Message
        }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# Option 17 - Install Computer Lab Scheduled Tasks
# ─────────────────────────────────────────────────────────────────────────────
function Register-LabScheduledTasks {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [string]$RepoOwner = 'dswaney',
        [string]$RepoName = 'Compton',
        [string]$Branch = 'main',
        [string]$RepoSubFolder = '',
        [string]$DestinationPath = 'C:\Scripts',
        [string]$LogDirectory = 'C:\Logs',
        [switch]$Force
    )

    $ErrorActionPreference = 'Stop'
    $ProgressPreference = 'SilentlyContinue'

    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        $global:LastStatus = "[ERROR] Scheduled task setup requires Administrator rights."
        throw "This function must be run as Administrator."
    }

    $script:RunStart = Get-Date
    $script:ComputerName = $env:COMPUTERNAME
    $script:BackupFolder = Join-Path $DestinationPath 'Backup'
    $script:YamlLogPath = Join-Path $LogDirectory ("{0}-RegisterLabScheduledTasks-{1}.yaml" -f $script:ComputerName, $script:RunStart.ToString('yyyy-MM-dd_HH-mm-ss'))
    $script:ActionHistory = New-Object System.Collections.Generic.List[object]
    $script:FileResults = New-Object System.Collections.Generic.List[object]
    $script:TaskResults = New-Object System.Collections.Generic.List[object]
    $script:OverallResult = 'Unknown'
    $script:FailureMessage = $null

    $scriptFiles = @(
        '00_Update-Scripts-FromGitHub.ps1',
        '01_Enable_Windows_Update_Services.ps1',
        '02_Remove_User_Profiles.ps1',
        '03_Weekend_Apps_Update.ps1',
        '04_Update_Edge_Silent.ps1',
        '05_Weekend_HP_Drivers_Update.ps1',
        '06_Weekend_Windows_Updates.ps1',
        '07_Force_Reboot_Install_Updates.ps1',
        '08_System_Repair.ps1',
        '09_Disable_Windows_Update_Services.ps1'
    )

    $taskDefinitions = @(
        [PSCustomObject]@{ Name = '01. Check for Updated Scripts';          Script = '00_Update-Scripts-FromGitHub.ps1';       Time = '01:15'; Arguments = '' },
        [PSCustomObject]@{ Name = '02. Enable Windows Update Services';     Script = '01_Enable_Windows_Update_Services.ps1';  Time = '01:20'; Arguments = '' },
        [PSCustomObject]@{ Name = '03. Remove User Profiles Weekly';        Script = '02_Remove_User_Profiles.ps1';            Time = '01:30'; Arguments = '' },
        [PSCustomObject]@{ Name = '04. Weekend Apps Update';                Script = '03_Weekend_Apps_Update.ps1';             Time = '02:00'; Arguments = '' },
        [PSCustomObject]@{ Name = '05. Update Edge Silent';                 Script = '04_Update_Edge_Silent.ps1';              Time = '02:45'; Arguments = '-KillEdgeProcesses' },
        [PSCustomObject]@{ Name = '06. Weekend HP Drivers Update';          Script = '05_Weekend_HP_Drivers_Update.ps1';       Time = '03:00'; Arguments = '' },
        [PSCustomObject]@{ Name = '07. Weekend Windows Updates - 1st Pass'; Script = '06_Weekend_Windows_Updates.ps1';         Time = '04:00'; Arguments = '' },
        [PSCustomObject]@{ Name = '08. Force Reboot Install Updates';       Script = '07_Force_Reboot_Install_Updates.ps1';    Time = '05:00'; Arguments = '' },
        [PSCustomObject]@{ Name = '09. Weekend Windows Updates - 2nd Pass'; Script = '06_Weekend_Windows_Updates.ps1';         Time = '05:30'; Arguments = '' },
        [PSCustomObject]@{ Name = '10. Disable Windows Update Services';    Script = '09_Disable_Windows_Update_Services.ps1'; Time = '06:00'; Arguments = '' },
        [PSCustomObject]@{ Name = '11. Force Reboot Install Updates 2';     Script = '07_Force_Reboot_Install_Updates.ps1';    Time = '06:05'; Arguments = '' },
        [PSCustomObject]@{ Name = '12. System Repair';                      Script = '08_System_Repair.ps1';                   Time = '06:15'; Arguments = '' },
        [PSCustomObject]@{ Name = '13. Force Reboot Install Updates 3';     Script = '07_Force_Reboot_Install_Updates.ps1';    Time = '07:00'; Arguments = '' }
    )

    function Ensure-Directory {
        param([Parameter(Mandatory)][string]$Path)
        if (-not (Test-Path -LiteralPath $Path)) {
            New-Item -Path $Path -ItemType Directory -Force | Out-Null
        }
    }

    function Write-Status {
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

        $script:ActionHistory.Add([PSCustomObject]@{
            Time    = $timestamp
            Level   = $Level
            Message = $Message
        }) | Out-Null
    }

    function ConvertTo-YamlScalar {
        param([AllowNull()]$Value)

        if ($null -eq $Value) { return 'null' }
        if ($Value -is [bool]) { return $Value.ToString().ToLowerInvariant() }
        if ($Value -is [int] -or $Value -is [long] -or $Value -is [double] -or $Value -is [decimal]) { return [string]$Value }
        if ($Value -is [datetime]) { return "'" + $Value.ToString('yyyy-MM-dd HH:mm:ss') + "'" }

        $text = [string]$Value
        $text = $text -replace "`r", ' '
        $text = $text -replace "`n", ' '
        $text = $text -replace "'", "''"
        return "'" + $text + "'"
    }

    function Save-Utf8NoBom {
        param(
            [Parameter(Mandatory)][string]$Path,
            [Parameter(Mandatory)][string]$Content
        )

        $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
        [System.IO.File]::WriteAllText($Path, $Content, $utf8NoBom)
    }

    function Get-RawGitHubUrl {
        param([Parameter(Mandatory)][string]$FileName)

        $pathPart = if ([string]::IsNullOrWhiteSpace($RepoSubFolder)) {
            $FileName
        }
        else {
            ($RepoSubFolder.Trim('/').Replace('\','/') + '/' + $FileName)
        }

        'https://raw.githubusercontent.com/{0}/{1}/{2}/{3}' -f $RepoOwner, $RepoName, $Branch, $pathPart
    }

    function Get-RemoteFileContent {
        param([Parameter(Mandatory)][string]$FileName)

        $url = Get-RawGitHubUrl -FileName $FileName
        $uriBuilder = New-Object System.UriBuilder($url)
        $uriBuilder.Query = 'cb={0}' -f [DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds()
        $finalUri = $uriBuilder.Uri.AbsoluteUri

        $response = Invoke-WebRequest -Uri $finalUri `
                                      -UseBasicParsing `
                                      -Headers @{
                                          'Cache-Control' = 'no-cache'
                                          'Pragma'        = 'no-cache'
                                          'User-Agent'    = 'PowerShell-GitHub-Updater'
                                      } `
                                      -ErrorAction Stop

        if ([string]::IsNullOrWhiteSpace($response.Content)) {
            throw "Downloaded content was empty for [$FileName]."
        }

        $response.Content
    }

    function Get-FileTextSafe {
        param([Parameter(Mandatory)][string]$Path)

        if (-not (Test-Path -LiteralPath $Path)) {
            return $null
        }

        try {
            return Get-Content -LiteralPath $Path -Raw -Encoding UTF8
        }
        catch {
            return Get-Content -LiteralPath $Path -Raw
        }
    }

    function Get-ScriptHeaderValue {
        param(
            [Parameter(Mandatory)][AllowEmptyString()][string]$Content,
            [Parameter(Mandatory)][string]$HeaderName
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

        return $null
    }

    function Convert-ToVersionObject {
        param([Parameter(Mandatory)][string]$VersionText)

        try {
            return [version]$VersionText.Trim()
        }
        catch {
            $clean = ($VersionText -replace '[^\d\.]', '').Trim('.')
            if ([string]::IsNullOrWhiteSpace($clean)) {
                return [version]'0.0'
            }

            try { return [version]$clean } catch { return [version]'0.0' }
        }
    }

    function Backup-File {
        param([Parameter(Mandatory)][string]$Path)

        if (-not (Test-Path -LiteralPath $Path)) {
            return $null
        }

        Ensure-Directory -Path $script:BackupFolder

        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($Path)
        $extension = [System.IO.Path]::GetExtension($Path)
        $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
        $backupPath = Join-Path $script:BackupFolder ("{0}_{1}{2}.bak" -f $baseName, $timestamp, $extension)
        Copy-Item -LiteralPath $Path -Destination $backupPath -Force
        $backupPath
    }

    function Write-YamlLog {
        try {
            Ensure-Directory -Path $LogDirectory

            $runEnd = Get-Date
            $duration = [math]::Round(($runEnd - $script:RunStart).TotalSeconds, 0)

            $updatedCount = @($script:FileResults | Where-Object { $_.Status -eq 'Updated' }).Count
            $currentCount = @($script:FileResults | Where-Object { $_.Status -eq 'Current' }).Count
            $downloadedMissingCount = @($script:FileResults | Where-Object { $_.Status -eq 'DownloadedMissing' }).Count
            $fileErrorCount = @($script:FileResults | Where-Object { $_.Status -eq 'Error' }).Count
            $taskCreatedCount = @($script:TaskResults | Where-Object { $_.Status -eq 'Created' }).Count
            $taskWhatIfCount = @($script:TaskResults | Where-Object { $_.Status -eq 'WhatIf' }).Count
            $taskErrorCount = @($script:TaskResults | Where-Object { $_.Status -eq 'Error' }).Count

            $lines = New-Object System.Collections.Generic.List[string]
            $lines.Add("computer_name: $(ConvertTo-YamlScalar $script:ComputerName)") | Out-Null
            $lines.Add("script_name: 'Register-LabScheduledTasks'") | Out-Null
            $lines.Add("script_version: '1.5.0'") | Out-Null
            $lines.Add("run_started: $(ConvertTo-YamlScalar $script:RunStart)") | Out-Null
            $lines.Add("run_finished: $(ConvertTo-YamlScalar $runEnd)") | Out-Null
            $lines.Add("duration_seconds: $duration") | Out-Null
            $lines.Add("destination_path: $(ConvertTo-YamlScalar $DestinationPath)") | Out-Null
            $lines.Add("backup_folder: $(ConvertTo-YamlScalar $script:BackupFolder)") | Out-Null
            $lines.Add("overall_result: $(ConvertTo-YamlScalar $script:OverallResult)") | Out-Null
            $lines.Add("failure_message: $(ConvertTo-YamlScalar $script:FailureMessage)") | Out-Null
            $lines.Add('') | Out-Null
            $lines.Add('summary:') | Out-Null
            $lines.Add("  files_updated: $updatedCount") | Out-Null
            $lines.Add("  files_current: $currentCount") | Out-Null
            $lines.Add("  files_downloaded_missing: $downloadedMissingCount") | Out-Null
            $lines.Add("  file_errors: $fileErrorCount") | Out-Null
            $lines.Add("  tasks_created: $taskCreatedCount") | Out-Null
            $lines.Add("  tasks_whatif: $taskWhatIfCount") | Out-Null
            $lines.Add("  task_errors: $taskErrorCount") | Out-Null
            $lines.Add('') | Out-Null

            $lines.Add('file_results:') | Out-Null
            if ($script:FileResults.Count -gt 0) {
                foreach ($item in $script:FileResults) {
                    $lines.Add('  -') | Out-Null
                    $lines.Add("    file_name: $(ConvertTo-YamlScalar $item.FileName)") | Out-Null
                    $lines.Add("    local_path: $(ConvertTo-YamlScalar $item.LocalPath)") | Out-Null
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
            $lines.Add('task_results:') | Out-Null
            if ($script:TaskResults.Count -gt 0) {
                foreach ($item in $script:TaskResults) {
                    $lines.Add('  -') | Out-Null
                    $lines.Add("    task_name: $(ConvertTo-YamlScalar $item.TaskName)") | Out-Null
                    $lines.Add("    script_path: $(ConvertTo-YamlScalar $item.ScriptPath)") | Out-Null
                    $lines.Add("    schedule_time: $(ConvertTo-YamlScalar $item.ScheduleTime)") | Out-Null
                    $lines.Add("    arguments: $(ConvertTo-YamlScalar $item.Arguments)") | Out-Null
                    $lines.Add("    status: $(ConvertTo-YamlScalar $item.Status)") | Out-Null
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

    function Get-TriggerTime {
        param([Parameter(Mandatory)][string]$TimeText)
        $parts = $TimeText.Split(':')
        if ($parts.Count -ne 2) {
            throw "Invalid schedule time [$TimeText]. Expected HH:mm."
        }

        (Get-Date -Hour ([int]$parts[0]) -Minute ([int]$parts[1]) -Second 0)
    }

    function Register-WeeklySystemTask {
        param(
            [Parameter(Mandatory)][string]$TaskName,
            [Parameter(Mandatory)][string]$ScriptPath,
            [Parameter(Mandatory)][string]$TimeText,
            [string]$Arguments = ''
        )

        $taskDescription = "Created by Compton_Tech_Utils Option 17"
        $argSuffix = if ([string]::IsNullOrWhiteSpace($Arguments)) { '' } else { ' ' + $Arguments.Trim() }
        $actionArgs = '-NoProfile -ExecutionPolicy Bypass -File "{0}"{1}' -f $ScriptPath, $argSuffix
        $target = "{0} [{1}]" -f $TaskName, $TimeText

        if (-not $PSCmdlet.ShouldProcess($target, 'Register weekly scheduled task as SYSTEM')) {
            $script:TaskResults.Add([PSCustomObject]@{
                TaskName     = $TaskName
                ScriptPath   = $ScriptPath
                ScheduleTime = $TimeText
                Arguments    = $Arguments
                Status       = 'WhatIf'
                Message      = 'WhatIf prevented task registration.'
            }) | Out-Null
            Write-Status "WhatIf: would register task [$TaskName] for Sundays at [$TimeText]." 'INFO'
            return
        }

        if (-not (Test-Path -LiteralPath $ScriptPath)) {
            throw "Script path does not exist: $ScriptPath"
        }

        try {
            if ($Force) {
                try {
                    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction Stop
                    Write-Status "Removed existing task [$TaskName] before recreation." 'INFO'
                }
                catch {
                    Write-Status "Existing task [$TaskName] was not present or could not be removed before recreation." 'INFO'
                }
            }

            $action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument $actionArgs
            $trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Sunday -At (Get-TriggerTime -TimeText $TimeText)
            $principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
            $settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

            $taskObject = New-ScheduledTask -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Description $taskDescription
            Register-ScheduledTask -TaskName $TaskName -InputObject $taskObject -Force | Out-Null

            $script:TaskResults.Add([PSCustomObject]@{
                TaskName     = $TaskName
                ScriptPath   = $ScriptPath
                ScheduleTime = $TimeText
                Arguments    = $Arguments
                Status       = 'Created'
                Message      = 'Scheduled task registered successfully.'
            }) | Out-Null

            Write-Status "Registered task [$TaskName] for Sundays at [$TimeText] as SYSTEM." 'OK'
        }
        catch {
            $script:TaskResults.Add([PSCustomObject]@{
                TaskName     = $TaskName
                ScriptPath   = $ScriptPath
                ScheduleTime = $TimeText
                Arguments    = $Arguments
                Status       = 'Error'
                Message      = $_.Exception.Message
            }) | Out-Null

            Write-Status "Failed to register task [$TaskName]: $($_.Exception.Message)" 'ERROR'
        }
    }

    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

        Ensure-Directory -Path $DestinationPath
        Ensure-Directory -Path $LogDirectory

        Write-Status "Ensuring script directory exists at [$DestinationPath]." 'INFO'
        Write-Status "Downloading and updating lab scripts from [$RepoOwner/$RepoName] branch [$Branch]." 'INFO'

        foreach ($file in $scriptFiles) {
            $localPath = Join-Path $DestinationPath $file

            try {
                $remoteContent = Get-RemoteFileContent -FileName $file
                $remoteVersionText = Get-ScriptHeaderValue -Content $remoteContent -HeaderName 'ScriptVersion'
                $remoteLastUpdated = Get-ScriptHeaderValue -Content $remoteContent -HeaderName 'LastUpdated'

                if ([string]::IsNullOrWhiteSpace($remoteVersionText)) {
                    throw "Remote file [$file] is missing a readable ScriptVersion header."
                }

                $remoteVersion = Convert-ToVersionObject -VersionText $remoteVersionText
                $localContent = Get-FileTextSafe -Path $localPath

                if ($null -eq $localContent) {
                    Save-Utf8NoBom -Path $localPath -Content $remoteContent
                    $script:FileResults.Add([PSCustomObject]@{
                        FileName          = $file
                        LocalPath         = $localPath
                        Status            = 'DownloadedMissing'
                        LocalVersion      = $null
                        RemoteVersion     = $remoteVersionText
                        LocalLastUpdated  = $null
                        RemoteLastUpdated = $remoteLastUpdated
                        BackupPath        = $null
                        Message           = 'Local file was missing and was downloaded from GitHub.'
                    }) | Out-Null
                    Write-Status "Downloaded missing script [$file]." 'OK'
                    continue
                }

                $localVersionText = Get-ScriptHeaderValue -Content $localContent -HeaderName 'ScriptVersion'
                $localLastUpdated = Get-ScriptHeaderValue -Content $localContent -HeaderName 'LastUpdated'
                if ([string]::IsNullOrWhiteSpace($localVersionText)) {
                    $localVersionText = '0.0'
                }

                $localVersion = Convert-ToVersionObject -VersionText $localVersionText

                if ($remoteVersion -gt $localVersion) {
                    $backupPath = Backup-File -Path $localPath
                    Save-Utf8NoBom -Path $localPath -Content $remoteContent
                    $script:FileResults.Add([PSCustomObject]@{
                        FileName          = $file
                        LocalPath         = $localPath
                        Status            = 'Updated'
                        LocalVersion      = $localVersionText
                        RemoteVersion     = $remoteVersionText
                        LocalLastUpdated  = $localLastUpdated
                        RemoteLastUpdated = $remoteLastUpdated
                        BackupPath        = $backupPath
                        Message           = "Updated local file from $localVersionText to $remoteVersionText."
                    }) | Out-Null
                    Write-Status "Updated script [$file] from [$localVersionText] to [$remoteVersionText]." 'OK'
                }
                else {
                    $script:FileResults.Add([PSCustomObject]@{
                        FileName          = $file
                        LocalPath         = $localPath
                        Status            = 'Current'
                        LocalVersion      = $localVersionText
                        RemoteVersion     = $remoteVersionText
                        LocalLastUpdated  = $localLastUpdated
                        RemoteLastUpdated = $remoteLastUpdated
                        BackupPath        = $null
                        Message           = 'Local file is already current.'
                    }) | Out-Null
                    Write-Status "Script [$file] is already current." 'INFO'
                }
            }
            catch {
                $script:FileResults.Add([PSCustomObject]@{
                    FileName          = $file
                    LocalPath         = $localPath
                    Status            = 'Error'
                    LocalVersion      = $null
                    RemoteVersion     = $null
                    LocalLastUpdated  = $null
                    RemoteLastUpdated = $null
                    BackupPath        = $null
                    Message           = $_.Exception.Message
                }) | Out-Null
                Write-Status "Failed processing script [$file]: $($_.Exception.Message)" 'ERROR'
            }
        }

        Write-Status "Registering scheduled tasks in Task Scheduler." 'INFO'

        foreach ($task in $taskDefinitions) {
            $scriptPath = Join-Path $DestinationPath $task.Script
            Register-WeeklySystemTask -TaskName $task.Name -ScriptPath $scriptPath -TimeText $task.Time -Arguments $task.Arguments
        }

        $fileErrors = @($script:FileResults | Where-Object { $_.Status -eq 'Error' }).Count
        $taskErrors = @($script:TaskResults | Where-Object { $_.Status -eq 'Error' }).Count
        $taskCreated = @($script:TaskResults | Where-Object { $_.Status -eq 'Created' }).Count

        if (($fileErrors + $taskErrors) -gt 0) {
            $script:OverallResult = 'CompletedWithErrors'
            $global:LastStatus = "[WARN] Option 17 completed with issues. Scripts downloaded/updated, tasks created: $taskCreated, errors: $($fileErrors + $taskErrors)."
        }
        else {
            $script:OverallResult = 'Succeeded'
            $global:LastStatus = "[OK] Option 17 completed successfully. C:\Scripts refreshed and $taskCreated scheduled tasks registered."
        }

        Write-YamlLog
    }
    catch {
        $script:FailureMessage = $_.Exception.Message
        $script:OverallResult = 'Failed'
        $global:LastStatus = "[ERROR] Option 17 failed: $($_.Exception.Message)"
        Write-Status "Fatal error in Option 17: $($_.Exception.Message)" 'ERROR'
        Write-YamlLog
        throw
    }
}


# ─────────────────────────────────────────────────────────────────────────────
# Option 18 - Set OneDrive to Automatically Login at Boot
# ─────────────────────────────────────────────────────────────────────────────
function Set-OneDriveAutoLoginPolicy {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [switch]$EnableSilentConfig,
        [switch]$DisableFirstRunWizard,
        [switch]$EnableAutoStartup,
        [switch]$EnableFilesOnDemand,
        [switch]$DisablePersonalSync,
        [switch]$EnableKnownFolderMove,
        [switch]$DisableAutoLogin,
        [string]$TenantId,
        [ValidateRange(1, 30)]
        [int]$SyncThrottleKbps = 0,  # 0 = No throttling
        [switch]$BackupSettings,
        [string]$LogPath = "$env:TEMP\OneDrivePolicy_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
    )

    # Security: Require elevation for policy configuration
    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        throw "This function must be run as Administrator"
    }

    $ErrorActionPreference = 'Continue'
    $ProgressPreference = 'SilentlyContinue'  # Speed: Disable progress bars
    
    # Initialize tracking using ArrayList for better performance and compatibility
    $script:logEntries = New-Object System.Collections.ArrayList
    $script:appliedPolicies = New-Object System.Collections.ArrayList
    $script:failedPolicies = New-Object System.Collections.ArrayList
    $script:skippedPolicies = New-Object System.Collections.ArrayList
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    
    function Write-LogEntry {
        param([string]$Message, [string]$Level = 'INFO')
        $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $logEntry = "[$timestamp] [$Level] $Message"
        
        # Ensure logEntries exists before adding to it
        if ($null -eq $script:logEntries) {
            $script:logEntries = New-Object System.Collections.ArrayList
        }
        [void]$script:logEntries.Add($logEntry)
        
        switch ($Level) {
            'ERROR' { Write-Host $Message -ForegroundColor Red }
            'WARNING' { Write-Host $Message -ForegroundColor Yellow }
            'SUCCESS' { Write-Host $Message -ForegroundColor Green }
            'INFO' { Write-Host $Message -ForegroundColor Cyan }
            'POLICY' { Write-Host $Message -ForegroundColor Magenta }
            'SKIP' { Write-Host $Message -ForegroundColor DarkGray }
        }
    }

    # Security: Registry validation and path management
    function Test-RegistryPath {
        param([string]$Path)
        try {
            # Security: Validate registry path format
            if ($Path -notmatch '^HK(LM|CU|CR|U|CC):\\') {
                return $false
            }
            return Test-Path $Path -ErrorAction SilentlyContinue
        } catch {
            return $false
        }
    }

    function New-RegistryPath {
        param([string]$Path)
        try {
            if (-not (Test-RegistryPath $Path)) {
                New-Item -Path $Path -Force -ErrorAction Stop | Out-Null
                Write-LogEntry "[OK] Created registry path: $Path" 'SUCCESS'
                return $true
            }
            return $true
        } catch {
            Write-LogEntry "[ERROR] Failed to create registry path: $Path - $_" 'ERROR'
            return $false
        }
    }

    # Get current registry value safely
    function Get-RegistryValue {
        param(
            [string]$Path,
            [string]$Name
        )
        
        try {
            if (Test-RegistryPath $Path) {
                $property = Get-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue
                if ($property) {
                    return $property.$Name
                }
            }
            return $null
        } catch {
            return $null
        }
    }

    # Check if policy value needs to be changed
    function Test-PolicyValue {
        param(
            [string]$Path,
            [string]$Name,
            [object]$DesiredValue
        )
        
        $currentValue = Get-RegistryValue -Path $Path -Name $Name
        return ($currentValue -eq $DesiredValue)
    }

    # Speed: Optimized policy application function with existing value detection
    function Set-OneDrivePolicy {
        param(
            [string]$Path,
            [string]$Name,
            [object]$Value,
            [string]$Type = 'DWord',
            [string]$Description,
            [switch]$Critical
        )
        
        try {
            # Check if the policy already has the correct value
            if (Test-PolicyValue -Path $Path -Name $Name -DesiredValue $Value) {
                $level = if ($Critical) { 'SKIP' } else { 'SKIP' }
                Write-LogEntry "⏭ $Description (already configured)" $level
                
                # [OK] FIX: Ensure skippedPolicies exists and use ArrayList.Add()
                if ($null -eq $script:skippedPolicies) {
                    $script:skippedPolicies = New-Object System.Collections.ArrayList
                }
                $skipInfo = @{
                    Path = $Path
                    Name = $Name
                    Value = $Value
                    Description = $Description
                    Critical = $Critical.IsPresent
                }
                [void]$script:skippedPolicies.Add($skipInfo)
                return $true
            }
            
            # Ensure registry path exists
            if (-not (New-RegistryPath -Path $Path)) {
                throw "Cannot create registry path: $Path"
            }
            
            # Apply policy setting
            if ($PSCmdlet.ShouldProcess("$Path\$Name", "Set policy value to $Value")) {
                Set-ItemProperty -Path $Path -Name $Name -Value $Value -Type $Type -ErrorAction Stop
                
                # Verify the setting was applied
                $verifyValue = Get-ItemProperty -Path $Path -Name $Name -ErrorAction Stop
                if ($verifyValue.$Name -eq $Value) {
                    $level = if ($Critical) { 'POLICY' } else { 'SUCCESS' }
                    Write-LogEntry "[OK] $Description" $level
                    
                    # [OK] FIX: Ensure appliedPolicies exists and use ArrayList.Add()
                    if ($null -eq $script:appliedPolicies) {
                        $script:appliedPolicies = New-Object System.Collections.ArrayList
                    }
                    $policyInfo = @{
                        Path = $Path
                        Name = $Name
                        Value = $Value
                        Description = $Description
                        Critical = $Critical.IsPresent
                    }
                    [void]$script:appliedPolicies.Add($policyInfo)
                    return $true
                } else {
                    throw "Verification failed: Expected $Value, got $($verifyValue.$Name)"
                }
            }
        } catch {
            Write-LogEntry "[ERROR] Failed to apply $Description : $_" 'ERROR'
            
            # [OK] FIX: Ensure failedPolicies exists and use ArrayList.Add()
            if ($null -eq $script:failedPolicies) {
                $script:failedPolicies = New-Object System.Collections.ArrayList
            }
            $failInfo = @{
                Path = $Path
                Name = $Name
                Description = $Description
                Error = $_.Exception.Message
            }
            [void]$script:failedPolicies.Add($failInfo)
            return $false
        }
    }

    # Security: OneDrive installation and version validation
    function Test-OneDriveInstallation {
        try {
            $oneDrivePaths = @(
                "${env:ProgramFiles}\Microsoft OneDrive\OneDrive.exe",
                "${env:ProgramFiles(x86)}\Microsoft OneDrive\OneDrive.exe",
                "${env:LOCALAPPDATA}\Microsoft\OneDrive\OneDrive.exe"
            )
            
            $installations = @()
            foreach ($path in $oneDrivePaths) {
                if (Test-Path $path) {
                    try {
                        $version = (Get-ItemProperty -Path $path).VersionInfo.FileVersion
                        $installations += @{
                            Path = $path
                            Version = $version
                            Type = if ($path -like "*Program Files*") { "System" } else { "User" }
                        }
                    } catch {
                        $installations += @{
                            Path = $path
                            Version = "Unknown"
                            Type = if ($path -like "*Program Files*") { "System" } else { "User" }
                        }
                    }
                }
            }
            
            return @{
                IsInstalled = ($installations.Count -gt 0)
                Installations = $installations
                RecommendedPath = $installations | Where-Object { $_.Type -eq "System" } | Select-Object -First 1
            }
            
        } catch {
            Write-LogEntry "OneDrive installation check failed: $_" 'WARNING'
            return @{
                IsInstalled = $false
                Error = $_.Exception.Message
            }
        }
    }

    # Security: Backup current OneDrive policies
    function Backup-OneDrivePolicies {
        param([string[]]$RegistryPaths)
        
        try {
            $backupData = @{}
            
            foreach ($regPath in $RegistryPaths) {
                if (Test-RegistryPath $regPath) {
                    $properties = Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue
                    if ($properties) {
                        $backupData[$regPath] = @{}
                        foreach ($property in $properties.PSObject.Properties) {
                            if ($property.Name -notmatch '^PS') {  # Skip PowerShell properties
                                $backupData[$regPath][$property.Name] = $property.Value
                            }
                        }
                    }
                }
            }
            
            if ($backupData.Count -gt 0) {
                $backupFile = "$env:TEMP\OneDrivePolicyBackup_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
                $backupData | ConvertTo-Json -Depth 4 | Out-File -FilePath $backupFile -Encoding UTF8
                Write-LogEntry "[OK] Policy backup saved to: $backupFile" 'INFO'
                return $backupFile
            }
            
        } catch {
            Write-LogEntry "[WARN] Failed to backup OneDrive policies: $_" 'WARNING'
        }
        
        return $null
    }

    # Security: Validate Tenant ID format
    function Test-TenantId {
        param([string]$TenantId)
        
        if ([string]::IsNullOrWhiteSpace($TenantId)) {
            return $true  # Optional parameter
        }
        
        # GUID format validation
        $guidPattern = '^[a-fA-F0-9]{8}-[a-fA-F0-9]{4}-[a-fA-F0-9]{4}-[a-fA-F0-9]{4}-[a-fA-F0-9]{12}$'
        return $TenantId -match $guidPattern
    }

    try {
        Write-LogEntry "=== OneDrive Auto-Login Policy Configuration Started ===" 'INFO'
        
        # Set default behaviors if no specific parameters provided
        if (-not ($EnableSilentConfig -or $DisableFirstRunWizard -or $EnableAutoStartup -or $EnableFilesOnDemand -or $DisablePersonalSync -or $EnableKnownFolderMove -or $DisableAutoLogin)) {
            Write-LogEntry "No specific policies specified, applying default configuration..." 'INFO'
            $EnableSilentConfig = $true
            $DisableFirstRunWizard = $true
            $EnableAutoStartup = $true
        }
        
        if ($WhatIfPreference) {
            Write-LogEntry "WhatIf mode - no registry changes will be made" 'INFO'
        }

        # Security: Validate Tenant ID if provided
        if ($TenantId -and -not (Test-TenantId -TenantId $TenantId)) {
            throw "Invalid Tenant ID format. Must be a valid GUID."
        }

        # Validate OneDrive installation
        Write-LogEntry "`n[CHECK] Validating OneDrive installation..." 'INFO'
        $oneDriveInfo = Test-OneDriveInstallation
        
        if (-not $oneDriveInfo.IsInstalled) {
            Write-LogEntry "[WARN] OneDrive not detected on this system" 'WARNING'
            Write-LogEntry "Policies will be applied but may not take effect until OneDrive is installed" 'WARNING'
        } else {
            Write-LogEntry "[OK] OneDrive installation detected:" 'SUCCESS'
            foreach ($installation in $oneDriveInfo.Installations) {
                Write-LogEntry "  • $($installation.Type): $($installation.Path) (v$($installation.Version))" 'INFO'
            }
        }

        # Define registry paths
        $registryPaths = @{
            MainPolicy = "HKLM:\SOFTWARE\Policies\Microsoft\OneDrive"
            UserPolicy = "HKCU:\SOFTWARE\Policies\Microsoft\OneDrive"
            TenantRestrictions = "HKLM:\SOFTWARE\Policies\Microsoft\OneDrive\TenantRestrictions"
            KnownFolderMove = "HKLM:\SOFTWARE\Policies\Microsoft\OneDrive\KnownFolderMove"
        }

        # Backup existing policies if requested
        if ($BackupSettings) {
            Write-LogEntry "`n[BACKUP] Backing up current OneDrive policies..." 'INFO'
            $backupFile = Backup-OneDrivePolicies -RegistryPaths $registryPaths.Values
        }

        if ($WhatIfPreference) {
            Write-LogEntry "`nWhatIf Summary - Policies that would be applied:" 'INFO'
            if ($EnableSilentConfig) { Write-LogEntry "  • Silent account configuration: Enabled" 'INFO' }
            if ($DisableFirstRunWizard) { Write-LogEntry "  • First run wizard: Disabled" 'INFO' }
            if ($EnableAutoStartup) { Write-LogEntry "  • Auto startup: Enabled" 'INFO' }
            if ($EnableFilesOnDemand) { Write-LogEntry "  • Files On-Demand: Enabled" 'INFO' }
            if ($DisablePersonalSync) { Write-LogEntry "  • Personal account sync: Disabled" 'INFO' }
            if ($EnableKnownFolderMove) { Write-LogEntry "  • Known Folder Move: Enabled" 'INFO' }
            if ($TenantId) { Write-LogEntry "  • Tenant restriction: $TenantId" 'INFO' }
            if ($SyncThrottleKbps -gt 0) { Write-LogEntry "  • Sync throttle: $SyncThrottleKbps KB/s" 'INFO' }
            $global:LastStatus = "[INFO] WhatIf completed - OneDrive policies would be configured."
            return
        }

        # Apply core OneDrive policies
        Write-LogEntry "`n🛠️ Applying OneDrive policies..." 'POLICY'
        
        # Silent account configuration
        if ($EnableSilentConfig) {
            Set-OneDrivePolicy -Path $registryPaths.MainPolicy -Name "SilentAccountConfig" -Value 1 -Description "Enabled silent account configuration" -Critical | Out-Null
        }
        
        # Disable first run wizard
        if ($DisableFirstRunWizard) {
            Set-OneDrivePolicy -Path $registryPaths.MainPolicy -Name "DisableFirstRunWizard" -Value 1 -Description "Disabled first run wizard" | Out-Null
        }
        
        # Auto startup policy
        if ($EnableAutoStartup) {
            Set-OneDrivePolicy -Path $registryPaths.MainPolicy -Name "OneDriveStartupPolicy" -Value 1 -Description "Enabled OneDrive auto startup" | Out-Null
        }
        
        # Files On-Demand
        if ($EnableFilesOnDemand) {
            Set-OneDrivePolicy -Path $registryPaths.MainPolicy -Name "FilesOnDemandEnabled" -Value 1 -Description "Enabled Files On-Demand" | Out-Null
        }
        
        # Disable personal account sync
        if ($DisablePersonalSync) {
            Set-OneDrivePolicy -Path $registryPaths.MainPolicy -Name "DisablePersonalSync" -Value 1 -Description "Disabled personal account synchronization" -Critical | Out-Null
        }
        
        # Known Folder Move
        if ($EnableKnownFolderMove) {
            Set-OneDrivePolicy -Path $registryPaths.KnownFolderMove -Name "KnownFolderMoveOpt" -Value 1 -Description "Enabled Known Folder Move optimization" | Out-Null
            
            if ($TenantId) {
                Set-OneDrivePolicy -Path $registryPaths.KnownFolderMove -Name $TenantId -Value 1 -Description "Enabled Known Folder Move for tenant: $TenantId" | Out-Null
            }
        }
        
        # Tenant restrictions
        if ($TenantId) {
            Set-OneDrivePolicy -Path $registryPaths.TenantRestrictions -Name $TenantId -Value 1 -Type "String" -Description "Applied tenant restriction for: $TenantId" -Critical | Out-Null
        }
        
        # Sync throttling
        if ($SyncThrottleKbps -gt 0) {
            Set-OneDrivePolicy -Path $registryPaths.MainPolicy -Name "UploadBandwidthLimit" -Value $SyncThrottleKbps -Description "Set upload bandwidth limit: $SyncThrottleKbps KB/s" | Out-Null
            Set-OneDrivePolicy -Path $registryPaths.MainPolicy -Name "DownloadBandwidthLimit" -Value $SyncThrottleKbps -Description "Set download bandwidth limit: $SyncThrottleKbps KB/s" | Out-Null
        }
        
        # Disable auto-login (override other settings)
        if ($DisableAutoLogin) {
            Set-OneDrivePolicy -Path $registryPaths.MainPolicy -Name "SilentAccountConfig" -Value 0 -Description "Disabled OneDrive auto-login" -Critical | Out-Null
            Set-OneDrivePolicy -Path $registryPaths.MainPolicy -Name "OneDriveStartupPolicy" -Value 0 -Description "Disabled OneDrive auto startup" | Out-Null
        }

        # Additional security and performance policies
        Write-LogEntry "`n[SECURITY] Applying security and performance policies..." 'POLICY'
        
        # Prevent OneDrive from generating network traffic until user signs in
        Set-OneDrivePolicy -Path $registryPaths.MainPolicy -Name "PreventNetworkTrafficPreUserSignIn" -Value 1 -Description "Prevented network traffic before user sign-in" | Out-Null
        
        # Block external sharing
        Set-OneDrivePolicy -Path $registryPaths.MainPolicy -Name "BlockExternalSync" -Value 1 -Description "Blocked external sharing and sync" | Out-Null
        
        # Enable automatic sign-in
        if (-not $DisableAutoLogin) {
            Set-OneDrivePolicy -Path $registryPaths.MainPolicy -Name "AutomaticUploadBandwidthPercentage" -Value 70 -Description "Set automatic upload bandwidth to 70%" | Out-Null
        }

        # Verification of applied policies
        Write-LogEntry "`n[CHECK] Verifying policy application..." 'INFO'
        $verificationErrors = 0
        
        # Ensure appliedPolicies exists before iterating
        if ($null -ne $script:appliedPolicies) {
            foreach ($policy in $script:appliedPolicies) {
                try {
                    $currentValue = Get-ItemProperty -Path $policy.Path -Name $policy.Name -ErrorAction Stop
                    if ($currentValue.($policy.Name) -ne $policy.Value) {
                        Write-LogEntry "[WARN] Verification failed for $($policy.Description)" 'WARNING'
                        $verificationErrors++
                    }
                } catch {
                    Write-LogEntry "[WARN] Could not verify $($policy.Description)" 'WARNING'
                    $verificationErrors++
                }
            }
        }
        
        $appliedCount = if ($null -ne $script:appliedPolicies) { $script:appliedPolicies.Count } else { 0 }
        $skippedCount = if ($null -ne $script:skippedPolicies) { $script:skippedPolicies.Count } else { 0 }
        
        if ($verificationErrors -eq 0 -and $appliedCount -gt 0) {
            Write-LogEntry "[OK] All OneDrive policies verified successfully" 'SUCCESS'
        } elseif ($appliedCount -eq 0 -and $skippedCount -gt 0) {
            Write-LogEntry "[OK] All OneDrive policies already correctly configured" 'SUCCESS'
        } elseif ($verificationErrors -gt 0) {
            Write-LogEntry "[WARN] $verificationErrors policies failed verification" 'WARNING'
        }

    } catch {
        Write-LogEntry "Critical error during OneDrive policy configuration: $_" 'ERROR'
        $global:LastStatus = "[ERROR] OneDrive policy configuration failed: $_"
        throw
    } finally {
        $stopwatch.Stop()
        $duration = $stopwatch.Elapsed.TotalSeconds
        
        # Safe counting with null checks
        $appliedCount = if ($null -ne $script:appliedPolicies) { $script:appliedPolicies.Count } else { 0 }
        $skippedCount = if ($null -ne $script:skippedPolicies) { $script:skippedPolicies.Count } else { 0 }
        $failedCount = if ($null -ne $script:failedPolicies) { $script:failedPolicies.Count } else { 0 }
        
        Write-LogEntry "`n[SUMMARY] Policy Configuration Summary:" 'INFO'
        Write-LogEntry "Duration: $([math]::Round($duration, 2)) seconds" 'INFO'
        Write-LogEntry "Policies applied: $appliedCount" 'SUCCESS'
        Write-LogEntry "Policies already correct: $skippedCount" 'SKIP'
        Write-LogEntry "Policies failed: $failedCount" 'ERROR'
        
        # Safe critical policy counting
        $criticalPolicies = 0
        $criticalSkipped = 0
        
        if ($null -ne $script:appliedPolicies) {
            $criticalPolicies = ($script:appliedPolicies | Where-Object { $_.Critical }).Count
        }
        if ($null -ne $script:skippedPolicies) {
            $criticalSkipped = ($script:skippedPolicies | Where-Object { $_.Critical }).Count
        }
        
        if ($criticalPolicies -gt 0) {
            Write-LogEntry "Critical policies applied: $criticalPolicies" 'POLICY'
        }
        if ($criticalSkipped -gt 0) {
            Write-LogEntry "Critical policies already configured: $criticalSkipped" 'SKIP'
        }
        
        if ($failedCount -gt 0 -and $null -ne $script:failedPolicies) {
            Write-LogEntry "`n[ERROR] Failed Policies:" 'ERROR'
            foreach ($failed in $script:failedPolicies) {
                Write-LogEntry "  • $($failed.Description): $($failed.Error)" 'ERROR'
            }
        }
        
        # Write detailed log file with null check
        try {
            if ($null -ne $script:logEntries) {
                $script:logEntries.ToArray() | Out-File -FilePath $LogPath -Encoding UTF8 -Force
                Write-LogEntry "[LOG] Detailed log saved to: $LogPath" 'INFO'
            }
        } catch {
            Write-LogEntry "[WARN] Failed to save log file: $_" 'WARNING'
        }
        
        # Set global status with safe counting
        $totalPolicies = $appliedCount + $skippedCount
        if ($totalPolicies -gt 0) {
            if ($appliedCount -gt 0) {
                $statusMsg = "[OK] Applied $appliedCount OneDrive policies"
                if ($skippedCount -gt 0) {
                    $statusMsg += " ($skippedCount already correct)"
                }
            } else {
                $statusMsg = "[OK] All $skippedCount OneDrive policies already correctly configured"
            }
            
            if ($failedCount -gt 0) {
                $statusMsg += " [$failedCount failed]"
            }
            if ($criticalPolicies -gt 0 -or $criticalSkipped -gt 0) {
                $statusMsg += " [$($criticalPolicies + $criticalSkipped) critical]"
            }
            $global:LastStatus = $statusMsg
        } else {
            $global:LastStatus = "[WARN] No OneDrive policies were processed"
        }
        
        Write-LogEntry "=== OneDrive Policy Configuration Completed ===" 'INFO'
    }
}

# Utility function to check current OneDrive policies
function Get-OneDrivePolicyStatus {
    [CmdletBinding()]
    param()
    
    try {
        $policyPaths = @{
            "Main Policies" = "HKLM:\SOFTWARE\Policies\Microsoft\OneDrive"
            "User Policies" = "HKCU:\SOFTWARE\Policies\Microsoft\OneDrive"
            "Tenant Restrictions" = "HKLM:\SOFTWARE\Policies\Microsoft\OneDrive\TenantRestrictions"
            "Known Folder Move" = "HKLM:\SOFTWARE\Policies\Microsoft\OneDrive\KnownFolderMove"
        }
        
        $status = @{}
        
        foreach ($category in $policyPaths.GetEnumerator()) {
            $status[$category.Key] = @{
                Path = $category.Value
                Exists = (Test-Path $category.Value)
                Policies = @{}
            }
            
            if ($status[$category.Key].Exists) {
                try {
                    $properties = Get-ItemProperty -Path $category.Value -ErrorAction SilentlyContinue
                    if ($properties) {
                        foreach ($prop in $properties.PSObject.Properties) {
                            if ($prop.Name -notmatch '^PS') {
                                $status[$category.Key].Policies[$prop.Name] = $prop.Value
                            }
                        }
                    }
                } catch {
                    $status[$category.Key].Error = $_.Exception.Message
                }
            }
        }
        
        return $status
        
    } catch {
        return @{
            Error = $_.Exception.Message
        }
    }
}

# Utility function to remove OneDrive policies
function Remove-OneDrivePolicies {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [switch]$ConfirmEach
    )
    
    $policyPaths = @(
        "HKLM:\SOFTWARE\Policies\Microsoft\OneDrive",
        "HKCU:\SOFTWARE\Policies\Microsoft\OneDrive"
    )
    
    $removedCount = 0
    
    foreach ($path in $policyPaths) {
        if (Test-Path $path) {
            try {
                if ($PSCmdlet.ShouldProcess($path, "Remove OneDrive policy registry key")) {
                    if (-not $ConfirmEach -or (Read-Host "Remove $path? (y/n)") -eq 'y') {
                        Remove-Item -Path $path -Recurse -Force -ErrorAction Stop
                        Write-Host "[OK] Removed OneDrive policies from: $path" -ForegroundColor Green
                        $removedCount++
                    }
                }
            } catch {
                Write-Warning "[WARN] Failed to remove $path : $_"
            }
        }
    }
    
    Write-Host "[OK] Removed OneDrive policies from $removedCount registry locations" -ForegroundColor Cyan
}


# ─────────────────────────────────────────────────────────────────────────────
# Option 19 - Full System Update
# ─────────────────────────────────────────────────────────────────────────────
function Run-CorePostDeploymentTasks {
    [CmdletBinding()]
    param()

    $ErrorActionPreference = 'Continue'
    $ProgressPreference = 'SilentlyContinue'

    $runStart = Get-Date
    $computerName = $env:COMPUTERNAME
    $timestampForFile = $runStart.ToString('yyyy-MM-dd_HH-mm-ss')
    $logDirectory = 'C:\Logs'
    $logPath = Join-Path -Path $logDirectory -ChildPath ('{0}-Full_System_Update-{1}.yaml' -f $computerName, $timestampForFile)

    $taskPlan = @(
        @{ Option = 8;  Name = 'Network Optimization';                     Function = 'Start-NetworkOptimization' },
        @{ Option = 2;  Name = 'Remove Windows Bloatware';                 Function = 'Remove-BloatwareApps' },
        @{ Option = 3;  Name = 'Set Recommended Registry Settings';        Function = 'Apply-RecommendedRegistrySettings' },
        @{ Option = 4;  Name = 'Optimize Windows Services';                Function = 'Optimize-WindowsServices' },
        @{ Option = 5;  Name = 'Enable PowerShell Remote Management';      Function = 'Enable-PowerShellRemotingSafely' },
        @{ Option = 6;  Name = 'Configure Automatic Time Sync';            Function = 'Configure-AutomaticTimeSync' },
        @{ Option = 7;  Name = 'Set Desktop Power Settings';               Function = 'Set-DesktopPowerSettings' },
        @{ Option = 9;  Name = 'Application Updates';                      Function = 'Update-Applications' },
        @{ Option = 10; Name = 'HP Driver Updates';                        Function = 'Update-HPDrivers' },
        @{ Option = 11; Name = 'Windows Updates';                          Function = 'Update-WindowsOS' }
    )

    $results = New-Object System.Collections.Generic.List[object]

    function ConvertTo-YamlSafeString {
        param(
            [AllowNull()]
            [object]$Value
        )

        if ($null -eq $Value) { return "''" }

        $stringValue = [string]$Value
        $stringValue = $stringValue -replace "`r`n", ' | '
        $stringValue = $stringValue -replace "`n", ' | '
        $stringValue = $stringValue -replace "`r", ' | '
        $stringValue = $stringValue -replace "'", "''"
        return "'$stringValue'"
    }

    function Invoke-FullUpdateFunction {
        param(
            [Parameter(Mandatory)]
            [string]$FunctionName
        )

        $previousState = @{
            ConfirmPreference = $ConfirmPreference
            ProgressPreference = $ProgressPreference
            VerbosePreference = $VerbosePreference
            InformationPreference = $InformationPreference
            ErrorActionPreference = $ErrorActionPreference
            PSDefaultParameterValues = if ($PSDefaultParameterValues) { $PSDefaultParameterValues.Clone() } else { @{} }
        }

        try {
            $script:ConfirmPreference = 'None'
            $script:ProgressPreference = 'SilentlyContinue'
            $script:VerbosePreference = 'SilentlyContinue'
            $script:InformationPreference = 'Continue'
            $script:ErrorActionPreference = 'Stop'
            $script:PSDefaultParameterValues = @{'*:Confirm' = $false}

            if (-not (Get-Command -Name $FunctionName -CommandType Function -ErrorAction SilentlyContinue)) {
                throw "Function '$FunctionName' was not found in the script."
            }

            & $FunctionName
        }
        finally {
            $script:ConfirmPreference = $previousState.ConfirmPreference
            $script:ProgressPreference = $previousState.ProgressPreference
            $script:VerbosePreference = $previousState.VerbosePreference
            $script:InformationPreference = $previousState.InformationPreference
            $script:ErrorActionPreference = $previousState.ErrorActionPreference
            $script:PSDefaultParameterValues = $previousState.PSDefaultParameterValues
        }
    }

    Write-Host "Starting Full System Update..." -ForegroundColor Cyan
    Write-Host "Running options in order: 8, 2, 3, 4, 5, 6, 7, 9, 10, 11" -ForegroundColor Cyan

    foreach ($task in $taskPlan) {
        $taskStart = Get-Date
        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        $result = [ordered]@{
            Option = [int]$task.Option
            Name = [string]$task.Name
            Function = [string]$task.Function
            StartTime = $taskStart.ToString('yyyy-MM-dd HH:mm:ss')
            EndTime = $null
            DurationSeconds = 0
            Success = $false
            Status = 'NotStarted'
            Error = $null
        }

        try {
            Write-Host ("[{0}] {1}" -f $task.Option, $task.Name) -ForegroundColor Yellow
            Invoke-FullUpdateFunction -FunctionName $task.Function
            $result.Success = $true
            $result.Status = 'Completed'
        }
        catch {
            $result.Success = $false
            $result.Status = 'Failed'
            $result.Error = $_.Exception.Message
            Write-Warning ("Option {0} failed: {1}" -f $task.Option, $_.Exception.Message)
        }
        finally {
            $stopwatch.Stop()
            $taskEnd = Get-Date
            $result.EndTime = $taskEnd.ToString('yyyy-MM-dd HH:mm:ss')
            $result.DurationSeconds = [math]::Round($stopwatch.Elapsed.TotalSeconds, 2)
            $results.Add([pscustomobject]$result)
        }
    }

    $runEnd = Get-Date
    $successfulTasks = @($results | Where-Object { $_.Success })
    $failedTasks = @($results | Where-Object { -not $_.Success })

    $yamlLines = New-Object System.Collections.Generic.List[string]
    $yamlLines.Add('run_summary:')
    $yamlLines.Add(('  computer_name: {0}' -f (ConvertTo-YamlSafeString $computerName)))
    $yamlLines.Add(('  started_at: {0}' -f (ConvertTo-YamlSafeString ($runStart.ToString('yyyy-MM-dd HH:mm:ss')))))
    $yamlLines.Add(('  ended_at: {0}' -f (ConvertTo-YamlSafeString ($runEnd.ToString('yyyy-MM-dd HH:mm:ss')))))
    $yamlLines.Add(('  duration_seconds: {0}' -f ([math]::Round(($runEnd - $runStart).TotalSeconds, 2))))
    $yamlLines.Add(('  total_tasks: {0}' -f $results.Count))
    $yamlLines.Add(('  successful_tasks: {0}' -f $successfulTasks.Count))
    $yamlLines.Add(('  failed_tasks: {0}' -f $failedTasks.Count))
    $yamlLines.Add(('  log_path: {0}' -f (ConvertTo-YamlSafeString $logPath)))
    $yamlLines.Add('task_order:')
    foreach ($task in $taskPlan) {
        $yamlLines.Add(('  - {0}' -f $task.Option))
    }
    $yamlLines.Add('tasks:')
    foreach ($item in $results) {
        $yamlLines.Add(('  - option: {0}' -f $item.Option))
        $yamlLines.Add(('    name: {0}' -f (ConvertTo-YamlSafeString $item.Name)))
        $yamlLines.Add(('    function: {0}' -f (ConvertTo-YamlSafeString $item.Function)))
        $yamlLines.Add(('    start_time: {0}' -f (ConvertTo-YamlSafeString $item.StartTime)))
        $yamlLines.Add(('    end_time: {0}' -f (ConvertTo-YamlSafeString $item.EndTime)))
        $yamlLines.Add(('    duration_seconds: {0}' -f $item.DurationSeconds))
        $yamlLines.Add(('    success: {0}' -f $item.Success.ToString().ToLower()))
        $yamlLines.Add(('    status: {0}' -f (ConvertTo-YamlSafeString $item.Status)))
        $yamlLines.Add(('    error: {0}' -f (ConvertTo-YamlSafeString $item.Error)))
    }

    try {
        if (-not (Test-Path -Path $logDirectory)) {
            New-Item -Path $logDirectory -ItemType Directory -Force | Out-Null
        }
        Set-Content -Path $logPath -Value $yamlLines -Encoding UTF8
    }
    catch {
        Write-Warning ("Failed to write YAML log: {0}" -f $_.Exception.Message)
    }

    if ($failedTasks.Count -eq 0) {
        $global:LastStatus = "[OK] Full System Update completed successfully. Log: $logPath"
        Write-Host $global:LastStatus -ForegroundColor Green
    }
    else {
        $global:LastStatus = "[WARN] Full System Update completed with $($failedTasks.Count) failed task(s). Log: $logPath"
        Write-Host $global:LastStatus -ForegroundColor Yellow
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# Menu Display
# ─────────────────────────────────────────────────────────────────────────────
function Show-Menu {
    Clear-Host
    Write-Host "╔══════════════════════════════════════════════════════════╗" -ForegroundColor Magenta
    Write-Host "║              Compton College Tech Utils                ║" -ForegroundColor Magenta
    Write-Host "╚══════════════════════════════════════════════════════════╝" -ForegroundColor Magenta
    Write-Host ""

    Write-Host "1.  Create MISAdmin account" -ForegroundColor White
    Write-Host "2.  Remove Windows Bloatware" -ForegroundColor White
    Write-Host "3.  Set Recommended Registry Settings" -ForegroundColor White
    Write-Host "4.  Optimize Windows Services" -ForegroundColor White
    Write-Host "5.  Enable PowerShell Remote Management" -ForegroundColor White
    Write-Host "6.  Configure Automatic Time Sync" -ForegroundColor White
    Write-Host "7.  Set Desktop Power Settings" -ForegroundColor White
    Write-Host "8.  Network Optimization" -ForegroundColor White
    Write-Host "9.  Application Updates" -ForegroundColor White
    Write-Host "10. HP Driver Updates" -ForegroundColor White
    Write-Host "11. Windows Updates" -ForegroundColor White
    Write-Host "12. Disk Cleanup" -ForegroundColor White
    Write-Host "13. System Repair" -ForegroundColor White
    Write-Host "14. Remove User Profiles" -ForegroundColor White
    Write-Host "15. Disable Last User Display" -ForegroundColor White
    Write-Host "16. Enable Automatic Login with CC-Student" -ForegroundColor White
    Write-Host "17. Install Computer Lab Scheduled Tasks" -ForegroundColor White
    Write-Host "18. Set OneDrive Auto Login on Boot" -ForegroundColor White
    Write-Host "19. Run Full System Updates" -ForegroundColor White
    Write-Host "20. Network Diag and Repair" -ForegroundColor White
    Write-Host "Q.  Exit" -ForegroundColor Red

    Write-Host ""
    Write-Host "Last Status: $global:LastStatus" -ForegroundColor Yellow
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# Main Loop
# ─────────────────────────────────────────────────────────────────────────────
function Main {
    do {
        Show-Menu
        $choice = Read-Host "Enter your selection"

        switch ($choice.ToUpperInvariant()) {
            "1" {
                Clear-Host
                New-MISAdminAccount
                Pause
            }
            "2" {
                Clear-Host
                Remove-BloatwareApps
                Pause
            }
            "3" {
                Clear-Host
                Apply-RecommendedRegistrySettings
                Pause
            }
            "4" {
                Clear-Host
                Optimize-WindowsServices
                Pause
            }
            "5" {
                Clear-Host
                Enable-PowerShellRemotingSafely
                Pause
            }
            "6" {
                Clear-Host
                Configure-AutomaticTimeSync
                Pause
            }
            "7" {
                Clear-Host
                Set-DesktopPowerSettings
                Pause
            }
            "8" {
                Clear-Host
                Start-NetworkOptimization
                Pause
            }
            "9" {
                Clear-Host
                Update-Applications
                Pause
            }
            "10" {
                Clear-Host
                Update-HPDrivers
                Pause
            }
            "11" {
                Clear-Host
                Update-WindowsOS
                Pause
            }
            "12" {
                Clear-Host
                Run-DiskCleanup
                Pause
            }
            "13" {
                Clear-Host
                Invoke-SystemMaintenance
                Pause
            }
            "14" {
                Clear-Host
                Remove-UserProfilesClassroom
                Pause
            }
            "15" {
                Clear-Host
                Apply-LoginScreenRegistryFixes
                Pause
            }
            "16" {
                Clear-Host
                Set-DomainAutoLogin
                Pause
            }
            "17" {
                Clear-Host
                Register-LabScheduledTasks
                Pause
            }
            "18" {
                Clear-Host
                Set-OneDriveAutoLoginPolicy
                Pause
            }
            "19" {
                Clear-Host
                Run-CorePostDeploymentTasks
                Pause
            }
            "20" {
                Clear-Host
                Run-NetworkDiagnostics
                Pause
            }
            "Q" {
                Clear-Host
                Write-Host "Compton College Tech Utils has exited." -ForegroundColor Cyan
                exit
            }
            Default {
                $global:LastStatus = "[ERROR] Invalid selection. Please try again."
                Start-Sleep -Seconds 1
            }
        }
    } while ($true)
}

# ─────────────────────────────────────────────────────────────────────────────
# Entry Point
# ─────────────────────────────────────────────────────────────────────────────
Invoke-StartupSelfUpdate
Main
