# ScriptVersion: 1.0
# LastUpdated: 2026-03-23

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
$script:BaseFileName = "{0}_{1}_RemoveUserProfiles" -f $script:ComputerName, $script:TimestampForFile
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
    exit 1
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
    exit 2
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
            Write-Log "Could not evaluate LastUseTime for $profileName: $($_.Exception.Message)" 'WARN'
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
    exit 2
}

exit 0