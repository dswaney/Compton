# ScriptName: 05_Weekend_HP_Drivers_Update.ps1
# ScriptVersion: 2.8.3
# LastUpdated: 2026-07-23
# Purpose: Weekend vendor driver update script with clean HP + Dell support,
#          HP-only HPCMSL maintenance after vendor detection,
#          YAML logging, colored output,
#          share/Internet fallback deployment for HPIA, .NET Desktop Runtime, and DCU,
#          section headers, progress display, and structured per-driver results.

[CmdletBinding()]
param([string]$WorkingRoot = 'C:\Temp\DriverUpdates',
    [string]$YamlLogFolder = 'C:\Logs',
    [switch]$IncludeSoftware,
    [switch]$IncludeBIOS,
    [switch]$All,
    [string]$HpiaSourceFolder = '\\filesvr\Labscripts\HPImageAssistant',
    [string]$HpiaInstallerSharePath = '\\filesvr\labscripts\HPImageAssistant.exe',
    [string]$HpiaInternetLandingPage = 'https://ftp.ext.hp.com/pub/caps-softpaq/cmit/HPIA.html',
    [string]$LocalHpiaFolder = 'C:\ProgramData\Compton\HPImageAssistant',
    [string]$DellCommandUpdateSharePath = '\\filesvr\labscripts\Dell-Command-Update-Windows-Universal-Application.exe',
    [string]$DotNetDesktopRuntimeSharePath = '\\filesvr\labscripts\windowsdesktop-runtime-8.0.25-win-x64.exe',
    [string]$DotNetDesktopRuntimeInternetUrl = 'https://aka.ms/dotnet/8.0/windowsdesktop-runtime-win-x64.exe',
    [string]$DellCommandUpdateWingetId = 'Dell.CommandUpdate.Universal'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# -----------------------------
# Script metadata
# -----------------------------
$script:ScriptName        = '05_Weekend_HP_Drivers_Update.ps1'
$script:ScriptVersion     = '2.8.3'
$script:StartTime         = Get-Date
$script:RunFailures       = New-Object System.Collections.Generic.List[string]
$script:InstalledList     = New-Object System.Collections.Generic.List[string]
$script:SkippedList       = New-Object System.Collections.Generic.List[string]
$script:YamlActionLines   = New-Object System.Collections.Generic.List[string]
$script:DriverResults     = New-Object System.Collections.Generic.List[object]
$script:DetectedVendor    = $null
$script:YamlPath          = $null
$script:ComputerName      = $env:COMPUTERNAME
$script:InstallAllUpdates  = [bool]$All
$script:InstallBiosFirmware = ([bool]$IncludeBIOS -or [bool]$All)

# -----------------------------
# Logging helpers
# -----------------------------
function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','OK','WARN','ERROR')][string]$Level = 'INFO'
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[{0}] [{1}] {2}" -f $timestamp, $Level.PadRight(5), $Message

    switch ($Level) {
        'INFO'  { Write-Host $line -ForegroundColor Cyan }
        'OK'    { Write-Host $line -ForegroundColor Green }
        'WARN'  { Write-Host $line -ForegroundColor Yellow }
        'ERROR' { Write-Host $line -ForegroundColor Red }
        default { Write-Host $line }
    }
}

function Write-Section {
    param([Parameter(Mandatory)][string]$Title)

    $border = ('=' * 72)
    Write-Host ''
    Write-Host $border -ForegroundColor Magenta
    Write-Host ("  {0}" -f $Title) -ForegroundColor Magenta
    Write-Host $border -ForegroundColor Magenta
    Add-YamlAction ("Section: {0}" -f $Title)
}

function Add-RunFailure {
    param([Parameter(Mandatory)][string]$Message)
    $script:RunFailures.Add($Message) | Out-Null
    Write-Log $Message 'WARN'
}

function Add-DriverResult {
    param(
        [Parameter(Mandatory)][string]$Vendor,
        [Parameter(Mandatory)][string]$Name,
        [string]$Id,
        [string]$Category,
        [ValidateSet('Detected','Installed','Downloaded','Skipped','Failed','Blocked')][string]$Status,
        [string]$Message = '',
        [string]$SourceReport = '',
        [string]$RawSummary = ''
    )

    $script:DriverResults.Add([pscustomobject]@{
        Vendor       = $Vendor
        Name         = $Name
        Id           = $Id
        Category     = $Category
        Status       = $Status
        Message      = $Message
        SourceReport = $SourceReport
        RawSummary   = $RawSummary
    }) | Out-Null
}

function ConvertTo-YamlSafeString {
    param([AllowNull()][object]$Value)

    if ($null -eq $Value) { return 'null' }

    $text = [string]$Value
    $text = $text -replace "`r", ''
    $text = $text -replace "`n", ' '
    $text = $text -replace "'", "''"
    return "'$text'"
}

function Initialize-YamlLog {
    param(
        [Parameter(Mandatory)][string]$ComputerName,
        [Parameter(Mandatory)][string]$YamlPath
    )

    $script:YamlPath = $YamlPath
    $script:YamlActionLines.Clear()

    Add-YamlAction 'Script initialized.'
    Add-YamlAction ("Working root: {0}" -f $WorkingRoot)
    Add-YamlAction ("Computer: {0}" -f $ComputerName)
}

function Add-YamlAction {
    param([Parameter(Mandatory)][string]$Text)
    $script:YamlActionLines.Add(("    - {0}" -f (ConvertTo-YamlSafeString $Text))) | Out-Null
}

function Save-YamlLog {
    param(
        [Parameter(Mandatory)][string]$Status
    )

    if ([string]::IsNullOrWhiteSpace($script:YamlPath)) {
        return
    }

    $endTime = Get-Date
    $duration = [math]::Round(($endTime - $script:StartTime).TotalSeconds, 0)

    $lines = New-Object System.Collections.Generic.List[string]

    $lines.Add('script:') | Out-Null
    $lines.Add(("  name: {0}" -f (ConvertTo-YamlSafeString $script:ScriptName))) | Out-Null
    $lines.Add(("  version: {0}" -f (ConvertTo-YamlSafeString $script:ScriptVersion))) | Out-Null
    $lines.Add(("  computer: {0}" -f (ConvertTo-YamlSafeString $script:ComputerName))) | Out-Null
    $lines.Add(("  started: {0}" -f (ConvertTo-YamlSafeString ($script:StartTime.ToString('s'))))) | Out-Null
    $lines.Add(("  ended: {0}" -f (ConvertTo-YamlSafeString ($endTime.ToString('s'))))) | Out-Null
    $lines.Add(("  duration_seconds: {0}" -f $duration)) | Out-Null

    $lines.Add('run:') | Out-Null
    $lines.Add(("  status: {0}" -f (ConvertTo-YamlSafeString $Status))) | Out-Null
    $lines.Add(("  vendor: {0}" -f (ConvertTo-YamlSafeString $script:DetectedVendor))) | Out-Null

    $lines.Add('  actions:') | Out-Null
    if ($script:YamlActionLines.Count -eq 0) {
        $lines.Add('    []') | Out-Null
    }
    else {
        foreach ($line in $script:YamlActionLines) {
            $lines.Add($line) | Out-Null
        }
    }

    $lines.Add('  installed:') | Out-Null
    if ($script:InstalledList.Count -eq 0) {
        $lines.Add('    []') | Out-Null
    }
    else {
        foreach ($item in $script:InstalledList) {
            $lines.Add(("    - {0}" -f (ConvertTo-YamlSafeString $item))) | Out-Null
        }
    }

    $lines.Add('  skipped:') | Out-Null
    if ($script:SkippedList.Count -eq 0) {
        $lines.Add('    []') | Out-Null
    }
    else {
        foreach ($item in $script:SkippedList) {
            $lines.Add(("    - {0}" -f (ConvertTo-YamlSafeString $item))) | Out-Null
        }
    }

    $lines.Add('  failures:') | Out-Null
    if ($script:RunFailures.Count -eq 0) {
        $lines.Add('    []') | Out-Null
    }
    else {
        foreach ($item in $script:RunFailures) {
            $lines.Add(("    - {0}" -f (ConvertTo-YamlSafeString $item))) | Out-Null
        }
    }

    $lines.Add('drivers:') | Out-Null
    if ($script:DriverResults.Count -eq 0) {
        $lines.Add('  []') | Out-Null
    }
    else {
        foreach ($driver in $script:DriverResults) {
            $lines.Add('  -') | Out-Null
            $lines.Add(("    vendor: {0}" -f (ConvertTo-YamlSafeString $driver.Vendor))) | Out-Null
            $lines.Add(("    name: {0}" -f (ConvertTo-YamlSafeString $driver.Name))) | Out-Null
            $lines.Add(("    id: {0}" -f (ConvertTo-YamlSafeString $driver.Id))) | Out-Null
            $lines.Add(("    category: {0}" -f (ConvertTo-YamlSafeString $driver.Category))) | Out-Null
            $lines.Add(("    status: {0}" -f (ConvertTo-YamlSafeString $driver.Status))) | Out-Null
            $lines.Add(("    message: {0}" -f (ConvertTo-YamlSafeString $driver.Message))) | Out-Null
            $lines.Add(("    source_report: {0}" -f (ConvertTo-YamlSafeString $driver.SourceReport))) | Out-Null
            $lines.Add(("    raw_summary: {0}" -f (ConvertTo-YamlSafeString $driver.RawSummary))) | Out-Null
        }
    }

    Set-Content -LiteralPath $script:YamlPath -Value $lines -Encoding UTF8
    Write-Host ("YAML log written successfully: {0}" -f $script:YamlPath) -ForegroundColor Green
}

# -----------------------------
# File/folder helpers
# -----------------------------
function Ensure-Folder {
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -Path $Path -ItemType Directory -Force | Out-Null
    }
}

function Ensure-WorkingFolderPermissions {
    param([Parameter(Mandatory)][string]$Path)

    try {
        & icacls.exe $Path '/grant' '*S-1-1-0:(OI)(CI)F' '/T' '/C' | Out-Null
    }
    catch {
        Write-Log ("Unable to relax working folder permissions on {0}: {1}" -f $Path, $_.Exception.Message) 'WARN'
    }
}

function Remove-WorkingFolderRobust {
    param(
        [Parameter(Mandatory)][string]$Path,
        [int]$RetryCount = 6,
        [int]$RetryDelaySeconds = 5
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        Write-Log ("Working folder already absent: {0}" -f $Path) 'OK'
        return
    }

    Write-Log ("Attempting to remove working folder: {0}" -f $Path) 'INFO'
    for ($i = 1; $i -le $RetryCount; $i++) {
        try {
            Remove-Item -LiteralPath $Path -Recurse -Force -ErrorAction Stop
            Write-Log 'Working folder removed successfully.' 'OK'
            return
        }
        catch {
            if ($i -eq $RetryCount) {
                Add-RunFailure ("Failed to remove working folder after retries: {0}" -f $_.Exception.Message)
                return
            }
            Start-Sleep -Seconds $RetryDelaySeconds
        }
    }
}

# -----------------------------
# Vendor detection
# -----------------------------
function Get-SystemManufacturer {
    try {
        return (Get-CimInstance Win32_ComputerSystem -ErrorAction Stop).Manufacturer
    }
    catch {
        throw "Unable to determine system manufacturer. $($_.Exception.Message)"
    }
}

function Get-DriverVendor {
    $manufacturer = Get-SystemManufacturer
    Write-Log ("Detected manufacturer: {0}" -f $manufacturer) 'INFO'

    if ($manufacturer -match 'Dell') { return 'Dell' }
    if ($manufacturer -match 'HP|Hewlett-Packard') { return 'HP' }

    throw "Unsupported manufacturer for this script: $manufacturer"
}

# -----------------------------
# HP Support - HP Image Assistant extracted-folder deployment
# -----------------------------
function Get-HPSystemModelInfo {
    [CmdletBinding()]
    param()

    $cs = Get-CimInstance Win32_ComputerSystem -ErrorAction Stop
    $csp = Get-CimInstance Win32_ComputerSystemProduct -ErrorAction SilentlyContinue
    $bb = Get-CimInstance Win32_BaseBoard -ErrorAction SilentlyContinue

    $platform = $null
    if ($bb -and $bb.Product) {
        $platform = $bb.Product.ToString().Trim().ToUpper()
        if ($platform.Length -gt 4) {
            $platform = $platform.Substring(0,4)
        }
    }

    $model = if ($cs.Model) { $cs.Model.ToString().Trim() } else { 'Unknown' }
    $sku = if ($csp -and $csp.Version) { $csp.Version.ToString().Trim() } else { 'Unknown' }

    $info = [pscustomobject]@{
        Manufacturer = $cs.Manufacturer
        Model        = $model
        SKU          = $sku
        Platform     = $platform
    }

    Write-Log ("Detected HP system model: {0}" -f $info.Model) 'INFO'
    Write-Log ("Detected HP platform/baseboard ID: {0}" -f $info.Platform) 'INFO'
    Add-YamlAction ("Detected HP system model: {0}" -f $info.Model)
    Add-YamlAction ("Detected HP platform/baseboard ID: {0}" -f $info.Platform)

    return $info
}

function Get-ExistingHPIAExecutable {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$PreferredFolder)

    $candidateFolders = New-Object System.Collections.Generic.List[string]

    if (-not [string]::IsNullOrWhiteSpace($PreferredFolder)) {
        $candidateFolders.Add($PreferredFolder) | Out-Null
    }

    foreach ($folder in @(
        'C:\Program Files\HP\HP Image Assistant',
        'C:\Program Files (x86)\HP\HP Image Assistant',
        'C:\SWSetup\HPImageAssistant',
        'C:\ProgramData\Compton\HPImageAssistant'
    )) {
        if (-not $candidateFolders.Contains($folder)) {
            $candidateFolders.Add($folder) | Out-Null
        }
    }

    foreach ($folder in $candidateFolders) {
        if ([string]::IsNullOrWhiteSpace($folder) -or -not (Test-Path -LiteralPath $folder)) {
            continue
        }

        $exe = Get-ChildItem -Path $folder -Recurse -Filter 'HPImageAssistant.exe' -File -ErrorAction SilentlyContinue |
            Sort-Object FullName |
            Select-Object -First 1

        if ($exe) {
            return $exe.FullName
        }
    }

    return $null
}

function Install-HPIAFromExtractedFolder {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SourceFolder,
        [Parameter(Mandatory)][string]$DestinationFolder
    )

    Write-Section 'HP Image Assistant Local Deployment'
    Write-Log ("HPIA source folder: {0}" -f $SourceFolder) 'INFO'
    Write-Log ("HPIA local folder: {0}" -f $DestinationFolder) 'INFO'
    Add-YamlAction ("HPIA source folder: {0}" -f $SourceFolder)
    Add-YamlAction ("HPIA local folder: {0}" -f $DestinationFolder)

    $existingExe = Get-ExistingHPIAExecutable -PreferredFolder $DestinationFolder
    if ($existingExe) {
        try {
            $existingVersion = (Get-Item -LiteralPath $existingExe -ErrorAction Stop).VersionInfo.FileVersion
            Write-Log ("HP Image Assistant is already installed/found at: {0} (Version: {1})" -f $existingExe, $existingVersion) 'OK'
            Add-YamlAction ("Skipped HPIA local deployment because HPImageAssistant.exe already exists: {0} (Version: {1})" -f $existingExe, $existingVersion)
        }
        catch {
            Write-Log ("HP Image Assistant is already installed/found at: {0}" -f $existingExe) 'OK'
            Add-YamlAction ("Skipped HPIA local deployment because HPImageAssistant.exe already exists: {0}" -f $existingExe)
        }

        return $existingExe
    }

    Write-Log 'HP Image Assistant was not found locally. Deploying from extracted source folder.' 'INFO'
    Add-YamlAction 'HP Image Assistant was not found locally. Deploying from extracted source folder.'

    if (-not (Test-Path -LiteralPath $SourceFolder)) {
        throw "HPIA source folder not found: $SourceFolder"
    }

    $sourceFiles = @(Get-ChildItem -Path $SourceFolder -Recurse -File -ErrorAction SilentlyContinue)
    Write-Log ("Source HPIA file count: {0}" -f $sourceFiles.Count) 'INFO'
    Add-YamlAction ("Source HPIA file count: {0}" -f $sourceFiles.Count)

    $sourceExe = Get-ChildItem -Path $SourceFolder -Recurse -Filter 'HPImageAssistant.exe' -File -ErrorAction SilentlyContinue |
        Sort-Object FullName |
        Select-Object -First 1

    if (-not $sourceExe) {
        throw "HPImageAssistant.exe was not found anywhere under source folder: $SourceFolder"
    }

    Write-Log ("Found source HPImageAssistant.exe: {0}" -f $sourceExe.FullName) 'OK'
    Add-YamlAction ("Found source HPImageAssistant.exe: {0}" -f $sourceExe.FullName)

    try {
        if (Test-Path -LiteralPath $DestinationFolder) {
            Write-Log ("Removing existing local HPIA folder before clean deployment: {0}" -f $DestinationFolder) 'INFO'
            Remove-Item -LiteralPath $DestinationFolder -Recurse -Force -ErrorAction Stop
        }

        New-Item -Path $DestinationFolder -ItemType Directory -Force | Out-Null

        Write-Log 'Copying extracted HPIA files locally with robocopy...' 'INFO'

        $roboLog = Join-Path $DestinationFolder 'HPIA_robocopy.log'
        $roboArgs = @(
            ('"{0}"' -f $SourceFolder),
            ('"{0}"' -f $DestinationFolder),
            '/E',
            '/COPY:DAT',
            '/R:3',
            '/W:5',
            '/NFL',
            '/NDL',
            '/NP',
            ('/LOG:"{0}"' -f $roboLog)
        )

        $robo = Start-Process -FilePath "$env:SystemRoot\System32\robocopy.exe" -ArgumentList ($roboArgs -join ' ') -Wait -PassThru -NoNewWindow

        # Robocopy exit codes 0-7 are success/non-fatal. 8+ indicates failure.
        if ($robo.ExitCode -ge 8) {
            throw "Robocopy failed copying HPIA files. Exit code: $($robo.ExitCode). Log: $roboLog"
        }

        Write-Log ("HPIA files copied locally with robocopy exit code {0}." -f $robo.ExitCode) 'OK'
        Add-YamlAction ("HPIA files copied locally with robocopy exit code {0}." -f $robo.ExitCode)

        try {
            Get-ChildItem -Path $DestinationFolder -Recurse -File -ErrorAction SilentlyContinue |
                ForEach-Object { Unblock-File -Path $_.FullName -ErrorAction SilentlyContinue }
        }
        catch {}

        $copiedFiles = @(Get-ChildItem -Path $DestinationFolder -Recurse -File -ErrorAction SilentlyContinue)
        Write-Log ("Local HPIA folder file count after copy: {0}" -f $copiedFiles.Count) 'INFO'
        Add-YamlAction ("Local HPIA folder file count after copy: {0}" -f $copiedFiles.Count)

        $localExeItem = Get-ChildItem -Path $DestinationFolder -Recurse -Filter 'HPImageAssistant.exe' -File -ErrorAction SilentlyContinue |
            Sort-Object FullName |
            Select-Object -First 1

        if (-not $localExeItem) {
            $sampleFiles = $copiedFiles | Select-Object -First 20 | ForEach-Object { $_.FullName }
            foreach ($sample in $sampleFiles) {
                Write-Log ("Local HPIA sample file: {0}" -f $sample) 'WARN'
            }

            throw "HPImageAssistant.exe was not found anywhere under local copy folder: $DestinationFolder"
        }

        $localExe = $localExeItem.FullName

        Write-Log ("Resolved local HPImageAssistant.exe location: {0}" -f $localExe) 'OK'
        Add-YamlAction ("Resolved local HPImageAssistant.exe location: {0}" -f $localExe)

        return $localExe
    }
    catch {
        throw "Failed to deploy HP Image Assistant locally: $($_.Exception.Message)"
    }
}


function Expand-HpiaSoftPaq {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$InstallerPath,
        [Parameter(Mandatory)][string]$DestinationFolder
    )

    Write-Section 'HP Image Assistant Installer Extraction'
    if (-not (Test-Path -LiteralPath $InstallerPath)) {
        throw "HPIA installer was not found: $InstallerPath"
    }

    if (Test-Path -LiteralPath $DestinationFolder) {
        Remove-Item -LiteralPath $DestinationFolder -Recurse -Force -ErrorAction Stop
    }
    Ensure-Folder -Path $DestinationFolder

    $arguments = @('-s', '-e', ("-f`"{0}`"" -f $DestinationFolder))
    $process = Start-Process -FilePath $InstallerPath -ArgumentList ($arguments -join ' ') -Wait -PassThru -NoNewWindow
    Write-Log ("HPIA SoftPaq extraction exit code: {0}" -f $process.ExitCode) 'INFO'

    if ($process.ExitCode -notin @(0, 3010, 1641)) {
        throw ("HPIA SoftPaq extraction failed with exit code {0}." -f $process.ExitCode)
    }

    Start-Sleep -Seconds 2
    $hpiaExe = Get-ExistingHPIAExecutable -PreferredFolder $DestinationFolder
    if (-not $hpiaExe) {
        throw "HPIA extraction completed, but HPImageAssistant.exe was not found under $DestinationFolder"
    }

    Write-Log ("HP Image Assistant extracted and verified: {0}" -f $hpiaExe) 'OK'
    Add-YamlAction ("HP Image Assistant extracted and verified: {0}" -f $hpiaExe)
    return $hpiaExe
}

function Get-LatestHpiaSoftPaqUrl {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$LandingPage)

    Write-Log ("Checking HP for the latest HPIA SoftPaq: {0}" -f $LandingPage) 'INFO'
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $response = Invoke-WebRequest -Uri $LandingPage -UseBasicParsing -TimeoutSec 60 -ErrorAction Stop
    $candidateLinks = New-Object System.Collections.Generic.List[string]

    foreach ($link in @($response.Links)) {
        if ($link.href -and $link.href -match '(?i)sp\d+\.exe(?:\?.*)?$') {
            $candidateLinks.Add([string]$link.href) | Out-Null
        }
    }
    if ($candidateLinks.Count -eq 0) {
        $matches = [regex]::Matches([string]$response.Content, '(?i)href\s*=\s*["'']([^"'']*sp\d+\.exe(?:\?[^"'']*)?)["'']')
        foreach ($match in $matches) { $candidateLinks.Add($match.Groups[1].Value) | Out-Null }
    }
    if ($candidateLinks.Count -eq 0) { throw 'The HP HPIA page did not contain a SoftPaq EXE download link.' }

    $resolved = foreach ($href in $candidateLinks) {
        try {
            $uri = [uri]::new([uri]$LandingPage, $href)
            $spMatch = [regex]::Match($uri.AbsoluteUri, '(?i)sp(?<number>\d+)\.exe')
            if ($spMatch.Success) {
                [pscustomobject]@{ Url = $uri.AbsoluteUri; SoftPaqNumber = [int]$spMatch.Groups['number'].Value }
            }
        } catch {}
    }
    $latest = $resolved | Sort-Object SoftPaqNumber -Descending | Select-Object -First 1
    if (-not $latest) { throw 'Unable to resolve a valid HPIA SoftPaq URL from the HP landing page.' }
    Add-YamlAction ("Latest HPIA SoftPaq resolved from HP: {0}" -f $latest.Url)
    return $latest.Url
}

function Ensure-HPIAAvailable {
    [CmdletBinding()]
    param()

    $existingExe = Get-ExistingHPIAExecutable -PreferredFolder $LocalHpiaFolder
    if ($existingExe) {
        Write-Log ("HP Image Assistant is already available locally: {0}" -f $existingExe) 'OK'
        Add-YamlAction 'HP Image Assistant was already available locally.'
        return $existingExe
    }

    if (Test-Path -LiteralPath $HpiaSourceFolder) {
        try { return Install-HPIAFromExtractedFolder -SourceFolder $HpiaSourceFolder -DestinationFolder $LocalHpiaFolder }
        catch { Write-Log ("Extracted HPIA share-folder deployment failed: {0}" -f $_.Exception.Message) 'WARN' }
    }

    if (Test-Path -LiteralPath $HpiaInstallerSharePath) {
        try {
            $localInstaller = Join-Path $WorkingRoot 'HPImageAssistant.exe'
            Copy-Item -LiteralPath $HpiaInstallerSharePath -Destination $localInstaller -Force -ErrorAction Stop
            return Expand-HpiaSoftPaq -InstallerPath $localInstaller -DestinationFolder $LocalHpiaFolder
        }
        catch { Write-Log ("HPIA share-installer deployment failed: {0}" -f $_.Exception.Message) 'WARN' }
    }

    Write-Log 'Downloading the latest HP Image Assistant SoftPaq from HP...' 'INFO'
    Add-YamlAction 'Falling back to an Internet download of HP Image Assistant.'
    $downloadUrl = Get-LatestHpiaSoftPaqUrl -LandingPage $HpiaInternetLandingPage
    $downloadPath = Join-Path $WorkingRoot 'HPImageAssistant-Internet.exe'
    Invoke-WebRequest -Uri $downloadUrl -OutFile $downloadPath -UseBasicParsing -TimeoutSec 300 -ErrorAction Stop
    if (-not (Test-Path -LiteralPath $downloadPath) -or (Get-Item -LiteralPath $downloadPath).Length -lt 1MB) {
        throw 'The downloaded HPIA installer is missing or unexpectedly small.'
    }
    return Expand-HpiaSoftPaq -InstallerPath $downloadPath -DestinationFolder $LocalHpiaFolder
}

function Get-HpiaExitStatus {
    param([int]$ExitCode)

    switch ($ExitCode) {
        0    { return 'success' }
        1    { return 'failed' }
        2    { return 'cancelled' }
        3    { return 'needs_reboot' }
        256  { return 'no_recommendations_or_success' }
        257  { return 'recommendations_found' }
        3010 { return 'needs_reboot' }
        3011 { return 'not_auto_installable_skipped' }
        3020 { return 'one_or_more_installations_failed' }
        4096 { return 'no_applicable_updates_or_platform_not_supported' }
        4097 { return 'invalid_parameters' }
        4098 { return 'no_internet_connection' }
        4099 { return 'invalid_softpaq_number_in_splist' }
        4100 { return 'softpaq_product_list_empty' }
        4101 { return 'parameter_no_longer_supported' }
        4102 { return 'tls_1_2_or_higher_required' }
        4103 { return 'remote_request_incomplete' }
        4104 { return 'generic_os_reference_used' }
        8193 { return 'hpia_analysis_or_report_generation_error' }
        default { return 'unknown' }
    }
}

function Get-HpiaRecommendationObjects {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ReportFolder,
        [datetime]$NotOlderThan = [datetime]::MinValue
    )

    $recommendations = New-Object System.Collections.Generic.List[object]

    function Add-HpiaCandidate {
        param(
            [Parameter(Mandatory)]$Value,
            [Parameter(Mandatory)][string]$SourceReport,
            [Parameter(Mandatory)][string]$SourceFormat
        )

        if ($null -eq $Value) { return }

        $raw = try {
            if ($Value -is [System.Xml.XmlNode]) { $Value.OuterXml }
            else { $Value | ConvertTo-Json -Depth 12 -Compress -ErrorAction Stop }
        }
        catch { $Value | Out-String }

        if ([string]::IsNullOrWhiteSpace($raw)) { return }
        if ($raw -notmatch '(?i)\bSP\s*[-:]?\s*\d{5,6}\b|SoftPaq') { return }

        $recommendations.Add([pscustomobject]@{
            Data         = $Value
            SourceReport = $SourceReport
            SourceFormat = $SourceFormat
            RawSummary   = (($raw -replace "`r|`n", ' ') -replace '\s+', ' ').Trim()
        }) | Out-Null
    }

    function Walk-HpiaJson {
        param(
            [AllowNull()]$Value,
            [Parameter(Mandatory)][string]$SourceReport,
            [int]$Depth = 0
        )

        if ($null -eq $Value -or $Depth -gt 20) { return }

        if ($Value -is [string] -or $Value -is [ValueType]) { return }

        if ($Value -is [System.Collections.IEnumerable] -and $Value -isnot [pscustomobject] -and $Value -isnot [System.Collections.IDictionary]) {
            foreach ($item in @($Value)) { Walk-HpiaJson -Value $item -SourceReport $SourceReport -Depth ($Depth + 1) }
            return
        }

        $propertyNames = @($Value.PSObject.Properties.Name)
        $hasSoftPaqProperty = @($propertyNames | Where-Object { $_ -match '(?i)softpaq|soft.?paq|spnumber|spid' }).Count -gt 0
        $serialized = try { $Value | ConvertTo-Json -Depth 4 -Compress -ErrorAction Stop } catch { '' }
        if ($hasSoftPaqProperty -or $serialized -match '(?i)\bSP\s*[-:]?\s*\d{5,6}\b') {
            Add-HpiaCandidate -Value $Value -SourceReport $SourceReport -SourceFormat 'JSON'
        }

        foreach ($property in $Value.PSObject.Properties) {
            Walk-HpiaJson -Value $property.Value -SourceReport $SourceReport -Depth ($Depth + 1)
        }
    }

    $jsonFiles = @(Get-ChildItem -Path $ReportFolder -Filter '*.json' -Recurse -File -ErrorAction SilentlyContinue)
    foreach ($jsonFile in $jsonFiles) {
        try {
            $json = Get-Content -LiteralPath $jsonFile.FullName -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            Write-Log ("Parsing HPIA JSON report recursively: {0}" -f $jsonFile.FullName) 'INFO'
            Add-YamlAction ("Parsing HPIA JSON report recursively: {0}" -f $jsonFile.FullName)
            Walk-HpiaJson -Value $json -SourceReport $jsonFile.FullName
        }
        catch {
            Write-Log ("Unable to parse HPIA JSON report {0}: {1}" -f $jsonFile.FullName, $_.Exception.Message) 'WARN'
            Add-YamlAction ("Unable to parse HPIA JSON report {0}: {1}" -f $jsonFile.FullName, $_.Exception.Message)
        }
    }

    $xmlFiles = @(Get-ChildItem -Path $ReportFolder -Filter '*.xml' -Recurse -File -ErrorAction SilentlyContinue)
    foreach ($xmlFile in $xmlFiles) {
        try {
            [xml]$xml = Get-Content -LiteralPath $xmlFile.FullName -Raw -ErrorAction Stop
            Write-Log ("Parsing HPIA XML report for SoftPaq-bearing elements: {0}" -f $xmlFile.FullName) 'INFO'
            Add-YamlAction ("Parsing HPIA XML report for SoftPaq-bearing elements: {0}" -f $xmlFile.FullName)

            $nodes = @($xml.SelectNodes('//*[' +
                'contains(translate(normalize-space(string(.)), "abcdefghijklmnopqrstuvwxyz", "ABCDEFGHIJKLMNOPQRSTUVWXYZ"), "SP")' +
                ' or @*[contains(translate(., "abcdefghijklmnopqrstuvwxyz", "ABCDEFGHIJKLMNOPQRSTUVWXYZ"), "SP")]' +
                ']'))

            foreach ($node in $nodes) {
                $ownText = $node.OuterXml
                if ($ownText -match '(?i)\bSP\s*[-:]?\s*\d{5,6}\b') {
                    # Prefer the smallest element containing the SoftPaq record; skip containers whose child also qualifies.
                    $qualifyingChild = @($node.ChildNodes | Where-Object {
                        $_ -is [System.Xml.XmlElement] -and $_.OuterXml -match '(?i)\bSP\s*[-:]?\s*\d{5,6}\b'
                    }).Count -gt 0
                    if (-not $qualifyingChild) {
                        Add-HpiaCandidate -Value $node -SourceReport $xmlFile.FullName -SourceFormat 'XML'
                    }
                }
            }
        }
        catch {
            Write-Log ("Unable to parse HPIA XML report {0}: {1}" -f $xmlFile.FullName, $_.Exception.Message) 'WARN'
            Add-YamlAction ("Unable to parse HPIA XML report {0}: {1}" -f $xmlFile.FullName, $_.Exception.Message)
        }
    }

    # Raw-text fallback: HP has changed HPIA report schemas between releases. Search every
    # generated JSON, XML, HTML, and log file for literal SPxxxxx references so a schema
    # change cannot silently hide applicable recommendations.
    $reportFiles = @(Get-ChildItem -Path $ReportFolder -Recurse -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Extension -in @('.json','.xml','.html','.htm','.log','.txt','.csv') })

    foreach ($reportFile in $reportFiles) {
        try {
            $rawReport = Get-Content -LiteralPath $reportFile.FullName -Raw -ErrorAction Stop
            $fileLength = $rawReport.Length
            $spMatches = [regex]::Matches($rawReport, '(?i)\bSP\s*[-_:]?\s*(\d{5,6})\b')
            $isCurrent = ($reportFile.LastWriteTime -ge $NotOlderThan)
            $isReadme = ($reportFile.Name -match '(?i)^Readme')
            $isDebugLog = ($reportFile.Name -match '(?i)^HP Image Assistant\.log$')
            $eligibleCandidateSource = ($isCurrent -and -not $isReadme -and -not $isDebugLog -and $reportFile.Extension -in @('.json','.xml','.html','.htm','.csv'))
            $sha256 = try { (Get-FileHash -LiteralPath $reportFile.FullName -Algorithm SHA256 -ErrorAction Stop).Hash } catch { 'unavailable' }

            Write-Log ("HPIA report inventory | File: {0} | Bytes: {1} | Modified: {2} | SHA256: {3} | SP references: {4} | Candidate source: {5}" -f $reportFile.FullName, $reportFile.Length, $reportFile.LastWriteTime.ToString('s'), $sha256, $spMatches.Count, $eligibleCandidateSource) 'INFO'
            Add-YamlAction ("HPIA report inventory. File={0}; Bytes={1}; Modified={2}; SHA256={3}; Characters={4}; SP references={5}; CurrentRun={6}; Readme={7}; DebugLog={8}; CandidateSource={9}" -f $reportFile.FullName, $reportFile.Length, $reportFile.LastWriteTime.ToString('s'), $sha256, $fileLength, $spMatches.Count, $isCurrent, $isReadme, $isDebugLog, $eligibleCandidateSource)

            if (-not $eligibleCandidateSource) {
                continue
            }

            $seenInFile = @{}
            foreach ($match in $spMatches) {
                $spNumber = $match.Groups[1].Value
                if ($seenInFile.ContainsKey($spNumber)) { continue }
                $seenInFile[$spNumber] = $true

                $contextStart = [Math]::Max(0, $match.Index - 350)
                $contextLength = [Math]::Min(900, $rawReport.Length - $contextStart)
                $context = $rawReport.Substring($contextStart, $contextLength)
                $context = (($context -replace "`r|`n", ' ') -replace '\s+', ' ').Trim()

                # A recommendation report can include current-state/status rows. Do not turn
                # explicit Pass/Installed/Current/No Action records into install candidates.
                if ($context -match '(?i)>\s*(Pass|Installed|Current|Up[- ]?to[- ]?date|No Action)\s*<' -or
                    $context -match '(?i)\b(Status|Result)\s*[:=]\s*["'']?(Pass|Installed|Current|Up[- ]?to[- ]?date|No Action)') {
                    Add-YamlAction ("Rejected raw SoftPaq context as non-recommendation status. SP{0}; Source={1}; Context={2}" -f $spNumber, $reportFile.FullName, $context)
                    continue
                }

                Add-HpiaCandidate -Value ([pscustomobject]@{
                    SoftPaqId = "SP$spNumber"
                    Name       = $context
                    Category   = 'Raw recommendation-report fallback'
                }) -SourceReport $reportFile.FullName -SourceFormat ('RAW-' + $reportFile.Extension.TrimStart('.').ToUpperInvariant())
            }
        }
        catch {
            Write-Log ("Unable to inventory/raw-scan HPIA report {0}: {1}" -f $reportFile.FullName, $_.Exception.Message) 'WARN'
            Add-YamlAction ("Unable to inventory/raw-scan HPIA report {0}: {1}" -f $reportFile.FullName, $_.Exception.Message)
        }
    }

    Write-Log ("HPIA parser produced {0} SoftPaq-bearing candidate record(s)." -f $recommendations.Count) 'INFO'
    Add-YamlAction ("HPIA parser produced {0} SoftPaq-bearing candidate record(s)." -f $recommendations.Count)

    # PowerShell 5.1 can throw 'Argument types do not match' when array-subexpression
    # syntax is used directly against a generic List[object]. Copy items into a normal
    # PowerShell array explicitly instead.
    $result = @()
    foreach ($item in $recommendations) { $result += $item }
    return $result
}

function Get-HpiaRecommendationValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Recommendation,
        [Parameter(Mandatory)][string[]]$PropertyNames
    )

    if ($Recommendation.PSObject.Properties.Name -contains 'Data') {
        $Recommendation = $Recommendation.Data
    }

    foreach ($prop in $PropertyNames) {
        try {
            if ($Recommendation.PSObject.Properties.Name -contains $prop) {
                $value = $Recommendation.$prop
                if ($null -ne $value -and -not [string]::IsNullOrWhiteSpace($value.ToString())) {
                    return $value.ToString()
                }
            }
        }
        catch {}
    }

    # XML fallback
    try {
        foreach ($prop in $PropertyNames) {
            $node = $Recommendation.SelectSingleNode('.//*[local-name()="' + $prop + '"]')
            if ($node -and -not [string]::IsNullOrWhiteSpace($node.InnerText)) {
                return $node.InnerText.Trim()
            }
        }
    }
    catch {}

    return $null
}

function Get-HpiaSoftPaqNumber {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Recommendation)

    $data = if ($Recommendation.PSObject.Properties.Name -contains 'Data') { $Recommendation.Data } else { $Recommendation }

    $candidate = Get-HpiaRecommendationValue -Recommendation $Recommendation -PropertyNames @(
        'SoftPaqId','SoftpaqId','SoftPaq','Softpaq','SoftPaqNumber','SoftpaqNumber','SP','Id','ID','Number'
    )

    if ($candidate -match '(?i)sp?(\d{5,6})') {
        return $matches[1]
    }

    $text = if ($data -is [System.Xml.XmlNode]) { $data.OuterXml } else { $data | Out-String }
    if ($text -match '(?i)sp(\d{5,6})') {
        return $matches[1]
    }

    if ($text -match '\b(\d{5,6})\b') {
        return $matches[1]
    }

    return $null
}

function Get-HpiaRecommendationCategory {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Recommendation)

    return (Get-HpiaRecommendationValue -Recommendation $Recommendation -PropertyNames @(
        'Category','Type','RecommendationType','ComponentType','Class','Group'
    ))
}

function Get-HpiaRecommendationName {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Recommendation)

    $data = if ($Recommendation.PSObject.Properties.Name -contains 'Data') { $Recommendation.Data } else { $Recommendation }

    $name = Get-HpiaRecommendationValue -Recommendation $Recommendation -PropertyNames @(
        'Name','Title','Component','ComponentName','Description','SoftPaqName','SoftpaqName'
    )

    if ($name) { return $name }

    $text = if ($data -is [System.Xml.XmlNode]) { $data.OuterXml } else { ($data | Out-String).Trim() }
    if ($text.Length -gt 160) {
        return $text.Substring(0,160)
    }

    return $text
}


function Test-HpiaBlockedUnattendedDriver {
    [CmdletBinding()]
    param(
        [AllowNull()][string]$Category,
        [AllowNull()][string]$Name
    )

    $combined = ("{0} {1}" -f $Category, $Name)

    # These driver families can affect boot-critical storage/controller behavior.
    # They are intentionally blocked from unattended HPIA installs.
    $blockedPattern = '(?i)\b(storage|chipset|rst|rapid storage|intel rapid|vmd|volume management device|nvme|nvm express|sata|raid|ahci|scsi|controller|management engine|mei|thunderbolt)\b'

    return ($combined -match $blockedPattern)
}

function New-HpiaDriverSPList {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ReportFolder,
        [Parameter(Mandatory)][string]$SPListPath,
        [datetime]$NotOlderThan = [datetime]::MinValue
    )

    $recommendations = @(Get-HpiaRecommendationObjects -ReportFolder $ReportFolder -NotOlderThan $NotOlderThan)
    Write-Log ("HPIA recommendation candidates parsed from reports: {0}" -f $recommendations.Count) 'INFO'
    Add-YamlAction ("HPIA recommendation candidates parsed from reports: {0}" -f $recommendations.Count)

    $selected = @()
    $seen = @{}
    $unparsed = 0

    foreach ($rec in $recommendations) {
        $category = Get-HpiaRecommendationCategory -Recommendation $rec
        $name = Get-HpiaRecommendationName -Recommendation $rec
        $sp = Get-HpiaSoftPaqNumber -Recommendation $rec
        $sourceReport = if ($rec.PSObject.Properties.Name -contains 'SourceReport') { [string]$rec.SourceReport } else { '' }
        $sourceFormat = if ($rec.PSObject.Properties.Name -contains 'SourceFormat') { [string]$rec.SourceFormat } else { '' }
        $rawSummary = if ($rec.PSObject.Properties.Name -contains 'RawSummary') { [string]$rec.RawSummary } else { '' }
        if ($rawSummary.Length -gt 1200) { $rawSummary = $rawSummary.Substring(0,1200) }

        if ([string]::IsNullOrWhiteSpace($name)) { $name = '(name not supplied by HPIA report)' }
        if ([string]::IsNullOrWhiteSpace($category)) { $category = '(category not supplied)' }

        if (-not $sp) {
            $unparsed++
            $detail = "Parser candidate did not yield a SoftPaq ID. Source format: $sourceFormat"
            Write-Log ("HPIA unparsed candidate | Name: {0} | Category: {1} | Source: {2}" -f $name, $category, $sourceReport) 'WARN'
            Add-YamlAction ("HPIA unparsed candidate. Name={0}; Category={1}; Source={2}; Raw={3}" -f $name, $category, $sourceReport, $rawSummary)
            Add-DriverResult -Vendor 'HP' -Name $name -Id '' -Category $category -Status 'Skipped' -Message $detail -SourceReport $sourceReport -RawSummary $rawSummary
            continue
        }

        $combined = ("{0} {1}" -f $category, $name)
        $decision = ''
        $status = 'Detected'

        if (-not $script:InstallBiosFirmware -and $combined -match '(?i)\bBIOS\b|Firmware') {
            $decision = 'BLOCKED: BIOS/Firmware is excluded in normal mode. Use -IncludeBIOS or -All.'
            $status = 'Blocked'
        }
        elseif (-not $script:InstallAllUpdates -and (Test-HpiaBlockedUnattendedDriver -Category $category -Name $name)) {
            $decision = 'BLOCKED: Storage/chipset/RST/VMD/NVMe/controller family is excluded in normal mode. Use -All.'
            $status = 'Blocked'
        }
        elseif (-not $script:InstallAllUpdates -and $combined -notmatch '(?i)Driver|Bluetooth|Audio|Graphics|Video|LAN|WLAN|Wireless|NIC|Touch|Fingerprint|Card Reader|Serial|USB|BIOS|Firmware' -and $category) {
            $decision = 'SKIPPED: Recommendation did not match the safe unattended allow-list. Use -All.'
            $status = 'Skipped'
        }
        elseif ($seen.ContainsKey($sp)) {
            $decision = 'SKIPPED: Duplicate SoftPaq already processed from another report/record.'
            $status = 'Skipped'
        }
        else {
            $seen[$sp] = $true
            $selected += $sp
            $decision = 'SELECTED: Added to HPIA SPList for installation.'
        }

        $logLine = "HPIA recommendation | SP$sp | Category: $category | Name: $name | Decision: $decision | Source: $sourceReport"
        Write-Log $logLine $(if ($status -eq 'Blocked') { 'WARN' } elseif ($status -eq 'Skipped') { 'INFO' } else { 'OK' })
        Add-YamlAction ("{0}; Format={1}; Raw={2}" -f $logLine, $sourceFormat, $rawSummary)
        Add-DriverResult -Vendor 'HP' -Name $name -Id $sp -Category $category -Status $status -Message $decision -SourceReport $sourceReport -RawSummary $rawSummary

        if ($status -in @('Blocked','Skipped')) {
            $script:SkippedList.Add(("SP{0} {1} -- {2}" -f $sp, $name, $decision)) | Out-Null
        }
    }

    if (@($selected).Count -gt 0) {
        Set-Content -LiteralPath $SPListPath -Value $selected -Encoding ASCII
        Write-Log ("Created HPIA SPList with {0} SoftPaq(s): {1}" -f @($selected).Count, $SPListPath) 'OK'
        Write-Log ("Selected SoftPaq IDs: {0}" -f ((@($selected) | ForEach-Object { "SP$_" }) -join ', ')) 'INFO'
        Add-YamlAction ("Created HPIA SPList with {0} SoftPaq(s): {1}" -f @($selected).Count, $SPListPath)
        Add-YamlAction ("Selected SoftPaq IDs: {0}" -f ((@($selected) | ForEach-Object { "SP$_" }) -join ', '))
    }
    else {
        $message = if ($recommendations.Count -gt 0 -and $unparsed -eq $recommendations.Count) {
            'HPIA reports contained candidate records, but no installable SoftPaq IDs could be extracted. Review the per-candidate parser diagnostics in this YAML log.'
        }
        elseif ($script:InstallAllUpdates) {
            'No unique installable HPIA SoftPaq recommendations were selected while -All was active.'
        }
        elseif ($script:InstallBiosFirmware) {
            'No installable HPIA SoftPaq recommendations were selected while BIOS/Firmware inclusion was active.'
        }
        else {
            'No safe HPIA SoftPaq recommendations were selected under the normal unattended policy.'
        }
        Write-Log $message 'WARN'
        Add-YamlAction $message
    }

    return @($selected)
}



# -----------------------------
# HP PowerShell module maintenance
# -----------------------------
function Initialize-NetworkDefaults {
    try {
        [Net.ServicePointManager]::SecurityProtocol =
            [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls13
    }
    catch {
        try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 }
        catch {}
    }

    try { [Net.ServicePointManager]::DefaultConnectionLimit = 64 }
    catch {}
}

function Get-HighestInstalledModuleVersion {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Name
    )

    try {
        $installed = Get-Module -Name $Name -ListAvailable -ErrorAction SilentlyContinue |
            Sort-Object Version -Descending |
            Select-Object -First 1

        if ($installed) { return [version]$installed.Version }
    }
    catch {
        Write-Log "Could not determine installed version for module ${Name}: $($_.Exception.Message)" 'WARN'
    }

    return $null
}

function Invoke-HPCMSLUpdateInFreshPowerShell {
    [CmdletBinding()]
    param(
        [ValidateSet('CurrentUser','AllUsers')]
        [string]$Scope = 'AllUsers'
    )

    Write-Log 'Starting isolated HPCMSL install/update using PowerShellGet/PSResourceGet when available...' 'INFO'
    Add-YamlAction 'Starting isolated HPCMSL install/update before HP driver updates.'

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

    if ([string]::IsNullOrWhiteSpace(`$InstallRoot)) { throw 'InstallRoot is blank.' }
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

    if (`$installed) { HelperLog ('Installed HPCMSL version: {0}' -f `$installed) }
    else { HelperLog 'HPCMSL is not currently installed.' }
    HelperLog ('Latest PSGallery HPCMSL version: {0}' -f `$online)

    if (`$installed -and `$installed -ge `$online) {
        HelperLog 'HPCMSL is already current. No update needed.'
        exit 0
    }

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
        if ('$Scope' -eq 'AllUsers') { `$installRoot = Join-Path `$env:ProgramFiles 'WindowsPowerShell\Modules' }
        else { `$installRoot = Join-Path `$HOME 'Documents\WindowsPowerShell\Modules' }
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
        Initialize-NetworkDefaults
        Set-Content -Path $helperPath -Value $helperScript -Encoding UTF8 -Force
        Write-Log "HPCMSL helper script created at $helperPath" 'INFO'

        $args = @('-NoProfile','-ExecutionPolicy','Bypass','-File',('"{0}"' -f $helperPath))
        $process = Start-Process -FilePath powershell.exe -ArgumentList ($args -join ' ') -Wait -PassThru -WindowStyle Hidden

        if (Test-Path -LiteralPath $helperLog) {
            Get-Content -LiteralPath $helperLog -ErrorAction SilentlyContinue | ForEach-Object {
                if (-not [string]::IsNullOrWhiteSpace($_)) { Write-Log "HPCMSL helper: $_" 'INFO' }
            }
        }

        if ($process.ExitCode -ne 0) {
            Write-Log "Isolated HPCMSL helper failed with exit code $($process.ExitCode). Continuing with HPIA driver updates." 'WARN'
            Add-YamlAction ("Isolated HPCMSL helper failed with exit code {0}. Continuing with HPIA driver updates." -f $process.ExitCode)
            return $false
        }

        $finalVersion = Get-HighestInstalledModuleVersion -Name 'HPCMSL'
        if ($finalVersion) {
            Write-Log "HPCMSL final installed version after isolated update: $finalVersion" 'OK'
            Add-YamlAction ("HPCMSL final installed version after isolated update: {0}" -f $finalVersion)
        }
        else {
            Write-Log 'HPCMSL helper completed, but this session could not verify the installed HPCMSL version.' 'WARN'
            Add-YamlAction 'HPCMSL helper completed, but installed version could not be verified.'
            return $false
        }

        return $true
    }
    catch {
        Write-Log "Failed to run isolated HPCMSL update helper: $($_.Exception.Message)" 'WARN'
        Add-YamlAction ("Failed to run isolated HPCMSL update helper: {0}" -f $_.Exception.Message)
        return $false
    }
    finally {
        Remove-Item -LiteralPath $helperPath -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $helperLog -Force -ErrorAction SilentlyContinue
    }
}

function Invoke-HPPowerShellModuleMaintenance {
    [CmdletBinding()]
    param()

    Write-Section 'HP PowerShell Module Maintenance'
    Write-Log 'HP vendor detected. Checking/installing HPCMSL before HP driver updates...' 'INFO'
    Add-YamlAction 'HP vendor detected; checking/installing HPCMSL before HP driver updates.'

    $hpcmslUpdated = Invoke-HPCMSLUpdateInFreshPowerShell -Scope AllUsers
    if ($hpcmslUpdated) {
        Write-Log 'HP PowerShell module maintenance completed.' 'OK'
        Add-YamlAction 'HP PowerShell module maintenance completed.'
    }
    else {
        Write-Log 'HP PowerShell module maintenance did not complete successfully. Continuing with HP Image Assistant driver updates.' 'WARN'
        Add-YamlAction 'HP PowerShell module maintenance did not complete successfully; continuing with HPIA driver updates.'
    }
}

function Invoke-HPDriverUpdates {
    Write-Section 'HP Driver Analysis and Filtered Installation'

    $policyName = if ($script:InstallAllUpdates) { 'ALL updates permitted (-All)' } elseif ($script:InstallBiosFirmware) { 'BIOS/Firmware permitted (-IncludeBIOS)' } else { 'Safe unattended drivers only (default)' }
    Write-Log ("HP update policy: {0}" -f $policyName) 'INFO'
    Write-Log ("BIOS/Firmware allowed: {0}; Storage/Chipset/Controller allowed: {1}; Software requested: {2}" -f $script:InstallBiosFirmware, $script:InstallAllUpdates, ([bool]$IncludeSoftware)) 'INFO'
    Add-YamlAction ("HP update policy: {0}" -f $policyName)
    Add-YamlAction ("BIOS/Firmware allowed: {0}; Storage/Chipset/Controller allowed: {1}; Software requested: {2}" -f $script:InstallBiosFirmware, $script:InstallAllUpdates, ([bool]$IncludeSoftware))

    $hpInfo = Get-HPSystemModelInfo

    $hpiaExe = Ensure-HPIAAvailable

    $hpiaReportRoot = Join-Path $YamlLogFolder 'HPIA'
    $hpiaRunStamp = $script:StartTime.ToString('yyyy-MM-dd_HHmmss')
    $hpiaReportFolder = Join-Path $hpiaReportRoot ("{0}_{1}" -f $script:ComputerName, $hpiaRunStamp)
    $hpiaDownloadFolder = Join-Path $WorkingRoot 'HPIADownloads'
    $hpiaExtractFolder = Join-Path $WorkingRoot 'HPIAExtracted'
    $hpiaSPListPath = Join-Path $WorkingRoot 'HPIA-Filtered-Updates.splist'

    Ensure-Folder -Path $hpiaReportFolder
    Add-YamlAction ("HPIA per-run report folder: {0}" -f $hpiaReportFolder)
    Write-Log ("HPIA reports for this run will be isolated in: {0}" -f $hpiaReportFolder) 'INFO'
    Ensure-Folder -Path $hpiaDownloadFolder
    Ensure-Folder -Path $hpiaExtractFolder
    Ensure-Folder -Path 'C:\Temp'

    # HPIA is more reliable under Task Scheduler/SYSTEM when TEMP/TMP are local and TLS 1.2 is forced.
    $env:TEMP = 'C:\Temp'
    $env:TMP  = 'C:\Temp'

    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Write-Log 'TLS 1.2 enabled for HP Image Assistant network operations.' 'INFO'
        Add-YamlAction 'TLS 1.2 enabled for HPIA network operations.'
    }
    catch {
        Write-Log ("Unable to force TLS 1.2 before HPIA run: {0}" -f $_.Exception.Message) 'WARN'
        Add-YamlAction ("Unable to force TLS 1.2 before HPIA run: {0}" -f $_.Exception.Message)
    }

    # Safer unattended HPIA workflow:
    # 1. Run /Action:List first to generate recommendations only.
    # 2. Parse the HPIA report and create a filtered SPList.
    # 3. Install only the filtered SPList.
    #
    # Blocked unattended families include BIOS, Firmware, Storage, Chipset,
    # Intel RST, Intel VMD, NVMe, RAID/AHCI/SATA/controller, MEI/Management Engine,
    # and Thunderbolt because those can affect boot or storage-controller behavior.
    Write-Log 'Running HP Image Assistant recommendation/list pass only. No drivers will be installed in this step.' 'INFO'
    Add-YamlAction 'Running HPIA /Action:List first so recommendations can be filtered before install.'

    $hpiaCategory = if ($script:InstallBiosFirmware -or $IncludeSoftware -or $script:InstallAllUpdates) { 'All' } else { 'Drivers' }

    $hpiaListArgs = @(
        '/Operation:Analyze',
        '/Action:List',
        ("/Category:{0}" -f $hpiaCategory),
        '/Selection:All',
        '/Silent',
        '/Noninteractive',
        '/Debug',
        "/ReportFolder:`"$hpiaReportFolder`"",
        "/SoftpaqDownloadFolder:`"$hpiaDownloadFolder`"",
        "/SoftpaqExtractFolder:`"$hpiaExtractFolder`""
    )

    Write-Log ("HPIA list command: {0} {1}" -f $hpiaExe, ($hpiaListArgs -join ' ')) 'INFO'
    Add-YamlAction ("HPIA list command: {0} {1}" -f $hpiaExe, ($hpiaListArgs -join ' '))

    Write-Progress -Activity 'HP Image Assistant' -Status ("Listing recommended drivers for {0}" -f $hpInfo.Model) -PercentComplete 25
    $hpiaListProc = Start-Process -FilePath $hpiaExe -ArgumentList ($hpiaListArgs -join ' ') -Wait -PassThru -NoNewWindow
    $listExitCode = [int]$hpiaListProc.ExitCode
    $listStatus = Get-HpiaExitStatus -ExitCode $listExitCode
    Write-Progress -Activity 'HP Image Assistant' -Completed

    Write-Log ("HP Image Assistant list pass completed with exit code {0} ({1})." -f $listExitCode, $listStatus) 'INFO'
    Add-YamlAction ("HPIA list pass completed with exit code {0} ({1})." -f $listExitCode, $listStatus)

    $reportFiles = @(Get-ChildItem -Path $hpiaReportFolder -File -Recurse -ErrorAction SilentlyContinue)
    foreach ($report in $reportFiles) {
        Add-YamlAction ("HPIA report generated: {0}" -f $report.FullName)
    }

    if ($reportFiles.Count -eq 0 -and $listExitCode -notin @(0,256,257,3010,3011,4096)) {
        throw "HPIA list pass failed or generated no reports. Exit code: $listExitCode ($listStatus). Review $hpiaReportFolder."
    }

    $selectedSoftPaqs = @(New-HpiaDriverSPList -ReportFolder $hpiaReportFolder -SPListPath $hpiaSPListPath -NotOlderThan $script:StartTime)

    if ($selectedSoftPaqs.Count -eq 0) {
        Write-Log 'No HP SoftPaqs were selected for installation. See the YAML driver records and actions for every parser and filtering decision.' 'INFO'
        Add-YamlAction 'No HP SoftPaqs were selected for installation. See driver records for complete parser and filtering decisions.'
        return
    }

    Write-Log ("Installing selected HP SoftPaq SPList. Count: {0}" -f $selectedSoftPaqs.Count) 'INFO'
    Add-YamlAction ("Installing selected HP SoftPaq SPList: {0}" -f $hpiaSPListPath)

    $hpiaInstallArgs = @(
        '/Operation:Analyze',
        '/Action:Install',
        ("/Category:{0}" -f $hpiaCategory),
        '/Selection:All',
        "/SPList:`"$hpiaSPListPath`"",
        '/Silent',
        '/Noninteractive',
        '/AutoCleanup',
        '/Debug',
        "/ReportFolder:`"$hpiaReportFolder`"",
        "/SoftpaqDownloadFolder:`"$hpiaDownloadFolder`"",
        "/SoftpaqExtractFolder:`"$hpiaExtractFolder`""
    )

    Write-Log ("HPIA filtered SPList install command: {0} {1}" -f $hpiaExe, ($hpiaInstallArgs -join ' ')) 'INFO'
    Add-YamlAction ("HPIA filtered SPList install command: {0} {1}" -f $hpiaExe, ($hpiaInstallArgs -join ' '))

    Write-Progress -Activity 'HP Image Assistant' -Status 'Installing filtered safe driver SoftPaqs' -PercentComplete 75
    $hpiaInstallProc = Start-Process -FilePath $hpiaExe -ArgumentList ($hpiaInstallArgs -join ' ') -Wait -PassThru -NoNewWindow
    $installExitCode = [int]$hpiaInstallProc.ExitCode
    $installStatus = Get-HpiaExitStatus -ExitCode $installExitCode
    Write-Progress -Activity 'HP Image Assistant' -Completed

    Write-Log ("HP Image Assistant filtered SPList install completed with exit code {0} ({1})." -f $installExitCode, $installStatus) 'INFO'
    Add-YamlAction ("HPIA filtered SPList install completed with exit code {0} ({1})." -f $installExitCode, $installStatus)

    $downloadedFiles = @(Get-ChildItem -Path $hpiaDownloadFolder -File -Recurse -ErrorAction SilentlyContinue)
    foreach ($file in $downloadedFiles) {
        Add-DriverResult -Vendor 'HP' -Name $file.Name -Id $null -Category 'Driver' -Status 'Detected' -Message ("Downloaded/processed by filtered HPIA SPList: {0}" -f $file.FullName)
    }

    switch ($installExitCode) {
        0 {
            Write-Log 'HPIA filtered SPList install completed successfully.' 'OK'
            Add-YamlAction 'HPIA filtered SPList install completed successfully.'
            return
        }
        256 {
            Write-Log 'HPIA completed successfully. No applicable filtered driver recommendations were found or no action was required.' 'OK'
            Add-YamlAction 'HPIA completed successfully with no applicable filtered driver recommendations or no action required.'
            return
        }
        257 {
            Write-Log 'HPIA completed and reported filtered driver actions.' 'OK'
            Add-YamlAction 'HPIA completed and reported filtered driver actions.'
            return
        }
        3010 {
            Write-Log 'HPIA filtered SPList install completed successfully. Reboot required.' 'WARN'
            Add-YamlAction 'HPIA filtered SPList install completed successfully and indicated reboot required.'
            return
        }
        3011 {
            Write-Log 'One or more filtered HPIA items were not auto-installable and were skipped.' 'WARN'
            Add-YamlAction 'One or more filtered HPIA items were not auto-installable and were skipped.'
            return
        }
        4096 {
            Write-Log 'HPIA completed but did not find applicable filtered driver updates for this platform.' 'OK'
            Add-YamlAction 'HPIA completed but did not find applicable filtered driver updates for this platform.'
            return
        }
        8193 {
            Write-Log 'HPIA filtered install returned 8193. Checking whether reports/logs were generated before failing the run.' 'WARN'
            Add-YamlAction 'HPIA filtered install returned 8193; checking for reports/logs before hard failure.'
            $generatedReports = @(Get-ChildItem -Path $hpiaReportFolder -File -Recurse -ErrorAction SilentlyContinue)
            if ($generatedReports.Count -gt 0) {
                Write-Log ("HPIA generated {0} report/log file(s) despite exit code 8193. Continuing so the scheduled workflow does not hard fail." -f $generatedReports.Count) 'WARN'
                Add-YamlAction ("HPIA generated {0} report/log file(s) despite exit code 8193. Continuing without hard failure." -f $generatedReports.Count)
                return
            }
            throw "HPIA failed with exit code 8193 and produced no report/log files in $hpiaReportFolder."
        }
        default {
            throw "HP Image Assistant filtered SPList install failed or returned an unexpected exit code: $installExitCode ($installStatus). Review reports in $hpiaReportFolder."
        }
    }
}

# -----------------------------
# Dell support
# -----------------------------
function Wait-ForFile {
    param(
        [Parameter(Mandatory)][string]$Path,
        [int]$TimeoutSeconds = 30
    )

    $start = Get-Date
    while (-not (Test-Path -LiteralPath $Path)) {
        Start-Sleep -Seconds 1
        if (((Get-Date) - $start).TotalSeconds -ge $TimeoutSeconds) {
            return $false
        }
    }
    return $true
}

function Get-DcuReportXmlSafely {
    param(
        [Parameter(Mandatory)][string]$Path,
        [int]$RetryCount = 5,
        [int]$RetryDelaySeconds = 2
    )

    if (-not (Wait-ForFile -Path $Path -TimeoutSeconds 20)) {
        throw "Dell DCU report file was not found: $Path"
    }

    Start-Sleep -Seconds 3

    for ($i = 1; $i -le $RetryCount; $i++) {
        try {
            $fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
            try {
                $sr = New-Object System.IO.StreamReader($fs)
                try {
                    $content = $sr.ReadToEnd()
                }
                finally {
                    $sr.Dispose()
                }
            }
            finally {
                $fs.Dispose()
            }

            $xml = New-Object System.Xml.XmlDocument
            $xml.LoadXml($content)
            return $xml
        }
        catch {
            if ($i -eq $RetryCount) {
                throw
            }
            Start-Sleep -Seconds $RetryDelaySeconds
        }
    }
}

function Get-DellNodeText {
    param(
        [Parameter(Mandatory)][xml]$Node,
        [Parameter(Mandatory)][string[]]$Names
    )

    foreach ($name in $Names) {
        try {
            $xpath = './/*[local-name()="' + $name + '"]'
            $child = $Node.SelectSingleNode($xpath)
            if ($child -and -not [string]::IsNullOrWhiteSpace($child.InnerText)) {
                return $child.InnerText.Trim()
            }
        }
        catch {}
    }

    return $null
}

function Get-DellReportItems {
    param([Parameter(Mandatory)][xml]$Xml)

    $items = @()
    try {
        $xpath = '//*[local-name()="Update" or local-name()="Package" or local-name()="SoftwareComponent" or local-name()="component" or local-name()="Device"]'
        $nodes = $Xml.SelectNodes($xpath)
        foreach ($node in $nodes) {
            $name = Get-DellNodeText -Node $node -Names @('Name','Title','PackageName')
            $version = Get-DellNodeText -Node $node -Names @('Version','PackageVersion')
            $category = Get-DellNodeText -Node $node -Names @('Category','Type')
            $id = Get-DellNodeText -Node $node -Names @('Id','PackageId','ReleaseId')

            if ($name -or $id) {
                $items += [pscustomobject]@{
                    Id       = $id
                    Name     = $name
                    Version  = $version
                    Category = $category
                }
            }
        }
    }
    catch {}

    return @($items)
}

function Get-DellDCUService {
    $candidateNames = @(
        'DellClientManagementService',
        'DellCommandUpdate',
        'DellUpdateService'
    )

    foreach ($name in $candidateNames) {
        $svc = Get-Service -Name $name -ErrorAction SilentlyContinue
        if ($svc) { return $svc }
    }

    $svc = Get-Service -ErrorAction SilentlyContinue |
        Where-Object { $_.DisplayName -match 'Dell.*Client.*Management|Dell.*Command.*Update' } |
        Select-Object -First 1

    return $svc
}

function Ensure-DellDCUService {
    param(
        [int]$TimeoutSeconds = 30
    )

    Write-Log 'Validating Dell Client Management Service...' 'INFO'
    Add-YamlAction 'Validating Dell Client Management Service.'

    $service = Get-DellDCUService
    if (-not $service) {
        throw 'Dell Client Management Service was not found. Dell Command | Update may need to be repaired or reinstalled.'
    }

    Write-Log ("Dell service detected: {0} ({1})" -f $service.DisplayName, $service.Name) 'OK'

    try {
        $wmiService = Get-CimInstance -ClassName Win32_Service -Filter ("Name='{0}'" -f $service.Name) -ErrorAction Stop
        if ($wmiService.StartMode -eq 'Disabled') {
            Write-Log 'Dell service startup type is Disabled. Setting it to Manual.' 'WARN'
            Set-Service -Name $service.Name -StartupType Manual -ErrorAction Stop
        }
    }
    catch {
        Write-Log ("Could not validate/set Dell service startup type: {0}" -f $_.Exception.Message) 'WARN'
    }

    $service.Refresh()
    if ($service.Status -ne 'Running') {
        Write-Log 'Starting Dell Client Management Service...' 'INFO'
        try {
            Start-Service -Name $service.Name -ErrorAction Stop
        }
        catch {
            Write-Log ("Start-Service failed. Attempting sc.exe recovery start. Error: {0}" -f $_.Exception.Message) 'WARN'
            & sc.exe start $service.Name | Out-Null
        }
    }

    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    do {
        Start-Sleep -Seconds 2
        $service = Get-Service -Name $service.Name -ErrorAction SilentlyContinue
        if ($service -and $service.Status -eq 'Running') {
            Write-Log 'Dell Client Management Service is running.' 'OK'
            Add-YamlAction 'Dell Client Management Service is running.'
            return $true
        }
    } while ((Get-Date) -lt $deadline)

    throw 'Dell Client Management Service did not reach the Running state before timeout.'
}

function Invoke-DellDCUCommandWithRetry {
    param(
        [Parameter(Mandatory)][string]$DcuCli,
        [Parameter(Mandatory)][string]$Arguments,
        [Parameter(Mandatory)][string]$OperationName,
        [int[]]$AcceptableExitCodes = @(0),
        [int]$MaxAttempts = 2
    )

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        Write-Log ("Dell DCU {0} attempt {1} of {2}..." -f $OperationName, $attempt, $MaxAttempts) 'INFO'
        Add-YamlAction ("Dell DCU {0} attempt {1} of {2}." -f $OperationName, $attempt, $MaxAttempts)

        $proc = Start-Process -FilePath $DcuCli -ArgumentList $Arguments -Wait -PassThru -NoNewWindow
        $exitCode = [int]$proc.ExitCode
        Write-Log ("Dell DCU {0} exit code: {1}" -f $OperationName, $exitCode) 'INFO'

        if ($AcceptableExitCodes -contains $exitCode) {
            return $proc
        }

        if ($exitCode -eq 3000) {
            Write-Log 'Dell DCU returned 3000, which normally indicates the Dell Client Management Service stopped or crashed.' 'WARN'
            Add-YamlAction 'Dell DCU returned 3000; attempting Dell service recovery before retry.'
        }
        else {
            Write-Log ("Dell DCU {0} returned non-success exit code {1}." -f $OperationName, $exitCode) 'WARN'
        }

        if ($attempt -lt $MaxAttempts) {
            Ensure-DellDCUService | Out-Null
            Start-Sleep -Seconds 5
            continue
        }

        throw "Dell DCU $OperationName failed after $MaxAttempts attempt(s). Last exit code: $exitCode"
    }
}


function Test-DotNetDesktopRuntime8 {
    [CmdletBinding()]
    param([version]$MinimumVersion = [version]'8.0.25')

    $dotnetExe = Join-Path $env:ProgramFiles 'dotnet\dotnet.exe'
    if (-not (Test-Path -LiteralPath $dotnetExe)) { return $false }
    try {
        foreach ($line in @(& $dotnetExe --list-runtimes 2>$null)) {
            if ($line -match '^Microsoft\.WindowsDesktop\.App\s+(?<Version>\d+\.\d+\.\d+)\s+') {
                try { if ([version]$Matches.Version -ge $MinimumVersion) { return $true } } catch {}
            }
        }
    } catch { Write-Log ("Unable to query installed .NET desktop runtimes: {0}" -f $_.Exception.Message) 'WARN' }
    return $false
}

function Invoke-InstallerAndValidateExitCode {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$FilePath,
        [Parameter(Mandatory)][string[]]$ArgumentList,
        [Parameter(Mandatory)][string]$DisplayName,
        [int[]]$SuccessExitCodes = @(0, 3010, 1641)
    )
    $process = Start-Process -FilePath $FilePath -ArgumentList $ArgumentList -Wait -PassThru -NoNewWindow
    Write-Log ("{0} installer exit code: {1}" -f $DisplayName, $process.ExitCode) 'INFO'
    if ($process.ExitCode -notin $SuccessExitCodes) { throw ("{0} installation failed with exit code {1}." -f $DisplayName, $process.ExitCode) }
    if ($process.ExitCode -in @(3010,1641)) { Add-YamlAction ("{0} installed; restart required." -f $DisplayName) }
    else { Add-YamlAction ("{0} installed successfully." -f $DisplayName) }
}

function Save-InternetFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][uri]$Uri,
        [Parameter(Mandatory)][string]$DestinationPath,
        [Parameter(Mandatory)][string]$DisplayName
    )
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    if (Test-Path -LiteralPath $DestinationPath) { Remove-Item -LiteralPath $DestinationPath -Force -ErrorAction SilentlyContinue }
    try { Start-BitsTransfer -Source $Uri.AbsoluteUri -Destination $DestinationPath -ErrorAction Stop }
    catch { Invoke-WebRequest -Uri $Uri.AbsoluteUri -OutFile $DestinationPath -UseBasicParsing -TimeoutSec 300 -ErrorAction Stop }
    if (-not (Test-Path -LiteralPath $DestinationPath) -or (Get-Item -LiteralPath $DestinationPath).Length -lt 1MB) {
        throw ("The downloaded {0} installer is missing or unexpectedly small." -f $DisplayName)
    }
}

function Ensure-DotNetDesktopRuntimeForDcu {
    [CmdletBinding()]
    param()

    if (Test-DotNetDesktopRuntime8 -MinimumVersion ([version]'8.0.25')) {
        Write-Log '.NET Desktop Runtime 8.0.25 or newer is already installed.' 'OK'
        return
    }

    $localInstaller = Join-Path $WorkingRoot 'windowsdesktop-runtime-8-win-x64.exe'
    if (Test-Path -LiteralPath $DotNetDesktopRuntimeSharePath) {
        Copy-Item -LiteralPath $DotNetDesktopRuntimeSharePath -Destination $localInstaller -Force -ErrorAction Stop
    } else {
        Write-Log 'The .NET share installer is unavailable; using the Microsoft Internet fallback.' 'WARN'
        Save-InternetFile -Uri ([uri]$DotNetDesktopRuntimeInternetUrl) -DestinationPath $localInstaller -DisplayName '.NET Desktop Runtime 8 x64'
    }

    Invoke-InstallerAndValidateExitCode -FilePath $localInstaller -ArgumentList @('/install','/quiet','/norestart') -DisplayName '.NET Desktop Runtime 8 x64'
    if (-not (Test-DotNetDesktopRuntime8 -MinimumVersion ([version]'8.0.25'))) {
        throw '.NET Desktop Runtime installation completed, but version 8.0.25 or newer was not detected.'
    }
}

function Ensure-DellCommandUpdateInstalled {
    [CmdletBinding()]
    param()

    $dcuCliPath = Join-Path ${env:ProgramFiles} 'Dell\CommandUpdate\dcu-cli.exe'
    if (Test-Path -LiteralPath $dcuCliPath) { return $dcuCliPath }

    Ensure-DotNetDesktopRuntimeForDcu
    if (Test-Path -LiteralPath $DellCommandUpdateSharePath) {
        $localDcuInstaller = Join-Path $WorkingRoot 'Dell-Command-Update-Windows-Universal-Application.exe'
        Copy-Item -LiteralPath $DellCommandUpdateSharePath -Destination $localDcuInstaller -Force -ErrorAction Stop
        Invoke-InstallerAndValidateExitCode -FilePath $localDcuInstaller -ArgumentList @('/s') -DisplayName 'Dell Command | Update'
    } else {
        Write-Log 'The Dell share installer is unavailable; using winget as the Internet fallback.' 'WARN'
        $winget = Get-Command -Name 'winget.exe' -ErrorAction SilentlyContinue
        if (-not $winget) { throw 'Dell Command | Update is not on the file share and winget.exe is unavailable.' }
        $wingetArgs = @('install','--id',$DellCommandUpdateWingetId,'--exact','--silent','--accept-package-agreements','--accept-source-agreements','--disable-interactivity')
        $wingetProcess = Start-Process -FilePath $winget.Source -ArgumentList $wingetArgs -Wait -PassThru -NoNewWindow
        if ($wingetProcess.ExitCode -ne 0) { throw ("winget failed to install Dell Command | Update. Exit code: {0}" -f $wingetProcess.ExitCode) }
    }

    Start-Sleep -Seconds 3
    if (-not (Test-Path -LiteralPath $dcuCliPath)) { throw "Dell Command | Update installed, but dcu-cli.exe was not found: $dcuCliPath" }
    return $dcuCliPath
}

function Invoke-DellDriverUpdates {
    Write-Section 'Dell Command Update Workflow'

    $dcuCli = Ensure-DellCommandUpdateInstalled

    Write-Log ("Using Dell Command | Update CLI: {0}" -f $dcuCli) 'OK'
    Add-YamlAction 'Using Dell Command | Update CLI.'

    Ensure-DellDCUService | Out-Null

    $dcuScanLog  = Join-Path $WorkingRoot 'Dell-DCU-Scan.log'
    $dcuApplyLog = Join-Path $WorkingRoot 'Dell-DCU-Apply.log'
    $dcuReport   = Join-Path $WorkingRoot 'Dell-DCU-ApplicableUpdates.xml'

    Write-Log 'Dell DCU Configure...' 'INFO'
    Write-Progress -Activity 'Dell Driver Update Workflow' -Status 'Configuring Dell Command Update' -PercentComplete 10
    $configureArgs = "/configure -silent -scheduleAuto -lockSettings=disable"
    Invoke-DellDCUCommandWithRetry -DcuCli $dcuCli -Arguments $configureArgs -OperationName 'Configure' -MaxAttempts 2 | Out-Null

    Write-Log 'Dell DCU Scan...' 'INFO'
    Write-Progress -Activity 'Dell Driver Update Workflow' -Status 'Scanning for updates' -PercentComplete 35
    $dellUpdateTypes = if ($script:InstallAllUpdates) { 'bios,firmware,driver,application,others' } elseif ($script:InstallBiosFirmware) { 'bios,firmware,driver,application' } else { 'driver,application' }
    Write-Log ("Dell update types selected: {0}" -f $dellUpdateTypes) 'INFO'
    Add-YamlAction ("Dell update types selected: {0}" -f $dellUpdateTypes)

    $scanArgs = "/scan -silent -updateType=$dellUpdateTypes -outputLog=""$dcuScanLog"" -report=""$dcuReport"""
    Invoke-DellDCUCommandWithRetry -DcuCli $dcuCli -Arguments $scanArgs -OperationName 'Scan' -MaxAttempts 2 | Out-Null
    Write-Log 'Waiting for Dell Command | Update to finish writing the report...' 'INFO'

    try {
        $xml = Get-DcuReportXmlSafely -Path $dcuReport
        $items = Get-DellReportItems -Xml $xml

        if ($items.Count -gt 0) {
            Add-YamlAction ("Dell DCU report parsed successfully. Updates detected: {0}" -f $items.Count)

            $total = $items.Count
            $index = 0

            foreach ($item in $items) {
                $index++
                $percent = 35 + [math]::Floor(($index / $total) * 35)
                $label = if ($item.Name) { $item.Name } elseif ($item.Id) { $item.Id } else { 'Dell update' }

                Write-Progress -Activity 'Dell Driver Update Workflow' -Status ("Parsing report: {0}" -f $label) -PercentComplete $percent

                if (-not $script:InstallBiosFirmware -and $item.Category -match 'BIOS|Firmware') {
                    $script:SkippedList.Add($label) | Out-Null
                    Add-DriverResult -Vendor 'Dell' -Name $item.Name -Id $item.Id -Category $item.Category -Status 'Blocked' -Message 'BIOS/Firmware update blocked by default script policy. Use -IncludeBIOS or -All to permit it.'
                    Write-Log ("Blocking Dell BIOS/Firmware update: {0}" -f $label) 'WARN'
                }
                else {
                    Add-DriverResult -Vendor 'Dell' -Name $item.Name -Id $item.Id -Category $item.Category -Status 'Detected' -Message 'Detected and permitted by the selected update policy.'
                }
            }
        }
        else {
            Add-YamlAction 'Dell DCU report parsed but returned no identifiable updates.'
        }
    }
    catch {
        Write-Log ("Failed to parse Dell DCU report: {0}" -f $_.Exception.Message) 'WARN'
        Add-YamlAction ("Failed to parse Dell DCU report: {0}" -f $_.Exception.Message)
    }

    Write-Log ("Dell DCU ApplyUpdates using update types: {0}" -f $dellUpdateTypes) 'INFO'
    Write-Progress -Activity 'Dell Driver Update Workflow' -Status 'Applying updates' -PercentComplete 85
    $applyArgs = "/applyUpdates -silent -updateType=$dellUpdateTypes -reboot=disable -outputLog=""$dcuApplyLog"""
    Invoke-DellDCUCommandWithRetry -DcuCli $dcuCli -Arguments $applyArgs -OperationName 'ApplyUpdates' -MaxAttempts 2 | Out-Null

    Write-Log (("Dell DCU logs: {0} ; {1} ; {2}") -f $dcuScanLog, $dcuApplyLog, $dcuReport) 'OK'

    Write-Log 'Running Dell BIOS settings script from the central share...' 'INFO'
    $dellBiosScript='\\filesvr\Labscripts\Set-Dell-BIOS-Settings.ps1'
    if (Test-Path -LiteralPath $dellBiosScript) {
        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $dellBiosScript
        if ($LASTEXITCODE -eq 0) {
            Write-Log 'Dell BIOS settings script completed successfully.' 'OK'
        } else {
            Write-Log ("Dell BIOS settings script exited with code {0}." -f $LASTEXITCODE) 'WARN'
        }
    } else {
        Write-Log ("Dell BIOS settings script not found: {0}" -f $dellBiosScript) 'WARN'
    }

    Write-Progress -Activity 'Dell Driver Update Workflow' -Completed
}
# -----------------------------
# Main
# -----------------------------
$finalStatus = 'success'

try {
    Write-Section 'Initialization'

    Ensure-Folder -Path $YamlLogFolder
    Ensure-Folder -Path $WorkingRoot
    Ensure-WorkingFolderPermissions -Path $WorkingRoot

    $yamlName = "{0}-{1}-{2}.yml" -f $script:ComputerName, '05_Weekend_Vendor_Drivers_Update', (Get-Date -Format 'yyyy-MM-dd_HHmmss')
    $yamlPath = Join-Path $YamlLogFolder $yamlName

    Write-Log ("YAML log will be written to: {0}" -f $yamlPath) 'INFO'
    Initialize-YamlLog -ComputerName $script:ComputerName -YamlPath $yamlPath

    Write-Log 'Initializing vendor driver update script...' 'INFO'
    if ($script:InstallAllUpdates) {
        Write-Log 'ALL-UPDATES OVERRIDE ENABLED: BIOS, firmware, storage, chipset, controller, and other normally excluded updates are permitted.' 'WARN'
        Add-YamlAction 'ALL-UPDATES OVERRIDE ENABLED.'
    }
    elseif ($script:InstallBiosFirmware) {
        Write-Log 'BIOS/FIRMWARE OVERRIDE ENABLED: BIOS and firmware updates are permitted; other high-risk driver exclusions remain active.' 'WARN'
        Add-YamlAction 'BIOS/FIRMWARE OVERRIDE ENABLED.'
    }
    else {
        Write-Log 'Safe scheduled-update policy enabled. BIOS, firmware, and boot-critical driver families remain excluded.' 'INFO'
        Add-YamlAction 'Safe scheduled-update policy enabled.'
    }


    Write-Section 'Vendor Detection'
    $script:DetectedVendor = Get-DriverVendor
    Write-Log ("Detected vendor workflow: {0}" -f $script:DetectedVendor) 'INFO'
    Write-Log ("Working root: {0}" -f $WorkingRoot) 'INFO'

    if ($script:DetectedVendor -eq 'HP') {
        Invoke-HPPowerShellModuleMaintenance
        Invoke-HPDriverUpdates
    }
    elseif ($script:DetectedVendor -eq 'Dell') {
        Invoke-DellDriverUpdates
    }

}
catch {
    $finalStatus = 'failed'
    Add-RunFailure ("Script failed: {0}" -f $_.Exception.Message)
}
finally {
    Write-Section 'Cleanup'
    Remove-WorkingFolderRobust -Path $WorkingRoot

    if ($script:RunFailures.Count -gt 0 -and $finalStatus -ne 'failed') {
        $finalStatus = 'completed_with_warnings'
    }

    if ($script:RunFailures.Count -gt 0) {
        Write-Log (("{0} driver update completed with one or more failures.") -f $script:DetectedVendor) 'WARN'
    }
    else {
        Write-Log (("{0} driver update script completed successfully.") -f $script:DetectedVendor) 'OK'
    }

    Save-YamlLog -Status $finalStatus
}# ScriptName: 05_Weekend_HP_Drivers_Update.ps1
# ScriptVersion: 2.7.0
# LastUpdated: 2026-07-21
# Purpose: Weekend vendor driver update script with clean HP + Dell support,
#          HP-only HPCMSL maintenance after vendor detection,
#          YAML logging, colored output,
#          share/Internet fallback deployment for HPIA, .NET Desktop Runtime, and DCU,
#          section headers, progress display, and structured per-driver results.

[CmdletBinding()]
param([string]$WorkingRoot = 'C:\Temp\DriverUpdates',
    [string]$YamlLogFolder = 'C:\Logs',
    [switch]$IncludeSoftware,
    [switch]$IncludeBIOS,
    [string]$HpiaSourceFolder = '\\filesvr\Labscripts\HPImageAssistant',
    [string]$HpiaInstallerSharePath = '\\filesvr\labscripts\HPImageAssistant.exe',
    [string]$HpiaInternetLandingPage = 'https://ftp.ext.hp.com/pub/caps-softpaq/cmit/HPIA.html',
    [string]$LocalHpiaFolder = 'C:\ProgramData\Compton\HPImageAssistant',
    [string]$DellCommandUpdateSharePath = '\\filesvr\labscripts\Dell-Command-Update-Windows-Universal-Application.exe',
    [string]$DotNetDesktopRuntimeSharePath = '\\filesvr\labscripts\windowsdesktop-runtime-8.0.25-win-x64.exe',
    [string]$DotNetDesktopRuntimeInternetUrl = 'https://aka.ms/dotnet/8.0/windowsdesktop-runtime-win-x64.exe',
    [string]$DellCommandUpdateWingetId = 'Dell.CommandUpdate.Universal'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# -----------------------------
# Script metadata
# -----------------------------
$script:ScriptName        = '05_Weekend_HP_Drivers_Update.ps1'
$script:ScriptVersion     = '2.7.0'
$script:StartTime         = Get-Date
$script:RunFailures       = New-Object System.Collections.Generic.List[string]
$script:InstalledList     = New-Object System.Collections.Generic.List[string]
$script:SkippedList       = New-Object System.Collections.Generic.List[string]
$script:YamlActionLines   = New-Object System.Collections.Generic.List[string]
$script:DriverResults     = New-Object System.Collections.Generic.List[object]
$script:DetectedVendor    = $null
$script:YamlPath          = $null
$script:ComputerName      = $env:COMPUTERNAME

# -----------------------------
# Logging helpers
# -----------------------------
function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','OK','WARN','ERROR')][string]$Level = 'INFO'
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[{0}] [{1}] {2}" -f $timestamp, $Level.PadRight(5), $Message

    switch ($Level) {
        'INFO'  { Write-Host $line -ForegroundColor Cyan }
        'OK'    { Write-Host $line -ForegroundColor Green }
        'WARN'  { Write-Host $line -ForegroundColor Yellow }
        'ERROR' { Write-Host $line -ForegroundColor Red }
        default { Write-Host $line }
    }
}

function Write-Section {
    param([Parameter(Mandatory)][string]$Title)

    $border = ('=' * 72)
    Write-Host ''
    Write-Host $border -ForegroundColor Magenta
    Write-Host ("  {0}" -f $Title) -ForegroundColor Magenta
    Write-Host $border -ForegroundColor Magenta
    Add-YamlAction ("Section: {0}" -f $Title)
}

function Add-RunFailure {
    param([Parameter(Mandatory)][string]$Message)
    $script:RunFailures.Add($Message) | Out-Null
    Write-Log $Message 'WARN'
}

function Add-DriverResult {
    param(
        [Parameter(Mandatory)][string]$Vendor,
        [Parameter(Mandatory)][string]$Name,
        [string]$Id,
        [string]$Category,
        [ValidateSet('Detected','Installed','Downloaded','Skipped','Failed','Blocked')][string]$Status,
        [string]$Message = ''
    )

    $script:DriverResults.Add([pscustomobject]@{
        Vendor   = $Vendor
        Name     = $Name
        Id       = $Id
        Category = $Category
        Status   = $Status
        Message  = $Message
    }) | Out-Null
}

function ConvertTo-YamlSafeString {
    param([AllowNull()][object]$Value)

    if ($null -eq $Value) { return 'null' }

    $text = [string]$Value
    $text = $text -replace "`r", ''
    $text = $text -replace "`n", ' '
    $text = $text -replace "'", "''"
    return "'$text'"
}

function Initialize-YamlLog {
    param(
        [Parameter(Mandatory)][string]$ComputerName,
        [Parameter(Mandatory)][string]$YamlPath
    )

    $script:YamlPath = $YamlPath
    $script:YamlActionLines.Clear()

    Add-YamlAction 'Script initialized.'
    Add-YamlAction ("Working root: {0}" -f $WorkingRoot)
    Add-YamlAction ("Computer: {0}" -f $ComputerName)
}

function Add-YamlAction {
    param([Parameter(Mandatory)][string]$Text)
    $script:YamlActionLines.Add(("    - {0}" -f (ConvertTo-YamlSafeString $Text))) | Out-Null
}

function Save-YamlLog {
    param(
        [Parameter(Mandatory)][string]$Status
    )

    if ([string]::IsNullOrWhiteSpace($script:YamlPath)) {
        return
    }

    $endTime = Get-Date
    $duration = [math]::Round(($endTime - $script:StartTime).TotalSeconds, 0)

    $lines = New-Object System.Collections.Generic.List[string]

    $lines.Add('script:') | Out-Null
    $lines.Add(("  name: {0}" -f (ConvertTo-YamlSafeString $script:ScriptName))) | Out-Null
    $lines.Add(("  version: {0}" -f (ConvertTo-YamlSafeString $script:ScriptVersion))) | Out-Null
    $lines.Add(("  computer: {0}" -f (ConvertTo-YamlSafeString $script:ComputerName))) | Out-Null
    $lines.Add(("  started: {0}" -f (ConvertTo-YamlSafeString ($script:StartTime.ToString('s'))))) | Out-Null
    $lines.Add(("  ended: {0}" -f (ConvertTo-YamlSafeString ($endTime.ToString('s'))))) | Out-Null
    $lines.Add(("  duration_seconds: {0}" -f $duration)) | Out-Null

    $lines.Add('run:') | Out-Null
    $lines.Add(("  status: {0}" -f (ConvertTo-YamlSafeString $Status))) | Out-Null
    $lines.Add(("  vendor: {0}" -f (ConvertTo-YamlSafeString $script:DetectedVendor))) | Out-Null

    $lines.Add('  actions:') | Out-Null
    if ($script:YamlActionLines.Count -eq 0) {
        $lines.Add('    []') | Out-Null
    }
    else {
        foreach ($line in $script:YamlActionLines) {
            $lines.Add($line) | Out-Null
        }
    }

    $lines.Add('  installed:') | Out-Null
    if ($script:InstalledList.Count -eq 0) {
        $lines.Add('    []') | Out-Null
    }
    else {
        foreach ($item in $script:InstalledList) {
            $lines.Add(("    - {0}" -f (ConvertTo-YamlSafeString $item))) | Out-Null
        }
    }

    $lines.Add('  skipped:') | Out-Null
    if ($script:SkippedList.Count -eq 0) {
        $lines.Add('    []') | Out-Null
    }
    else {
        foreach ($item in $script:SkippedList) {
            $lines.Add(("    - {0}" -f (ConvertTo-YamlSafeString $item))) | Out-Null
        }
    }

    $lines.Add('  failures:') | Out-Null
    if ($script:RunFailures.Count -eq 0) {
        $lines.Add('    []') | Out-Null
    }
    else {
        foreach ($item in $script:RunFailures) {
            $lines.Add(("    - {0}" -f (ConvertTo-YamlSafeString $item))) | Out-Null
        }
    }

    $lines.Add('drivers:') | Out-Null
    if ($script:DriverResults.Count -eq 0) {
        $lines.Add('  []') | Out-Null
    }
    else {
        foreach ($driver in $script:DriverResults) {
            $lines.Add('  -') | Out-Null
            $lines.Add(("    vendor: {0}" -f (ConvertTo-YamlSafeString $driver.Vendor))) | Out-Null
            $lines.Add(("    name: {0}" -f (ConvertTo-YamlSafeString $driver.Name))) | Out-Null
            $lines.Add(("    id: {0}" -f (ConvertTo-YamlSafeString $driver.Id))) | Out-Null
            $lines.Add(("    category: {0}" -f (ConvertTo-YamlSafeString $driver.Category))) | Out-Null
            $lines.Add(("    status: {0}" -f (ConvertTo-YamlSafeString $driver.Status))) | Out-Null
            $lines.Add(("    message: {0}" -f (ConvertTo-YamlSafeString $driver.Message))) | Out-Null
        }
    }

    Set-Content -LiteralPath $script:YamlPath -Value $lines -Encoding UTF8
    Write-Host ("YAML log written successfully: {0}" -f $script:YamlPath) -ForegroundColor Green
}

# -----------------------------
# File/folder helpers
# -----------------------------
function Ensure-Folder {
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -Path $Path -ItemType Directory -Force | Out-Null
    }
}

function Ensure-WorkingFolderPermissions {
    param([Parameter(Mandatory)][string]$Path)

    try {
        & icacls.exe $Path '/grant' '*S-1-1-0:(OI)(CI)F' '/T' '/C' | Out-Null
    }
    catch {
        Write-Log ("Unable to relax working folder permissions on {0}: {1}" -f $Path, $_.Exception.Message) 'WARN'
    }
}

function Remove-WorkingFolderRobust {
    param(
        [Parameter(Mandatory)][string]$Path,
        [int]$RetryCount = 6,
        [int]$RetryDelaySeconds = 5
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        Write-Log ("Working folder already absent: {0}" -f $Path) 'OK'
        return
    }

    Write-Log ("Attempting to remove working folder: {0}" -f $Path) 'INFO'
    for ($i = 1; $i -le $RetryCount; $i++) {
        try {
            Remove-Item -LiteralPath $Path -Recurse -Force -ErrorAction Stop
            Write-Log 'Working folder removed successfully.' 'OK'
            return
        }
        catch {
            if ($i -eq $RetryCount) {
                Add-RunFailure ("Failed to remove working folder after retries: {0}" -f $_.Exception.Message)
                return
            }
            Start-Sleep -Seconds $RetryDelaySeconds
        }
    }
}

# -----------------------------
# Vendor detection
# -----------------------------
function Get-SystemManufacturer {
    try {
        return (Get-CimInstance Win32_ComputerSystem -ErrorAction Stop).Manufacturer
    }
    catch {
        throw "Unable to determine system manufacturer. $($_.Exception.Message)"
    }
}

function Get-DriverVendor {
    $manufacturer = Get-SystemManufacturer
    Write-Log ("Detected manufacturer: {0}" -f $manufacturer) 'INFO'

    if ($manufacturer -match 'Dell') { return 'Dell' }
    if ($manufacturer -match 'HP|Hewlett-Packard') { return 'HP' }

    throw "Unsupported manufacturer for this script: $manufacturer"
}

# -----------------------------
# HP Support - HP Image Assistant extracted-folder deployment
# -----------------------------
function Get-HPSystemModelInfo {
    [CmdletBinding()]
    param()

    $cs = Get-CimInstance Win32_ComputerSystem -ErrorAction Stop
    $csp = Get-CimInstance Win32_ComputerSystemProduct -ErrorAction SilentlyContinue
    $bb = Get-CimInstance Win32_BaseBoard -ErrorAction SilentlyContinue

    $platform = $null
    if ($bb -and $bb.Product) {
        $platform = $bb.Product.ToString().Trim().ToUpper()
        if ($platform.Length -gt 4) {
            $platform = $platform.Substring(0,4)
        }
    }

    $model = if ($cs.Model) { $cs.Model.ToString().Trim() } else { 'Unknown' }
    $sku = if ($csp -and $csp.Version) { $csp.Version.ToString().Trim() } else { 'Unknown' }

    $info = [pscustomobject]@{
        Manufacturer = $cs.Manufacturer
        Model        = $model
        SKU          = $sku
        Platform     = $platform
    }

    Write-Log ("Detected HP system model: {0}" -f $info.Model) 'INFO'
    Write-Log ("Detected HP platform/baseboard ID: {0}" -f $info.Platform) 'INFO'
    Add-YamlAction ("Detected HP system model: {0}" -f $info.Model)
    Add-YamlAction ("Detected HP platform/baseboard ID: {0}" -f $info.Platform)

    return $info
}

function Get-ExistingHPIAExecutable {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$PreferredFolder)

    $candidateFolders = New-Object System.Collections.Generic.List[string]

    if (-not [string]::IsNullOrWhiteSpace($PreferredFolder)) {
        $candidateFolders.Add($PreferredFolder) | Out-Null
    }

    foreach ($folder in @(
        'C:\Program Files\HP\HP Image Assistant',
        'C:\Program Files (x86)\HP\HP Image Assistant',
        'C:\SWSetup\HPImageAssistant',
        'C:\ProgramData\Compton\HPImageAssistant'
    )) {
        if (-not $candidateFolders.Contains($folder)) {
            $candidateFolders.Add($folder) | Out-Null
        }
    }

    foreach ($folder in $candidateFolders) {
        if ([string]::IsNullOrWhiteSpace($folder) -or -not (Test-Path -LiteralPath $folder)) {
            continue
        }

        $exe = Get-ChildItem -Path $folder -Recurse -Filter 'HPImageAssistant.exe' -File -ErrorAction SilentlyContinue |
            Sort-Object FullName |
            Select-Object -First 1

        if ($exe) {
            return $exe.FullName
        }
    }

    return $null
}

function Install-HPIAFromExtractedFolder {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SourceFolder,
        [Parameter(Mandatory)][string]$DestinationFolder
    )

    Write-Section 'HP Image Assistant Local Deployment'
    Write-Log ("HPIA source folder: {0}" -f $SourceFolder) 'INFO'
    Write-Log ("HPIA local folder: {0}" -f $DestinationFolder) 'INFO'
    Add-YamlAction ("HPIA source folder: {0}" -f $SourceFolder)
    Add-YamlAction ("HPIA local folder: {0}" -f $DestinationFolder)

    $existingExe = Get-ExistingHPIAExecutable -PreferredFolder $DestinationFolder
    if ($existingExe) {
        try {
            $existingVersion = (Get-Item -LiteralPath $existingExe -ErrorAction Stop).VersionInfo.FileVersion
            Write-Log ("HP Image Assistant is already installed/found at: {0} (Version: {1})" -f $existingExe, $existingVersion) 'OK'
            Add-YamlAction ("Skipped HPIA local deployment because HPImageAssistant.exe already exists: {0} (Version: {1})" -f $existingExe, $existingVersion)
        }
        catch {
            Write-Log ("HP Image Assistant is already installed/found at: {0}" -f $existingExe) 'OK'
            Add-YamlAction ("Skipped HPIA local deployment because HPImageAssistant.exe already exists: {0}" -f $existingExe)
        }

        return $existingExe
    }

    Write-Log 'HP Image Assistant was not found locally. Deploying from extracted source folder.' 'INFO'
    Add-YamlAction 'HP Image Assistant was not found locally. Deploying from extracted source folder.'

    if (-not (Test-Path -LiteralPath $SourceFolder)) {
        throw "HPIA source folder not found: $SourceFolder"
    }

    $sourceFiles = @(Get-ChildItem -Path $SourceFolder -Recurse -File -ErrorAction SilentlyContinue)
    Write-Log ("Source HPIA file count: {0}" -f $sourceFiles.Count) 'INFO'
    Add-YamlAction ("Source HPIA file count: {0}" -f $sourceFiles.Count)

    $sourceExe = Get-ChildItem -Path $SourceFolder -Recurse -Filter 'HPImageAssistant.exe' -File -ErrorAction SilentlyContinue |
        Sort-Object FullName |
        Select-Object -First 1

    if (-not $sourceExe) {
        throw "HPImageAssistant.exe was not found anywhere under source folder: $SourceFolder"
    }

    Write-Log ("Found source HPImageAssistant.exe: {0}" -f $sourceExe.FullName) 'OK'
    Add-YamlAction ("Found source HPImageAssistant.exe: {0}" -f $sourceExe.FullName)

    try {
        if (Test-Path -LiteralPath $DestinationFolder) {
            Write-Log ("Removing existing local HPIA folder before clean deployment: {0}" -f $DestinationFolder) 'INFO'
            Remove-Item -LiteralPath $DestinationFolder -Recurse -Force -ErrorAction Stop
        }

        New-Item -Path $DestinationFolder -ItemType Directory -Force | Out-Null

        Write-Log 'Copying extracted HPIA files locally with robocopy...' 'INFO'

        $roboLog = Join-Path $DestinationFolder 'HPIA_robocopy.log'
        $roboArgs = @(
            ('"{0}"' -f $SourceFolder),
            ('"{0}"' -f $DestinationFolder),
            '/E',
            '/COPY:DAT',
            '/R:3',
            '/W:5',
            '/NFL',
            '/NDL',
            '/NP',
            ('/LOG:"{0}"' -f $roboLog)
        )

        $robo = Start-Process -FilePath "$env:SystemRoot\System32\robocopy.exe" -ArgumentList ($roboArgs -join ' ') -Wait -PassThru -NoNewWindow

        # Robocopy exit codes 0-7 are success/non-fatal. 8+ indicates failure.
        if ($robo.ExitCode -ge 8) {
            throw "Robocopy failed copying HPIA files. Exit code: $($robo.ExitCode). Log: $roboLog"
        }

        Write-Log ("HPIA files copied locally with robocopy exit code {0}." -f $robo.ExitCode) 'OK'
        Add-YamlAction ("HPIA files copied locally with robocopy exit code {0}." -f $robo.ExitCode)

        try {
            Get-ChildItem -Path $DestinationFolder -Recurse -File -ErrorAction SilentlyContinue |
                ForEach-Object { Unblock-File -Path $_.FullName -ErrorAction SilentlyContinue }
        }
        catch {}

        $copiedFiles = @(Get-ChildItem -Path $DestinationFolder -Recurse -File -ErrorAction SilentlyContinue)
        Write-Log ("Local HPIA folder file count after copy: {0}" -f $copiedFiles.Count) 'INFO'
        Add-YamlAction ("Local HPIA folder file count after copy: {0}" -f $copiedFiles.Count)

        $localExeItem = Get-ChildItem -Path $DestinationFolder -Recurse -Filter 'HPImageAssistant.exe' -File -ErrorAction SilentlyContinue |
            Sort-Object FullName |
            Select-Object -First 1

        if (-not $localExeItem) {
            $sampleFiles = $copiedFiles | Select-Object -First 20 | ForEach-Object { $_.FullName }
            foreach ($sample in $sampleFiles) {
                Write-Log ("Local HPIA sample file: {0}" -f $sample) 'WARN'
            }

            throw "HPImageAssistant.exe was not found anywhere under local copy folder: $DestinationFolder"
        }

        $localExe = $localExeItem.FullName

        Write-Log ("Resolved local HPImageAssistant.exe location: {0}" -f $localExe) 'OK'
        Add-YamlAction ("Resolved local HPImageAssistant.exe location: {0}" -f $localExe)

        return $localExe
    }
    catch {
        throw "Failed to deploy HP Image Assistant locally: $($_.Exception.Message)"
    }
}


function Expand-HpiaSoftPaq {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$InstallerPath,
        [Parameter(Mandatory)][string]$DestinationFolder
    )

    Write-Section 'HP Image Assistant Installer Extraction'
    if (-not (Test-Path -LiteralPath $InstallerPath)) {
        throw "HPIA installer was not found: $InstallerPath"
    }

    if (Test-Path -LiteralPath $DestinationFolder) {
        Remove-Item -LiteralPath $DestinationFolder -Recurse -Force -ErrorAction Stop
    }
    Ensure-Folder -Path $DestinationFolder

    $arguments = @('-s', '-e', ("-f`"{0}`"" -f $DestinationFolder))
    $process = Start-Process -FilePath $InstallerPath -ArgumentList ($arguments -join ' ') -Wait -PassThru -NoNewWindow
    Write-Log ("HPIA SoftPaq extraction exit code: {0}" -f $process.ExitCode) 'INFO'

    if ($process.ExitCode -notin @(0, 3010, 1641)) {
        throw ("HPIA SoftPaq extraction failed with exit code {0}." -f $process.ExitCode)
    }

    Start-Sleep -Seconds 2
    $hpiaExe = Get-ExistingHPIAExecutable -PreferredFolder $DestinationFolder
    if (-not $hpiaExe) {
        throw "HPIA extraction completed, but HPImageAssistant.exe was not found under $DestinationFolder"
    }

    Write-Log ("HP Image Assistant extracted and verified: {0}" -f $hpiaExe) 'OK'
    Add-YamlAction ("HP Image Assistant extracted and verified: {0}" -f $hpiaExe)
    return $hpiaExe
}

function Get-LatestHpiaSoftPaqUrl {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$LandingPage)

    Write-Log ("Checking HP for the latest HPIA SoftPaq: {0}" -f $LandingPage) 'INFO'
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $response = Invoke-WebRequest -Uri $LandingPage -UseBasicParsing -TimeoutSec 60 -ErrorAction Stop
    $candidateLinks = New-Object System.Collections.Generic.List[string]

    foreach ($link in @($response.Links)) {
        if ($link.href -and $link.href -match '(?i)sp\d+\.exe(?:\?.*)?$') {
            $candidateLinks.Add([string]$link.href) | Out-Null
        }
    }
    if ($candidateLinks.Count -eq 0) {
        $matches = [regex]::Matches([string]$response.Content, '(?i)href\s*=\s*["'']([^"'']*sp\d+\.exe(?:\?[^"'']*)?)["'']')
        foreach ($match in $matches) { $candidateLinks.Add($match.Groups[1].Value) | Out-Null }
    }
    if ($candidateLinks.Count -eq 0) { throw 'The HP HPIA page did not contain a SoftPaq EXE download link.' }

    $resolved = foreach ($href in $candidateLinks) {
        try {
            $uri = [uri]::new([uri]$LandingPage, $href)
            $spMatch = [regex]::Match($uri.AbsoluteUri, '(?i)sp(?<number>\d+)\.exe')
            if ($spMatch.Success) {
                [pscustomobject]@{ Url = $uri.AbsoluteUri; SoftPaqNumber = [int]$spMatch.Groups['number'].Value }
            }
        } catch {}
    }
    $latest = $resolved | Sort-Object SoftPaqNumber -Descending | Select-Object -First 1
    if (-not $latest) { throw 'Unable to resolve a valid HPIA SoftPaq URL from the HP landing page.' }
    Add-YamlAction ("Latest HPIA SoftPaq resolved from HP: {0}" -f $latest.Url)
    return $latest.Url
}

function Ensure-HPIAAvailable {
    [CmdletBinding()]
    param()

    $existingExe = Get-ExistingHPIAExecutable -PreferredFolder $LocalHpiaFolder
    if ($existingExe) {
        Write-Log ("HP Image Assistant is already available locally: {0}" -f $existingExe) 'OK'
        Add-YamlAction 'HP Image Assistant was already available locally.'
        return $existingExe
    }

    if (Test-Path -LiteralPath $HpiaSourceFolder) {
        try { return Install-HPIAFromExtractedFolder -SourceFolder $HpiaSourceFolder -DestinationFolder $LocalHpiaFolder }
        catch { Write-Log ("Extracted HPIA share-folder deployment failed: {0}" -f $_.Exception.Message) 'WARN' }
    }

    if (Test-Path -LiteralPath $HpiaInstallerSharePath) {
        try {
            $localInstaller = Join-Path $WorkingRoot 'HPImageAssistant.exe'
            Copy-Item -LiteralPath $HpiaInstallerSharePath -Destination $localInstaller -Force -ErrorAction Stop
            return Expand-HpiaSoftPaq -InstallerPath $localInstaller -DestinationFolder $LocalHpiaFolder
        }
        catch { Write-Log ("HPIA share-installer deployment failed: {0}" -f $_.Exception.Message) 'WARN' }
    }

    Write-Log 'Downloading the latest HP Image Assistant SoftPaq from HP...' 'INFO'
    Add-YamlAction 'Falling back to an Internet download of HP Image Assistant.'
    $downloadUrl = Get-LatestHpiaSoftPaqUrl -LandingPage $HpiaInternetLandingPage
    $downloadPath = Join-Path $WorkingRoot 'HPImageAssistant-Internet.exe'
    Invoke-WebRequest -Uri $downloadUrl -OutFile $downloadPath -UseBasicParsing -TimeoutSec 300 -ErrorAction Stop
    if (-not (Test-Path -LiteralPath $downloadPath) -or (Get-Item -LiteralPath $downloadPath).Length -lt 1MB) {
        throw 'The downloaded HPIA installer is missing or unexpectedly small.'
    }
    return Expand-HpiaSoftPaq -InstallerPath $downloadPath -DestinationFolder $LocalHpiaFolder
}

function Get-HpiaExitStatus {
    param([int]$ExitCode)

    switch ($ExitCode) {
        0    { return 'success' }
        1    { return 'failed' }
        2    { return 'cancelled' }
        3    { return 'needs_reboot' }
        256  { return 'no_recommendations_or_success' }
        257  { return 'recommendations_found' }
        3010 { return 'needs_reboot' }
        3011 { return 'not_auto_installable_skipped' }
        4096 { return 'no_applicable_updates_or_platform_not_supported' }
        4097 { return 'invalid_parameters' }
        8193 { return 'hpia_analysis_or_report_generation_error' }
        default { return 'unknown' }
    }
}

function Get-HpiaRecommendationObjects {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ReportFolder
    )

    # Use a flexible PowerShell array instead of a strongly typed generic list.
    # HPIA reports can contain mixed object types from JSON and XML parsing.
    $recommendations = @()

    $jsonFiles = @(Get-ChildItem -Path $ReportFolder -Filter '*.json' -Recurse -File -ErrorAction SilentlyContinue)
    foreach ($jsonFile in $jsonFiles) {
        try {
            $json = Get-Content -LiteralPath $jsonFile.FullName -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop

            if ($json.HPIA -and $json.HPIA.Recommendations) {
                foreach ($rec in @($json.HPIA.Recommendations)) {
                    $recommendations += $rec
                    try {
                        Write-Log ("HPIA JSON recommendation type: {0}" -f $rec.GetType().FullName) 'INFO'
                    }
                    catch {}
                }
            }
            elseif ($json.Recommendations) {
                foreach ($rec in @($json.Recommendations)) {
                    $recommendations += $rec
                    try {
                        Write-Log ("HPIA JSON recommendation type: {0}" -f $rec.GetType().FullName) 'INFO'
                    }
                    catch {}
                }
            }
        }
        catch {
            Write-Log ("Unable to parse HPIA JSON report {0}: {1}" -f $jsonFile.FullName, $_.Exception.Message) 'WARN'
        }
    }

    $xmlFiles = @(Get-ChildItem -Path $ReportFolder -Filter '*.xml' -Recurse -File -ErrorAction SilentlyContinue)
    foreach ($xmlFile in $xmlFiles) {
        try {
            [xml]$xml = Get-Content -LiteralPath $xmlFile.FullName -Raw -ErrorAction Stop

            $nodes = @($xml.SelectNodes('//*[contains(translate(local-name(), "ABCDEFGHIJKLMNOPQRSTUVWXYZ", "abcdefghijklmnopqrstuvwxyz"), "recommend")]'))
            foreach ($node in $nodes) {
                $recommendations += $node
                try {
                    Write-Log ("HPIA XML recommendation node type: {0}" -f $node.GetType().FullName) 'INFO'
                }
                catch {}
            }
        }
        catch {
            Write-Log ("Unable to parse HPIA XML report {0}: {1}" -f $xmlFile.FullName, $_.Exception.Message) 'WARN'
        }
    }

    return @($recommendations)
}

function Get-HpiaRecommendationValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Recommendation,
        [Parameter(Mandatory)][string[]]$PropertyNames
    )

    foreach ($prop in $PropertyNames) {
        try {
            if ($Recommendation.PSObject.Properties.Name -contains $prop) {
                $value = $Recommendation.$prop
                if ($null -ne $value -and -not [string]::IsNullOrWhiteSpace($value.ToString())) {
                    return $value.ToString()
                }
            }
        }
        catch {}
    }

    # XML fallback
    try {
        foreach ($prop in $PropertyNames) {
            $node = $Recommendation.SelectSingleNode('.//*[local-name()="' + $prop + '"]')
            if ($node -and -not [string]::IsNullOrWhiteSpace($node.InnerText)) {
                return $node.InnerText.Trim()
            }
        }
    }
    catch {}

    return $null
}

function Get-HpiaSoftPaqNumber {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Recommendation)

    $candidate = Get-HpiaRecommendationValue -Recommendation $Recommendation -PropertyNames @(
        'SoftPaqId','SoftpaqId','SoftPaq','Softpaq','SoftPaqNumber','SoftpaqNumber','SP','Id','ID','Number'
    )

    if ($candidate -match '(?i)sp?(\d{5,6})') {
        return $matches[1]
    }

    $text = ($Recommendation | Out-String)
    if ($text -match '(?i)sp(\d{5,6})') {
        return $matches[1]
    }

    if ($text -match '\b(\d{5,6})\b') {
        return $matches[1]
    }

    return $null
}

function Get-HpiaRecommendationCategory {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Recommendation)

    return (Get-HpiaRecommendationValue -Recommendation $Recommendation -PropertyNames @(
        'Category','Type','RecommendationType','ComponentType','Class','Group'
    ))
}

function Get-HpiaRecommendationName {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Recommendation)

    $name = Get-HpiaRecommendationValue -Recommendation $Recommendation -PropertyNames @(
        'Name','Title','Component','ComponentName','Description','SoftPaqName','SoftpaqName'
    )

    if ($name) { return $name }

    $text = ($Recommendation | Out-String).Trim()
    if ($text.Length -gt 160) {
        return $text.Substring(0,160)
    }

    return $text
}


function Test-HpiaBlockedUnattendedDriver {
    [CmdletBinding()]
    param(
        [AllowNull()][string]$Category,
        [AllowNull()][string]$Name
    )

    $combined = ("{0} {1}" -f $Category, $Name)

    # These driver families can affect boot-critical storage/controller behavior.
    # They are intentionally blocked from unattended HPIA installs.
    $blockedPattern = '(?i)\b(storage|chipset|rst|rapid storage|intel rapid|vmd|volume management device|nvme|nvm express|sata|raid|ahci|scsi|controller|management engine|mei|thunderbolt)\b'

    return ($combined -match $blockedPattern)
}

function New-HpiaDriverSPList {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ReportFolder,
        [Parameter(Mandatory)][string]$SPListPath
    )

    $recommendations = @(Get-HpiaRecommendationObjects -ReportFolder $ReportFolder)
    Write-Log ("HPIA recommendations parsed from reports: {0}" -f $recommendations.Count) 'INFO'
    Add-YamlAction ("HPIA recommendations parsed from reports: {0}" -f $recommendations.Count)

    $selected = @()
    $seen = @{}

    foreach ($rec in $recommendations) {
        $category = Get-HpiaRecommendationCategory -Recommendation $rec
        $name = Get-HpiaRecommendationName -Recommendation $rec
        $sp = Get-HpiaSoftPaqNumber -Recommendation $rec

        if (-not $sp) {
            continue
        }

        # Exclude BIOS/Firmware explicitly.
        $combined = ("{0} {1}" -f $category, $name)
        if ($combined -match '(?i)\bBIOS\b|Firmware') {
            $script:SkippedList.Add(("SP{0} {1}" -f $sp, $name)) | Out-Null
            Add-DriverResult -Vendor 'HP' -Name $name -Id $sp -Category $category -Status 'Blocked' -Message 'Excluded because it appears to be BIOS/Firmware.'
            continue
        }

        # Block unattended installation of boot-critical driver families.
        if (Test-HpiaBlockedUnattendedDriver -Category $category -Name $name) {
            $script:SkippedList.Add(("SP{0} {1}" -f $sp, $name)) | Out-Null
            Add-DriverResult -Vendor 'HP' -Name $name -Id $sp -Category $category -Status 'Blocked' -Message 'Excluded from unattended install because it appears to be storage/chipset/RST/VMD/NVMe/controller related.'
            continue
        }

        # Prefer safer driver-like recommendations, but avoid storage/chipset/controller families above.
        if ($combined -notmatch '(?i)Driver|Bluetooth|Audio|Graphics|Video|LAN|WLAN|Wireless|NIC|Touch|Fingerprint|Card Reader|Serial|USB' -and $category) {
            $script:SkippedList.Add(("SP{0} {1}" -f $sp, $name)) | Out-Null
            Add-DriverResult -Vendor 'HP' -Name $name -Id $sp -Category $category -Status 'Skipped' -Message 'Excluded because it did not appear to be a safe unattended driver recommendation.'
            continue
        }

        if (-not $seen.ContainsKey($sp)) {
            $seen[$sp] = $true
            $selected += $sp
            Add-DriverResult -Vendor 'HP' -Name $name -Id $sp -Category $category -Status 'Detected' -Message 'Selected for HPIA SPList install.'
        }
    }

    if (@($selected).Count -gt 0) {
        Set-Content -LiteralPath $SPListPath -Value $selected -Encoding ASCII
        Write-Log ("Created filtered HPIA SPList with {0} SoftPaqs: {1}" -f @($selected).Count, $SPListPath) 'OK'
        Add-YamlAction ("Created filtered HPIA SPList with {0} SoftPaqs: {1}" -f @($selected).Count, $SPListPath)
    }
    else {
        Write-Log 'No non-BIOS/Firmware driver SoftPaq recommendations were selected from HPIA reports.' 'OK'
        Add-YamlAction 'No non-BIOS/Firmware driver SoftPaq recommendations were selected from HPIA reports.'
    }

    return @($selected)
}



# -----------------------------
# HP PowerShell module maintenance
# -----------------------------
function Initialize-NetworkDefaults {
    try {
        [Net.ServicePointManager]::SecurityProtocol =
            [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls13
    }
    catch {
        try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 }
        catch {}
    }

    try { [Net.ServicePointManager]::DefaultConnectionLimit = 64 }
    catch {}
}

function Get-HighestInstalledModuleVersion {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Name
    )

    try {
        $installed = Get-Module -Name $Name -ListAvailable -ErrorAction SilentlyContinue |
            Sort-Object Version -Descending |
            Select-Object -First 1

        if ($installed) { return [version]$installed.Version }
    }
    catch {
        Write-Log "Could not determine installed version for module ${Name}: $($_.Exception.Message)" 'WARN'
    }

    return $null
}

function Invoke-HPCMSLUpdateInFreshPowerShell {
    [CmdletBinding()]
    param(
        [ValidateSet('CurrentUser','AllUsers')]
        [string]$Scope = 'AllUsers'
    )

    Write-Log 'Starting isolated HPCMSL install/update using PowerShellGet/PSResourceGet when available...' 'INFO'
    Add-YamlAction 'Starting isolated HPCMSL install/update before HP driver updates.'

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

    if ([string]::IsNullOrWhiteSpace(`$InstallRoot)) { throw 'InstallRoot is blank.' }
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

    if (`$installed) { HelperLog ('Installed HPCMSL version: {0}' -f `$installed) }
    else { HelperLog 'HPCMSL is not currently installed.' }
    HelperLog ('Latest PSGallery HPCMSL version: {0}' -f `$online)

    if (`$installed -and `$installed -ge `$online) {
        HelperLog 'HPCMSL is already current. No update needed.'
        exit 0
    }

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
        if ('$Scope' -eq 'AllUsers') { `$installRoot = Join-Path `$env:ProgramFiles 'WindowsPowerShell\Modules' }
        else { `$installRoot = Join-Path `$HOME 'Documents\WindowsPowerShell\Modules' }
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
        Initialize-NetworkDefaults
        Set-Content -Path $helperPath -Value $helperScript -Encoding UTF8 -Force
        Write-Log "HPCMSL helper script created at $helperPath" 'INFO'

        $args = @('-NoProfile','-ExecutionPolicy','Bypass','-File',('"{0}"' -f $helperPath))
        $process = Start-Process -FilePath powershell.exe -ArgumentList ($args -join ' ') -Wait -PassThru -WindowStyle Hidden

        if (Test-Path -LiteralPath $helperLog) {
            Get-Content -LiteralPath $helperLog -ErrorAction SilentlyContinue | ForEach-Object {
                if (-not [string]::IsNullOrWhiteSpace($_)) { Write-Log "HPCMSL helper: $_" 'INFO' }
            }
        }

        if ($process.ExitCode -ne 0) {
            Write-Log "Isolated HPCMSL helper failed with exit code $($process.ExitCode). Continuing with HPIA driver updates." 'WARN'
            Add-YamlAction ("Isolated HPCMSL helper failed with exit code {0}. Continuing with HPIA driver updates." -f $process.ExitCode)
            return $false
        }

        $finalVersion = Get-HighestInstalledModuleVersion -Name 'HPCMSL'
        if ($finalVersion) {
            Write-Log "HPCMSL final installed version after isolated update: $finalVersion" 'OK'
            Add-YamlAction ("HPCMSL final installed version after isolated update: {0}" -f $finalVersion)
        }
        else {
            Write-Log 'HPCMSL helper completed, but this session could not verify the installed HPCMSL version.' 'WARN'
            Add-YamlAction 'HPCMSL helper completed, but installed version could not be verified.'
            return $false
        }

        return $true
    }
    catch {
        Write-Log "Failed to run isolated HPCMSL update helper: $($_.Exception.Message)" 'WARN'
        Add-YamlAction ("Failed to run isolated HPCMSL update helper: {0}" -f $_.Exception.Message)
        return $false
    }
    finally {
        Remove-Item -LiteralPath $helperPath -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $helperLog -Force -ErrorAction SilentlyContinue
    }
}

function Invoke-HPPowerShellModuleMaintenance {
    [CmdletBinding()]
    param()

    Write-Section 'HP PowerShell Module Maintenance'
    Write-Log 'HP vendor detected. Checking/installing HPCMSL before HP driver updates...' 'INFO'
    Add-YamlAction 'HP vendor detected; checking/installing HPCMSL before HP driver updates.'

    $hpcmslUpdated = Invoke-HPCMSLUpdateInFreshPowerShell -Scope AllUsers
    if ($hpcmslUpdated) {
        Write-Log 'HP PowerShell module maintenance completed.' 'OK'
        Add-YamlAction 'HP PowerShell module maintenance completed.'
    }
    else {
        Write-Log 'HP PowerShell module maintenance did not complete successfully. Continuing with HP Image Assistant driver updates.' 'WARN'
        Add-YamlAction 'HP PowerShell module maintenance did not complete successfully; continuing with HPIA driver updates.'
    }
}

function Invoke-HPDriverUpdates {
    Write-Section 'HP Driver Analysis and Filtered Installation'

    $hpInfo = Get-HPSystemModelInfo

    $hpiaExe = Ensure-HPIAAvailable

    $hpiaReportFolder = Join-Path $YamlLogFolder 'HPIA'
    $hpiaDownloadFolder = Join-Path $WorkingRoot 'HPIADownloads'
    $hpiaExtractFolder = Join-Path $WorkingRoot 'HPIAExtracted'
    $hpiaSPListPath = Join-Path $WorkingRoot 'HPIA-Filtered-Safe-Drivers.splist'

    Ensure-Folder -Path $hpiaReportFolder
    Ensure-Folder -Path $hpiaDownloadFolder
    Ensure-Folder -Path $hpiaExtractFolder
    Ensure-Folder -Path 'C:\Temp'

    # HPIA is more reliable under Task Scheduler/SYSTEM when TEMP/TMP are local and TLS 1.2 is forced.
    $env:TEMP = 'C:\Temp'
    $env:TMP  = 'C:\Temp'

    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Write-Log 'TLS 1.2 enabled for HP Image Assistant network operations.' 'INFO'
        Add-YamlAction 'TLS 1.2 enabled for HPIA network operations.'
    }
    catch {
        Write-Log ("Unable to force TLS 1.2 before HPIA run: {0}" -f $_.Exception.Message) 'WARN'
        Add-YamlAction ("Unable to force TLS 1.2 before HPIA run: {0}" -f $_.Exception.Message)
    }

    # Safer unattended HPIA workflow:
    # 1. Run /Action:List first to generate recommendations only.
    # 2. Parse the HPIA report and create a filtered SPList.
    # 3. Install only the filtered SPList.
    #
    # Blocked unattended families include BIOS, Firmware, Storage, Chipset,
    # Intel RST, Intel VMD, NVMe, RAID/AHCI/SATA/controller, MEI/Management Engine,
    # and Thunderbolt because those can affect boot or storage-controller behavior.
    Write-Log 'Running HP Image Assistant recommendation/list pass only. No drivers will be installed in this step.' 'INFO'
    Add-YamlAction 'Running HPIA /Action:List first so recommendations can be filtered before install.'

    $hpiaListArgs = @(
        '/Operation:Analyze',
        '/Action:List',
        '/Category:Drivers',
        '/Selection:All',
        '/Silent',
        '/Noninteractive',
        '/Debug',
        "/ReportFolder:`"$hpiaReportFolder`"",
        "/SoftpaqDownloadFolder:`"$hpiaDownloadFolder`"",
        "/SoftpaqExtractFolder:`"$hpiaExtractFolder`""
    )

    Write-Log ("HPIA list command: {0} {1}" -f $hpiaExe, ($hpiaListArgs -join ' ')) 'INFO'
    Add-YamlAction ("HPIA list command: {0} {1}" -f $hpiaExe, ($hpiaListArgs -join ' '))

    Write-Progress -Activity 'HP Image Assistant' -Status ("Listing recommended drivers for {0}" -f $hpInfo.Model) -PercentComplete 25
    $hpiaListProc = Start-Process -FilePath $hpiaExe -ArgumentList ($hpiaListArgs -join ' ') -Wait -PassThru -NoNewWindow
    $listExitCode = [int]$hpiaListProc.ExitCode
    $listStatus = Get-HpiaExitStatus -ExitCode $listExitCode
    Write-Progress -Activity 'HP Image Assistant' -Completed

    Write-Log ("HP Image Assistant list pass completed with exit code {0} ({1})." -f $listExitCode, $listStatus) 'INFO'
    Add-YamlAction ("HPIA list pass completed with exit code {0} ({1})." -f $listExitCode, $listStatus)

    $reportFiles = @(Get-ChildItem -Path $hpiaReportFolder -File -Recurse -ErrorAction SilentlyContinue)
    foreach ($report in $reportFiles) {
        Add-YamlAction ("HPIA report generated: {0}" -f $report.FullName)
    }

    if ($reportFiles.Count -eq 0 -and $listExitCode -notin @(0,256,257,3010,3011,4096)) {
        throw "HPIA list pass failed or generated no reports. Exit code: $listExitCode ($listStatus). Review $hpiaReportFolder."
    }

    $selectedSoftPaqs = @(New-HpiaDriverSPList -ReportFolder $hpiaReportFolder -SPListPath $hpiaSPListPath)

    if ($selectedSoftPaqs.Count -eq 0) {
        Write-Log 'No safe HP driver SoftPaqs were selected for unattended installation after filtering.' 'OK'
        Add-YamlAction 'No safe HP driver SoftPaqs were selected for unattended installation after filtering.'
        return
    }

    Write-Log ("Installing filtered safe HP driver SPList. Count: {0}" -f $selectedSoftPaqs.Count) 'INFO'
    Add-YamlAction ("Installing filtered safe HP driver SPList: {0}" -f $hpiaSPListPath)

    $hpiaInstallArgs = @(
        "/SPList:`"$hpiaSPListPath`"",
        '/Silent',
        '/Noninteractive',
        '/AutoCleanup',
        '/Debug',
        "/ReportFolder:`"$hpiaReportFolder`"",
        "/SoftpaqDownloadFolder:`"$hpiaDownloadFolder`"",
        "/SoftpaqExtractFolder:`"$hpiaExtractFolder`""
    )

    Write-Log ("HPIA filtered SPList install command: {0} {1}" -f $hpiaExe, ($hpiaInstallArgs -join ' ')) 'INFO'
    Add-YamlAction ("HPIA filtered SPList install command: {0} {1}" -f $hpiaExe, ($hpiaInstallArgs -join ' '))

    Write-Progress -Activity 'HP Image Assistant' -Status 'Installing filtered safe driver SoftPaqs' -PercentComplete 75
    $hpiaInstallProc = Start-Process -FilePath $hpiaExe -ArgumentList ($hpiaInstallArgs -join ' ') -Wait -PassThru -NoNewWindow
    $installExitCode = [int]$hpiaInstallProc.ExitCode
    $installStatus = Get-HpiaExitStatus -ExitCode $installExitCode
    Write-Progress -Activity 'HP Image Assistant' -Completed

    Write-Log ("HP Image Assistant filtered SPList install completed with exit code {0} ({1})." -f $installExitCode, $installStatus) 'INFO'
    Add-YamlAction ("HPIA filtered SPList install completed with exit code {0} ({1})." -f $installExitCode, $installStatus)

    $downloadedFiles = @(Get-ChildItem -Path $hpiaDownloadFolder -File -Recurse -ErrorAction SilentlyContinue)
    foreach ($file in $downloadedFiles) {
        Add-DriverResult -Vendor 'HP' -Name $file.Name -Id $null -Category 'Driver' -Status 'Detected' -Message ("Downloaded/processed by filtered HPIA SPList: {0}" -f $file.FullName)
    }

    switch ($installExitCode) {
        0 {
            Write-Log 'HPIA filtered SPList install completed successfully.' 'OK'
            Add-YamlAction 'HPIA filtered SPList install completed successfully.'
            return
        }
        256 {
            Write-Log 'HPIA completed successfully. No applicable filtered driver recommendations were found or no action was required.' 'OK'
            Add-YamlAction 'HPIA completed successfully with no applicable filtered driver recommendations or no action required.'
            return
        }
        257 {
            Write-Log 'HPIA completed and reported filtered driver actions.' 'OK'
            Add-YamlAction 'HPIA completed and reported filtered driver actions.'
            return
        }
        3010 {
            Write-Log 'HPIA filtered SPList install completed successfully. Reboot required.' 'WARN'
            Add-YamlAction 'HPIA filtered SPList install completed successfully and indicated reboot required.'
            return
        }
        3011 {
            Write-Log 'One or more filtered HPIA items were not auto-installable and were skipped.' 'WARN'
            Add-YamlAction 'One or more filtered HPIA items were not auto-installable and were skipped.'
            return
        }
        4096 {
            Write-Log 'HPIA completed but did not find applicable filtered driver updates for this platform.' 'OK'
            Add-YamlAction 'HPIA completed but did not find applicable filtered driver updates for this platform.'
            return
        }
        8193 {
            Write-Log 'HPIA filtered install returned 8193. Checking whether reports/logs were generated before failing the run.' 'WARN'
            Add-YamlAction 'HPIA filtered install returned 8193; checking for reports/logs before hard failure.'
            $generatedReports = @(Get-ChildItem -Path $hpiaReportFolder -File -Recurse -ErrorAction SilentlyContinue)
            if ($generatedReports.Count -gt 0) {
                Write-Log ("HPIA generated {0} report/log file(s) despite exit code 8193. Continuing so the scheduled workflow does not hard fail." -f $generatedReports.Count) 'WARN'
                Add-YamlAction ("HPIA generated {0} report/log file(s) despite exit code 8193. Continuing without hard failure." -f $generatedReports.Count)
                return
            }
            throw "HPIA failed with exit code 8193 and produced no report/log files in $hpiaReportFolder."
        }
        default {
            throw "HP Image Assistant filtered SPList install failed or returned an unexpected exit code: $installExitCode ($installStatus). Review reports in $hpiaReportFolder."
        }
    }
}

# -----------------------------
# Dell support
# -----------------------------
function Wait-ForFile {
    param(
        [Parameter(Mandatory)][string]$Path,
        [int]$TimeoutSeconds = 30
    )

    $start = Get-Date
    while (-not (Test-Path -LiteralPath $Path)) {
        Start-Sleep -Seconds 1
        if (((Get-Date) - $start).TotalSeconds -ge $TimeoutSeconds) {
            return $false
        }
    }
    return $true
}

function Get-DcuReportXmlSafely {
    param(
        [Parameter(Mandatory)][string]$Path,
        [int]$RetryCount = 5,
        [int]$RetryDelaySeconds = 2
    )

    if (-not (Wait-ForFile -Path $Path -TimeoutSeconds 20)) {
        throw "Dell DCU report file was not found: $Path"
    }

    Start-Sleep -Seconds 3

    for ($i = 1; $i -le $RetryCount; $i++) {
        try {
            $fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
            try {
                $sr = New-Object System.IO.StreamReader($fs)
                try {
                    $content = $sr.ReadToEnd()
                }
                finally {
                    $sr.Dispose()
                }
            }
            finally {
                $fs.Dispose()
            }

            $xml = New-Object System.Xml.XmlDocument
            $xml.LoadXml($content)
            return $xml
        }
        catch {
            if ($i -eq $RetryCount) {
                throw
            }
            Start-Sleep -Seconds $RetryDelaySeconds
        }
    }
}

function Get-DellNodeText {
    param(
        [Parameter(Mandatory)][xml]$Node,
        [Parameter(Mandatory)][string[]]$Names
    )

    foreach ($name in $Names) {
        try {
            $xpath = './/*[local-name()="' + $name + '"]'
            $child = $Node.SelectSingleNode($xpath)
            if ($child -and -not [string]::IsNullOrWhiteSpace($child.InnerText)) {
                return $child.InnerText.Trim()
            }
        }
        catch {}
    }

    return $null
}

function Get-DellReportItems {
    param([Parameter(Mandatory)][xml]$Xml)

    $items = @()
    try {
        $xpath = '//*[local-name()="Update" or local-name()="Package" or local-name()="SoftwareComponent" or local-name()="component" or local-name()="Device"]'
        $nodes = $Xml.SelectNodes($xpath)
        foreach ($node in $nodes) {
            $name = Get-DellNodeText -Node $node -Names @('Name','Title','PackageName')
            $version = Get-DellNodeText -Node $node -Names @('Version','PackageVersion')
            $category = Get-DellNodeText -Node $node -Names @('Category','Type')
            $id = Get-DellNodeText -Node $node -Names @('Id','PackageId','ReleaseId')

            if ($name -or $id) {
                $items += [pscustomobject]@{
                    Id       = $id
                    Name     = $name
                    Version  = $version
                    Category = $category
                }
            }
        }
    }
    catch {}

    return @($items)
}

function Get-DellDCUService {
    $candidateNames = @(
        'DellClientManagementService',
        'DellCommandUpdate',
        'DellUpdateService'
    )

    foreach ($name in $candidateNames) {
        $svc = Get-Service -Name $name -ErrorAction SilentlyContinue
        if ($svc) { return $svc }
    }

    $svc = Get-Service -ErrorAction SilentlyContinue |
        Where-Object { $_.DisplayName -match 'Dell.*Client.*Management|Dell.*Command.*Update' } |
        Select-Object -First 1

    return $svc
}

function Ensure-DellDCUService {
    param(
        [int]$TimeoutSeconds = 30
    )

    Write-Log 'Validating Dell Client Management Service...' 'INFO'
    Add-YamlAction 'Validating Dell Client Management Service.'

    $service = Get-DellDCUService
    if (-not $service) {
        throw 'Dell Client Management Service was not found. Dell Command | Update may need to be repaired or reinstalled.'
    }

    Write-Log ("Dell service detected: {0} ({1})" -f $service.DisplayName, $service.Name) 'OK'

    try {
        $wmiService = Get-CimInstance -ClassName Win32_Service -Filter ("Name='{0}'" -f $service.Name) -ErrorAction Stop
        if ($wmiService.StartMode -eq 'Disabled') {
            Write-Log 'Dell service startup type is Disabled. Setting it to Manual.' 'WARN'
            Set-Service -Name $service.Name -StartupType Manual -ErrorAction Stop
        }
    }
    catch {
        Write-Log ("Could not validate/set Dell service startup type: {0}" -f $_.Exception.Message) 'WARN'
    }

    $service.Refresh()
    if ($service.Status -ne 'Running') {
        Write-Log 'Starting Dell Client Management Service...' 'INFO'
        try {
            Start-Service -Name $service.Name -ErrorAction Stop
        }
        catch {
            Write-Log ("Start-Service failed. Attempting sc.exe recovery start. Error: {0}" -f $_.Exception.Message) 'WARN'
            & sc.exe start $service.Name | Out-Null
        }
    }

    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    do {
        Start-Sleep -Seconds 2
        $service = Get-Service -Name $service.Name -ErrorAction SilentlyContinue
        if ($service -and $service.Status -eq 'Running') {
            Write-Log 'Dell Client Management Service is running.' 'OK'
            Add-YamlAction 'Dell Client Management Service is running.'
            return $true
        }
    } while ((Get-Date) -lt $deadline)

    throw 'Dell Client Management Service did not reach the Running state before timeout.'
}

function Invoke-DellDCUCommandWithRetry {
    param(
        [Parameter(Mandatory)][string]$DcuCli,
        [Parameter(Mandatory)][string]$Arguments,
        [Parameter(Mandatory)][string]$OperationName,
        [int[]]$AcceptableExitCodes = @(0),
        [int]$MaxAttempts = 2
    )

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        Write-Log ("Dell DCU {0} attempt {1} of {2}..." -f $OperationName, $attempt, $MaxAttempts) 'INFO'
        Add-YamlAction ("Dell DCU {0} attempt {1} of {2}." -f $OperationName, $attempt, $MaxAttempts)

        $proc = Start-Process -FilePath $DcuCli -ArgumentList $Arguments -Wait -PassThru -NoNewWindow
        $exitCode = [int]$proc.ExitCode
        Write-Log ("Dell DCU {0} exit code: {1}" -f $OperationName, $exitCode) 'INFO'

        if ($AcceptableExitCodes -contains $exitCode) {
            return $proc
        }

        if ($exitCode -eq 3000) {
            Write-Log 'Dell DCU returned 3000, which normally indicates the Dell Client Management Service stopped or crashed.' 'WARN'
            Add-YamlAction 'Dell DCU returned 3000; attempting Dell service recovery before retry.'
        }
        else {
            Write-Log ("Dell DCU {0} returned non-success exit code {1}." -f $OperationName, $exitCode) 'WARN'
        }

        if ($attempt -lt $MaxAttempts) {
            Ensure-DellDCUService | Out-Null
            Start-Sleep -Seconds 5
            continue
        }

        throw "Dell DCU $OperationName failed after $MaxAttempts attempt(s). Last exit code: $exitCode"
    }
}


function Test-DotNetDesktopRuntime8 {
    [CmdletBinding()]
    param([version]$MinimumVersion = [version]'8.0.25')

    $dotnetExe = Join-Path $env:ProgramFiles 'dotnet\dotnet.exe'
    if (-not (Test-Path -LiteralPath $dotnetExe)) { return $false }
    try {
        foreach ($line in @(& $dotnetExe --list-runtimes 2>$null)) {
            if ($line -match '^Microsoft\.WindowsDesktop\.App\s+(?<Version>\d+\.\d+\.\d+)\s+') {
                try { if ([version]$Matches.Version -ge $MinimumVersion) { return $true } } catch {}
            }
        }
    } catch { Write-Log ("Unable to query installed .NET desktop runtimes: {0}" -f $_.Exception.Message) 'WARN' }
    return $false
}

function Invoke-InstallerAndValidateExitCode {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$FilePath,
        [Parameter(Mandatory)][string[]]$ArgumentList,
        [Parameter(Mandatory)][string]$DisplayName,
        [int[]]$SuccessExitCodes = @(0, 3010, 1641)
    )
    $process = Start-Process -FilePath $FilePath -ArgumentList $ArgumentList -Wait -PassThru -NoNewWindow
    Write-Log ("{0} installer exit code: {1}" -f $DisplayName, $process.ExitCode) 'INFO'
    if ($process.ExitCode -notin $SuccessExitCodes) { throw ("{0} installation failed with exit code {1}." -f $DisplayName, $process.ExitCode) }
    if ($process.ExitCode -in @(3010,1641)) { Add-YamlAction ("{0} installed; restart required." -f $DisplayName) }
    else { Add-YamlAction ("{0} installed successfully." -f $DisplayName) }
}

function Save-InternetFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][uri]$Uri,
        [Parameter(Mandatory)][string]$DestinationPath,
        [Parameter(Mandatory)][string]$DisplayName
    )
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    if (Test-Path -LiteralPath $DestinationPath) { Remove-Item -LiteralPath $DestinationPath -Force -ErrorAction SilentlyContinue }
    try { Start-BitsTransfer -Source $Uri.AbsoluteUri -Destination $DestinationPath -ErrorAction Stop }
    catch { Invoke-WebRequest -Uri $Uri.AbsoluteUri -OutFile $DestinationPath -UseBasicParsing -TimeoutSec 300 -ErrorAction Stop }
    if (-not (Test-Path -LiteralPath $DestinationPath) -or (Get-Item -LiteralPath $DestinationPath).Length -lt 1MB) {
        throw ("The downloaded {0} installer is missing or unexpectedly small." -f $DisplayName)
    }
}

function Ensure-DotNetDesktopRuntimeForDcu {
    [CmdletBinding()]
    param()

    if (Test-DotNetDesktopRuntime8 -MinimumVersion ([version]'8.0.25')) {
        Write-Log '.NET Desktop Runtime 8.0.25 or newer is already installed.' 'OK'
        return
    }

    $localInstaller = Join-Path $WorkingRoot 'windowsdesktop-runtime-8-win-x64.exe'
    if (Test-Path -LiteralPath $DotNetDesktopRuntimeSharePath) {
        Copy-Item -LiteralPath $DotNetDesktopRuntimeSharePath -Destination $localInstaller -Force -ErrorAction Stop
    } else {
        Write-Log 'The .NET share installer is unavailable; using the Microsoft Internet fallback.' 'WARN'
        Save-InternetFile -Uri ([uri]$DotNetDesktopRuntimeInternetUrl) -DestinationPath $localInstaller -DisplayName '.NET Desktop Runtime 8 x64'
    }

    Invoke-InstallerAndValidateExitCode -FilePath $localInstaller -ArgumentList @('/install','/quiet','/norestart') -DisplayName '.NET Desktop Runtime 8 x64'
    if (-not (Test-DotNetDesktopRuntime8 -MinimumVersion ([version]'8.0.25'))) {
        throw '.NET Desktop Runtime installation completed, but version 8.0.25 or newer was not detected.'
    }
}

function Ensure-DellCommandUpdateInstalled {
    [CmdletBinding()]
    param()

    $dcuCliPath = Join-Path ${env:ProgramFiles} 'Dell\CommandUpdate\dcu-cli.exe'
    if (Test-Path -LiteralPath $dcuCliPath) { return $dcuCliPath }

    Ensure-DotNetDesktopRuntimeForDcu
    if (Test-Path -LiteralPath $DellCommandUpdateSharePath) {
        $localDcuInstaller = Join-Path $WorkingRoot 'Dell-Command-Update-Windows-Universal-Application.exe'
        Copy-Item -LiteralPath $DellCommandUpdateSharePath -Destination $localDcuInstaller -Force -ErrorAction Stop
        Invoke-InstallerAndValidateExitCode -FilePath $localDcuInstaller -ArgumentList @('/s') -DisplayName 'Dell Command | Update'
    } else {
        Write-Log 'The Dell share installer is unavailable; using winget as the Internet fallback.' 'WARN'
        $winget = Get-Command -Name 'winget.exe' -ErrorAction SilentlyContinue
        if (-not $winget) { throw 'Dell Command | Update is not on the file share and winget.exe is unavailable.' }
        $wingetArgs = @('install','--id',$DellCommandUpdateWingetId,'--exact','--silent','--accept-package-agreements','--accept-source-agreements','--disable-interactivity')
        $wingetProcess = Start-Process -FilePath $winget.Source -ArgumentList $wingetArgs -Wait -PassThru -NoNewWindow
        if ($wingetProcess.ExitCode -ne 0) { throw ("winget failed to install Dell Command | Update. Exit code: {0}" -f $wingetProcess.ExitCode) }
    }

    Start-Sleep -Seconds 3
    if (-not (Test-Path -LiteralPath $dcuCliPath)) { throw "Dell Command | Update installed, but dcu-cli.exe was not found: $dcuCliPath" }
    return $dcuCliPath
}

function Invoke-DellDriverUpdates {
    Write-Section 'Dell Command Update Workflow'

    $dcuCli = Ensure-DellCommandUpdateInstalled

    Write-Log ("Using Dell Command | Update CLI: {0}" -f $dcuCli) 'OK'
    Add-YamlAction 'Using Dell Command | Update CLI.'

    Ensure-DellDCUService | Out-Null

    $dcuScanLog  = Join-Path $WorkingRoot 'Dell-DCU-Scan.log'
    $dcuApplyLog = Join-Path $WorkingRoot 'Dell-DCU-Apply.log'
    $dcuReport   = Join-Path $WorkingRoot 'Dell-DCU-ApplicableUpdates.xml'

    Write-Log 'Dell DCU Configure...' 'INFO'
    Write-Progress -Activity 'Dell Driver Update Workflow' -Status 'Configuring Dell Command Update' -PercentComplete 10
    $configureArgs = "/configure -silent -scheduleAuto -lockSettings=disable"
    Invoke-DellDCUCommandWithRetry -DcuCli $dcuCli -Arguments $configureArgs -OperationName 'Configure' -MaxAttempts 2 | Out-Null

    Write-Log 'Dell DCU Scan...' 'INFO'
    Write-Progress -Activity 'Dell Driver Update Workflow' -Status 'Scanning for updates' -PercentComplete 35
    $scanArgs = "/scan -silent -updateType=driver,application -outputLog=""$dcuScanLog"" -report=""$dcuReport"""
    Invoke-DellDCUCommandWithRetry -DcuCli $dcuCli -Arguments $scanArgs -OperationName 'Scan' -MaxAttempts 2 | Out-Null
    Write-Log 'Waiting for Dell Command | Update to finish writing the report...' 'INFO'

    try {
        $xml = Get-DcuReportXmlSafely -Path $dcuReport
        $items = Get-DellReportItems -Xml $xml

        if ($items.Count -gt 0) {
            Add-YamlAction ("Dell DCU report parsed successfully. Updates detected: {0}" -f $items.Count)

            $total = $items.Count
            $index = 0

            foreach ($item in $items) {
                $index++
                $percent = 35 + [math]::Floor(($index / $total) * 35)
                $label = if ($item.Name) { $item.Name } elseif ($item.Id) { $item.Id } else { 'Dell update' }

                Write-Progress -Activity 'Dell Driver Update Workflow' -Status ("Parsing report: {0}" -f $label) -PercentComplete $percent

                if ($item.Category -match 'BIOS|Firmware') {
                    $script:SkippedList.Add($label) | Out-Null
                    Add-DriverResult -Vendor 'Dell' -Name $item.Name -Id $item.Id -Category $item.Category -Status 'Blocked' -Message 'BIOS/Firmware update blocked by script policy.'
                    Write-Log ("Blocking Dell BIOS/Firmware update: {0}" -f $label) 'WARN'
                }
                else {
                    Add-DriverResult -Vendor 'Dell' -Name $item.Name -Id $item.Id -Category $item.Category -Status 'Detected' -Message 'Detected in Dell DCU report.'
                }
            }
        }
        else {
            Add-YamlAction 'Dell DCU report parsed but returned no identifiable updates.'
        }
    }
    catch {
        Write-Log ("Failed to parse Dell DCU report: {0}" -f $_.Exception.Message) 'WARN'
        Add-YamlAction ("Failed to parse Dell DCU report: {0}" -f $_.Exception.Message)
    }

    Write-Log 'Dell DCU ApplyUpdates for driver/application updates only. BIOS/Firmware are excluded by -updateType.' 'INFO'
    Write-Progress -Activity 'Dell Driver Update Workflow' -Status 'Applying updates' -PercentComplete 85
    $applyArgs = "/applyUpdates -silent -updateType=driver,application -reboot=disable -outputLog=""$dcuApplyLog"""
    Invoke-DellDCUCommandWithRetry -DcuCli $dcuCli -Arguments $applyArgs -OperationName 'ApplyUpdates' -MaxAttempts 2 | Out-Null

    Write-Log (("Dell DCU logs: {0} ; {1} ; {2}") -f $dcuScanLog, $dcuApplyLog, $dcuReport) 'OK'
    Write-Progress -Activity 'Dell Driver Update Workflow' -Completed
}
# -----------------------------
# Main
# -----------------------------
$finalStatus = 'success'

try {
    Write-Section 'Initialization'

    Ensure-Folder -Path $YamlLogFolder
    Ensure-Folder -Path $WorkingRoot
    Ensure-WorkingFolderPermissions -Path $WorkingRoot

    $yamlName = "{0}-{1}-{2}.yml" -f $script:ComputerName, '05_Weekend_Vendor_Drivers_Update', (Get-Date -Format 'yyyy-MM-dd_HHmmss')
    $yamlPath = Join-Path $YamlLogFolder $yamlName

    Write-Log ("YAML log will be written to: {0}" -f $yamlPath) 'INFO'
    Initialize-YamlLog -ComputerName $script:ComputerName -YamlPath $yamlPath

    Write-Log 'Initializing vendor driver update script...' 'INFO'


    Write-Section 'Vendor Detection'
    $script:DetectedVendor = Get-DriverVendor
    Write-Log ("Detected vendor workflow: {0}" -f $script:DetectedVendor) 'INFO'
    Write-Log ("Working root: {0}" -f $WorkingRoot) 'INFO'

    if ($script:DetectedVendor -eq 'HP') {
        Invoke-HPPowerShellModuleMaintenance
        Invoke-HPDriverUpdates
    }
    elseif ($script:DetectedVendor -eq 'Dell') {
        Invoke-DellDriverUpdates
    }

}
catch {
    $finalStatus = 'failed'
    Add-RunFailure ("Script failed: {0}" -f $_.Exception.Message)
}
finally {
    Write-Section 'Cleanup'
    Remove-WorkingFolderRobust -Path $WorkingRoot

    if ($script:RunFailures.Count -gt 0 -and $finalStatus -ne 'failed') {
        $finalStatus = 'completed_with_warnings'
    }

    if ($script:RunFailures.Count -gt 0) {
        Write-Log (("{0} driver update completed with one or more failures.") -f $script:DetectedVendor) 'WARN'
    }
    else {
        Write-Log (("{0} driver update script completed successfully.") -f $script:DetectedVendor) 'OK'
    }

    Save-YamlLog -Status $finalStatus
}# ScriptName: 05_Weekend_HP_Drivers_Update.ps1
# ScriptVersion: 2.6.0
# LastUpdated: 2026-06-29
# Purpose: Weekend vendor driver update script with clean HP + Dell support,
#          HP-only HPCMSL maintenance after vendor detection,
#          YAML logging, colored output,
#          section headers, progress display, and structured per-driver results.

[CmdletBinding()]
param([string]$WorkingRoot = 'C:\Temp\DriverUpdates',
    [string]$YamlLogFolder = 'C:\Logs',
    [switch]$IncludeSoftware,
    [switch]$IncludeBIOS,
    [string]$HpiaSourceFolder = '\\filesvr\Labscripts\HPImageAssistant',
    [string]$LocalHpiaFolder = 'C:\ProgramData\Compton\HPImageAssistant'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# -----------------------------
# Script metadata
# -----------------------------
$script:ScriptName        = '05_Weekend_HP_Drivers_Update.ps1'
$script:ScriptVersion     = '2.6.0'
$script:StartTime         = Get-Date
$script:RunFailures       = New-Object System.Collections.Generic.List[string]
$script:InstalledList     = New-Object System.Collections.Generic.List[string]
$script:SkippedList       = New-Object System.Collections.Generic.List[string]
$script:YamlActionLines   = New-Object System.Collections.Generic.List[string]
$script:DriverResults     = New-Object System.Collections.Generic.List[object]
$script:DetectedVendor    = $null
$script:YamlPath          = $null
$script:ComputerName      = $env:COMPUTERNAME

# -----------------------------
# Logging helpers
# -----------------------------
function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','OK','WARN','ERROR')][string]$Level = 'INFO'
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[{0}] [{1}] {2}" -f $timestamp, $Level.PadRight(5), $Message

    switch ($Level) {
        'INFO'  { Write-Host $line -ForegroundColor Cyan }
        'OK'    { Write-Host $line -ForegroundColor Green }
        'WARN'  { Write-Host $line -ForegroundColor Yellow }
        'ERROR' { Write-Host $line -ForegroundColor Red }
        default { Write-Host $line }
    }
}

function Write-Section {
    param([Parameter(Mandatory)][string]$Title)

    $border = ('=' * 72)
    Write-Host ''
    Write-Host $border -ForegroundColor Magenta
    Write-Host ("  {0}" -f $Title) -ForegroundColor Magenta
    Write-Host $border -ForegroundColor Magenta
    Add-YamlAction ("Section: {0}" -f $Title)
}

function Add-RunFailure {
    param([Parameter(Mandatory)][string]$Message)
    $script:RunFailures.Add($Message) | Out-Null
    Write-Log $Message 'WARN'
}

function Add-DriverResult {
    param(
        [Parameter(Mandatory)][string]$Vendor,
        [Parameter(Mandatory)][string]$Name,
        [string]$Id,
        [string]$Category,
        [ValidateSet('Detected','Installed','Downloaded','Skipped','Failed','Blocked')][string]$Status,
        [string]$Message = ''
    )

    $script:DriverResults.Add([pscustomobject]@{
        Vendor   = $Vendor
        Name     = $Name
        Id       = $Id
        Category = $Category
        Status   = $Status
        Message  = $Message
    }) | Out-Null
}

function ConvertTo-YamlSafeString {
    param([AllowNull()][object]$Value)

    if ($null -eq $Value) { return 'null' }

    $text = [string]$Value
    $text = $text -replace "`r", ''
    $text = $text -replace "`n", ' '
    $text = $text -replace "'", "''"
    return "'$text'"
}

function Initialize-YamlLog {
    param(
        [Parameter(Mandatory)][string]$ComputerName,
        [Parameter(Mandatory)][string]$YamlPath
    )

    $script:YamlPath = $YamlPath
    $script:YamlActionLines.Clear()

    Add-YamlAction 'Script initialized.'
    Add-YamlAction ("Working root: {0}" -f $WorkingRoot)
    Add-YamlAction ("Computer: {0}" -f $ComputerName)
}

function Add-YamlAction {
    param([Parameter(Mandatory)][string]$Text)
    $script:YamlActionLines.Add(("    - {0}" -f (ConvertTo-YamlSafeString $Text))) | Out-Null
}

function Save-YamlLog {
    param(
        [Parameter(Mandatory)][string]$Status
    )

    if ([string]::IsNullOrWhiteSpace($script:YamlPath)) {
        return
    }

    $endTime = Get-Date
    $duration = [math]::Round(($endTime - $script:StartTime).TotalSeconds, 0)

    $lines = New-Object System.Collections.Generic.List[string]

    $lines.Add('script:') | Out-Null
    $lines.Add(("  name: {0}" -f (ConvertTo-YamlSafeString $script:ScriptName))) | Out-Null
    $lines.Add(("  version: {0}" -f (ConvertTo-YamlSafeString $script:ScriptVersion))) | Out-Null
    $lines.Add(("  computer: {0}" -f (ConvertTo-YamlSafeString $script:ComputerName))) | Out-Null
    $lines.Add(("  started: {0}" -f (ConvertTo-YamlSafeString ($script:StartTime.ToString('s'))))) | Out-Null
    $lines.Add(("  ended: {0}" -f (ConvertTo-YamlSafeString ($endTime.ToString('s'))))) | Out-Null
    $lines.Add(("  duration_seconds: {0}" -f $duration)) | Out-Null

    $lines.Add('run:') | Out-Null
    $lines.Add(("  status: {0}" -f (ConvertTo-YamlSafeString $Status))) | Out-Null
    $lines.Add(("  vendor: {0}" -f (ConvertTo-YamlSafeString $script:DetectedVendor))) | Out-Null

    $lines.Add('  actions:') | Out-Null
    if ($script:YamlActionLines.Count -eq 0) {
        $lines.Add('    []') | Out-Null
    }
    else {
        foreach ($line in $script:YamlActionLines) {
            $lines.Add($line) | Out-Null
        }
    }

    $lines.Add('  installed:') | Out-Null
    if ($script:InstalledList.Count -eq 0) {
        $lines.Add('    []') | Out-Null
    }
    else {
        foreach ($item in $script:InstalledList) {
            $lines.Add(("    - {0}" -f (ConvertTo-YamlSafeString $item))) | Out-Null
        }
    }

    $lines.Add('  skipped:') | Out-Null
    if ($script:SkippedList.Count -eq 0) {
        $lines.Add('    []') | Out-Null
    }
    else {
        foreach ($item in $script:SkippedList) {
            $lines.Add(("    - {0}" -f (ConvertTo-YamlSafeString $item))) | Out-Null
        }
    }

    $lines.Add('  failures:') | Out-Null
    if ($script:RunFailures.Count -eq 0) {
        $lines.Add('    []') | Out-Null
    }
    else {
        foreach ($item in $script:RunFailures) {
            $lines.Add(("    - {0}" -f (ConvertTo-YamlSafeString $item))) | Out-Null
        }
    }

    $lines.Add('drivers:') | Out-Null
    if ($script:DriverResults.Count -eq 0) {
        $lines.Add('  []') | Out-Null
    }
    else {
        foreach ($driver in $script:DriverResults) {
            $lines.Add('  -') | Out-Null
            $lines.Add(("    vendor: {0}" -f (ConvertTo-YamlSafeString $driver.Vendor))) | Out-Null
            $lines.Add(("    name: {0}" -f (ConvertTo-YamlSafeString $driver.Name))) | Out-Null
            $lines.Add(("    id: {0}" -f (ConvertTo-YamlSafeString $driver.Id))) | Out-Null
            $lines.Add(("    category: {0}" -f (ConvertTo-YamlSafeString $driver.Category))) | Out-Null
            $lines.Add(("    status: {0}" -f (ConvertTo-YamlSafeString $driver.Status))) | Out-Null
            $lines.Add(("    message: {0}" -f (ConvertTo-YamlSafeString $driver.Message))) | Out-Null
        }
    }

    Set-Content -LiteralPath $script:YamlPath -Value $lines -Encoding UTF8
    Write-Host ("YAML log written successfully: {0}" -f $script:YamlPath) -ForegroundColor Green
}

# -----------------------------
# File/folder helpers
# -----------------------------
function Ensure-Folder {
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -Path $Path -ItemType Directory -Force | Out-Null
    }
}

function Ensure-WorkingFolderPermissions {
    param([Parameter(Mandatory)][string]$Path)

    try {
        & icacls.exe $Path '/grant' '*S-1-1-0:(OI)(CI)F' '/T' '/C' | Out-Null
    }
    catch {
        Write-Log ("Unable to relax working folder permissions on {0}: {1}" -f $Path, $_.Exception.Message) 'WARN'
    }
}

function Remove-WorkingFolderRobust {
    param(
        [Parameter(Mandatory)][string]$Path,
        [int]$RetryCount = 6,
        [int]$RetryDelaySeconds = 5
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        Write-Log ("Working folder already absent: {0}" -f $Path) 'OK'
        return
    }

    Write-Log ("Attempting to remove working folder: {0}" -f $Path) 'INFO'
    for ($i = 1; $i -le $RetryCount; $i++) {
        try {
            Remove-Item -LiteralPath $Path -Recurse -Force -ErrorAction Stop
            Write-Log 'Working folder removed successfully.' 'OK'
            return
        }
        catch {
            if ($i -eq $RetryCount) {
                Add-RunFailure ("Failed to remove working folder after retries: {0}" -f $_.Exception.Message)
                return
            }
            Start-Sleep -Seconds $RetryDelaySeconds
        }
    }
}

# -----------------------------
# Vendor detection
# -----------------------------
function Get-SystemManufacturer {
    try {
        return (Get-CimInstance Win32_ComputerSystem -ErrorAction Stop).Manufacturer
    }
    catch {
        throw "Unable to determine system manufacturer. $($_.Exception.Message)"
    }
}

function Get-DriverVendor {
    $manufacturer = Get-SystemManufacturer
    Write-Log ("Detected manufacturer: {0}" -f $manufacturer) 'INFO'

    if ($manufacturer -match 'Dell') { return 'Dell' }
    if ($manufacturer -match 'HP|Hewlett-Packard') { return 'HP' }

    throw "Unsupported manufacturer for this script: $manufacturer"
}

# -----------------------------
# HP Support - HP Image Assistant extracted-folder deployment
# -----------------------------
function Get-HPSystemModelInfo {
    [CmdletBinding()]
    param()

    $cs = Get-CimInstance Win32_ComputerSystem -ErrorAction Stop
    $csp = Get-CimInstance Win32_ComputerSystemProduct -ErrorAction SilentlyContinue
    $bb = Get-CimInstance Win32_BaseBoard -ErrorAction SilentlyContinue

    $platform = $null
    if ($bb -and $bb.Product) {
        $platform = $bb.Product.ToString().Trim().ToUpper()
        if ($platform.Length -gt 4) {
            $platform = $platform.Substring(0,4)
        }
    }

    $model = if ($cs.Model) { $cs.Model.ToString().Trim() } else { 'Unknown' }
    $sku = if ($csp -and $csp.Version) { $csp.Version.ToString().Trim() } else { 'Unknown' }

    $info = [pscustomobject]@{
        Manufacturer = $cs.Manufacturer
        Model        = $model
        SKU          = $sku
        Platform     = $platform
    }

    Write-Log ("Detected HP system model: {0}" -f $info.Model) 'INFO'
    Write-Log ("Detected HP platform/baseboard ID: {0}" -f $info.Platform) 'INFO'
    Add-YamlAction ("Detected HP system model: {0}" -f $info.Model)
    Add-YamlAction ("Detected HP platform/baseboard ID: {0}" -f $info.Platform)

    return $info
}

function Get-ExistingHPIAExecutable {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$PreferredFolder)

    $candidateFolders = New-Object System.Collections.Generic.List[string]

    if (-not [string]::IsNullOrWhiteSpace($PreferredFolder)) {
        $candidateFolders.Add($PreferredFolder) | Out-Null
    }

    foreach ($folder in @(
        'C:\Program Files\HP\HP Image Assistant',
        'C:\Program Files (x86)\HP\HP Image Assistant',
        'C:\SWSetup\HPImageAssistant',
        'C:\ProgramData\Compton\HPImageAssistant'
    )) {
        if (-not $candidateFolders.Contains($folder)) {
            $candidateFolders.Add($folder) | Out-Null
        }
    }

    foreach ($folder in $candidateFolders) {
        if ([string]::IsNullOrWhiteSpace($folder) -or -not (Test-Path -LiteralPath $folder)) {
            continue
        }

        $exe = Get-ChildItem -Path $folder -Recurse -Filter 'HPImageAssistant.exe' -File -ErrorAction SilentlyContinue |
            Sort-Object FullName |
            Select-Object -First 1

        if ($exe) {
            return $exe.FullName
        }
    }

    return $null
}

function Install-HPIAFromExtractedFolder {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SourceFolder,
        [Parameter(Mandatory)][string]$DestinationFolder
    )

    Write-Section 'HP Image Assistant Local Deployment'
    Write-Log ("HPIA source folder: {0}" -f $SourceFolder) 'INFO'
    Write-Log ("HPIA local folder: {0}" -f $DestinationFolder) 'INFO'
    Add-YamlAction ("HPIA source folder: {0}" -f $SourceFolder)
    Add-YamlAction ("HPIA local folder: {0}" -f $DestinationFolder)

    $existingExe = Get-ExistingHPIAExecutable -PreferredFolder $DestinationFolder
    if ($existingExe) {
        try {
            $existingVersion = (Get-Item -LiteralPath $existingExe -ErrorAction Stop).VersionInfo.FileVersion
            Write-Log ("HP Image Assistant is already installed/found at: {0} (Version: {1})" -f $existingExe, $existingVersion) 'OK'
            Add-YamlAction ("Skipped HPIA local deployment because HPImageAssistant.exe already exists: {0} (Version: {1})" -f $existingExe, $existingVersion)
        }
        catch {
            Write-Log ("HP Image Assistant is already installed/found at: {0}" -f $existingExe) 'OK'
            Add-YamlAction ("Skipped HPIA local deployment because HPImageAssistant.exe already exists: {0}" -f $existingExe)
        }

        return $existingExe
    }

    Write-Log 'HP Image Assistant was not found locally. Deploying from extracted source folder.' 'INFO'
    Add-YamlAction 'HP Image Assistant was not found locally. Deploying from extracted source folder.'

    if (-not (Test-Path -LiteralPath $SourceFolder)) {
        throw "HPIA source folder not found: $SourceFolder"
    }

    $sourceFiles = @(Get-ChildItem -Path $SourceFolder -Recurse -File -ErrorAction SilentlyContinue)
    Write-Log ("Source HPIA file count: {0}" -f $sourceFiles.Count) 'INFO'
    Add-YamlAction ("Source HPIA file count: {0}" -f $sourceFiles.Count)

    $sourceExe = Get-ChildItem -Path $SourceFolder -Recurse -Filter 'HPImageAssistant.exe' -File -ErrorAction SilentlyContinue |
        Sort-Object FullName |
        Select-Object -First 1

    if (-not $sourceExe) {
        throw "HPImageAssistant.exe was not found anywhere under source folder: $SourceFolder"
    }

    Write-Log ("Found source HPImageAssistant.exe: {0}" -f $sourceExe.FullName) 'OK'
    Add-YamlAction ("Found source HPImageAssistant.exe: {0}" -f $sourceExe.FullName)

    try {
        if (Test-Path -LiteralPath $DestinationFolder) {
            Write-Log ("Removing existing local HPIA folder before clean deployment: {0}" -f $DestinationFolder) 'INFO'
            Remove-Item -LiteralPath $DestinationFolder -Recurse -Force -ErrorAction Stop
        }

        New-Item -Path $DestinationFolder -ItemType Directory -Force | Out-Null

        Write-Log 'Copying extracted HPIA files locally with robocopy...' 'INFO'

        $roboLog = Join-Path $DestinationFolder 'HPIA_robocopy.log'
        $roboArgs = @(
            ('"{0}"' -f $SourceFolder),
            ('"{0}"' -f $DestinationFolder),
            '/E',
            '/COPY:DAT',
            '/R:3',
            '/W:5',
            '/NFL',
            '/NDL',
            '/NP',
            ('/LOG:"{0}"' -f $roboLog)
        )

        $robo = Start-Process -FilePath "$env:SystemRoot\System32\robocopy.exe" -ArgumentList ($roboArgs -join ' ') -Wait -PassThru -NoNewWindow

        # Robocopy exit codes 0-7 are success/non-fatal. 8+ indicates failure.
        if ($robo.ExitCode -ge 8) {
            throw "Robocopy failed copying HPIA files. Exit code: $($robo.ExitCode). Log: $roboLog"
        }

        Write-Log ("HPIA files copied locally with robocopy exit code {0}." -f $robo.ExitCode) 'OK'
        Add-YamlAction ("HPIA files copied locally with robocopy exit code {0}." -f $robo.ExitCode)

        try {
            Get-ChildItem -Path $DestinationFolder -Recurse -File -ErrorAction SilentlyContinue |
                ForEach-Object { Unblock-File -Path $_.FullName -ErrorAction SilentlyContinue }
        }
        catch {}

        $copiedFiles = @(Get-ChildItem -Path $DestinationFolder -Recurse -File -ErrorAction SilentlyContinue)
        Write-Log ("Local HPIA folder file count after copy: {0}" -f $copiedFiles.Count) 'INFO'
        Add-YamlAction ("Local HPIA folder file count after copy: {0}" -f $copiedFiles.Count)

        $localExeItem = Get-ChildItem -Path $DestinationFolder -Recurse -Filter 'HPImageAssistant.exe' -File -ErrorAction SilentlyContinue |
            Sort-Object FullName |
            Select-Object -First 1

        if (-not $localExeItem) {
            $sampleFiles = $copiedFiles | Select-Object -First 20 | ForEach-Object { $_.FullName }
            foreach ($sample in $sampleFiles) {
                Write-Log ("Local HPIA sample file: {0}" -f $sample) 'WARN'
            }

            throw "HPImageAssistant.exe was not found anywhere under local copy folder: $DestinationFolder"
        }

        $localExe = $localExeItem.FullName

        Write-Log ("Resolved local HPImageAssistant.exe location: {0}" -f $localExe) 'OK'
        Add-YamlAction ("Resolved local HPImageAssistant.exe location: {0}" -f $localExe)

        return $localExe
    }
    catch {
        throw "Failed to deploy HP Image Assistant locally: $($_.Exception.Message)"
    }
}

function Get-HpiaExitStatus {
    param([int]$ExitCode)

    switch ($ExitCode) {
        0    { return 'success' }
        1    { return 'failed' }
        2    { return 'cancelled' }
        3    { return 'needs_reboot' }
        256  { return 'no_recommendations_or_success' }
        257  { return 'recommendations_found' }
        3010 { return 'needs_reboot' }
        3011 { return 'not_auto_installable_skipped' }
        4096 { return 'no_applicable_updates_or_platform_not_supported' }
        4097 { return 'invalid_parameters' }
        8193 { return 'hpia_analysis_or_report_generation_error' }
        default { return 'unknown' }
    }
}

function Get-HpiaRecommendationObjects {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ReportFolder
    )

    # Use a flexible PowerShell array instead of a strongly typed generic list.
    # HPIA reports can contain mixed object types from JSON and XML parsing.
    $recommendations = @()

    $jsonFiles = @(Get-ChildItem -Path $ReportFolder -Filter '*.json' -Recurse -File -ErrorAction SilentlyContinue)
    foreach ($jsonFile in $jsonFiles) {
        try {
            $json = Get-Content -LiteralPath $jsonFile.FullName -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop

            if ($json.HPIA -and $json.HPIA.Recommendations) {
                foreach ($rec in @($json.HPIA.Recommendations)) {
                    $recommendations += $rec
                    try {
                        Write-Log ("HPIA JSON recommendation type: {0}" -f $rec.GetType().FullName) 'INFO'
                    }
                    catch {}
                }
            }
            elseif ($json.Recommendations) {
                foreach ($rec in @($json.Recommendations)) {
                    $recommendations += $rec
                    try {
                        Write-Log ("HPIA JSON recommendation type: {0}" -f $rec.GetType().FullName) 'INFO'
                    }
                    catch {}
                }
            }
        }
        catch {
            Write-Log ("Unable to parse HPIA JSON report {0}: {1}" -f $jsonFile.FullName, $_.Exception.Message) 'WARN'
        }
    }

    $xmlFiles = @(Get-ChildItem -Path $ReportFolder -Filter '*.xml' -Recurse -File -ErrorAction SilentlyContinue)
    foreach ($xmlFile in $xmlFiles) {
        try {
            [xml]$xml = Get-Content -LiteralPath $xmlFile.FullName -Raw -ErrorAction Stop

            $nodes = @($xml.SelectNodes('//*[contains(translate(local-name(), "ABCDEFGHIJKLMNOPQRSTUVWXYZ", "abcdefghijklmnopqrstuvwxyz"), "recommend")]'))
            foreach ($node in $nodes) {
                $recommendations += $node
                try {
                    Write-Log ("HPIA XML recommendation node type: {0}" -f $node.GetType().FullName) 'INFO'
                }
                catch {}
            }
        }
        catch {
            Write-Log ("Unable to parse HPIA XML report {0}: {1}" -f $xmlFile.FullName, $_.Exception.Message) 'WARN'
        }
    }

    return @($recommendations)
}

function Get-HpiaRecommendationValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Recommendation,
        [Parameter(Mandatory)][string[]]$PropertyNames
    )

    foreach ($prop in $PropertyNames) {
        try {
            if ($Recommendation.PSObject.Properties.Name -contains $prop) {
                $value = $Recommendation.$prop
                if ($null -ne $value -and -not [string]::IsNullOrWhiteSpace($value.ToString())) {
                    return $value.ToString()
                }
            }
        }
        catch {}
    }

    # XML fallback
    try {
        foreach ($prop in $PropertyNames) {
            $node = $Recommendation.SelectSingleNode('.//*[local-name()="' + $prop + '"]')
            if ($node -and -not [string]::IsNullOrWhiteSpace($node.InnerText)) {
                return $node.InnerText.Trim()
            }
        }
    }
    catch {}

    return $null
}

function Get-HpiaSoftPaqNumber {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Recommendation)

    $candidate = Get-HpiaRecommendationValue -Recommendation $Recommendation -PropertyNames @(
        'SoftPaqId','SoftpaqId','SoftPaq','Softpaq','SoftPaqNumber','SoftpaqNumber','SP','Id','ID','Number'
    )

    if ($candidate -match '(?i)sp?(\d{5,6})') {
        return $matches[1]
    }

    $text = ($Recommendation | Out-String)
    if ($text -match '(?i)sp(\d{5,6})') {
        return $matches[1]
    }

    if ($text -match '\b(\d{5,6})\b') {
        return $matches[1]
    }

    return $null
}

function Get-HpiaRecommendationCategory {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Recommendation)

    return (Get-HpiaRecommendationValue -Recommendation $Recommendation -PropertyNames @(
        'Category','Type','RecommendationType','ComponentType','Class','Group'
    ))
}

function Get-HpiaRecommendationName {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Recommendation)

    $name = Get-HpiaRecommendationValue -Recommendation $Recommendation -PropertyNames @(
        'Name','Title','Component','ComponentName','Description','SoftPaqName','SoftpaqName'
    )

    if ($name) { return $name }

    $text = ($Recommendation | Out-String).Trim()
    if ($text.Length -gt 160) {
        return $text.Substring(0,160)
    }

    return $text
}


function Test-HpiaBlockedUnattendedDriver {
    [CmdletBinding()]
    param(
        [AllowNull()][string]$Category,
        [AllowNull()][string]$Name
    )

    $combined = ("{0} {1}" -f $Category, $Name)

    # These driver families can affect boot-critical storage/controller behavior.
    # They are intentionally blocked from unattended HPIA installs.
    $blockedPattern = '(?i)\b(storage|chipset|rst|rapid storage|intel rapid|vmd|volume management device|nvme|nvm express|sata|raid|ahci|scsi|controller|management engine|mei|thunderbolt)\b'

    return ($combined -match $blockedPattern)
}

function New-HpiaDriverSPList {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ReportFolder,
        [Parameter(Mandatory)][string]$SPListPath
    )

    $recommendations = @(Get-HpiaRecommendationObjects -ReportFolder $ReportFolder)
    Write-Log ("HPIA recommendations parsed from reports: {0}" -f $recommendations.Count) 'INFO'
    Add-YamlAction ("HPIA recommendations parsed from reports: {0}" -f $recommendations.Count)

    $selected = @()
    $seen = @{}

    foreach ($rec in $recommendations) {
        $category = Get-HpiaRecommendationCategory -Recommendation $rec
        $name = Get-HpiaRecommendationName -Recommendation $rec
        $sp = Get-HpiaSoftPaqNumber -Recommendation $rec

        if (-not $sp) {
            continue
        }

        # Exclude BIOS/Firmware explicitly.
        $combined = ("{0} {1}" -f $category, $name)
        if ($combined -match '(?i)\bBIOS\b|Firmware') {
            $script:SkippedList.Add(("SP{0} {1}" -f $sp, $name)) | Out-Null
            Add-DriverResult -Vendor 'HP' -Name $name -Id $sp -Category $category -Status 'Blocked' -Message 'Excluded because it appears to be BIOS/Firmware.'
            continue
        }

        # Block unattended installation of boot-critical driver families.
        if (Test-HpiaBlockedUnattendedDriver -Category $category -Name $name) {
            $script:SkippedList.Add(("SP{0} {1}" -f $sp, $name)) | Out-Null
            Add-DriverResult -Vendor 'HP' -Name $name -Id $sp -Category $category -Status 'Blocked' -Message 'Excluded from unattended install because it appears to be storage/chipset/RST/VMD/NVMe/controller related.'
            continue
        }

        # Prefer safer driver-like recommendations, but avoid storage/chipset/controller families above.
        if ($combined -notmatch '(?i)Driver|Bluetooth|Audio|Graphics|Video|LAN|WLAN|Wireless|NIC|Touch|Fingerprint|Card Reader|Serial|USB' -and $category) {
            $script:SkippedList.Add(("SP{0} {1}" -f $sp, $name)) | Out-Null
            Add-DriverResult -Vendor 'HP' -Name $name -Id $sp -Category $category -Status 'Skipped' -Message 'Excluded because it did not appear to be a safe unattended driver recommendation.'
            continue
        }

        if (-not $seen.ContainsKey($sp)) {
            $seen[$sp] = $true
            $selected += $sp
            Add-DriverResult -Vendor 'HP' -Name $name -Id $sp -Category $category -Status 'Detected' -Message 'Selected for HPIA SPList install.'
        }
    }

    if (@($selected).Count -gt 0) {
        Set-Content -LiteralPath $SPListPath -Value $selected -Encoding ASCII
        Write-Log ("Created filtered HPIA SPList with {0} SoftPaqs: {1}" -f @($selected).Count, $SPListPath) 'OK'
        Add-YamlAction ("Created filtered HPIA SPList with {0} SoftPaqs: {1}" -f @($selected).Count, $SPListPath)
    }
    else {
        Write-Log 'No non-BIOS/Firmware driver SoftPaq recommendations were selected from HPIA reports.' 'OK'
        Add-YamlAction 'No non-BIOS/Firmware driver SoftPaq recommendations were selected from HPIA reports.'
    }

    return @($selected)
}



# -----------------------------
# HP PowerShell module maintenance
# -----------------------------
function Initialize-NetworkDefaults {
    try {
        [Net.ServicePointManager]::SecurityProtocol =
            [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls13
    }
    catch {
        try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 }
        catch {}
    }

    try { [Net.ServicePointManager]::DefaultConnectionLimit = 64 }
    catch {}
}

function Get-HighestInstalledModuleVersion {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Name
    )

    try {
        $installed = Get-Module -Name $Name -ListAvailable -ErrorAction SilentlyContinue |
            Sort-Object Version -Descending |
            Select-Object -First 1

        if ($installed) { return [version]$installed.Version }
    }
    catch {
        Write-Log "Could not determine installed version for module ${Name}: $($_.Exception.Message)" 'WARN'
    }

    return $null
}

function Invoke-HPCMSLUpdateInFreshPowerShell {
    [CmdletBinding()]
    param(
        [ValidateSet('CurrentUser','AllUsers')]
        [string]$Scope = 'AllUsers'
    )

    Write-Log 'Starting isolated HPCMSL install/update using PowerShellGet/PSResourceGet when available...' 'INFO'
    Add-YamlAction 'Starting isolated HPCMSL install/update before HP driver updates.'

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

    if ([string]::IsNullOrWhiteSpace(`$InstallRoot)) { throw 'InstallRoot is blank.' }
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

    if (`$installed) { HelperLog ('Installed HPCMSL version: {0}' -f `$installed) }
    else { HelperLog 'HPCMSL is not currently installed.' }
    HelperLog ('Latest PSGallery HPCMSL version: {0}' -f `$online)

    if (`$installed -and `$installed -ge `$online) {
        HelperLog 'HPCMSL is already current. No update needed.'
        exit 0
    }

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
        if ('$Scope' -eq 'AllUsers') { `$installRoot = Join-Path `$env:ProgramFiles 'WindowsPowerShell\Modules' }
        else { `$installRoot = Join-Path `$HOME 'Documents\WindowsPowerShell\Modules' }
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
        Initialize-NetworkDefaults
        Set-Content -Path $helperPath -Value $helperScript -Encoding UTF8 -Force
        Write-Log "HPCMSL helper script created at $helperPath" 'INFO'

        $args = @('-NoProfile','-ExecutionPolicy','Bypass','-File',('"{0}"' -f $helperPath))
        $process = Start-Process -FilePath powershell.exe -ArgumentList ($args -join ' ') -Wait -PassThru -WindowStyle Hidden

        if (Test-Path -LiteralPath $helperLog) {
            Get-Content -LiteralPath $helperLog -ErrorAction SilentlyContinue | ForEach-Object {
                if (-not [string]::IsNullOrWhiteSpace($_)) { Write-Log "HPCMSL helper: $_" 'INFO' }
            }
        }

        if ($process.ExitCode -ne 0) {
            Write-Log "Isolated HPCMSL helper failed with exit code $($process.ExitCode). Continuing with HPIA driver updates." 'WARN'
            Add-YamlAction ("Isolated HPCMSL helper failed with exit code {0}. Continuing with HPIA driver updates." -f $process.ExitCode)
            return $false
        }

        $finalVersion = Get-HighestInstalledModuleVersion -Name 'HPCMSL'
        if ($finalVersion) {
            Write-Log "HPCMSL final installed version after isolated update: $finalVersion" 'OK'
            Add-YamlAction ("HPCMSL final installed version after isolated update: {0}" -f $finalVersion)
        }
        else {
            Write-Log 'HPCMSL helper completed, but this session could not verify the installed HPCMSL version.' 'WARN'
            Add-YamlAction 'HPCMSL helper completed, but installed version could not be verified.'
            return $false
        }

        return $true
    }
    catch {
        Write-Log "Failed to run isolated HPCMSL update helper: $($_.Exception.Message)" 'WARN'
        Add-YamlAction ("Failed to run isolated HPCMSL update helper: {0}" -f $_.Exception.Message)
        return $false
    }
    finally {
        Remove-Item -LiteralPath $helperPath -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $helperLog -Force -ErrorAction SilentlyContinue
    }
}

function Invoke-HPPowerShellModuleMaintenance {
    [CmdletBinding()]
    param()

    Write-Section 'HP PowerShell Module Maintenance'
    Write-Log 'HP vendor detected. Checking/installing HPCMSL before HP driver updates...' 'INFO'
    Add-YamlAction 'HP vendor detected; checking/installing HPCMSL before HP driver updates.'

    $hpcmslUpdated = Invoke-HPCMSLUpdateInFreshPowerShell -Scope AllUsers
    if ($hpcmslUpdated) {
        Write-Log 'HP PowerShell module maintenance completed.' 'OK'
        Add-YamlAction 'HP PowerShell module maintenance completed.'
    }
    else {
        Write-Log 'HP PowerShell module maintenance did not complete successfully. Continuing with HP Image Assistant driver updates.' 'WARN'
        Add-YamlAction 'HP PowerShell module maintenance did not complete successfully; continuing with HPIA driver updates.'
    }
}

function Invoke-HPDriverUpdates {
    Write-Section 'HP Driver Analysis and Filtered Installation'

    $hpInfo = Get-HPSystemModelInfo

    $hpiaExe = Install-HPIAFromExtractedFolder -SourceFolder $HpiaSourceFolder -DestinationFolder $LocalHpiaFolder

    $hpiaReportFolder = Join-Path $YamlLogFolder 'HPIA'
    $hpiaDownloadFolder = Join-Path $WorkingRoot 'HPIADownloads'
    $hpiaExtractFolder = Join-Path $WorkingRoot 'HPIAExtracted'
    $hpiaSPListPath = Join-Path $WorkingRoot 'HPIA-Filtered-Safe-Drivers.splist'

    Ensure-Folder -Path $hpiaReportFolder
    Ensure-Folder -Path $hpiaDownloadFolder
    Ensure-Folder -Path $hpiaExtractFolder
    Ensure-Folder -Path 'C:\Temp'

    # HPIA is more reliable under Task Scheduler/SYSTEM when TEMP/TMP are local and TLS 1.2 is forced.
    $env:TEMP = 'C:\Temp'
    $env:TMP  = 'C:\Temp'

    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Write-Log 'TLS 1.2 enabled for HP Image Assistant network operations.' 'INFO'
        Add-YamlAction 'TLS 1.2 enabled for HPIA network operations.'
    }
    catch {
        Write-Log ("Unable to force TLS 1.2 before HPIA run: {0}" -f $_.Exception.Message) 'WARN'
        Add-YamlAction ("Unable to force TLS 1.2 before HPIA run: {0}" -f $_.Exception.Message)
    }

    # Safer unattended HPIA workflow:
    # 1. Run /Action:List first to generate recommendations only.
    # 2. Parse the HPIA report and create a filtered SPList.
    # 3. Install only the filtered SPList.
    #
    # Blocked unattended families include BIOS, Firmware, Storage, Chipset,
    # Intel RST, Intel VMD, NVMe, RAID/AHCI/SATA/controller, MEI/Management Engine,
    # and Thunderbolt because those can affect boot or storage-controller behavior.
    Write-Log 'Running HP Image Assistant recommendation/list pass only. No drivers will be installed in this step.' 'INFO'
    Add-YamlAction 'Running HPIA /Action:List first so recommendations can be filtered before install.'

    $hpiaListArgs = @(
        '/Operation:Analyze',
        '/Action:List',
        '/Category:Drivers',
        '/Selection:All',
        '/Silent',
        '/Noninteractive',
        '/Debug',
        "/ReportFolder:`"$hpiaReportFolder`"",
        "/SoftpaqDownloadFolder:`"$hpiaDownloadFolder`"",
        "/SoftpaqExtractFolder:`"$hpiaExtractFolder`""
    )

    Write-Log ("HPIA list command: {0} {1}" -f $hpiaExe, ($hpiaListArgs -join ' ')) 'INFO'
    Add-YamlAction ("HPIA list command: {0} {1}" -f $hpiaExe, ($hpiaListArgs -join ' '))

    Write-Progress -Activity 'HP Image Assistant' -Status ("Listing recommended drivers for {0}" -f $hpInfo.Model) -PercentComplete 25
    $hpiaListProc = Start-Process -FilePath $hpiaExe -ArgumentList ($hpiaListArgs -join ' ') -Wait -PassThru -NoNewWindow
    $listExitCode = [int]$hpiaListProc.ExitCode
    $listStatus = Get-HpiaExitStatus -ExitCode $listExitCode
    Write-Progress -Activity 'HP Image Assistant' -Completed

    Write-Log ("HP Image Assistant list pass completed with exit code {0} ({1})." -f $listExitCode, $listStatus) 'INFO'
    Add-YamlAction ("HPIA list pass completed with exit code {0} ({1})." -f $listExitCode, $listStatus)

    $reportFiles = @(Get-ChildItem -Path $hpiaReportFolder -File -Recurse -ErrorAction SilentlyContinue)
    foreach ($report in $reportFiles) {
        Add-YamlAction ("HPIA report generated: {0}" -f $report.FullName)
    }

    if ($reportFiles.Count -eq 0 -and $listExitCode -notin @(0,256,257,3010,3011,4096)) {
        throw "HPIA list pass failed or generated no reports. Exit code: $listExitCode ($listStatus). Review $hpiaReportFolder."
    }

    $selectedSoftPaqs = @(New-HpiaDriverSPList -ReportFolder $hpiaReportFolder -SPListPath $hpiaSPListPath)

    if ($selectedSoftPaqs.Count -eq 0) {
        Write-Log 'No safe HP driver SoftPaqs were selected for unattended installation after filtering.' 'OK'
        Add-YamlAction 'No safe HP driver SoftPaqs were selected for unattended installation after filtering.'
        return
    }

    Write-Log ("Installing filtered safe HP driver SPList. Count: {0}" -f $selectedSoftPaqs.Count) 'INFO'
    Add-YamlAction ("Installing filtered safe HP driver SPList: {0}" -f $hpiaSPListPath)

    $hpiaInstallArgs = @(
        "/SPList:`"$hpiaSPListPath`"",
        '/Silent',
        '/Noninteractive',
        '/AutoCleanup',
        '/Debug',
        "/ReportFolder:`"$hpiaReportFolder`"",
        "/SoftpaqDownloadFolder:`"$hpiaDownloadFolder`"",
        "/SoftpaqExtractFolder:`"$hpiaExtractFolder`""
    )

    Write-Log ("HPIA filtered SPList install command: {0} {1}" -f $hpiaExe, ($hpiaInstallArgs -join ' ')) 'INFO'
    Add-YamlAction ("HPIA filtered SPList install command: {0} {1}" -f $hpiaExe, ($hpiaInstallArgs -join ' '))

    Write-Progress -Activity 'HP Image Assistant' -Status 'Installing filtered safe driver SoftPaqs' -PercentComplete 75
    $hpiaInstallProc = Start-Process -FilePath $hpiaExe -ArgumentList ($hpiaInstallArgs -join ' ') -Wait -PassThru -NoNewWindow
    $installExitCode = [int]$hpiaInstallProc.ExitCode
    $installStatus = Get-HpiaExitStatus -ExitCode $installExitCode
    Write-Progress -Activity 'HP Image Assistant' -Completed

    Write-Log ("HP Image Assistant filtered SPList install completed with exit code {0} ({1})." -f $installExitCode, $installStatus) 'INFO'
    Add-YamlAction ("HPIA filtered SPList install completed with exit code {0} ({1})." -f $installExitCode, $installStatus)

    $downloadedFiles = @(Get-ChildItem -Path $hpiaDownloadFolder -File -Recurse -ErrorAction SilentlyContinue)
    foreach ($file in $downloadedFiles) {
        Add-DriverResult -Vendor 'HP' -Name $file.Name -Id $null -Category 'Driver' -Status 'Detected' -Message ("Downloaded/processed by filtered HPIA SPList: {0}" -f $file.FullName)
    }

    switch ($installExitCode) {
        0 {
            Write-Log 'HPIA filtered SPList install completed successfully.' 'OK'
            Add-YamlAction 'HPIA filtered SPList install completed successfully.'
            return
        }
        256 {
            Write-Log 'HPIA completed successfully. No applicable filtered driver recommendations were found or no action was required.' 'OK'
            Add-YamlAction 'HPIA completed successfully with no applicable filtered driver recommendations or no action required.'
            return
        }
        257 {
            Write-Log 'HPIA completed and reported filtered driver actions.' 'OK'
            Add-YamlAction 'HPIA completed and reported filtered driver actions.'
            return
        }
        3010 {
            Write-Log 'HPIA filtered SPList install completed successfully. Reboot required.' 'WARN'
            Add-YamlAction 'HPIA filtered SPList install completed successfully and indicated reboot required.'
            return
        }
        3011 {
            Write-Log 'One or more filtered HPIA items were not auto-installable and were skipped.' 'WARN'
            Add-YamlAction 'One or more filtered HPIA items were not auto-installable and were skipped.'
            return
        }
        4096 {
            Write-Log 'HPIA completed but did not find applicable filtered driver updates for this platform.' 'OK'
            Add-YamlAction 'HPIA completed but did not find applicable filtered driver updates for this platform.'
            return
        }
        8193 {
            Write-Log 'HPIA filtered install returned 8193. Checking whether reports/logs were generated before failing the run.' 'WARN'
            Add-YamlAction 'HPIA filtered install returned 8193; checking for reports/logs before hard failure.'
            $generatedReports = @(Get-ChildItem -Path $hpiaReportFolder -File -Recurse -ErrorAction SilentlyContinue)
            if ($generatedReports.Count -gt 0) {
                Write-Log ("HPIA generated {0} report/log file(s) despite exit code 8193. Continuing so the scheduled workflow does not hard fail." -f $generatedReports.Count) 'WARN'
                Add-YamlAction ("HPIA generated {0} report/log file(s) despite exit code 8193. Continuing without hard failure." -f $generatedReports.Count)
                return
            }
            throw "HPIA failed with exit code 8193 and produced no report/log files in $hpiaReportFolder."
        }
        default {
            throw "HP Image Assistant filtered SPList install failed or returned an unexpected exit code: $installExitCode ($installStatus). Review reports in $hpiaReportFolder."
        }
    }
}

# -----------------------------
# Dell support
# -----------------------------
function Wait-ForFile {
    param(
        [Parameter(Mandatory)][string]$Path,
        [int]$TimeoutSeconds = 30
    )

    $start = Get-Date
    while (-not (Test-Path -LiteralPath $Path)) {
        Start-Sleep -Seconds 1
        if (((Get-Date) - $start).TotalSeconds -ge $TimeoutSeconds) {
            return $false
        }
    }
    return $true
}

function Get-DcuReportXmlSafely {
    param(
        [Parameter(Mandatory)][string]$Path,
        [int]$RetryCount = 5,
        [int]$RetryDelaySeconds = 2
    )

    if (-not (Wait-ForFile -Path $Path -TimeoutSeconds 20)) {
        throw "Dell DCU report file was not found: $Path"
    }

    Start-Sleep -Seconds 3

    for ($i = 1; $i -le $RetryCount; $i++) {
        try {
            $fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
            try {
                $sr = New-Object System.IO.StreamReader($fs)
                try {
                    $content = $sr.ReadToEnd()
                }
                finally {
                    $sr.Dispose()
                }
            }
            finally {
                $fs.Dispose()
            }

            $xml = New-Object System.Xml.XmlDocument
            $xml.LoadXml($content)
            return $xml
        }
        catch {
            if ($i -eq $RetryCount) {
                throw
            }
            Start-Sleep -Seconds $RetryDelaySeconds
        }
    }
}

function Get-DellNodeText {
    param(
        [Parameter(Mandatory)][xml]$Node,
        [Parameter(Mandatory)][string[]]$Names
    )

    foreach ($name in $Names) {
        try {
            $xpath = './/*[local-name()="' + $name + '"]'
            $child = $Node.SelectSingleNode($xpath)
            if ($child -and -not [string]::IsNullOrWhiteSpace($child.InnerText)) {
                return $child.InnerText.Trim()
            }
        }
        catch {}
    }

    return $null
}

function Get-DellReportItems {
    param([Parameter(Mandatory)][xml]$Xml)

    $items = @()
    try {
        $xpath = '//*[local-name()="Update" or local-name()="Package" or local-name()="SoftwareComponent" or local-name()="component" or local-name()="Device"]'
        $nodes = $Xml.SelectNodes($xpath)
        foreach ($node in $nodes) {
            $name = Get-DellNodeText -Node $node -Names @('Name','Title','PackageName')
            $version = Get-DellNodeText -Node $node -Names @('Version','PackageVersion')
            $category = Get-DellNodeText -Node $node -Names @('Category','Type')
            $id = Get-DellNodeText -Node $node -Names @('Id','PackageId','ReleaseId')

            if ($name -or $id) {
                $items += [pscustomobject]@{
                    Id       = $id
                    Name     = $name
                    Version  = $version
                    Category = $category
                }
            }
        }
    }
    catch {}

    return @($items)
}

function Get-DellDCUService {
    $candidateNames = @(
        'DellClientManagementService',
        'DellCommandUpdate',
        'DellUpdateService'
    )

    foreach ($name in $candidateNames) {
        $svc = Get-Service -Name $name -ErrorAction SilentlyContinue
        if ($svc) { return $svc }
    }

    $svc = Get-Service -ErrorAction SilentlyContinue |
        Where-Object { $_.DisplayName -match 'Dell.*Client.*Management|Dell.*Command.*Update' } |
        Select-Object -First 1

    return $svc
}

function Ensure-DellDCUService {
    param(
        [int]$TimeoutSeconds = 30
    )

    Write-Log 'Validating Dell Client Management Service...' 'INFO'
    Add-YamlAction 'Validating Dell Client Management Service.'

    $service = Get-DellDCUService
    if (-not $service) {
        throw 'Dell Client Management Service was not found. Dell Command | Update may need to be repaired or reinstalled.'
    }

    Write-Log ("Dell service detected: {0} ({1})" -f $service.DisplayName, $service.Name) 'OK'

    try {
        $wmiService = Get-CimInstance -ClassName Win32_Service -Filter ("Name='{0}'" -f $service.Name) -ErrorAction Stop
        if ($wmiService.StartMode -eq 'Disabled') {
            Write-Log 'Dell service startup type is Disabled. Setting it to Manual.' 'WARN'
            Set-Service -Name $service.Name -StartupType Manual -ErrorAction Stop
        }
    }
    catch {
        Write-Log ("Could not validate/set Dell service startup type: {0}" -f $_.Exception.Message) 'WARN'
    }

    $service.Refresh()
    if ($service.Status -ne 'Running') {
        Write-Log 'Starting Dell Client Management Service...' 'INFO'
        try {
            Start-Service -Name $service.Name -ErrorAction Stop
        }
        catch {
            Write-Log ("Start-Service failed. Attempting sc.exe recovery start. Error: {0}" -f $_.Exception.Message) 'WARN'
            & sc.exe start $service.Name | Out-Null
        }
    }

    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    do {
        Start-Sleep -Seconds 2
        $service = Get-Service -Name $service.Name -ErrorAction SilentlyContinue
        if ($service -and $service.Status -eq 'Running') {
            Write-Log 'Dell Client Management Service is running.' 'OK'
            Add-YamlAction 'Dell Client Management Service is running.'
            return $true
        }
    } while ((Get-Date) -lt $deadline)

    throw 'Dell Client Management Service did not reach the Running state before timeout.'
}

function Invoke-DellDCUCommandWithRetry {
    param(
        [Parameter(Mandatory)][string]$DcuCli,
        [Parameter(Mandatory)][string]$Arguments,
        [Parameter(Mandatory)][string]$OperationName,
        [int[]]$AcceptableExitCodes = @(0),
        [int]$MaxAttempts = 2
    )

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        Write-Log ("Dell DCU {0} attempt {1} of {2}..." -f $OperationName, $attempt, $MaxAttempts) 'INFO'
        Add-YamlAction ("Dell DCU {0} attempt {1} of {2}." -f $OperationName, $attempt, $MaxAttempts)

        $proc = Start-Process -FilePath $DcuCli -ArgumentList $Arguments -Wait -PassThru -NoNewWindow
        $exitCode = [int]$proc.ExitCode
        Write-Log ("Dell DCU {0} exit code: {1}" -f $OperationName, $exitCode) 'INFO'

        if ($AcceptableExitCodes -contains $exitCode) {
            return $proc
        }

        if ($exitCode -eq 3000) {
            Write-Log 'Dell DCU returned 3000, which normally indicates the Dell Client Management Service stopped or crashed.' 'WARN'
            Add-YamlAction 'Dell DCU returned 3000; attempting Dell service recovery before retry.'
        }
        else {
            Write-Log ("Dell DCU {0} returned non-success exit code {1}." -f $OperationName, $exitCode) 'WARN'
        }

        if ($attempt -lt $MaxAttempts) {
            Ensure-DellDCUService | Out-Null
            Start-Sleep -Seconds 5
            continue
        }

        throw "Dell DCU $OperationName failed after $MaxAttempts attempt(s). Last exit code: $exitCode"
    }
}

function Invoke-DellDriverUpdates {
    Write-Section 'Dell Command Update Workflow'

    $dcuCli = Join-Path ${env:ProgramFiles} 'Dell\CommandUpdate\dcu-cli.exe'
    if (-not (Test-Path -LiteralPath $dcuCli)) {
        throw "Dell Command | Update CLI was not found: $dcuCli"
    }

    Write-Log ("Using Dell Command | Update CLI: {0}" -f $dcuCli) 'OK'
    Add-YamlAction 'Using Dell Command | Update CLI.'

    Ensure-DellDCUService | Out-Null

    $dcuScanLog  = Join-Path $WorkingRoot 'Dell-DCU-Scan.log'
    $dcuApplyLog = Join-Path $WorkingRoot 'Dell-DCU-Apply.log'
    $dcuReport   = Join-Path $WorkingRoot 'Dell-DCU-ApplicableUpdates.xml'

    Write-Log 'Dell DCU Configure...' 'INFO'
    Write-Progress -Activity 'Dell Driver Update Workflow' -Status 'Configuring Dell Command Update' -PercentComplete 10
    $configureArgs = "/configure -silent -scheduleAuto -lockSettings=disable"
    Invoke-DellDCUCommandWithRetry -DcuCli $dcuCli -Arguments $configureArgs -OperationName 'Configure' -MaxAttempts 2 | Out-Null

    Write-Log 'Dell DCU Scan...' 'INFO'
    Write-Progress -Activity 'Dell Driver Update Workflow' -Status 'Scanning for updates' -PercentComplete 35
    $scanArgs = "/scan -silent -updateType=driver,application -outputLog=""$dcuScanLog"" -report=""$dcuReport"""
    Invoke-DellDCUCommandWithRetry -DcuCli $dcuCli -Arguments $scanArgs -OperationName 'Scan' -MaxAttempts 2 | Out-Null
    Write-Log 'Waiting for Dell Command | Update to finish writing the report...' 'INFO'

    try {
        $xml = Get-DcuReportXmlSafely -Path $dcuReport
        $items = Get-DellReportItems -Xml $xml

        if ($items.Count -gt 0) {
            Add-YamlAction ("Dell DCU report parsed successfully. Updates detected: {0}" -f $items.Count)

            $total = $items.Count
            $index = 0

            foreach ($item in $items) {
                $index++
                $percent = 35 + [math]::Floor(($index / $total) * 35)
                $label = if ($item.Name) { $item.Name } elseif ($item.Id) { $item.Id } else { 'Dell update' }

                Write-Progress -Activity 'Dell Driver Update Workflow' -Status ("Parsing report: {0}" -f $label) -PercentComplete $percent

                if ($item.Category -match 'BIOS|Firmware') {
                    $script:SkippedList.Add($label) | Out-Null
                    Add-DriverResult -Vendor 'Dell' -Name $item.Name -Id $item.Id -Category $item.Category -Status 'Blocked' -Message 'BIOS/Firmware update blocked by script policy.'
                    Write-Log ("Blocking Dell BIOS/Firmware update: {0}" -f $label) 'WARN'
                }
                else {
                    Add-DriverResult -Vendor 'Dell' -Name $item.Name -Id $item.Id -Category $item.Category -Status 'Detected' -Message 'Detected in Dell DCU report.'
                }
            }
        }
        else {
            Add-YamlAction 'Dell DCU report parsed but returned no identifiable updates.'
        }
    }
    catch {
        Write-Log ("Failed to parse Dell DCU report: {0}" -f $_.Exception.Message) 'WARN'
        Add-YamlAction ("Failed to parse Dell DCU report: {0}" -f $_.Exception.Message)
    }

    Write-Log 'Dell DCU ApplyUpdates for driver/application updates only. BIOS/Firmware are excluded by -updateType.' 'INFO'
    Write-Progress -Activity 'Dell Driver Update Workflow' -Status 'Applying updates' -PercentComplete 85
    $applyArgs = "/applyUpdates -silent -updateType=driver,application -reboot=disable -outputLog=""$dcuApplyLog"""
    Invoke-DellDCUCommandWithRetry -DcuCli $dcuCli -Arguments $applyArgs -OperationName 'ApplyUpdates' -MaxAttempts 2 | Out-Null

    Write-Log (("Dell DCU logs: {0} ; {1} ; {2}") -f $dcuScanLog, $dcuApplyLog, $dcuReport) 'OK'
    Write-Progress -Activity 'Dell Driver Update Workflow' -Completed
}
# -----------------------------
# Main
# -----------------------------
$finalStatus = 'success'

try {
    Write-Section 'Initialization'

    Ensure-Folder -Path $YamlLogFolder
    Ensure-Folder -Path $WorkingRoot
    Ensure-WorkingFolderPermissions -Path $WorkingRoot

    $yamlName = "{0}-{1}-{2}.yml" -f $script:ComputerName, '05_Weekend_Vendor_Drivers_Update', (Get-Date -Format 'yyyy-MM-dd_HHmmss')
    $yamlPath = Join-Path $YamlLogFolder $yamlName

    Write-Log ("YAML log will be written to: {0}" -f $yamlPath) 'INFO'
    Initialize-YamlLog -ComputerName $script:ComputerName -YamlPath $yamlPath

    Write-Log 'Initializing vendor driver update script...' 'INFO'


    Write-Section 'Vendor Detection'
    $script:DetectedVendor = Get-DriverVendor
    Write-Log ("Detected vendor workflow: {0}" -f $script:DetectedVendor) 'INFO'
    Write-Log ("Working root: {0}" -f $WorkingRoot) 'INFO'

    if ($script:DetectedVendor -eq 'HP') {
        Invoke-HPPowerShellModuleMaintenance
        Invoke-HPDriverUpdates
    }
    elseif ($script:DetectedVendor -eq 'Dell') {
        Invoke-DellDriverUpdates
    }

}
catch {
    $finalStatus = 'failed'
    Add-RunFailure ("Script failed: {0}" -f $_.Exception.Message)
}
finally {
    Write-Section 'Cleanup'
    Remove-WorkingFolderRobust -Path $WorkingRoot

    if ($script:RunFailures.Count -gt 0 -and $finalStatus -ne 'failed') {
        $finalStatus = 'completed_with_warnings'
    }

    if ($script:RunFailures.Count -gt 0) {
        Write-Log (("{0} driver update completed with one or more failures.") -f $script:DetectedVendor) 'WARN'
    }
    else {
        Write-Log (("{0} driver update script completed successfully.") -f $script:DetectedVendor) 'OK'
    }

    Save-YamlLog -Status $finalStatus
}# ScriptName: 05_Weekend_HP_Drivers_Update.ps1
# ScriptVersion: 2.5.2
# LastUpdated: 2026-06-15
# Purpose: Weekend vendor driver update script with clean HP + Dell support,
#          HP-only HPCMSL maintenance after vendor detection,
#          YAML logging, colored output,
#          section headers, progress display, and structured per-driver results.

[CmdletBinding()]
param([string]$WorkingRoot = 'C:\Temp\DriverUpdates',
    [string]$YamlLogFolder = 'C:\Logs',
    [switch]$IncludeSoftware,
    [switch]$IncludeBIOS,
    [string]$HpiaSourceFolder = '\\filesvr\Labscripts\HPImageAssistant',
    [string]$LocalHpiaFolder = 'C:\ProgramData\Compton\HPImageAssistant'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# -----------------------------
# Script metadata
# -----------------------------
$script:ScriptName        = '05_Weekend_HP_Drivers_Update.ps1'
$script:ScriptVersion     = '2.5.2'
$script:StartTime         = Get-Date
$script:RunFailures       = New-Object System.Collections.Generic.List[string]
$script:InstalledList     = New-Object System.Collections.Generic.List[string]
$script:SkippedList       = New-Object System.Collections.Generic.List[string]
$script:YamlActionLines   = New-Object System.Collections.Generic.List[string]
$script:DriverResults     = New-Object System.Collections.Generic.List[object]
$script:DetectedVendor    = $null
$script:YamlPath          = $null
$script:ComputerName      = $env:COMPUTERNAME

# -----------------------------
# Logging helpers
# -----------------------------
function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','OK','WARN','ERROR')][string]$Level = 'INFO'
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[{0}] [{1}] {2}" -f $timestamp, $Level.PadRight(5), $Message

    switch ($Level) {
        'INFO'  { Write-Host $line -ForegroundColor Cyan }
        'OK'    { Write-Host $line -ForegroundColor Green }
        'WARN'  { Write-Host $line -ForegroundColor Yellow }
        'ERROR' { Write-Host $line -ForegroundColor Red }
        default { Write-Host $line }
    }
}

function Write-Section {
    param([Parameter(Mandatory)][string]$Title)

    $border = ('=' * 72)
    Write-Host ''
    Write-Host $border -ForegroundColor Magenta
    Write-Host ("  {0}" -f $Title) -ForegroundColor Magenta
    Write-Host $border -ForegroundColor Magenta
    Add-YamlAction ("Section: {0}" -f $Title)
}

function Add-RunFailure {
    param([Parameter(Mandatory)][string]$Message)
    $script:RunFailures.Add($Message) | Out-Null
    Write-Log $Message 'WARN'
}

function Add-DriverResult {
    param(
        [Parameter(Mandatory)][string]$Vendor,
        [Parameter(Mandatory)][string]$Name,
        [string]$Id,
        [string]$Category,
        [ValidateSet('Detected','Installed','Downloaded','Skipped','Failed','Blocked')][string]$Status,
        [string]$Message = ''
    )

    $script:DriverResults.Add([pscustomobject]@{
        Vendor   = $Vendor
        Name     = $Name
        Id       = $Id
        Category = $Category
        Status   = $Status
        Message  = $Message
    }) | Out-Null
}

function ConvertTo-YamlSafeString {
    param([AllowNull()][object]$Value)

    if ($null -eq $Value) { return 'null' }

    $text = [string]$Value
    $text = $text -replace "`r", ''
    $text = $text -replace "`n", ' '
    $text = $text -replace "'", "''"
    return "'$text'"
}

function Initialize-YamlLog {
    param(
        [Parameter(Mandatory)][string]$ComputerName,
        [Parameter(Mandatory)][string]$YamlPath
    )

    $script:YamlPath = $YamlPath
    $script:YamlActionLines.Clear()

    Add-YamlAction 'Script initialized.'
    Add-YamlAction ("Working root: {0}" -f $WorkingRoot)
    Add-YamlAction ("Computer: {0}" -f $ComputerName)
}

function Add-YamlAction {
    param([Parameter(Mandatory)][string]$Text)
    $script:YamlActionLines.Add(("    - {0}" -f (ConvertTo-YamlSafeString $Text))) | Out-Null
}

function Save-YamlLog {
    param(
        [Parameter(Mandatory)][string]$Status
    )

    if ([string]::IsNullOrWhiteSpace($script:YamlPath)) {
        return
    }

    $endTime = Get-Date
    $duration = [math]::Round(($endTime - $script:StartTime).TotalSeconds, 0)

    $lines = New-Object System.Collections.Generic.List[string]

    $lines.Add('script:') | Out-Null
    $lines.Add(("  name: {0}" -f (ConvertTo-YamlSafeString $script:ScriptName))) | Out-Null
    $lines.Add(("  version: {0}" -f (ConvertTo-YamlSafeString $script:ScriptVersion))) | Out-Null
    $lines.Add(("  computer: {0}" -f (ConvertTo-YamlSafeString $script:ComputerName))) | Out-Null
    $lines.Add(("  started: {0}" -f (ConvertTo-YamlSafeString ($script:StartTime.ToString('s'))))) | Out-Null
    $lines.Add(("  ended: {0}" -f (ConvertTo-YamlSafeString ($endTime.ToString('s'))))) | Out-Null
    $lines.Add(("  duration_seconds: {0}" -f $duration)) | Out-Null

    $lines.Add('run:') | Out-Null
    $lines.Add(("  status: {0}" -f (ConvertTo-YamlSafeString $Status))) | Out-Null
    $lines.Add(("  vendor: {0}" -f (ConvertTo-YamlSafeString $script:DetectedVendor))) | Out-Null

    $lines.Add('  actions:') | Out-Null
    if ($script:YamlActionLines.Count -eq 0) {
        $lines.Add('    []') | Out-Null
    }
    else {
        foreach ($line in $script:YamlActionLines) {
            $lines.Add($line) | Out-Null
        }
    }

    $lines.Add('  installed:') | Out-Null
    if ($script:InstalledList.Count -eq 0) {
        $lines.Add('    []') | Out-Null
    }
    else {
        foreach ($item in $script:InstalledList) {
            $lines.Add(("    - {0}" -f (ConvertTo-YamlSafeString $item))) | Out-Null
        }
    }

    $lines.Add('  skipped:') | Out-Null
    if ($script:SkippedList.Count -eq 0) {
        $lines.Add('    []') | Out-Null
    }
    else {
        foreach ($item in $script:SkippedList) {
            $lines.Add(("    - {0}" -f (ConvertTo-YamlSafeString $item))) | Out-Null
        }
    }

    $lines.Add('  failures:') | Out-Null
    if ($script:RunFailures.Count -eq 0) {
        $lines.Add('    []') | Out-Null
    }
    else {
        foreach ($item in $script:RunFailures) {
            $lines.Add(("    - {0}" -f (ConvertTo-YamlSafeString $item))) | Out-Null
        }
    }

    $lines.Add('drivers:') | Out-Null
    if ($script:DriverResults.Count -eq 0) {
        $lines.Add('  []') | Out-Null
    }
    else {
        foreach ($driver in $script:DriverResults) {
            $lines.Add('  -') | Out-Null
            $lines.Add(("    vendor: {0}" -f (ConvertTo-YamlSafeString $driver.Vendor))) | Out-Null
            $lines.Add(("    name: {0}" -f (ConvertTo-YamlSafeString $driver.Name))) | Out-Null
            $lines.Add(("    id: {0}" -f (ConvertTo-YamlSafeString $driver.Id))) | Out-Null
            $lines.Add(("    category: {0}" -f (ConvertTo-YamlSafeString $driver.Category))) | Out-Null
            $lines.Add(("    status: {0}" -f (ConvertTo-YamlSafeString $driver.Status))) | Out-Null
            $lines.Add(("    message: {0}" -f (ConvertTo-YamlSafeString $driver.Message))) | Out-Null
        }
    }

    Set-Content -LiteralPath $script:YamlPath -Value $lines -Encoding UTF8
    Write-Host ("YAML log written successfully: {0}" -f $script:YamlPath) -ForegroundColor Green
}

# -----------------------------
# File/folder helpers
# -----------------------------
function Ensure-Folder {
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -Path $Path -ItemType Directory -Force | Out-Null
    }
}

function Ensure-WorkingFolderPermissions {
    param([Parameter(Mandatory)][string]$Path)

    try {
        & icacls.exe $Path '/grant' '*S-1-1-0:(OI)(CI)F' '/T' '/C' | Out-Null
    }
    catch {
        Write-Log ("Unable to relax working folder permissions on {0}: {1}" -f $Path, $_.Exception.Message) 'WARN'
    }
}

function Remove-WorkingFolderRobust {
    param(
        [Parameter(Mandatory)][string]$Path,
        [int]$RetryCount = 6,
        [int]$RetryDelaySeconds = 5
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        Write-Log ("Working folder already absent: {0}" -f $Path) 'OK'
        return
    }

    Write-Log ("Attempting to remove working folder: {0}" -f $Path) 'INFO'
    for ($i = 1; $i -le $RetryCount; $i++) {
        try {
            Remove-Item -LiteralPath $Path -Recurse -Force -ErrorAction Stop
            Write-Log 'Working folder removed successfully.' 'OK'
            return
        }
        catch {
            if ($i -eq $RetryCount) {
                Add-RunFailure ("Failed to remove working folder after retries: {0}" -f $_.Exception.Message)
                return
            }
            Start-Sleep -Seconds $RetryDelaySeconds
        }
    }
}

# -----------------------------
# Vendor detection
# -----------------------------
function Get-SystemManufacturer {
    try {
        return (Get-CimInstance Win32_ComputerSystem -ErrorAction Stop).Manufacturer
    }
    catch {
        throw "Unable to determine system manufacturer. $($_.Exception.Message)"
    }
}

function Get-DriverVendor {
    $manufacturer = Get-SystemManufacturer
    Write-Log ("Detected manufacturer: {0}" -f $manufacturer) 'INFO'

    if ($manufacturer -match 'Dell') { return 'Dell' }
    if ($manufacturer -match 'HP|Hewlett-Packard') { return 'HP' }

    throw "Unsupported manufacturer for this script: $manufacturer"
}

# -----------------------------
# HP Support - HP Image Assistant extracted-folder deployment
# -----------------------------
function Get-HPSystemModelInfo {
    [CmdletBinding()]
    param()

    $cs = Get-CimInstance Win32_ComputerSystem -ErrorAction Stop
    $csp = Get-CimInstance Win32_ComputerSystemProduct -ErrorAction SilentlyContinue
    $bb = Get-CimInstance Win32_BaseBoard -ErrorAction SilentlyContinue

    $platform = $null
    if ($bb -and $bb.Product) {
        $platform = $bb.Product.ToString().Trim().ToUpper()
        if ($platform.Length -gt 4) {
            $platform = $platform.Substring(0,4)
        }
    }

    $model = if ($cs.Model) { $cs.Model.ToString().Trim() } else { 'Unknown' }
    $sku = if ($csp -and $csp.Version) { $csp.Version.ToString().Trim() } else { 'Unknown' }

    $info = [pscustomobject]@{
        Manufacturer = $cs.Manufacturer
        Model        = $model
        SKU          = $sku
        Platform     = $platform
    }

    Write-Log ("Detected HP system model: {0}" -f $info.Model) 'INFO'
    Write-Log ("Detected HP platform/baseboard ID: {0}" -f $info.Platform) 'INFO'
    Add-YamlAction ("Detected HP system model: {0}" -f $info.Model)
    Add-YamlAction ("Detected HP platform/baseboard ID: {0}" -f $info.Platform)

    return $info
}

function Get-ExistingHPIAExecutable {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$PreferredFolder)

    $candidateFolders = New-Object System.Collections.Generic.List[string]

    if (-not [string]::IsNullOrWhiteSpace($PreferredFolder)) {
        $candidateFolders.Add($PreferredFolder) | Out-Null
    }

    foreach ($folder in @(
        'C:\Program Files\HP\HP Image Assistant',
        'C:\Program Files (x86)\HP\HP Image Assistant',
        'C:\SWSetup\HPImageAssistant',
        'C:\ProgramData\Compton\HPImageAssistant'
    )) {
        if (-not $candidateFolders.Contains($folder)) {
            $candidateFolders.Add($folder) | Out-Null
        }
    }

    foreach ($folder in $candidateFolders) {
        if ([string]::IsNullOrWhiteSpace($folder) -or -not (Test-Path -LiteralPath $folder)) {
            continue
        }

        $exe = Get-ChildItem -Path $folder -Recurse -Filter 'HPImageAssistant.exe' -File -ErrorAction SilentlyContinue |
            Sort-Object FullName |
            Select-Object -First 1

        if ($exe) {
            return $exe.FullName
        }
    }

    return $null
}

function Install-HPIAFromExtractedFolder {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SourceFolder,
        [Parameter(Mandatory)][string]$DestinationFolder
    )

    Write-Section 'HP Image Assistant Local Deployment'
    Write-Log ("HPIA source folder: {0}" -f $SourceFolder) 'INFO'
    Write-Log ("HPIA local folder: {0}" -f $DestinationFolder) 'INFO'
    Add-YamlAction ("HPIA source folder: {0}" -f $SourceFolder)
    Add-YamlAction ("HPIA local folder: {0}" -f $DestinationFolder)

    $existingExe = Get-ExistingHPIAExecutable -PreferredFolder $DestinationFolder
    if ($existingExe) {
        try {
            $existingVersion = (Get-Item -LiteralPath $existingExe -ErrorAction Stop).VersionInfo.FileVersion
            Write-Log ("HP Image Assistant is already installed/found at: {0} (Version: {1})" -f $existingExe, $existingVersion) 'OK'
            Add-YamlAction ("Skipped HPIA local deployment because HPImageAssistant.exe already exists: {0} (Version: {1})" -f $existingExe, $existingVersion)
        }
        catch {
            Write-Log ("HP Image Assistant is already installed/found at: {0}" -f $existingExe) 'OK'
            Add-YamlAction ("Skipped HPIA local deployment because HPImageAssistant.exe already exists: {0}" -f $existingExe)
        }

        return $existingExe
    }

    Write-Log 'HP Image Assistant was not found locally. Deploying from extracted source folder.' 'INFO'
    Add-YamlAction 'HP Image Assistant was not found locally. Deploying from extracted source folder.'

    if (-not (Test-Path -LiteralPath $SourceFolder)) {
        throw "HPIA source folder not found: $SourceFolder"
    }

    $sourceFiles = @(Get-ChildItem -Path $SourceFolder -Recurse -File -ErrorAction SilentlyContinue)
    Write-Log ("Source HPIA file count: {0}" -f $sourceFiles.Count) 'INFO'
    Add-YamlAction ("Source HPIA file count: {0}" -f $sourceFiles.Count)

    $sourceExe = Get-ChildItem -Path $SourceFolder -Recurse -Filter 'HPImageAssistant.exe' -File -ErrorAction SilentlyContinue |
        Sort-Object FullName |
        Select-Object -First 1

    if (-not $sourceExe) {
        throw "HPImageAssistant.exe was not found anywhere under source folder: $SourceFolder"
    }

    Write-Log ("Found source HPImageAssistant.exe: {0}" -f $sourceExe.FullName) 'OK'
    Add-YamlAction ("Found source HPImageAssistant.exe: {0}" -f $sourceExe.FullName)

    try {
        if (Test-Path -LiteralPath $DestinationFolder) {
            Write-Log ("Removing existing local HPIA folder before clean deployment: {0}" -f $DestinationFolder) 'INFO'
            Remove-Item -LiteralPath $DestinationFolder -Recurse -Force -ErrorAction Stop
        }

        New-Item -Path $DestinationFolder -ItemType Directory -Force | Out-Null

        Write-Log 'Copying extracted HPIA files locally with robocopy...' 'INFO'

        $roboLog = Join-Path $DestinationFolder 'HPIA_robocopy.log'
        $roboArgs = @(
            ('"{0}"' -f $SourceFolder),
            ('"{0}"' -f $DestinationFolder),
            '/E',
            '/COPY:DAT',
            '/R:3',
            '/W:5',
            '/NFL',
            '/NDL',
            '/NP',
            ('/LOG:"{0}"' -f $roboLog)
        )

        $robo = Start-Process -FilePath "$env:SystemRoot\System32\robocopy.exe" -ArgumentList ($roboArgs -join ' ') -Wait -PassThru -NoNewWindow

        # Robocopy exit codes 0-7 are success/non-fatal. 8+ indicates failure.
        if ($robo.ExitCode -ge 8) {
            throw "Robocopy failed copying HPIA files. Exit code: $($robo.ExitCode). Log: $roboLog"
        }

        Write-Log ("HPIA files copied locally with robocopy exit code {0}." -f $robo.ExitCode) 'OK'
        Add-YamlAction ("HPIA files copied locally with robocopy exit code {0}." -f $robo.ExitCode)

        try {
            Get-ChildItem -Path $DestinationFolder -Recurse -File -ErrorAction SilentlyContinue |
                ForEach-Object { Unblock-File -Path $_.FullName -ErrorAction SilentlyContinue }
        }
        catch {}

        $copiedFiles = @(Get-ChildItem -Path $DestinationFolder -Recurse -File -ErrorAction SilentlyContinue)
        Write-Log ("Local HPIA folder file count after copy: {0}" -f $copiedFiles.Count) 'INFO'
        Add-YamlAction ("Local HPIA folder file count after copy: {0}" -f $copiedFiles.Count)

        $localExeItem = Get-ChildItem -Path $DestinationFolder -Recurse -Filter 'HPImageAssistant.exe' -File -ErrorAction SilentlyContinue |
            Sort-Object FullName |
            Select-Object -First 1

        if (-not $localExeItem) {
            $sampleFiles = $copiedFiles | Select-Object -First 20 | ForEach-Object { $_.FullName }
            foreach ($sample in $sampleFiles) {
                Write-Log ("Local HPIA sample file: {0}" -f $sample) 'WARN'
            }

            throw "HPImageAssistant.exe was not found anywhere under local copy folder: $DestinationFolder"
        }

        $localExe = $localExeItem.FullName

        Write-Log ("Resolved local HPImageAssistant.exe location: {0}" -f $localExe) 'OK'
        Add-YamlAction ("Resolved local HPImageAssistant.exe location: {0}" -f $localExe)

        return $localExe
    }
    catch {
        throw "Failed to deploy HP Image Assistant locally: $($_.Exception.Message)"
    }
}

function Get-HpiaExitStatus {
    param([int]$ExitCode)

    switch ($ExitCode) {
        0    { return 'success' }
        1    { return 'failed' }
        2    { return 'cancelled' }
        3    { return 'needs_reboot' }
        256  { return 'no_recommendations_or_success' }
        257  { return 'recommendations_found' }
        3010 { return 'needs_reboot' }
        3011 { return 'not_auto_installable_skipped' }
        4096 { return 'no_applicable_updates_or_platform_not_supported' }
        4097 { return 'invalid_parameters' }
        8193 { return 'hpia_analysis_or_report_generation_error' }
        default { return 'unknown' }
    }
}

function Get-HpiaRecommendationObjects {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ReportFolder
    )

    # Use a flexible PowerShell array instead of a strongly typed generic list.
    # HPIA reports can contain mixed object types from JSON and XML parsing.
    $recommendations = @()

    $jsonFiles = @(Get-ChildItem -Path $ReportFolder -Filter '*.json' -Recurse -File -ErrorAction SilentlyContinue)
    foreach ($jsonFile in $jsonFiles) {
        try {
            $json = Get-Content -LiteralPath $jsonFile.FullName -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop

            if ($json.HPIA -and $json.HPIA.Recommendations) {
                foreach ($rec in @($json.HPIA.Recommendations)) {
                    $recommendations += $rec
                    try {
                        Write-Log ("HPIA JSON recommendation type: {0}" -f $rec.GetType().FullName) 'INFO'
                    }
                    catch {}
                }
            }
            elseif ($json.Recommendations) {
                foreach ($rec in @($json.Recommendations)) {
                    $recommendations += $rec
                    try {
                        Write-Log ("HPIA JSON recommendation type: {0}" -f $rec.GetType().FullName) 'INFO'
                    }
                    catch {}
                }
            }
        }
        catch {
            Write-Log ("Unable to parse HPIA JSON report {0}: {1}" -f $jsonFile.FullName, $_.Exception.Message) 'WARN'
        }
    }

    $xmlFiles = @(Get-ChildItem -Path $ReportFolder -Filter '*.xml' -Recurse -File -ErrorAction SilentlyContinue)
    foreach ($xmlFile in $xmlFiles) {
        try {
            [xml]$xml = Get-Content -LiteralPath $xmlFile.FullName -Raw -ErrorAction Stop

            $nodes = @($xml.SelectNodes('//*[contains(translate(local-name(), "ABCDEFGHIJKLMNOPQRSTUVWXYZ", "abcdefghijklmnopqrstuvwxyz"), "recommend")]'))
            foreach ($node in $nodes) {
                $recommendations += $node
                try {
                    Write-Log ("HPIA XML recommendation node type: {0}" -f $node.GetType().FullName) 'INFO'
                }
                catch {}
            }
        }
        catch {
            Write-Log ("Unable to parse HPIA XML report {0}: {1}" -f $xmlFile.FullName, $_.Exception.Message) 'WARN'
        }
    }

    return @($recommendations)
}

function Get-HpiaRecommendationValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Recommendation,
        [Parameter(Mandatory)][string[]]$PropertyNames
    )

    foreach ($prop in $PropertyNames) {
        try {
            if ($Recommendation.PSObject.Properties.Name -contains $prop) {
                $value = $Recommendation.$prop
                if ($null -ne $value -and -not [string]::IsNullOrWhiteSpace($value.ToString())) {
                    return $value.ToString()
                }
            }
        }
        catch {}
    }

    # XML fallback
    try {
        foreach ($prop in $PropertyNames) {
            $node = $Recommendation.SelectSingleNode('.//*[local-name()="' + $prop + '"]')
            if ($node -and -not [string]::IsNullOrWhiteSpace($node.InnerText)) {
                return $node.InnerText.Trim()
            }
        }
    }
    catch {}

    return $null
}

function Get-HpiaSoftPaqNumber {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Recommendation)

    $candidate = Get-HpiaRecommendationValue -Recommendation $Recommendation -PropertyNames @(
        'SoftPaqId','SoftpaqId','SoftPaq','Softpaq','SoftPaqNumber','SoftpaqNumber','SP','Id','ID','Number'
    )

    if ($candidate -match '(?i)sp?(\d{5,6})') {
        return $matches[1]
    }

    $text = ($Recommendation | Out-String)
    if ($text -match '(?i)sp(\d{5,6})') {
        return $matches[1]
    }

    if ($text -match '\b(\d{5,6})\b') {
        return $matches[1]
    }

    return $null
}

function Get-HpiaRecommendationCategory {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Recommendation)

    return (Get-HpiaRecommendationValue -Recommendation $Recommendation -PropertyNames @(
        'Category','Type','RecommendationType','ComponentType','Class','Group'
    ))
}

function Get-HpiaRecommendationName {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Recommendation)

    $name = Get-HpiaRecommendationValue -Recommendation $Recommendation -PropertyNames @(
        'Name','Title','Component','ComponentName','Description','SoftPaqName','SoftpaqName'
    )

    if ($name) { return $name }

    $text = ($Recommendation | Out-String).Trim()
    if ($text.Length -gt 160) {
        return $text.Substring(0,160)
    }

    return $text
}

function New-HpiaDriverSPList {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ReportFolder,
        [Parameter(Mandatory)][string]$SPListPath
    )

    $recommendations = @(Get-HpiaRecommendationObjects -ReportFolder $ReportFolder)
    Write-Log ("HPIA recommendations parsed from reports: {0}" -f $recommendations.Count) 'INFO'
    Add-YamlAction ("HPIA recommendations parsed from reports: {0}" -f $recommendations.Count)

    $selected = @()
    $seen = @{}

    foreach ($rec in $recommendations) {
        $category = Get-HpiaRecommendationCategory -Recommendation $rec
        $name = Get-HpiaRecommendationName -Recommendation $rec
        $sp = Get-HpiaSoftPaqNumber -Recommendation $rec

        if (-not $sp) {
            continue
        }

        # Exclude BIOS/Firmware explicitly.
        $combined = ("{0} {1}" -f $category, $name)
        if ($combined -match '(?i)\bBIOS\b|Firmware') {
            $script:SkippedList.Add(("SP{0} {1}" -f $sp, $name)) | Out-Null
            Add-DriverResult -Vendor 'HP' -Name $name -Id $sp -Category $category -Status 'Blocked' -Message 'Excluded because it appears to be BIOS/Firmware.'
            continue
        }

        # Prefer driver-like recommendations, but allow blank category if the SoftPaq was recommended and not BIOS/Firmware.
        if ($combined -notmatch '(?i)Driver|Bluetooth|Chipset|Audio|Graphics|Video|LAN|WLAN|Wireless|NIC|Touch|Fingerprint|Card Reader|Storage|Serial|USB|Management Engine' -and $category) {
            $script:SkippedList.Add(("SP{0} {1}" -f $sp, $name)) | Out-Null
            Add-DriverResult -Vendor 'HP' -Name $name -Id $sp -Category $category -Status 'Skipped' -Message 'Excluded because it did not appear to be a driver recommendation.'
            continue
        }

        if (-not $seen.ContainsKey($sp)) {
            $seen[$sp] = $true
            $selected += $sp
            Add-DriverResult -Vendor 'HP' -Name $name -Id $sp -Category $category -Status 'Detected' -Message 'Selected for HPIA SPList install.'
        }
    }

    if (@($selected).Count -gt 0) {
        Set-Content -LiteralPath $SPListPath -Value $selected -Encoding ASCII
        Write-Log ("Created filtered HPIA SPList with {0} SoftPaqs: {1}" -f @($selected).Count, $SPListPath) 'OK'
        Add-YamlAction ("Created filtered HPIA SPList with {0} SoftPaqs: {1}" -f @($selected).Count, $SPListPath)
    }
    else {
        Write-Log 'No non-BIOS/Firmware driver SoftPaq recommendations were selected from HPIA reports.' 'OK'
        Add-YamlAction 'No non-BIOS/Firmware driver SoftPaq recommendations were selected from HPIA reports.'
    }

    return @($selected)
}



# -----------------------------
# HP PowerShell module maintenance
# -----------------------------
function Initialize-NetworkDefaults {
    try {
        [Net.ServicePointManager]::SecurityProtocol =
            [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls13
    }
    catch {
        try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 }
        catch {}
    }

    try { [Net.ServicePointManager]::DefaultConnectionLimit = 64 }
    catch {}
}

function Get-HighestInstalledModuleVersion {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Name
    )

    try {
        $installed = Get-Module -Name $Name -ListAvailable -ErrorAction SilentlyContinue |
            Sort-Object Version -Descending |
            Select-Object -First 1

        if ($installed) { return [version]$installed.Version }
    }
    catch {
        Write-Log "Could not determine installed version for module ${Name}: $($_.Exception.Message)" 'WARN'
    }

    return $null
}

function Invoke-HPCMSLUpdateInFreshPowerShell {
    [CmdletBinding()]
    param(
        [ValidateSet('CurrentUser','AllUsers')]
        [string]$Scope = 'AllUsers'
    )

    Write-Log 'Starting isolated HPCMSL install/update using PowerShellGet/PSResourceGet when available...' 'INFO'
    Add-YamlAction 'Starting isolated HPCMSL install/update before HP driver updates.'

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

    if ([string]::IsNullOrWhiteSpace(`$InstallRoot)) { throw 'InstallRoot is blank.' }
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

    if (`$installed) { HelperLog ('Installed HPCMSL version: {0}' -f `$installed) }
    else { HelperLog 'HPCMSL is not currently installed.' }
    HelperLog ('Latest PSGallery HPCMSL version: {0}' -f `$online)

    if (`$installed -and `$installed -ge `$online) {
        HelperLog 'HPCMSL is already current. No update needed.'
        exit 0
    }

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
        if ('$Scope' -eq 'AllUsers') { `$installRoot = Join-Path `$env:ProgramFiles 'WindowsPowerShell\Modules' }
        else { `$installRoot = Join-Path `$HOME 'Documents\WindowsPowerShell\Modules' }
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
        Initialize-NetworkDefaults
        Set-Content -Path $helperPath -Value $helperScript -Encoding UTF8 -Force
        Write-Log "HPCMSL helper script created at $helperPath" 'INFO'

        $args = @('-NoProfile','-ExecutionPolicy','Bypass','-File',('"{0}"' -f $helperPath))
        $process = Start-Process -FilePath powershell.exe -ArgumentList ($args -join ' ') -Wait -PassThru -WindowStyle Hidden

        if (Test-Path -LiteralPath $helperLog) {
            Get-Content -LiteralPath $helperLog -ErrorAction SilentlyContinue | ForEach-Object {
                if (-not [string]::IsNullOrWhiteSpace($_)) { Write-Log "HPCMSL helper: $_" 'INFO' }
            }
        }

        if ($process.ExitCode -ne 0) {
            Write-Log "Isolated HPCMSL helper failed with exit code $($process.ExitCode). Continuing with HPIA driver updates." 'WARN'
            Add-YamlAction ("Isolated HPCMSL helper failed with exit code {0}. Continuing with HPIA driver updates." -f $process.ExitCode)
            return $false
        }

        $finalVersion = Get-HighestInstalledModuleVersion -Name 'HPCMSL'
        if ($finalVersion) {
            Write-Log "HPCMSL final installed version after isolated update: $finalVersion" 'OK'
            Add-YamlAction ("HPCMSL final installed version after isolated update: {0}" -f $finalVersion)
        }
        else {
            Write-Log 'HPCMSL helper completed, but this session could not verify the installed HPCMSL version.' 'WARN'
            Add-YamlAction 'HPCMSL helper completed, but installed version could not be verified.'
            return $false
        }

        return $true
    }
    catch {
        Write-Log "Failed to run isolated HPCMSL update helper: $($_.Exception.Message)" 'WARN'
        Add-YamlAction ("Failed to run isolated HPCMSL update helper: {0}" -f $_.Exception.Message)
        return $false
    }
    finally {
        Remove-Item -LiteralPath $helperPath -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $helperLog -Force -ErrorAction SilentlyContinue
    }
}

function Invoke-HPPowerShellModuleMaintenance {
    [CmdletBinding()]
    param()

    Write-Section 'HP PowerShell Module Maintenance'
    Write-Log 'HP vendor detected. Checking/installing HPCMSL before HP driver updates...' 'INFO'
    Add-YamlAction 'HP vendor detected; checking/installing HPCMSL before HP driver updates.'

    $hpcmslUpdated = Invoke-HPCMSLUpdateInFreshPowerShell -Scope AllUsers
    if ($hpcmslUpdated) {
        Write-Log 'HP PowerShell module maintenance completed.' 'OK'
        Add-YamlAction 'HP PowerShell module maintenance completed.'
    }
    else {
        Write-Log 'HP PowerShell module maintenance did not complete successfully. Continuing with HP Image Assistant driver updates.' 'WARN'
        Add-YamlAction 'HP PowerShell module maintenance did not complete successfully; continuing with HPIA driver updates.'
    }
}

function Invoke-HPDriverUpdates {
    Write-Section 'HP Driver Analysis and Installation'

    $hpInfo = Get-HPSystemModelInfo

    $hpiaExe = Install-HPIAFromExtractedFolder -SourceFolder $HpiaSourceFolder -DestinationFolder $LocalHpiaFolder

    $hpiaReportFolder = Join-Path $YamlLogFolder 'HPIA'
    $hpiaDownloadFolder = Join-Path $WorkingRoot 'HPIADownloads'
    $hpiaExtractFolder = Join-Path $WorkingRoot 'HPIAExtracted'

    Ensure-Folder -Path $hpiaReportFolder
    Ensure-Folder -Path $hpiaDownloadFolder
    Ensure-Folder -Path $hpiaExtractFolder
    Ensure-Folder -Path 'C:\Temp'

    # HPIA is more reliable under Task Scheduler/SYSTEM when TEMP/TMP are local and TLS 1.2 is forced.
    $env:TEMP = 'C:\Temp'
    $env:TMP  = 'C:\Temp'

    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Write-Log 'TLS 1.2 enabled for HP Image Assistant network operations.' 'INFO'
        Add-YamlAction 'TLS 1.2 enabled for HPIA network operations.'
    }
    catch {
        Write-Log ("Unable to force TLS 1.2 before HPIA run: {0}" -f $_.Exception.Message) 'WARN'
        Add-YamlAction ("Unable to force TLS 1.2 before HPIA run: {0}" -f $_.Exception.Message)
    }

    # Production unattended HPIA workflow:
    # - Avoids /Action:List because it has been returning 8193 on this platform.
    # - Uses /Category:Drivers to avoid BIOS/Firmware/Dock/Thunderbolt firmware during unattended runs.
    # - Uses /Noninteractive, /AutoCleanup, and /Debug for scheduled task reliability and better logs.
    Write-Log 'Running HP Image Assistant production analyze/install pass for drivers only...' 'INFO'
    Add-YamlAction 'Running HPIA production analyze/install pass for drivers only.'

    $hpiaArgs = @(
        '/Operation:Analyze',
        '/Action:Install',
        '/Category:Drivers',
        '/Selection:All',
        '/Silent',
        '/Noninteractive',
        '/AutoCleanup',
        '/Debug',
        "/ReportFolder:`"$hpiaReportFolder`"",
        "/SoftpaqDownloadFolder:`"$hpiaDownloadFolder`"",
        "/SoftpaqExtractFolder:`"$hpiaExtractFolder`""
    )

    Write-Log ("HPIA production command: {0} {1}" -f $hpiaExe, ($hpiaArgs -join ' ')) 'INFO'
    Add-YamlAction ("HPIA production command: {0} {1}" -f $hpiaExe, ($hpiaArgs -join ' '))

    Write-Progress -Activity 'HP Image Assistant' -Status ("Installing recommended drivers for {0}" -f $hpInfo.Model) -PercentComplete 50

    $hpiaProc = Start-Process -FilePath $hpiaExe -ArgumentList ($hpiaArgs -join ' ') -Wait -PassThru -NoNewWindow
    $exitCode = [int]$hpiaProc.ExitCode
    $status = Get-HpiaExitStatus -ExitCode $exitCode

    Write-Progress -Activity 'HP Image Assistant' -Completed

    Write-Log ("HP Image Assistant production pass completed with exit code {0} ({1})." -f $exitCode, $status) 'INFO'
    Add-YamlAction ("HP Image Assistant production pass completed with exit code {0} ({1})." -f $exitCode, $status)

    $downloadedFiles = @(Get-ChildItem -Path $hpiaDownloadFolder -File -Recurse -ErrorAction SilentlyContinue)
    foreach ($file in $downloadedFiles) {
        Add-DriverResult -Vendor 'HP' -Name $file.Name -Id $null -Category 'Driver' -Status 'Detected' -Message ("Downloaded/processed by HPIA: {0}" -f $file.FullName)
    }

    $reportFiles = @(Get-ChildItem -Path $hpiaReportFolder -File -Recurse -ErrorAction SilentlyContinue)
    foreach ($report in $reportFiles) {
        Add-YamlAction ("HPIA report generated: {0}" -f $report.FullName)
    }

    switch ($exitCode) {
        0 {
            Write-Log 'HPIA completed successfully.' 'OK'
            Add-YamlAction 'HPIA completed successfully.'
            return
        }

        256 {
            Write-Log 'HPIA completed successfully. No applicable driver recommendations were found or no action was required.' 'OK'
            Add-YamlAction 'HPIA completed successfully with no applicable driver recommendations or no action required.'
            return
        }

        257 {
            Write-Log 'HPIA completed and reported recommendations/driver actions.' 'OK'
            Add-YamlAction 'HPIA completed and reported recommendations/driver actions.'
            return
        }

        3010 {
            Write-Log 'HPIA completed successfully. Reboot required.' 'WARN'
            Add-YamlAction 'HPIA completed successfully and indicated reboot required.'
            return
        }

        3011 {
            Write-Log 'One or more HPIA items were not auto-installable and were skipped.' 'WARN'
            Add-YamlAction 'One or more HPIA items were not auto-installable and were skipped.'
            return
        }

        4096 {
            Write-Log 'HPIA completed but did not find applicable driver updates for this platform.' 'OK'
            Add-YamlAction 'HPIA completed but did not find applicable driver updates for this platform.'
            return
        }

        8193 {
            Write-Log 'HPIA returned 8193. Checking whether reports/logs were generated before failing the run.' 'WARN'
            Add-YamlAction 'HPIA returned 8193. Checking whether reports/logs were generated before failing the run.'

            $generatedReports = @(Get-ChildItem -Path $hpiaReportFolder -File -Recurse -ErrorAction SilentlyContinue)
            if ($generatedReports.Count -gt 0) {
                Write-Log ("HPIA generated {0} report/log file(s) despite exit code 8193. Continuing so the scheduled workflow does not hard fail." -f $generatedReports.Count) 'WARN'
                Add-YamlAction ("HPIA generated {0} report/log file(s) despite exit code 8193. Continuing without hard failure." -f $generatedReports.Count)
                return
            }

            throw "HPIA failed with exit code 8193 and produced no report/log files in $hpiaReportFolder."
        }

        default {
            throw "HP Image Assistant production pass failed or returned an unexpected exit code: $exitCode ($status). Review reports in $hpiaReportFolder."
        }
    }
}

# -----------------------------
# Dell support
# -----------------------------
function Wait-ForFile {
    param(
        [Parameter(Mandatory)][string]$Path,
        [int]$TimeoutSeconds = 30
    )

    $start = Get-Date
    while (-not (Test-Path -LiteralPath $Path)) {
        Start-Sleep -Seconds 1
        if (((Get-Date) - $start).TotalSeconds -ge $TimeoutSeconds) {
            return $false
        }
    }
    return $true
}

function Get-DcuReportXmlSafely {
    param(
        [Parameter(Mandatory)][string]$Path,
        [int]$RetryCount = 5,
        [int]$RetryDelaySeconds = 2
    )

    if (-not (Wait-ForFile -Path $Path -TimeoutSeconds 20)) {
        throw "Dell DCU report file was not found: $Path"
    }

    Start-Sleep -Seconds 3

    for ($i = 1; $i -le $RetryCount; $i++) {
        try {
            $fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
            try {
                $sr = New-Object System.IO.StreamReader($fs)
                try {
                    $content = $sr.ReadToEnd()
                }
                finally {
                    $sr.Dispose()
                }
            }
            finally {
                $fs.Dispose()
            }

            $xml = New-Object System.Xml.XmlDocument
            $xml.LoadXml($content)
            return $xml
        }
        catch {
            if ($i -eq $RetryCount) {
                throw
            }
            Start-Sleep -Seconds $RetryDelaySeconds
        }
    }
}

function Get-DellNodeText {
    param(
        [Parameter(Mandatory)][xml]$Node,
        [Parameter(Mandatory)][string[]]$Names
    )

    foreach ($name in $Names) {
        try {
            $xpath = './/*[local-name()="' + $name + '"]'
            $child = $Node.SelectSingleNode($xpath)
            if ($child -and -not [string]::IsNullOrWhiteSpace($child.InnerText)) {
                return $child.InnerText.Trim()
            }
        }
        catch {}
    }

    return $null
}

function Get-DellReportItems {
    param([Parameter(Mandatory)][xml]$Xml)

    $items = @()
    try {
        $xpath = '//*[local-name()="Update" or local-name()="Package" or local-name()="SoftwareComponent" or local-name()="component" or local-name()="Device"]'
        $nodes = $Xml.SelectNodes($xpath)
        foreach ($node in $nodes) {
            $name = Get-DellNodeText -Node $node -Names @('Name','Title','PackageName')
            $version = Get-DellNodeText -Node $node -Names @('Version','PackageVersion')
            $category = Get-DellNodeText -Node $node -Names @('Category','Type')
            $id = Get-DellNodeText -Node $node -Names @('Id','PackageId','ReleaseId')

            if ($name -or $id) {
                $items += [pscustomobject]@{
                    Id       = $id
                    Name     = $name
                    Version  = $version
                    Category = $category
                }
            }
        }
    }
    catch {}

    return @($items)
}

function Get-DellDCUService {
    $candidateNames = @(
        'DellClientManagementService',
        'DellCommandUpdate',
        'DellUpdateService'
    )

    foreach ($name in $candidateNames) {
        $svc = Get-Service -Name $name -ErrorAction SilentlyContinue
        if ($svc) { return $svc }
    }

    $svc = Get-Service -ErrorAction SilentlyContinue |
        Where-Object { $_.DisplayName -match 'Dell.*Client.*Management|Dell.*Command.*Update' } |
        Select-Object -First 1

    return $svc
}

function Ensure-DellDCUService {
    param(
        [int]$TimeoutSeconds = 30
    )

    Write-Log 'Validating Dell Client Management Service...' 'INFO'
    Add-YamlAction 'Validating Dell Client Management Service.'

    $service = Get-DellDCUService
    if (-not $service) {
        throw 'Dell Client Management Service was not found. Dell Command | Update may need to be repaired or reinstalled.'
    }

    Write-Log ("Dell service detected: {0} ({1})" -f $service.DisplayName, $service.Name) 'OK'

    try {
        $wmiService = Get-CimInstance -ClassName Win32_Service -Filter ("Name='{0}'" -f $service.Name) -ErrorAction Stop
        if ($wmiService.StartMode -eq 'Disabled') {
            Write-Log 'Dell service startup type is Disabled. Setting it to Manual.' 'WARN'
            Set-Service -Name $service.Name -StartupType Manual -ErrorAction Stop
        }
    }
    catch {
        Write-Log ("Could not validate/set Dell service startup type: {0}" -f $_.Exception.Message) 'WARN'
    }

    $service.Refresh()
    if ($service.Status -ne 'Running') {
        Write-Log 'Starting Dell Client Management Service...' 'INFO'
        try {
            Start-Service -Name $service.Name -ErrorAction Stop
        }
        catch {
            Write-Log ("Start-Service failed. Attempting sc.exe recovery start. Error: {0}" -f $_.Exception.Message) 'WARN'
            & sc.exe start $service.Name | Out-Null
        }
    }

    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    do {
        Start-Sleep -Seconds 2
        $service = Get-Service -Name $service.Name -ErrorAction SilentlyContinue
        if ($service -and $service.Status -eq 'Running') {
            Write-Log 'Dell Client Management Service is running.' 'OK'
            Add-YamlAction 'Dell Client Management Service is running.'
            return $true
        }
    } while ((Get-Date) -lt $deadline)

    throw 'Dell Client Management Service did not reach the Running state before timeout.'
}

function Invoke-DellDCUCommandWithRetry {
    param(
        [Parameter(Mandatory)][string]$DcuCli,
        [Parameter(Mandatory)][string]$Arguments,
        [Parameter(Mandatory)][string]$OperationName,
        [int[]]$AcceptableExitCodes = @(0),
        [int]$MaxAttempts = 2
    )

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        Write-Log ("Dell DCU {0} attempt {1} of {2}..." -f $OperationName, $attempt, $MaxAttempts) 'INFO'
        Add-YamlAction ("Dell DCU {0} attempt {1} of {2}." -f $OperationName, $attempt, $MaxAttempts)

        $proc = Start-Process -FilePath $DcuCli -ArgumentList $Arguments -Wait -PassThru -NoNewWindow
        $exitCode = [int]$proc.ExitCode
        Write-Log ("Dell DCU {0} exit code: {1}" -f $OperationName, $exitCode) 'INFO'

        if ($AcceptableExitCodes -contains $exitCode) {
            return $proc
        }

        if ($exitCode -eq 3000) {
            Write-Log 'Dell DCU returned 3000, which normally indicates the Dell Client Management Service stopped or crashed.' 'WARN'
            Add-YamlAction 'Dell DCU returned 3000; attempting Dell service recovery before retry.'
        }
        else {
            Write-Log ("Dell DCU {0} returned non-success exit code {1}." -f $OperationName, $exitCode) 'WARN'
        }

        if ($attempt -lt $MaxAttempts) {
            Ensure-DellDCUService | Out-Null
            Start-Sleep -Seconds 5
            continue
        }

        throw "Dell DCU $OperationName failed after $MaxAttempts attempt(s). Last exit code: $exitCode"
    }
}

function Invoke-DellDriverUpdates {
    Write-Section 'Dell Command Update Workflow'

    $dcuCli = Join-Path ${env:ProgramFiles} 'Dell\CommandUpdate\dcu-cli.exe'
    if (-not (Test-Path -LiteralPath $dcuCli)) {
        throw "Dell Command | Update CLI was not found: $dcuCli"
    }

    Write-Log ("Using Dell Command | Update CLI: {0}" -f $dcuCli) 'OK'
    Add-YamlAction 'Using Dell Command | Update CLI.'

    Ensure-DellDCUService | Out-Null

    $dcuScanLog  = Join-Path $WorkingRoot 'Dell-DCU-Scan.log'
    $dcuApplyLog = Join-Path $WorkingRoot 'Dell-DCU-Apply.log'
    $dcuReport   = Join-Path $WorkingRoot 'Dell-DCU-ApplicableUpdates.xml'

    Write-Log 'Dell DCU Configure...' 'INFO'
    Write-Progress -Activity 'Dell Driver Update Workflow' -Status 'Configuring Dell Command Update' -PercentComplete 10
    $configureArgs = "/configure -silent -scheduleAuto -lockSettings=disable"
    Invoke-DellDCUCommandWithRetry -DcuCli $dcuCli -Arguments $configureArgs -OperationName 'Configure' -MaxAttempts 2 | Out-Null

    Write-Log 'Dell DCU Scan...' 'INFO'
    Write-Progress -Activity 'Dell Driver Update Workflow' -Status 'Scanning for updates' -PercentComplete 35
    $scanArgs = "/scan -silent -outputLog=""$dcuScanLog"" -report=""$dcuReport"""
    Invoke-DellDCUCommandWithRetry -DcuCli $dcuCli -Arguments $scanArgs -OperationName 'Scan' -MaxAttempts 2 | Out-Null
    Write-Log 'Waiting for Dell Command | Update to finish writing the report...' 'INFO'

    try {
        $xml = Get-DcuReportXmlSafely -Path $dcuReport
        $items = Get-DellReportItems -Xml $xml

        if ($items.Count -gt 0) {
            Add-YamlAction ("Dell DCU report parsed successfully. Updates detected: {0}" -f $items.Count)

            $total = $items.Count
            $index = 0

            foreach ($item in $items) {
                $index++
                $percent = 35 + [math]::Floor(($index / $total) * 35)
                $label = if ($item.Name) { $item.Name } elseif ($item.Id) { $item.Id } else { 'Dell update' }

                Write-Progress -Activity 'Dell Driver Update Workflow' -Status ("Parsing report: {0}" -f $label) -PercentComplete $percent

                if ($item.Category -match 'BIOS|Firmware') {
                    $script:SkippedList.Add($label) | Out-Null
                    Add-DriverResult -Vendor 'Dell' -Name $item.Name -Id $item.Id -Category $item.Category -Status 'Blocked' -Message 'BIOS/Firmware update blocked by script policy.'
                    Write-Log ("Blocking Dell BIOS/Firmware update: {0}" -f $label) 'WARN'
                }
                else {
                    Add-DriverResult -Vendor 'Dell' -Name $item.Name -Id $item.Id -Category $item.Category -Status 'Detected' -Message 'Detected in Dell DCU report.'
                }
            }
        }
        else {
            Add-YamlAction 'Dell DCU report parsed but returned no identifiable updates.'
        }
    }
    catch {
        Write-Log ("Failed to parse Dell DCU report: {0}" -f $_.Exception.Message) 'WARN'
        Add-YamlAction ("Failed to parse Dell DCU report: {0}" -f $_.Exception.Message)
    }

    Write-Log 'Dell DCU ApplyUpdates...' 'INFO'
    Write-Progress -Activity 'Dell Driver Update Workflow' -Status 'Applying updates' -PercentComplete 85
    $applyArgs = "/applyUpdates -silent -reboot=disable -outputLog=""$dcuApplyLog"""
    Invoke-DellDCUCommandWithRetry -DcuCli $dcuCli -Arguments $applyArgs -OperationName 'ApplyUpdates' -MaxAttempts 2 | Out-Null

    Write-Log (("Dell DCU logs: {0} ; {1} ; {2}") -f $dcuScanLog, $dcuApplyLog, $dcuReport) 'OK'
    Write-Progress -Activity 'Dell Driver Update Workflow' -Completed
}
# -----------------------------
# Main
# -----------------------------
$finalStatus = 'success'

try {
    Write-Section 'Initialization'

    Ensure-Folder -Path $YamlLogFolder
    Ensure-Folder -Path $WorkingRoot
    Ensure-WorkingFolderPermissions -Path $WorkingRoot

    $yamlName = "{0}-{1}-{2}.yml" -f $script:ComputerName, '05_Weekend_Vendor_Drivers_Update', (Get-Date -Format 'yyyy-MM-dd_HHmmss')
    $yamlPath = Join-Path $YamlLogFolder $yamlName

    Write-Log ("YAML log will be written to: {0}" -f $yamlPath) 'INFO'
    Initialize-YamlLog -ComputerName $script:ComputerName -YamlPath $yamlPath

    Write-Log 'Initializing vendor driver update script...' 'INFO'


    Write-Section 'Vendor Detection'
    $script:DetectedVendor = Get-DriverVendor
    Write-Log ("Detected vendor workflow: {0}" -f $script:DetectedVendor) 'INFO'
    Write-Log ("Working root: {0}" -f $WorkingRoot) 'INFO'

    if ($script:DetectedVendor -eq 'HP') {
        Invoke-HPPowerShellModuleMaintenance
        Invoke-HPDriverUpdates
    }
    elseif ($script:DetectedVendor -eq 'Dell') {
        Invoke-DellDriverUpdates
    }

}
catch {
    $finalStatus = 'failed'
    Add-RunFailure ("Script failed: {0}" -f $_.Exception.Message)
}
finally {
    Write-Section 'Cleanup'
    Remove-WorkingFolderRobust -Path $WorkingRoot

    if ($script:RunFailures.Count -gt 0 -and $finalStatus -ne 'failed') {
        $finalStatus = 'completed_with_warnings'
    }

    if ($script:RunFailures.Count -gt 0) {
        Write-Log (("{0} driver update completed with one or more failures.") -f $script:DetectedVendor) 'WARN'
    }
    else {
        Write-Log (("{0} driver update script completed successfully.") -f $script:DetectedVendor) 'OK'
    }

    Save-YamlLog -Status $finalStatus
}# ScriptName: 05_Weekend_HP_Drivers_Update.ps1
# ScriptVersion: 2.5.1
# LastUpdated: 2026-06-15
# Purpose: Weekend vendor driver update script with clean HP + Dell support,
#          YAML logging, colored output,
#          section headers, progress display, and structured per-driver results.

[CmdletBinding()]
param([string]$WorkingRoot = 'C:\Temp\DriverUpdates',
    [string]$YamlLogFolder = 'C:\Logs',
    [switch]$IncludeSoftware,
    [switch]$IncludeBIOS,
    [string]$HpiaSourceFolder = '\\filesvr\Labscripts\HPImageAssistant',
    [string]$LocalHpiaFolder = 'C:\ProgramData\Compton\HPImageAssistant'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# -----------------------------
# Script metadata
# -----------------------------
$script:ScriptName        = '05_Weekend_HP_Drivers_Update.ps1'
$script:ScriptVersion     = '2.5.1'
$script:StartTime         = Get-Date
$script:RunFailures       = New-Object System.Collections.Generic.List[string]
$script:InstalledList     = New-Object System.Collections.Generic.List[string]
$script:SkippedList       = New-Object System.Collections.Generic.List[string]
$script:YamlActionLines   = New-Object System.Collections.Generic.List[string]
$script:DriverResults     = New-Object System.Collections.Generic.List[object]
$script:DetectedVendor    = $null
$script:YamlPath          = $null
$script:ComputerName      = $env:COMPUTERNAME

# -----------------------------
# Logging helpers
# -----------------------------
function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','OK','WARN','ERROR')][string]$Level = 'INFO'
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[{0}] [{1}] {2}" -f $timestamp, $Level.PadRight(5), $Message

    switch ($Level) {
        'INFO'  { Write-Host $line -ForegroundColor Cyan }
        'OK'    { Write-Host $line -ForegroundColor Green }
        'WARN'  { Write-Host $line -ForegroundColor Yellow }
        'ERROR' { Write-Host $line -ForegroundColor Red }
        default { Write-Host $line }
    }
}

function Write-Section {
    param([Parameter(Mandatory)][string]$Title)

    $border = ('=' * 72)
    Write-Host ''
    Write-Host $border -ForegroundColor Magenta
    Write-Host ("  {0}" -f $Title) -ForegroundColor Magenta
    Write-Host $border -ForegroundColor Magenta
    Add-YamlAction ("Section: {0}" -f $Title)
}

function Add-RunFailure {
    param([Parameter(Mandatory)][string]$Message)
    $script:RunFailures.Add($Message) | Out-Null
    Write-Log $Message 'WARN'
}

function Add-DriverResult {
    param(
        [Parameter(Mandatory)][string]$Vendor,
        [Parameter(Mandatory)][string]$Name,
        [string]$Id,
        [string]$Category,
        [ValidateSet('Detected','Installed','Downloaded','Skipped','Failed','Blocked')][string]$Status,
        [string]$Message = ''
    )

    $script:DriverResults.Add([pscustomobject]@{
        Vendor   = $Vendor
        Name     = $Name
        Id       = $Id
        Category = $Category
        Status   = $Status
        Message  = $Message
    }) | Out-Null
}

function ConvertTo-YamlSafeString {
    param([AllowNull()][object]$Value)

    if ($null -eq $Value) { return 'null' }

    $text = [string]$Value
    $text = $text -replace "`r", ''
    $text = $text -replace "`n", ' '
    $text = $text -replace "'", "''"
    return "'$text'"
}

function Initialize-YamlLog {
    param(
        [Parameter(Mandatory)][string]$ComputerName,
        [Parameter(Mandatory)][string]$YamlPath
    )

    $script:YamlPath = $YamlPath
    $script:YamlActionLines.Clear()

    Add-YamlAction 'Script initialized.'
    Add-YamlAction ("Working root: {0}" -f $WorkingRoot)
    Add-YamlAction ("Computer: {0}" -f $ComputerName)
}

function Add-YamlAction {
    param([Parameter(Mandatory)][string]$Text)
    $script:YamlActionLines.Add(("    - {0}" -f (ConvertTo-YamlSafeString $Text))) | Out-Null
}

function Save-YamlLog {
    param(
        [Parameter(Mandatory)][string]$Status
    )

    if ([string]::IsNullOrWhiteSpace($script:YamlPath)) {
        return
    }

    $endTime = Get-Date
    $duration = [math]::Round(($endTime - $script:StartTime).TotalSeconds, 0)

    $lines = New-Object System.Collections.Generic.List[string]

    $lines.Add('script:') | Out-Null
    $lines.Add(("  name: {0}" -f (ConvertTo-YamlSafeString $script:ScriptName))) | Out-Null
    $lines.Add(("  version: {0}" -f (ConvertTo-YamlSafeString $script:ScriptVersion))) | Out-Null
    $lines.Add(("  computer: {0}" -f (ConvertTo-YamlSafeString $script:ComputerName))) | Out-Null
    $lines.Add(("  started: {0}" -f (ConvertTo-YamlSafeString ($script:StartTime.ToString('s'))))) | Out-Null
    $lines.Add(("  ended: {0}" -f (ConvertTo-YamlSafeString ($endTime.ToString('s'))))) | Out-Null
    $lines.Add(("  duration_seconds: {0}" -f $duration)) | Out-Null

    $lines.Add('run:') | Out-Null
    $lines.Add(("  status: {0}" -f (ConvertTo-YamlSafeString $Status))) | Out-Null
    $lines.Add(("  vendor: {0}" -f (ConvertTo-YamlSafeString $script:DetectedVendor))) | Out-Null

    $lines.Add('  actions:') | Out-Null
    if ($script:YamlActionLines.Count -eq 0) {
        $lines.Add('    []') | Out-Null
    }
    else {
        foreach ($line in $script:YamlActionLines) {
            $lines.Add($line) | Out-Null
        }
    }

    $lines.Add('  installed:') | Out-Null
    if ($script:InstalledList.Count -eq 0) {
        $lines.Add('    []') | Out-Null
    }
    else {
        foreach ($item in $script:InstalledList) {
            $lines.Add(("    - {0}" -f (ConvertTo-YamlSafeString $item))) | Out-Null
        }
    }

    $lines.Add('  skipped:') | Out-Null
    if ($script:SkippedList.Count -eq 0) {
        $lines.Add('    []') | Out-Null
    }
    else {
        foreach ($item in $script:SkippedList) {
            $lines.Add(("    - {0}" -f (ConvertTo-YamlSafeString $item))) | Out-Null
        }
    }

    $lines.Add('  failures:') | Out-Null
    if ($script:RunFailures.Count -eq 0) {
        $lines.Add('    []') | Out-Null
    }
    else {
        foreach ($item in $script:RunFailures) {
            $lines.Add(("    - {0}" -f (ConvertTo-YamlSafeString $item))) | Out-Null
        }
    }

    $lines.Add('drivers:') | Out-Null
    if ($script:DriverResults.Count -eq 0) {
        $lines.Add('  []') | Out-Null
    }
    else {
        foreach ($driver in $script:DriverResults) {
            $lines.Add('  -') | Out-Null
            $lines.Add(("    vendor: {0}" -f (ConvertTo-YamlSafeString $driver.Vendor))) | Out-Null
            $lines.Add(("    name: {0}" -f (ConvertTo-YamlSafeString $driver.Name))) | Out-Null
            $lines.Add(("    id: {0}" -f (ConvertTo-YamlSafeString $driver.Id))) | Out-Null
            $lines.Add(("    category: {0}" -f (ConvertTo-YamlSafeString $driver.Category))) | Out-Null
            $lines.Add(("    status: {0}" -f (ConvertTo-YamlSafeString $driver.Status))) | Out-Null
            $lines.Add(("    message: {0}" -f (ConvertTo-YamlSafeString $driver.Message))) | Out-Null
        }
    }

    Set-Content -LiteralPath $script:YamlPath -Value $lines -Encoding UTF8
    Write-Host ("YAML log written successfully: {0}" -f $script:YamlPath) -ForegroundColor Green
}

# -----------------------------
# File/folder helpers
# -----------------------------
function Ensure-Folder {
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -Path $Path -ItemType Directory -Force | Out-Null
    }
}

function Ensure-WorkingFolderPermissions {
    param([Parameter(Mandatory)][string]$Path)

    try {
        & icacls.exe $Path '/grant' '*S-1-1-0:(OI)(CI)F' '/T' '/C' | Out-Null
    }
    catch {
        Write-Log ("Unable to relax working folder permissions on {0}: {1}" -f $Path, $_.Exception.Message) 'WARN'
    }
}

function Remove-WorkingFolderRobust {
    param(
        [Parameter(Mandatory)][string]$Path,
        [int]$RetryCount = 6,
        [int]$RetryDelaySeconds = 5
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        Write-Log ("Working folder already absent: {0}" -f $Path) 'OK'
        return
    }

    Write-Log ("Attempting to remove working folder: {0}" -f $Path) 'INFO'
    for ($i = 1; $i -le $RetryCount; $i++) {
        try {
            Remove-Item -LiteralPath $Path -Recurse -Force -ErrorAction Stop
            Write-Log 'Working folder removed successfully.' 'OK'
            return
        }
        catch {
            if ($i -eq $RetryCount) {
                Add-RunFailure ("Failed to remove working folder after retries: {0}" -f $_.Exception.Message)
                return
            }
            Start-Sleep -Seconds $RetryDelaySeconds
        }
    }
}

# -----------------------------
# Vendor detection
# -----------------------------
function Get-SystemManufacturer {
    try {
        return (Get-CimInstance Win32_ComputerSystem -ErrorAction Stop).Manufacturer
    }
    catch {
        throw "Unable to determine system manufacturer. $($_.Exception.Message)"
    }
}

function Get-DriverVendor {
    $manufacturer = Get-SystemManufacturer
    Write-Log ("Detected manufacturer: {0}" -f $manufacturer) 'INFO'

    if ($manufacturer -match 'Dell') { return 'Dell' }
    if ($manufacturer -match 'HP|Hewlett-Packard') { return 'HP' }

    throw "Unsupported manufacturer for this script: $manufacturer"
}

# -----------------------------
# HP Support - HP Image Assistant extracted-folder deployment
# -----------------------------
function Get-HPSystemModelInfo {
    [CmdletBinding()]
    param()

    $cs = Get-CimInstance Win32_ComputerSystem -ErrorAction Stop
    $csp = Get-CimInstance Win32_ComputerSystemProduct -ErrorAction SilentlyContinue
    $bb = Get-CimInstance Win32_BaseBoard -ErrorAction SilentlyContinue

    $platform = $null
    if ($bb -and $bb.Product) {
        $platform = $bb.Product.ToString().Trim().ToUpper()
        if ($platform.Length -gt 4) {
            $platform = $platform.Substring(0,4)
        }
    }

    $model = if ($cs.Model) { $cs.Model.ToString().Trim() } else { 'Unknown' }
    $sku = if ($csp -and $csp.Version) { $csp.Version.ToString().Trim() } else { 'Unknown' }

    $info = [pscustomobject]@{
        Manufacturer = $cs.Manufacturer
        Model        = $model
        SKU          = $sku
        Platform     = $platform
    }

    Write-Log ("Detected HP system model: {0}" -f $info.Model) 'INFO'
    Write-Log ("Detected HP platform/baseboard ID: {0}" -f $info.Platform) 'INFO'
    Add-YamlAction ("Detected HP system model: {0}" -f $info.Model)
    Add-YamlAction ("Detected HP platform/baseboard ID: {0}" -f $info.Platform)

    return $info
}

function Get-ExistingHPIAExecutable {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$PreferredFolder)

    $candidateFolders = New-Object System.Collections.Generic.List[string]

    if (-not [string]::IsNullOrWhiteSpace($PreferredFolder)) {
        $candidateFolders.Add($PreferredFolder) | Out-Null
    }

    foreach ($folder in @(
        'C:\Program Files\HP\HP Image Assistant',
        'C:\Program Files (x86)\HP\HP Image Assistant',
        'C:\SWSetup\HPImageAssistant',
        'C:\ProgramData\Compton\HPImageAssistant'
    )) {
        if (-not $candidateFolders.Contains($folder)) {
            $candidateFolders.Add($folder) | Out-Null
        }
    }

    foreach ($folder in $candidateFolders) {
        if ([string]::IsNullOrWhiteSpace($folder) -or -not (Test-Path -LiteralPath $folder)) {
            continue
        }

        $exe = Get-ChildItem -Path $folder -Recurse -Filter 'HPImageAssistant.exe' -File -ErrorAction SilentlyContinue |
            Sort-Object FullName |
            Select-Object -First 1

        if ($exe) {
            return $exe.FullName
        }
    }

    return $null
}

function Install-HPIAFromExtractedFolder {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SourceFolder,
        [Parameter(Mandatory)][string]$DestinationFolder
    )

    Write-Section 'HP Image Assistant Local Deployment'
    Write-Log ("HPIA source folder: {0}" -f $SourceFolder) 'INFO'
    Write-Log ("HPIA local folder: {0}" -f $DestinationFolder) 'INFO'
    Add-YamlAction ("HPIA source folder: {0}" -f $SourceFolder)
    Add-YamlAction ("HPIA local folder: {0}" -f $DestinationFolder)

    $existingExe = Get-ExistingHPIAExecutable -PreferredFolder $DestinationFolder
    if ($existingExe) {
        try {
            $existingVersion = (Get-Item -LiteralPath $existingExe -ErrorAction Stop).VersionInfo.FileVersion
            Write-Log ("HP Image Assistant is already installed/found at: {0} (Version: {1})" -f $existingExe, $existingVersion) 'OK'
            Add-YamlAction ("Skipped HPIA local deployment because HPImageAssistant.exe already exists: {0} (Version: {1})" -f $existingExe, $existingVersion)
        }
        catch {
            Write-Log ("HP Image Assistant is already installed/found at: {0}" -f $existingExe) 'OK'
            Add-YamlAction ("Skipped HPIA local deployment because HPImageAssistant.exe already exists: {0}" -f $existingExe)
        }

        return $existingExe
    }

    Write-Log 'HP Image Assistant was not found locally. Deploying from extracted source folder.' 'INFO'
    Add-YamlAction 'HP Image Assistant was not found locally. Deploying from extracted source folder.'

    if (-not (Test-Path -LiteralPath $SourceFolder)) {
        throw "HPIA source folder not found: $SourceFolder"
    }

    $sourceFiles = @(Get-ChildItem -Path $SourceFolder -Recurse -File -ErrorAction SilentlyContinue)
    Write-Log ("Source HPIA file count: {0}" -f $sourceFiles.Count) 'INFO'
    Add-YamlAction ("Source HPIA file count: {0}" -f $sourceFiles.Count)

    $sourceExe = Get-ChildItem -Path $SourceFolder -Recurse -Filter 'HPImageAssistant.exe' -File -ErrorAction SilentlyContinue |
        Sort-Object FullName |
        Select-Object -First 1

    if (-not $sourceExe) {
        throw "HPImageAssistant.exe was not found anywhere under source folder: $SourceFolder"
    }

    Write-Log ("Found source HPImageAssistant.exe: {0}" -f $sourceExe.FullName) 'OK'
    Add-YamlAction ("Found source HPImageAssistant.exe: {0}" -f $sourceExe.FullName)

    try {
        if (Test-Path -LiteralPath $DestinationFolder) {
            Write-Log ("Removing existing local HPIA folder before clean deployment: {0}" -f $DestinationFolder) 'INFO'
            Remove-Item -LiteralPath $DestinationFolder -Recurse -Force -ErrorAction Stop
        }

        New-Item -Path $DestinationFolder -ItemType Directory -Force | Out-Null

        Write-Log 'Copying extracted HPIA files locally with robocopy...' 'INFO'

        $roboLog = Join-Path $DestinationFolder 'HPIA_robocopy.log'
        $roboArgs = @(
            ('"{0}"' -f $SourceFolder),
            ('"{0}"' -f $DestinationFolder),
            '/E',
            '/COPY:DAT',
            '/R:3',
            '/W:5',
            '/NFL',
            '/NDL',
            '/NP',
            ('/LOG:"{0}"' -f $roboLog)
        )

        $robo = Start-Process -FilePath "$env:SystemRoot\System32\robocopy.exe" -ArgumentList ($roboArgs -join ' ') -Wait -PassThru -NoNewWindow

        # Robocopy exit codes 0-7 are success/non-fatal. 8+ indicates failure.
        if ($robo.ExitCode -ge 8) {
            throw "Robocopy failed copying HPIA files. Exit code: $($robo.ExitCode). Log: $roboLog"
        }

        Write-Log ("HPIA files copied locally with robocopy exit code {0}." -f $robo.ExitCode) 'OK'
        Add-YamlAction ("HPIA files copied locally with robocopy exit code {0}." -f $robo.ExitCode)

        try {
            Get-ChildItem -Path $DestinationFolder -Recurse -File -ErrorAction SilentlyContinue |
                ForEach-Object { Unblock-File -Path $_.FullName -ErrorAction SilentlyContinue }
        }
        catch {}

        $copiedFiles = @(Get-ChildItem -Path $DestinationFolder -Recurse -File -ErrorAction SilentlyContinue)
        Write-Log ("Local HPIA folder file count after copy: {0}" -f $copiedFiles.Count) 'INFO'
        Add-YamlAction ("Local HPIA folder file count after copy: {0}" -f $copiedFiles.Count)

        $localExeItem = Get-ChildItem -Path $DestinationFolder -Recurse -Filter 'HPImageAssistant.exe' -File -ErrorAction SilentlyContinue |
            Sort-Object FullName |
            Select-Object -First 1

        if (-not $localExeItem) {
            $sampleFiles = $copiedFiles | Select-Object -First 20 | ForEach-Object { $_.FullName }
            foreach ($sample in $sampleFiles) {
                Write-Log ("Local HPIA sample file: {0}" -f $sample) 'WARN'
            }

            throw "HPImageAssistant.exe was not found anywhere under local copy folder: $DestinationFolder"
        }

        $localExe = $localExeItem.FullName

        Write-Log ("Resolved local HPImageAssistant.exe location: {0}" -f $localExe) 'OK'
        Add-YamlAction ("Resolved local HPImageAssistant.exe location: {0}" -f $localExe)

        return $localExe
    }
    catch {
        throw "Failed to deploy HP Image Assistant locally: $($_.Exception.Message)"
    }
}

function Get-HpiaExitStatus {
    param([int]$ExitCode)

    switch ($ExitCode) {
        0    { return 'success' }
        1    { return 'failed' }
        2    { return 'cancelled' }
        3    { return 'needs_reboot' }
        256  { return 'no_recommendations_or_success' }
        257  { return 'recommendations_found' }
        3010 { return 'needs_reboot' }
        3011 { return 'not_auto_installable_skipped' }
        4096 { return 'no_applicable_updates_or_platform_not_supported' }
        4097 { return 'invalid_parameters' }
        8193 { return 'hpia_analysis_or_report_generation_error' }
        default { return 'unknown' }
    }
}

function Get-HpiaRecommendationObjects {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ReportFolder
    )

    # Use a flexible PowerShell array instead of a strongly typed generic list.
    # HPIA reports can contain mixed object types from JSON and XML parsing.
    $recommendations = @()

    $jsonFiles = @(Get-ChildItem -Path $ReportFolder -Filter '*.json' -Recurse -File -ErrorAction SilentlyContinue)
    foreach ($jsonFile in $jsonFiles) {
        try {
            $json = Get-Content -LiteralPath $jsonFile.FullName -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop

            if ($json.HPIA -and $json.HPIA.Recommendations) {
                foreach ($rec in @($json.HPIA.Recommendations)) {
                    $recommendations += $rec
                    try {
                        Write-Log ("HPIA JSON recommendation type: {0}" -f $rec.GetType().FullName) 'INFO'
                    }
                    catch {}
                }
            }
            elseif ($json.Recommendations) {
                foreach ($rec in @($json.Recommendations)) {
                    $recommendations += $rec
                    try {
                        Write-Log ("HPIA JSON recommendation type: {0}" -f $rec.GetType().FullName) 'INFO'
                    }
                    catch {}
                }
            }
        }
        catch {
            Write-Log ("Unable to parse HPIA JSON report {0}: {1}" -f $jsonFile.FullName, $_.Exception.Message) 'WARN'
        }
    }

    $xmlFiles = @(Get-ChildItem -Path $ReportFolder -Filter '*.xml' -Recurse -File -ErrorAction SilentlyContinue)
    foreach ($xmlFile in $xmlFiles) {
        try {
            [xml]$xml = Get-Content -LiteralPath $xmlFile.FullName -Raw -ErrorAction Stop

            $nodes = @($xml.SelectNodes('//*[contains(translate(local-name(), "ABCDEFGHIJKLMNOPQRSTUVWXYZ", "abcdefghijklmnopqrstuvwxyz"), "recommend")]'))
            foreach ($node in $nodes) {
                $recommendations += $node
                try {
                    Write-Log ("HPIA XML recommendation node type: {0}" -f $node.GetType().FullName) 'INFO'
                }
                catch {}
            }
        }
        catch {
            Write-Log ("Unable to parse HPIA XML report {0}: {1}" -f $xmlFile.FullName, $_.Exception.Message) 'WARN'
        }
    }

    return @($recommendations)
}

function Get-HpiaRecommendationValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Recommendation,
        [Parameter(Mandatory)][string[]]$PropertyNames
    )

    foreach ($prop in $PropertyNames) {
        try {
            if ($Recommendation.PSObject.Properties.Name -contains $prop) {
                $value = $Recommendation.$prop
                if ($null -ne $value -and -not [string]::IsNullOrWhiteSpace($value.ToString())) {
                    return $value.ToString()
                }
            }
        }
        catch {}
    }

    # XML fallback
    try {
        foreach ($prop in $PropertyNames) {
            $node = $Recommendation.SelectSingleNode('.//*[local-name()="' + $prop + '"]')
            if ($node -and -not [string]::IsNullOrWhiteSpace($node.InnerText)) {
                return $node.InnerText.Trim()
            }
        }
    }
    catch {}

    return $null
}

function Get-HpiaSoftPaqNumber {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Recommendation)

    $candidate = Get-HpiaRecommendationValue -Recommendation $Recommendation -PropertyNames @(
        'SoftPaqId','SoftpaqId','SoftPaq','Softpaq','SoftPaqNumber','SoftpaqNumber','SP','Id','ID','Number'
    )

    if ($candidate -match '(?i)sp?(\d{5,6})') {
        return $matches[1]
    }

    $text = ($Recommendation | Out-String)
    if ($text -match '(?i)sp(\d{5,6})') {
        return $matches[1]
    }

    if ($text -match '\b(\d{5,6})\b') {
        return $matches[1]
    }

    return $null
}

function Get-HpiaRecommendationCategory {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Recommendation)

    return (Get-HpiaRecommendationValue -Recommendation $Recommendation -PropertyNames @(
        'Category','Type','RecommendationType','ComponentType','Class','Group'
    ))
}

function Get-HpiaRecommendationName {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Recommendation)

    $name = Get-HpiaRecommendationValue -Recommendation $Recommendation -PropertyNames @(
        'Name','Title','Component','ComponentName','Description','SoftPaqName','SoftpaqName'
    )

    if ($name) { return $name }

    $text = ($Recommendation | Out-String).Trim()
    if ($text.Length -gt 160) {
        return $text.Substring(0,160)
    }

    return $text
}

function New-HpiaDriverSPList {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ReportFolder,
        [Parameter(Mandatory)][string]$SPListPath
    )

    $recommendations = @(Get-HpiaRecommendationObjects -ReportFolder $ReportFolder)
    Write-Log ("HPIA recommendations parsed from reports: {0}" -f $recommendations.Count) 'INFO'
    Add-YamlAction ("HPIA recommendations parsed from reports: {0}" -f $recommendations.Count)

    $selected = @()
    $seen = @{}

    foreach ($rec in $recommendations) {
        $category = Get-HpiaRecommendationCategory -Recommendation $rec
        $name = Get-HpiaRecommendationName -Recommendation $rec
        $sp = Get-HpiaSoftPaqNumber -Recommendation $rec

        if (-not $sp) {
            continue
        }

        # Exclude BIOS/Firmware explicitly.
        $combined = ("{0} {1}" -f $category, $name)
        if ($combined -match '(?i)\bBIOS\b|Firmware') {
            $script:SkippedList.Add(("SP{0} {1}" -f $sp, $name)) | Out-Null
            Add-DriverResult -Vendor 'HP' -Name $name -Id $sp -Category $category -Status 'Blocked' -Message 'Excluded because it appears to be BIOS/Firmware.'
            continue
        }

        # Prefer driver-like recommendations, but allow blank category if the SoftPaq was recommended and not BIOS/Firmware.
        if ($combined -notmatch '(?i)Driver|Bluetooth|Chipset|Audio|Graphics|Video|LAN|WLAN|Wireless|NIC|Touch|Fingerprint|Card Reader|Storage|Serial|USB|Management Engine' -and $category) {
            $script:SkippedList.Add(("SP{0} {1}" -f $sp, $name)) | Out-Null
            Add-DriverResult -Vendor 'HP' -Name $name -Id $sp -Category $category -Status 'Skipped' -Message 'Excluded because it did not appear to be a driver recommendation.'
            continue
        }

        if (-not $seen.ContainsKey($sp)) {
            $seen[$sp] = $true
            $selected += $sp
            Add-DriverResult -Vendor 'HP' -Name $name -Id $sp -Category $category -Status 'Detected' -Message 'Selected for HPIA SPList install.'
        }
    }

    if (@($selected).Count -gt 0) {
        Set-Content -LiteralPath $SPListPath -Value $selected -Encoding ASCII
        Write-Log ("Created filtered HPIA SPList with {0} SoftPaqs: {1}" -f @($selected).Count, $SPListPath) 'OK'
        Add-YamlAction ("Created filtered HPIA SPList with {0} SoftPaqs: {1}" -f @($selected).Count, $SPListPath)
    }
    else {
        Write-Log 'No non-BIOS/Firmware driver SoftPaq recommendations were selected from HPIA reports.' 'OK'
        Add-YamlAction 'No non-BIOS/Firmware driver SoftPaq recommendations were selected from HPIA reports.'
    }

    return @($selected)
}


function Invoke-HPDriverUpdates {
    Write-Section 'HP Driver Analysis and Installation'

    $hpInfo = Get-HPSystemModelInfo

    $hpiaExe = Install-HPIAFromExtractedFolder -SourceFolder $HpiaSourceFolder -DestinationFolder $LocalHpiaFolder

    $hpiaReportFolder = Join-Path $YamlLogFolder 'HPIA'
    $hpiaDownloadFolder = Join-Path $WorkingRoot 'HPIADownloads'
    $hpiaExtractFolder = Join-Path $WorkingRoot 'HPIAExtracted'

    Ensure-Folder -Path $hpiaReportFolder
    Ensure-Folder -Path $hpiaDownloadFolder
    Ensure-Folder -Path $hpiaExtractFolder
    Ensure-Folder -Path 'C:\Temp'

    # HPIA is more reliable under Task Scheduler/SYSTEM when TEMP/TMP are local and TLS 1.2 is forced.
    $env:TEMP = 'C:\Temp'
    $env:TMP  = 'C:\Temp'

    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Write-Log 'TLS 1.2 enabled for HP Image Assistant network operations.' 'INFO'
        Add-YamlAction 'TLS 1.2 enabled for HPIA network operations.'
    }
    catch {
        Write-Log ("Unable to force TLS 1.2 before HPIA run: {0}" -f $_.Exception.Message) 'WARN'
        Add-YamlAction ("Unable to force TLS 1.2 before HPIA run: {0}" -f $_.Exception.Message)
    }

    # Production unattended HPIA workflow:
    # - Avoids /Action:List because it has been returning 8193 on this platform.
    # - Uses /Category:Drivers to avoid BIOS/Firmware/Dock/Thunderbolt firmware during unattended runs.
    # - Uses /Noninteractive, /AutoCleanup, and /Debug for scheduled task reliability and better logs.
    Write-Log 'Running HP Image Assistant production analyze/install pass for drivers only...' 'INFO'
    Add-YamlAction 'Running HPIA production analyze/install pass for drivers only.'

    $hpiaArgs = @(
        '/Operation:Analyze',
        '/Action:Install',
        '/Category:Drivers',
        '/Selection:All',
        '/Silent',
        '/Noninteractive',
        '/AutoCleanup',
        '/Debug',
        "/ReportFolder:`"$hpiaReportFolder`"",
        "/SoftpaqDownloadFolder:`"$hpiaDownloadFolder`"",
        "/SoftpaqExtractFolder:`"$hpiaExtractFolder`""
    )

    Write-Log ("HPIA production command: {0} {1}" -f $hpiaExe, ($hpiaArgs -join ' ')) 'INFO'
    Add-YamlAction ("HPIA production command: {0} {1}" -f $hpiaExe, ($hpiaArgs -join ' '))

    Write-Progress -Activity 'HP Image Assistant' -Status ("Installing recommended drivers for {0}" -f $hpInfo.Model) -PercentComplete 50

    $hpiaProc = Start-Process -FilePath $hpiaExe -ArgumentList ($hpiaArgs -join ' ') -Wait -PassThru -NoNewWindow
    $exitCode = [int]$hpiaProc.ExitCode
    $status = Get-HpiaExitStatus -ExitCode $exitCode

    Write-Progress -Activity 'HP Image Assistant' -Completed

    Write-Log ("HP Image Assistant production pass completed with exit code {0} ({1})." -f $exitCode, $status) 'INFO'
    Add-YamlAction ("HP Image Assistant production pass completed with exit code {0} ({1})." -f $exitCode, $status)

    $downloadedFiles = @(Get-ChildItem -Path $hpiaDownloadFolder -File -Recurse -ErrorAction SilentlyContinue)
    foreach ($file in $downloadedFiles) {
        Add-DriverResult -Vendor 'HP' -Name $file.Name -Id $null -Category 'Driver' -Status 'Detected' -Message ("Downloaded/processed by HPIA: {0}" -f $file.FullName)
    }

    $reportFiles = @(Get-ChildItem -Path $hpiaReportFolder -File -Recurse -ErrorAction SilentlyContinue)
    foreach ($report in $reportFiles) {
        Add-YamlAction ("HPIA report generated: {0}" -f $report.FullName)
    }

    switch ($exitCode) {
        0 {
            Write-Log 'HPIA completed successfully.' 'OK'
            Add-YamlAction 'HPIA completed successfully.'
            return
        }

        256 {
            Write-Log 'HPIA completed successfully. No applicable driver recommendations were found or no action was required.' 'OK'
            Add-YamlAction 'HPIA completed successfully with no applicable driver recommendations or no action required.'
            return
        }

        257 {
            Write-Log 'HPIA completed and reported recommendations/driver actions.' 'OK'
            Add-YamlAction 'HPIA completed and reported recommendations/driver actions.'
            return
        }

        3010 {
            Write-Log 'HPIA completed successfully. Reboot required.' 'WARN'
            Add-YamlAction 'HPIA completed successfully and indicated reboot required.'
            return
        }

        3011 {
            Write-Log 'One or more HPIA items were not auto-installable and were skipped.' 'WARN'
            Add-YamlAction 'One or more HPIA items were not auto-installable and were skipped.'
            return
        }

        4096 {
            Write-Log 'HPIA completed but did not find applicable driver updates for this platform.' 'OK'
            Add-YamlAction 'HPIA completed but did not find applicable driver updates for this platform.'
            return
        }

        8193 {
            Write-Log 'HPIA returned 8193. Checking whether reports/logs were generated before failing the run.' 'WARN'
            Add-YamlAction 'HPIA returned 8193. Checking whether reports/logs were generated before failing the run.'

            $generatedReports = @(Get-ChildItem -Path $hpiaReportFolder -File -Recurse -ErrorAction SilentlyContinue)
            if ($generatedReports.Count -gt 0) {
                Write-Log ("HPIA generated {0} report/log file(s) despite exit code 8193. Continuing so the scheduled workflow does not hard fail." -f $generatedReports.Count) 'WARN'
                Add-YamlAction ("HPIA generated {0} report/log file(s) despite exit code 8193. Continuing without hard failure." -f $generatedReports.Count)
                return
            }

            throw "HPIA failed with exit code 8193 and produced no report/log files in $hpiaReportFolder."
        }

        default {
            throw "HP Image Assistant production pass failed or returned an unexpected exit code: $exitCode ($status). Review reports in $hpiaReportFolder."
        }
    }
}

# -----------------------------
# Dell support
# -----------------------------
function Wait-ForFile {
    param(
        [Parameter(Mandatory)][string]$Path,
        [int]$TimeoutSeconds = 30
    )

    $start = Get-Date
    while (-not (Test-Path -LiteralPath $Path)) {
        Start-Sleep -Seconds 1
        if (((Get-Date) - $start).TotalSeconds -ge $TimeoutSeconds) {
            return $false
        }
    }
    return $true
}

function Get-DcuReportXmlSafely {
    param(
        [Parameter(Mandatory)][string]$Path,
        [int]$RetryCount = 5,
        [int]$RetryDelaySeconds = 2
    )

    if (-not (Wait-ForFile -Path $Path -TimeoutSeconds 20)) {
        throw "Dell DCU report file was not found: $Path"
    }

    Start-Sleep -Seconds 3

    for ($i = 1; $i -le $RetryCount; $i++) {
        try {
            $fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
            try {
                $sr = New-Object System.IO.StreamReader($fs)
                try {
                    $content = $sr.ReadToEnd()
                }
                finally {
                    $sr.Dispose()
                }
            }
            finally {
                $fs.Dispose()
            }

            $xml = New-Object System.Xml.XmlDocument
            $xml.LoadXml($content)
            return $xml
        }
        catch {
            if ($i -eq $RetryCount) {
                throw
            }
            Start-Sleep -Seconds $RetryDelaySeconds
        }
    }
}

function Get-DellNodeText {
    param(
        [Parameter(Mandatory)][xml]$Node,
        [Parameter(Mandatory)][string[]]$Names
    )

    foreach ($name in $Names) {
        try {
            $xpath = './/*[local-name()="' + $name + '"]'
            $child = $Node.SelectSingleNode($xpath)
            if ($child -and -not [string]::IsNullOrWhiteSpace($child.InnerText)) {
                return $child.InnerText.Trim()
            }
        }
        catch {}
    }

    return $null
}

function Get-DellReportItems {
    param([Parameter(Mandatory)][xml]$Xml)

    $items = @()
    try {
        $xpath = '//*[local-name()="Update" or local-name()="Package" or local-name()="SoftwareComponent" or local-name()="component" or local-name()="Device"]'
        $nodes = $Xml.SelectNodes($xpath)
        foreach ($node in $nodes) {
            $name = Get-DellNodeText -Node $node -Names @('Name','Title','PackageName')
            $version = Get-DellNodeText -Node $node -Names @('Version','PackageVersion')
            $category = Get-DellNodeText -Node $node -Names @('Category','Type')
            $id = Get-DellNodeText -Node $node -Names @('Id','PackageId','ReleaseId')

            if ($name -or $id) {
                $items += [pscustomobject]@{
                    Id       = $id
                    Name     = $name
                    Version  = $version
                    Category = $category
                }
            }
        }
    }
    catch {}

    return @($items)
}

function Get-DellDCUService {
    $candidateNames = @(
        'DellClientManagementService',
        'DellCommandUpdate',
        'DellUpdateService'
    )

    foreach ($name in $candidateNames) {
        $svc = Get-Service -Name $name -ErrorAction SilentlyContinue
        if ($svc) { return $svc }
    }

    $svc = Get-Service -ErrorAction SilentlyContinue |
        Where-Object { $_.DisplayName -match 'Dell.*Client.*Management|Dell.*Command.*Update' } |
        Select-Object -First 1

    return $svc
}

function Ensure-DellDCUService {
    param(
        [int]$TimeoutSeconds = 30
    )

    Write-Log 'Validating Dell Client Management Service...' 'INFO'
    Add-YamlAction 'Validating Dell Client Management Service.'

    $service = Get-DellDCUService
    if (-not $service) {
        throw 'Dell Client Management Service was not found. Dell Command | Update may need to be repaired or reinstalled.'
    }

    Write-Log ("Dell service detected: {0} ({1})" -f $service.DisplayName, $service.Name) 'OK'

    try {
        $wmiService = Get-CimInstance -ClassName Win32_Service -Filter ("Name='{0}'" -f $service.Name) -ErrorAction Stop
        if ($wmiService.StartMode -eq 'Disabled') {
            Write-Log 'Dell service startup type is Disabled. Setting it to Manual.' 'WARN'
            Set-Service -Name $service.Name -StartupType Manual -ErrorAction Stop
        }
    }
    catch {
        Write-Log ("Could not validate/set Dell service startup type: {0}" -f $_.Exception.Message) 'WARN'
    }

    $service.Refresh()
    if ($service.Status -ne 'Running') {
        Write-Log 'Starting Dell Client Management Service...' 'INFO'
        try {
            Start-Service -Name $service.Name -ErrorAction Stop
        }
        catch {
            Write-Log ("Start-Service failed. Attempting sc.exe recovery start. Error: {0}" -f $_.Exception.Message) 'WARN'
            & sc.exe start $service.Name | Out-Null
        }
    }

    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    do {
        Start-Sleep -Seconds 2
        $service = Get-Service -Name $service.Name -ErrorAction SilentlyContinue
        if ($service -and $service.Status -eq 'Running') {
            Write-Log 'Dell Client Management Service is running.' 'OK'
            Add-YamlAction 'Dell Client Management Service is running.'
            return $true
        }
    } while ((Get-Date) -lt $deadline)

    throw 'Dell Client Management Service did not reach the Running state before timeout.'
}

function Invoke-DellDCUCommandWithRetry {
    param(
        [Parameter(Mandatory)][string]$DcuCli,
        [Parameter(Mandatory)][string]$Arguments,
        [Parameter(Mandatory)][string]$OperationName,
        [int[]]$AcceptableExitCodes = @(0),
        [int]$MaxAttempts = 2
    )

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        Write-Log ("Dell DCU {0} attempt {1} of {2}..." -f $OperationName, $attempt, $MaxAttempts) 'INFO'
        Add-YamlAction ("Dell DCU {0} attempt {1} of {2}." -f $OperationName, $attempt, $MaxAttempts)

        $proc = Start-Process -FilePath $DcuCli -ArgumentList $Arguments -Wait -PassThru -NoNewWindow
        $exitCode = [int]$proc.ExitCode
        Write-Log ("Dell DCU {0} exit code: {1}" -f $OperationName, $exitCode) 'INFO'

        if ($AcceptableExitCodes -contains $exitCode) {
            return $proc
        }

        if ($exitCode -eq 3000) {
            Write-Log 'Dell DCU returned 3000, which normally indicates the Dell Client Management Service stopped or crashed.' 'WARN'
            Add-YamlAction 'Dell DCU returned 3000; attempting Dell service recovery before retry.'
        }
        else {
            Write-Log ("Dell DCU {0} returned non-success exit code {1}." -f $OperationName, $exitCode) 'WARN'
        }

        if ($attempt -lt $MaxAttempts) {
            Ensure-DellDCUService | Out-Null
            Start-Sleep -Seconds 5
            continue
        }

        throw "Dell DCU $OperationName failed after $MaxAttempts attempt(s). Last exit code: $exitCode"
    }
}

function Invoke-DellDriverUpdates {
    Write-Section 'Dell Command Update Workflow'

    $dcuCli = Join-Path ${env:ProgramFiles} 'Dell\CommandUpdate\dcu-cli.exe'
    if (-not (Test-Path -LiteralPath $dcuCli)) {
        throw "Dell Command | Update CLI was not found: $dcuCli"
    }

    Write-Log ("Using Dell Command | Update CLI: {0}" -f $dcuCli) 'OK'
    Add-YamlAction 'Using Dell Command | Update CLI.'

    Ensure-DellDCUService | Out-Null

    $dcuScanLog  = Join-Path $WorkingRoot 'Dell-DCU-Scan.log'
    $dcuApplyLog = Join-Path $WorkingRoot 'Dell-DCU-Apply.log'
    $dcuReport   = Join-Path $WorkingRoot 'Dell-DCU-ApplicableUpdates.xml'

    Write-Log 'Dell DCU Configure...' 'INFO'
    Write-Progress -Activity 'Dell Driver Update Workflow' -Status 'Configuring Dell Command Update' -PercentComplete 10
    $configureArgs = "/configure -silent -scheduleAuto -lockSettings=disable"
    Invoke-DellDCUCommandWithRetry -DcuCli $dcuCli -Arguments $configureArgs -OperationName 'Configure' -MaxAttempts 2 | Out-Null

    Write-Log 'Dell DCU Scan...' 'INFO'
    Write-Progress -Activity 'Dell Driver Update Workflow' -Status 'Scanning for updates' -PercentComplete 35
    $scanArgs = "/scan -silent -outputLog=""$dcuScanLog"" -report=""$dcuReport"""
    Invoke-DellDCUCommandWithRetry -DcuCli $dcuCli -Arguments $scanArgs -OperationName 'Scan' -MaxAttempts 2 | Out-Null
    Write-Log 'Waiting for Dell Command | Update to finish writing the report...' 'INFO'

    try {
        $xml = Get-DcuReportXmlSafely -Path $dcuReport
        $items = Get-DellReportItems -Xml $xml

        if ($items.Count -gt 0) {
            Add-YamlAction ("Dell DCU report parsed successfully. Updates detected: {0}" -f $items.Count)

            $total = $items.Count
            $index = 0

            foreach ($item in $items) {
                $index++
                $percent = 35 + [math]::Floor(($index / $total) * 35)
                $label = if ($item.Name) { $item.Name } elseif ($item.Id) { $item.Id } else { 'Dell update' }

                Write-Progress -Activity 'Dell Driver Update Workflow' -Status ("Parsing report: {0}" -f $label) -PercentComplete $percent

                if ($item.Category -match 'BIOS|Firmware') {
                    $script:SkippedList.Add($label) | Out-Null
                    Add-DriverResult -Vendor 'Dell' -Name $item.Name -Id $item.Id -Category $item.Category -Status 'Blocked' -Message 'BIOS/Firmware update blocked by script policy.'
                    Write-Log ("Blocking Dell BIOS/Firmware update: {0}" -f $label) 'WARN'
                }
                else {
                    Add-DriverResult -Vendor 'Dell' -Name $item.Name -Id $item.Id -Category $item.Category -Status 'Detected' -Message 'Detected in Dell DCU report.'
                }
            }
        }
        else {
            Add-YamlAction 'Dell DCU report parsed but returned no identifiable updates.'
        }
    }
    catch {
        Write-Log ("Failed to parse Dell DCU report: {0}" -f $_.Exception.Message) 'WARN'
        Add-YamlAction ("Failed to parse Dell DCU report: {0}" -f $_.Exception.Message)
    }

    Write-Log 'Dell DCU ApplyUpdates...' 'INFO'
    Write-Progress -Activity 'Dell Driver Update Workflow' -Status 'Applying updates' -PercentComplete 85
    $applyArgs = "/applyUpdates -silent -reboot=disable -outputLog=""$dcuApplyLog"""
    Invoke-DellDCUCommandWithRetry -DcuCli $dcuCli -Arguments $applyArgs -OperationName 'ApplyUpdates' -MaxAttempts 2 | Out-Null

    Write-Log (("Dell DCU logs: {0} ; {1} ; {2}") -f $dcuScanLog, $dcuApplyLog, $dcuReport) 'OK'
    Write-Progress -Activity 'Dell Driver Update Workflow' -Completed
}
# -----------------------------
# Main
# -----------------------------
$finalStatus = 'success'

try {
    Write-Section 'Initialization'

    Ensure-Folder -Path $YamlLogFolder
    Ensure-Folder -Path $WorkingRoot
    Ensure-WorkingFolderPermissions -Path $WorkingRoot

    $yamlName = "{0}-{1}-{2}.yml" -f $script:ComputerName, '05_Weekend_Vendor_Drivers_Update', (Get-Date -Format 'yyyy-MM-dd_HHmmss')
    $yamlPath = Join-Path $YamlLogFolder $yamlName

    Write-Log ("YAML log will be written to: {0}" -f $yamlPath) 'INFO'
    Initialize-YamlLog -ComputerName $script:ComputerName -YamlPath $yamlPath

    Write-Log 'Initializing vendor driver update script...' 'INFO'


    Write-Section 'Vendor Detection'
    $script:DetectedVendor = Get-DriverVendor
    Write-Log ("Detected vendor workflow: {0}" -f $script:DetectedVendor) 'INFO'
    Write-Log ("Working root: {0}" -f $WorkingRoot) 'INFO'

    if ($script:DetectedVendor -eq 'HP') {
        Invoke-HPDriverUpdates
    }
    elseif ($script:DetectedVendor -eq 'Dell') {
        Invoke-DellDriverUpdates
    }

}
catch {
    $finalStatus = 'failed'
    Add-RunFailure ("Script failed: {0}" -f $_.Exception.Message)
}
finally {
    Write-Section 'Cleanup'
    Remove-WorkingFolderRobust -Path $WorkingRoot

    if ($script:RunFailures.Count -gt 0 -and $finalStatus -ne 'failed') {
        $finalStatus = 'completed_with_warnings'
    }

    if ($script:RunFailures.Count -gt 0) {
        Write-Log (("{0} driver update completed with one or more failures.") -f $script:DetectedVendor) 'WARN'
    }
    else {
        Write-Log (("{0} driver update script completed successfully.") -f $script:DetectedVendor) 'OK'
    }

    Save-YamlLog -Status $finalStatus
}
