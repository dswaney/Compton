# =====================================================================
# ScriptName: 05_Weekend_HP_Drivers_Update.ps1
# ScriptVersion: 1.4
# LastUpdated: 2026-04-07
# =====================================================================

[CmdletBinding()]
param(
    [switch]$IncludeBIOS = $false, # Ignored in this hardened version; BIOS updates are blocked
    [switch]$IncludeSoftware = $false,
    [switch]$SuspendBitLockerForBIOS = $false, # Ignored in this hardened version; BIOS updates are blocked
    [string]$WorkingRoot = 'C:\Temp\HPDrivers',
    [string]$LogPath = 'C:\Temp\HPDrivers\HP-Driver-Update.log',
    [string]$YamlLogFolder = 'C:\Logs',
    [int]$CleanupRetryCount = 12,
    [int]$CleanupRetryDelaySeconds = 10
)

$ErrorActionPreference = 'Stop'

$script:RunStart = Get-Date
$script:ComputerName = $env:COMPUTERNAME
$script:YamlLogPath = $null
$script:OverallResult = 'Unknown'
$script:FailureMessage = $null
$script:CleanupResult = $false
$script:DetectedSoftPaqs = New-Object System.Collections.Generic.List[object]
$script:InstalledSoftPaqResults = New-Object System.Collections.Generic.List[object]

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
        if (-not [string]::IsNullOrWhiteSpace($logDir) -and -not (Test-Path -LiteralPath $logDir)) {
            New-Item -Path $logDir -ItemType Directory -Force | Out-Null
        }
        Add-Content -Path $LogPath -Value $line -Encoding UTF8
    }
    catch {
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

function Test-IsHPSystem {
    try {
        $manufacturer = (Get-CimInstance -ClassName Win32_ComputerSystem).Manufacturer
        return ($manufacturer -match 'HP|Hewlett-Packard')
    }
    catch {
        return $false
    }
}

function Ensure-Folder {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -Path $Path -ItemType Directory -Force | Out-Null
    }
}

function ConvertTo-YamlSafeValue {
    param(
        [AllowNull()]
        [object]$Value
    )

    if ($null -eq $Value) {
        return 'null'
    }

    $text = [string]$Value
    $text = $text -replace "`r", ' '
    $text = $text -replace "`n", ' '
    $text = $text -replace '"', '\"'
    return '"' + $text + '"'
}

function Initialize-YamlLog {
    Ensure-Folder -Path $YamlLogFolder

    $timestamp = Get-Date -Format 'yyyy-MM-dd_HHmmss'
    $fileName = "$($script:ComputerName)-HPDrivers-$timestamp.yml"
    $script:YamlLogPath = Join-Path $YamlLogFolder $fileName

    Write-Log "YAML log will be written to: $($script:YamlLogPath)" 'INFO'
}

function Add-DetectedSoftPaq {
    param(
        [string]$Id,
        [string]$Name,
        [string]$Version,
        [string]$Category
    )

    $script:DetectedSoftPaqs.Add([PSCustomObject]@{
        Id       = $Id
        Name     = $Name
        Version  = $Version
        Category = $Category
    }) | Out-Null
}

function Add-InstalledSoftPaqResult {
    param(
        [string]$Id,
        [string]$Name,
        [string]$Version,
        [string]$Status,
        [string]$Message
    )

    $script:InstalledSoftPaqResults.Add([PSCustomObject]@{
        Id      = $Id
        Name    = $Name
        Version = $Version
        Status  = $Status
        Message = $Message
    }) | Out-Null
}

function Write-YamlLog {
    try {
        if ([string]::IsNullOrWhiteSpace($script:YamlLogPath)) {
            Initialize-YamlLog
        }

        $runEnd = Get-Date
        $duration = [math]::Round(($runEnd - $script:RunStart).TotalSeconds, 0)

        $yamlLines = New-Object System.Collections.Generic.List[string]

        $yamlLines.Add('computer_name: ' + (ConvertTo-YamlSafeValue $script:ComputerName)) | Out-Null
        $yamlLines.Add('script_name: "05_Weekend_HP_Drivers_Update.ps1"') | Out-Null
        $yamlLines.Add('script_version: "1.4"') | Out-Null
        $yamlLines.Add('run_started: ' + (ConvertTo-YamlSafeValue ($script:RunStart.ToString('yyyy-MM-dd HH:mm:ss')))) | Out-Null
        $yamlLines.Add('run_finished: ' + (ConvertTo-YamlSafeValue ($runEnd.ToString('yyyy-MM-dd HH:mm:ss')))) | Out-Null
        $yamlLines.Add('duration_seconds: ' + $duration) | Out-Null

        $yamlLines.Add('options:') | Out-Null
        $yamlLines.Add('  include_bios: ' + ($(if ($IncludeBIOS) { 'true' } else { 'false' }))) | Out-Null
        $yamlLines.Add('  include_software: ' + ($(if ($IncludeSoftware) { 'true' } else { 'false' }))) | Out-Null
        $yamlLines.Add('  suspend_bitlocker_for_bios: ' + ($(if ($SuspendBitLockerForBIOS) { 'true' } else { 'false' }))) | Out-Null
        $yamlLines.Add('  working_root: ' + (ConvertTo-YamlSafeValue $WorkingRoot)) | Out-Null
        $yamlLines.Add('  cleanup_retry_count: ' + $CleanupRetryCount) | Out-Null
        $yamlLines.Add('  cleanup_retry_delay_seconds: ' + $CleanupRetryDelaySeconds) | Out-Null

        $yamlLines.Add('cleanup_successful: ' + ($(if ($script:CleanupResult) { 'true' } else { 'false' }))) | Out-Null
        $yamlLines.Add('overall_result: ' + (ConvertTo-YamlSafeValue $script:OverallResult)) | Out-Null

        if (-not [string]::IsNullOrWhiteSpace($script:FailureMessage)) {
            $yamlLines.Add('failure_message: ' + (ConvertTo-YamlSafeValue $script:FailureMessage)) | Out-Null
        }
        else {
            $yamlLines.Add('failure_message: null') | Out-Null
        }

        $yamlLines.Add('detected_softpaqs:') | Out-Null
        if ($script:DetectedSoftPaqs.Count -gt 0) {
            foreach ($item in $script:DetectedSoftPaqs) {
                $yamlLines.Add('  - id: ' + (ConvertTo-YamlSafeValue $item.Id)) | Out-Null
                $yamlLines.Add('    name: ' + (ConvertTo-YamlSafeValue $item.Name)) | Out-Null
                $yamlLines.Add('    version: ' + (ConvertTo-YamlSafeValue $item.Version)) | Out-Null
                $yamlLines.Add('    category: ' + (ConvertTo-YamlSafeValue $item.Category)) | Out-Null
            }
        }
        else {
            $yamlLines.Add('  - id: null') | Out-Null
            $yamlLines.Add('    name: "No applicable HP SoftPaq updates detected"') | Out-Null
            $yamlLines.Add('    version: null') | Out-Null
            $yamlLines.Add('    category: null') | Out-Null
        }

        $yamlLines.Add('install_results:') | Out-Null
        if ($script:InstalledSoftPaqResults.Count -gt 0) {
            foreach ($item in $script:InstalledSoftPaqResults) {
                $yamlLines.Add('  - id: ' + (ConvertTo-YamlSafeValue $item.Id)) | Out-Null
                $yamlLines.Add('    name: ' + (ConvertTo-YamlSafeValue $item.Name)) | Out-Null
                $yamlLines.Add('    version: ' + (ConvertTo-YamlSafeValue $item.Version)) | Out-Null
                $yamlLines.Add('    status: ' + (ConvertTo-YamlSafeValue $item.Status)) | Out-Null
                $yamlLines.Add('    message: ' + (ConvertTo-YamlSafeValue $item.Message)) | Out-Null
            }
        }
        else {
            $yamlLines.Add('  - id: null') | Out-Null
            $yamlLines.Add('    name: "No SoftPaq install operations were performed"') | Out-Null
            $yamlLines.Add('    version: null') | Out-Null
            $yamlLines.Add('    status: "None"') | Out-Null
            $yamlLines.Add('    message: "None"') | Out-Null
        }

        Set-Content -Path $script:YamlLogPath -Value $yamlLines -Encoding UTF8
        Write-Log "YAML log written successfully: $($script:YamlLogPath)" 'OK'
    }
    catch {
        Write-Log "Failed to write YAML log: $($_.Exception.Message)" 'WARN'
    }
}

function Save-PowerSettings {
    $settings = [ordered]@{}

    $settings.DisplayTimeoutDC = (
        Get-CimInstance -Namespace root\cimv2\power -Class Win32_PowerSettingDataIndex |
        Where-Object InstanceID -EQ 'Microsoft:PowerSettingDataIndex\{381b4222-f694-41f0-9685-ff5bb260df2e}\DC\{3c0bc021-c8a8-4e07-a973-6b14cbcb2b7e}'
    ).SettingIndexValue / 60

    $settings.DisplayTimeoutAC = (
        Get-CimInstance -Namespace root\cimv2\power -Class Win32_PowerSettingDataIndex |
        Where-Object InstanceID -EQ 'Microsoft:PowerSettingDataIndex\{381b4222-f694-41f0-9685-ff5bb260df2e}\AC\{3c0bc021-c8a8-4e07-a973-6b14cbcb2b7e}'
    ).SettingIndexValue / 60

    $settings.SleepTimeoutDC = (
        Get-CimInstance -Namespace root\cimv2\power -Class Win32_PowerSettingDataIndex |
        Where-Object InstanceID -EQ 'Microsoft:PowerSettingDataIndex\{381b4222-f694-41f0-9685-ff5bb260df2e}\DC\{29f6c1db-86da-48c5-9fdb-f2b67b1f44da}'
    ).SettingIndexValue / 60

    $settings.SleepTimeoutAC = (
        Get-CimInstance -Namespace root\cimv2\power -Class Win32_PowerSettingDataIndex |
        Where-Object InstanceID -EQ 'Microsoft:PowerSettingDataIndex\{381b4222-f694-41f0-9685-ff5bb260df2e}\AC\{29f6c1db-86da-48c5-9fdb-f2b67b1f44da}'
    ).SettingIndexValue / 60

    return [PSCustomObject]$settings
}

function Set-UnlimitedPowerTimeouts {
    Write-Log "Temporarily disabling monitor and sleep timeouts..." 'INFO'
    powercfg -change -monitor-timeout-dc 0 | Out-Null
    powercfg -change -monitor-timeout-ac 0 | Out-Null
    powercfg -change -standby-timeout-dc 0 | Out-Null
    powercfg -change -standby-timeout-ac 0 | Out-Null
}

function Restore-PowerSettings {
    param($Saved)

    if ($null -eq $Saved) { return }

    Write-Log "Restoring previous power timeout settings..." 'INFO'
    powercfg -change -monitor-timeout-dc $Saved.DisplayTimeoutDC | Out-Null
    powercfg -change -monitor-timeout-ac $Saved.DisplayTimeoutAC | Out-Null
    powercfg -change -standby-timeout-dc $Saved.SleepTimeoutDC | Out-Null
    powercfg -change -standby-timeout-ac $Saved.SleepTimeoutAC | Out-Null
}

function Ensure-Tls12 {
    try {
        [Net.ServicePointManager]::SecurityProtocol =
            [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
        Write-Log "Enabled TLS 1.2 for PowerShell Gallery access." 'OK'
    }
    catch {
        Write-Log "Could not explicitly enable TLS 1.2: $($_.Exception.Message)" 'WARN'
    }
}

function Ensure-NuGetProvider {
    Write-Log "Ensuring NuGet provider is installed..." 'INFO'
    try {
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope AllUsers | Out-Null
        Write-Log "NuGet provider is ready." 'OK'
    }
    catch {
        Write-Log "NuGet provider installation/check failed: $($_.Exception.Message)" 'WARN'
    }
}

function Ensure-PSGalleryTrusted {
    try {
        $repo = Get-PSRepository -Name 'PSGallery' -ErrorAction Stop
        if ($repo.InstallationPolicy -ne 'Trusted') {
            Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted
            Write-Log "Set PSGallery repository to Trusted." 'OK'
        }
        else {
            Write-Log "PSGallery repository already Trusted." 'INFO'
        }
    }
    catch {
        Write-Log "Could not validate/set PSGallery trust: $($_.Exception.Message)" 'WARN'
    }
}

function Install-ModuleIfPossible {
    param(
        [Parameter(Mandatory)][string]$Name,
        [switch]$AllowClobber
    )

    $args = @{
        Name        = $Name
        Scope       = 'AllUsers'
        Force       = $true
        ErrorAction = 'Stop'
    }

    if ($AllowClobber) {
        $args['AllowClobber'] = $true
    }

    Install-Module @args | Out-Null
}

function Ensure-PackageTooling {
    Write-Log "Ensuring PowerShell package tooling is current enough for HPCMSL..." 'INFO'

    Ensure-Tls12
    Ensure-NuGetProvider
    Ensure-PSGalleryTrusted

    try {
        Install-ModuleIfPossible -Name 'PowerShellGet' -AllowClobber
        Write-Log "PowerShellGet installed/updated." 'OK'
    }
    catch {
        Write-Log "PowerShellGet update failed: $($_.Exception.Message)" 'WARN'
    }

    try {
        Install-ModuleIfPossible -Name 'Microsoft.PowerShell.PSResourceGet'
        Write-Log "Microsoft.PowerShell.PSResourceGet installed/updated." 'OK'
    }
    catch {
        Write-Log "PSResourceGet install/update failed: $($_.Exception.Message)" 'WARN'
    }

    try {
        Import-Module PowerShellGet -Force -ErrorAction Stop
    }
    catch {
        Write-Log "Could not import PowerShellGet: $($_.Exception.Message)" 'WARN'
    }

    try {
        Import-Module Microsoft.PowerShell.PSResourceGet -Force -ErrorAction Stop
    }
    catch {
        Write-Log "Could not import PSResourceGet yet: $($_.Exception.Message)" 'WARN'
    }
}

function Ensure-HPCMSL {
    Write-Log "Ensuring HP CMSL is available..." 'INFO'

    Ensure-PackageTooling

    $moduleLoaded = $false

    if (-not (Get-Module -ListAvailable -Name HPCMSL)) {
        try {
            if (Get-Command -Name Install-PSResource -ErrorAction SilentlyContinue) {
                Write-Log "Installing HPCMSL with Install-PSResource..." 'INFO'
                Install-PSResource -Name HPCMSL -Scope AllUsers -TrustRepository -Quiet -AcceptLicense -ErrorAction Stop | Out-Null
            }
            else {
                Write-Log "Install-PSResource not available. Falling back to Install-Module for HPCMSL..." 'WARN'
                Install-ModuleIfPossible -Name 'HPCMSL' -AllowClobber
            }
        }
        catch {
            Write-Log "Primary HPCMSL install attempt failed: $($_.Exception.Message)" 'WARN'

            try {
                Write-Log "Trying fallback HPCMSL install with Install-Module..." 'INFO'
                Install-ModuleIfPossible -Name 'HPCMSL' -AllowClobber
            }
            catch {
                throw "Failed to install HPCMSL. $($_.Exception.Message)"
            }
        }
    }
    else {
        Write-Log "HPCMSL already present on system." 'INFO'
    }

    try {
        Import-Module HPCMSL -Force -ErrorAction Stop
        $moduleLoaded = $true
    }
    catch {
        try {
            Import-Module HP.Softpaq -Force -ErrorAction Stop
            $moduleLoaded = $true
        }
        catch {
            throw "HPCMSL/HP.Softpaq could not be imported after installation. $($_.Exception.Message)"
        }
    }

    if ($moduleLoaded) {
        Write-Log "HP CMSL imported successfully." 'OK'
    }
}

function Get-HPSoftpaqCategories {
    $categories = @('Driver')

    if ($IncludeBIOS) {
        Write-Log "IncludeBIOS was requested, but this script is hardened to block BIOS updates. BIOS updates will not be installed." 'WARN'
    }

    if ($IncludeSoftware) {
        $categories += @('Diagnostic', 'Dock', 'Software', 'Utility')
    }

    return $categories
}

function Get-DriverList {
    $categories = Get-HPSoftpaqCategories
    Write-Log "Querying HP SoftPaq list for categories: $($categories -join ', ')" 'INFO'

    $list = Get-SoftpaqList -Category $categories

    if (-not $list) {
        Write-Log "No applicable HP SoftPaq updates were returned." 'OK'
        return @()
    }

    foreach ($item in $list) {
        $category = $null
        if ($item.PSObject.Properties.Name -contains 'Category') {
            $category = [string]$item.Category
        }

        Add-DetectedSoftPaq -Id ([string]$item.Id) -Name ([string]$item.Name) -Version ([string]$item.Version) -Category $category
        Write-Log "Detected: [$($item.Id)] $($item.Name) Version $($item.Version)" 'INFO'
    }

    return @($list)
}


function Exclude-BiosAndFirmwareSoftpaqs {
    param([object[]]$Softpaqs)

    if (-not $Softpaqs) {
        return @()
    }

    $blockedPattern = '(?i)\b(BIOS|Firmware|UEFI|System\s+Firmware|Embedded\s+Controller|EC\s+Firmware|Thunderbolt\s+Firmware|Dock\s+Firmware)\b'
    $allowed = @()

    foreach ($item in $Softpaqs) {
        $category = ''
        $name = ''
        $id = ''
        $version = ''

        if ($null -ne $item) {
            if ($item.PSObject.Properties.Name -contains 'Category' -and $null -ne $item.Category) {
                $category = [string]$item.Category
            }
            if ($item.PSObject.Properties.Name -contains 'Name' -and $null -ne $item.Name) {
                $name = [string]$item.Name
            }
            if ($item.PSObject.Properties.Name -contains 'Id' -and $null -ne $item.Id) {
                $id = [string]$item.Id
            }
            if ($item.PSObject.Properties.Name -contains 'Version' -and $null -ne $item.Version) {
                $version = [string]$item.Version
            }
        }

        $isBlocked = $false
        if ($category -match '(?i)\bBIOS\b' -or $category -match '(?i)\bFirmware\b' -or $name -match $blockedPattern) {
            $isBlocked = $true
        }

        if ($isBlocked) {
            Write-Log "Blocking BIOS/Firmware SoftPaq: [$id] $name (Category: $category)" 'WARN'
            Add-InstalledSoftPaqResult -Id $id -Name $name -Version $version -Status 'Blocked' -Message 'Skipped because BIOS/Firmware updates are not allowed in this script'
        }
        else {
            $allowed += $item
        }
    }

    return $allowed
}

function Install-SoftpaqList {
    param([object[]]$Softpaqs)

    $failures = 0

    foreach ($item in $Softpaqs) {
        try {
            Write-Log "Installing SoftPaq [$($item.Id)] $($item.Name)..." 'INFO'
            Get-Softpaq -Number $item.Id -Action SilentInstall | Out-Null
            Write-Log "Installed SoftPaq [$($item.Id)] $($item.Name)." 'OK'
            Add-InstalledSoftPaqResult -Id ([string]$item.Id) -Name ([string]$item.Name) -Version ([string]$item.Version) -Status 'Succeeded' -Message 'Installed successfully'
        }
        catch {
            $msg = $_.Exception.Message
            Write-Log "Failed SoftPaq [$($item.Id)] $($item.Name): $msg" 'WARN'
            Add-InstalledSoftPaqResult -Id ([string]$item.Id) -Name ([string]$item.Name) -Version ([string]$item.Version) -Status 'Failed' -Message $msg
            $failures++
        }
    }

    return $failures
}

function Suspend-BitLockerIfNeeded {
    if ($SuspendBitLockerForBIOS) {
        Write-Log "SuspendBitLockerForBIOS was requested, but this script is hardened to block BIOS updates. BitLocker suspend will not be used." 'WARN'
    }
    return
}

function Remove-WorkingFolderRobust {
    param(
        [Parameter(Mandatory)][string]$Path,
        [int]$RetryCount = 12,
        [int]$RetryDelaySeconds = 10
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        Write-Log "Working folder already absent: $Path" 'OK'
        return $true
    }

    Write-Log "Attempting to remove working folder: $Path" 'INFO'

    for ($i = 1; $i -le $RetryCount; $i++) {
        try {
            Start-Sleep -Seconds 2
            Remove-Item -LiteralPath $Path -Recurse -Force -ErrorAction Stop

            if (-not (Test-Path -LiteralPath $Path)) {
                Write-Log "Working folder removed successfully." 'OK'
                return $true
            }
        }
        catch {
            Write-Log "Cleanup attempt $i/$RetryCount failed: $($_.Exception.Message)" 'WARN'
        }

        if ($i -lt $RetryCount) {
            Start-Sleep -Seconds $RetryDelaySeconds
        }
    }

    Write-Log "Working folder still exists after cleanup attempts: $Path" 'ERROR'
    return $false
}

# Main
if (-not (Test-IsAdministrator)) {
    Write-Error "Please run this script as Administrator."
    exit 1
}

Initialize-YamlLog

if (-not (Test-IsHPSystem)) {
    Write-Log "This is not an HP or Hewlett-Packard system. Skipping HP driver update." 'WARN'
    $script:OverallResult = 'SkippedNonHPSystem'
    Write-YamlLog
    exit 0
}

Ensure-Folder -Path $WorkingRoot

$SavedPower = $null
$OriginalLocation = (Get-Location).Path
$Failures = 0
$CleanupOk = $false

try {
    Write-Log "Initializing HP driver update script..." 'INFO'
    Write-Log "Working root: $WorkingRoot" 'INFO'

    $SavedPower = Save-PowerSettings
    Set-UnlimitedPowerTimeouts
    Ensure-HPCMSL
    Suspend-BitLockerIfNeeded

    Set-Location -Path $WorkingRoot

    $softpaqs = Get-DriverList
    $softpaqs = @(Exclude-BiosAndFirmwareSoftpaqs -Softpaqs $softpaqs)
    if ($softpaqs.Count -eq 0) {
        $CleanupOk = Remove-WorkingFolderRobust -Path $WorkingRoot -RetryCount $CleanupRetryCount -RetryDelaySeconds $CleanupRetryDelaySeconds
        $script:CleanupResult = $CleanupOk
        $script:OverallResult = if ($CleanupOk) { 'SucceededNoApplicableUpdates' } else { 'SucceededNoApplicableUpdatesCleanupIncomplete' }
        Write-YamlLog
        exit $(if ($CleanupOk) { 0 } else { 2 })
    }

    $Failures = Install-SoftpaqList -Softpaqs $softpaqs
}
catch {
    $script:FailureMessage = $_.Exception.Message
    Write-Log "Script failed: $($_.Exception.Message)" 'ERROR'
    $Failures++
}
finally {
    try {
        Set-Location -Path $OriginalLocation
    }
    catch {
    }

    try {
        Restore-PowerSettings -Saved $SavedPower
    }
    catch {
        Write-Log "Failed restoring power settings: $($_.Exception.Message)" 'WARN'
    }

    $CleanupOk = Remove-WorkingFolderRobust -Path $WorkingRoot -RetryCount $CleanupRetryCount -RetryDelaySeconds $CleanupRetryDelaySeconds
    $script:CleanupResult = $CleanupOk
}

if ($Failures -eq 0 -and $CleanupOk) {
    Write-Log "HP driver update script completed successfully." 'OK'
    $script:OverallResult = 'Succeeded'
    Write-YamlLog
    exit 0
}
elseif ($Failures -eq 0 -and -not $CleanupOk) {
    Write-Log "HP driver update succeeded, but cleanup was incomplete." 'WARN'
    $script:OverallResult = 'SucceededCleanupIncomplete'
    Write-YamlLog
    exit 2
}
else {
    Write-Log "HP driver update completed with one or more failures." 'WARN'
    if ([string]::IsNullOrWhiteSpace($script:FailureMessage)) {
        $script:FailureMessage = 'One or more SoftPaq installations failed.'
    }
    $script:OverallResult = 'Failed'
    Write-YamlLog
    exit 3
}
