# =====================================================================
# ScriptName: 06_Weekend_Windows_Updates.ps1
# ScriptVersion: 1.3
# LastUpdated: 2026-03-24
# Purpose: Installs Windows Updates using PSWindowsUpdate, writes a
#          standard log and a YAML audit log in C:\Logs, and explicitly
#          reboots if Windows reports that a reboot is required.
# =====================================================================

[CmdletBinding()]
param(
    [switch]$ResetWUComponentsFirst = $false,
    [int]$OperationTimeoutSeconds = 1800,
    [int]$RebootDelaySeconds = 30,
    [string]$LogPath = "$env:SystemDrive\Logs\Weekend-Windows-Updates.log",
    [string]$YamlLogFolder = "$env:SystemDrive\Logs"
)

$ErrorActionPreference = 'Stop'

$script:RunStart       = Get-Date
$script:ComputerName   = $env:COMPUTERNAME
$script:YamlLogPath    = $null
$script:UpdateEntries  = New-Object System.Collections.Generic.List[object]
$script:RawUpdateLines = New-Object System.Collections.Generic.List[string]
$script:OverallResult  = 'Unknown'
$script:RebootRequired = $false
$script:FailureMessage = $null

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
        Add-Content -Path $LogPath -Value $line
    }
    catch {
        # Do not fail script because logging failed
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

function Ensure-Folder {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

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
    $fileName = "$($script:ComputerName)-WindowsUpdates-$timestamp.yml"
    $script:YamlLogPath = Join-Path $YamlLogFolder $fileName

    Write-Log "YAML log will be written to: $($script:YamlLogPath)" 'INFO'
}

function Add-UpdateEntry {
    param(
        [string]$Title,
        [string]$KB,
        [string]$Size,
        [string]$Status,
        [string]$Result,
        [string]$Source = 'PSWindowsUpdate'
    )

    $entry = [PSCustomObject]@{
        Title  = $Title
        KB     = $KB
        Size   = $Size
        Status = $Status
        Result = $Result
        Source = $Source
    }

    $script:UpdateEntries.Add($entry) | Out-Null
}

function Try-ParseUpdateObject {
    param(
        [Parameter(Mandatory)]
        [object]$Item
    )

    if ($null -eq $Item) {
        return $false
    }

    $properties = $Item.PSObject.Properties.Name
    if (-not $properties -or $properties.Count -eq 0) {
        return $false
    }

    $title = $null
    foreach ($name in @('Title','KBArticleTitle')) {
        if ($properties -contains $name -and -not [string]::IsNullOrWhiteSpace([string]$Item.$name)) {
            $title = [string]$Item.$name
            break
        }
    }

    $kb = $null
    foreach ($name in @('KB','KBArticleIDs','KBArticleID')) {
        if ($properties -contains $name -and $null -ne $Item.$name) {
            $value = $Item.$name
            if ($value -is [System.Array]) {
                $kb = (($value | ForEach-Object { [string]$_ }) -join ', ')
            }
            else {
                $kb = [string]$value
            }
            if (-not [string]::IsNullOrWhiteSpace($kb)) {
                break
            }
        }
    }

    $size = $null
    foreach ($name in @('Size','MaxDownloadSize')) {
        if ($properties -contains $name -and $null -ne $Item.$name) {
            $size = [string]$Item.$name
            if (-not [string]::IsNullOrWhiteSpace($size)) {
                break
            }
        }
    }

    $status = $null
    foreach ($name in @('Status','Result','UpdateState')) {
        if ($properties -contains $name -and $null -ne $Item.$name) {
            $status = [string]$Item.$name
            if (-not [string]::IsNullOrWhiteSpace($status)) {
                break
            }
        }
    }

    $result = $null
    foreach ($name in @('Result','Status','HResult')) {
        if ($properties -contains $name -and $null -ne $Item.$name) {
            $result = [string]$Item.$name
            if (-not [string]::IsNullOrWhiteSpace($result)) {
                break
            }
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($title) -or -not [string]::IsNullOrWhiteSpace($kb)) {
        Add-UpdateEntry -Title $title -KB $kb -Size $size -Status $status -Result $result
        return $true
    }

    return $false
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
        $yamlLines.Add('script_name: "06_Weekend_Windows_Updates.ps1"') | Out-Null
        $yamlLines.Add('script_version: "1.3"') | Out-Null
        $yamlLines.Add('run_started: ' + (ConvertTo-YamlSafeValue ($script:RunStart.ToString('yyyy-MM-dd HH:mm:ss')))) | Out-Null
        $yamlLines.Add('run_finished: ' + (ConvertTo-YamlSafeValue ($runEnd.ToString('yyyy-MM-dd HH:mm:ss')))) | Out-Null
        $yamlLines.Add('duration_seconds: ' + $duration) | Out-Null
        $yamlLines.Add('reset_wu_components_first: ' + ($(if ($ResetWUComponentsFirst) { 'true' } else { 'false' }))) | Out-Null
        $yamlLines.Add('reboot_required: ' + ($(if ($script:RebootRequired) { 'true' } else { 'false' }))) | Out-Null
        $yamlLines.Add('overall_result: ' + (ConvertTo-YamlSafeValue $script:OverallResult)) | Out-Null

        if (-not [string]::IsNullOrWhiteSpace($script:FailureMessage)) {
            $yamlLines.Add('failure_message: ' + (ConvertTo-YamlSafeValue $script:FailureMessage)) | Out-Null
        } else {
            $yamlLines.Add('failure_message: null') | Out-Null
        }

        $yamlLines.Add('updates:') | Out-Null

        if ($script:UpdateEntries.Count -gt 0) {
            foreach ($entry in $script:UpdateEntries) {
                $yamlLines.Add('  - title: '  + (ConvertTo-YamlSafeValue $entry.Title))  | Out-Null
                $yamlLines.Add('    kb: '     + (ConvertTo-YamlSafeValue $entry.KB))     | Out-Null
                $yamlLines.Add('    size: '   + (ConvertTo-YamlSafeValue $entry.Size))   | Out-Null
                $yamlLines.Add('    status: ' + (ConvertTo-YamlSafeValue $entry.Status)) | Out-Null
                $yamlLines.Add('    result: ' + (ConvertTo-YamlSafeValue $entry.Result)) | Out-Null
                $yamlLines.Add('    source: ' + (ConvertTo-YamlSafeValue $entry.Source)) | Out-Null
            }
        }
        else {
            $yamlLines.Add('  - title: "No structured update entries captured"') | Out-Null
            $yamlLines.Add('    kb: null') | Out-Null
            $yamlLines.Add('    size: null') | Out-Null
            $yamlLines.Add('    status: "None"') | Out-Null
            $yamlLines.Add('    result: "None"') | Out-Null
            $yamlLines.Add('    source: "Script"') | Out-Null
        }

        $yamlLines.Add('raw_output:') | Out-Null
        if ($script:RawUpdateLines.Count -gt 0) {
            foreach ($line in $script:RawUpdateLines) {
                $yamlLines.Add('  - ' + (ConvertTo-YamlSafeValue $line)) | Out-Null
            }
        }
        else {
            $yamlLines.Add('  - "No raw update output captured"') | Out-Null
        }

        Set-Content -Path $script:YamlLogPath -Value $yamlLines -Encoding UTF8
        Write-Log "YAML log written successfully: $($script:YamlLogPath)" 'OK'
    }
    catch {
        Write-Log "Failed to write YAML log: $($_.Exception.Message)" 'WARN'
    }
}

function Ensure-PSWindowsUpdate {
    Write-Log "Ensuring PSWindowsUpdate module is available..." 'INFO'

    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    }
    catch {}

    if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {
        Write-Log "Installing NuGet package provider..." 'INFO'
        Install-PackageProvider -Name NuGet -Force -Scope AllUsers | Out-Null
    }

    try {
        $repo = Get-PSRepository -Name 'PSGallery' -ErrorAction Stop
        if ($repo.InstallationPolicy -ne 'Trusted') {
            Write-Log "Setting PSGallery as Trusted..." 'INFO'
            Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted
        }
    }
    catch {
        Write-Log "Could not validate PSGallery repository settings: $($_.Exception.Message)" 'WARN'
    }

    if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
        Write-Log "Installing PSWindowsUpdate module..." 'INFO'
        Install-Module -Name PSWindowsUpdate -Force -AllowClobber -Scope AllUsers
    }

    Import-Module PSWindowsUpdate -Force
    Write-Log "PSWindowsUpdate module imported successfully." 'OK'
}

function Reset-WUComponentsSafe {
    Write-Log "Resetting Windows Update components..." 'INFO'

    if (-not (Get-Command -Name Reset-WUComponents -ErrorAction SilentlyContinue)) {
        throw "Reset-WUComponents command was not found after importing PSWindowsUpdate."
    }

    Reset-WUComponents -Verbose *>&1 | ForEach-Object {
        Write-Log "$_" 'INFO'
    }

    Write-Log "Windows Update components reset complete." 'OK'
}

function Install-AvailableWindowsUpdates {
    Write-Log "Scanning for and installing available Windows Updates..." 'INFO'

    if (-not (Get-Command -Name Install-WindowsUpdate -ErrorAction SilentlyContinue)) {
        throw "Install-WindowsUpdate command was not found."
    }

    $results = Install-WindowsUpdate -MicrosoftUpdate -AcceptAll -IgnoreReboot -Verbose *>&1

    if ($results) {
        foreach ($item in $results) {
            $line = [string]$item
            if (-not [string]::IsNullOrWhiteSpace($line)) {
                $script:RawUpdateLines.Add($line) | Out-Null
                Write-Log $line 'INFO'
            }

            [void](Try-ParseUpdateObject -Item $item)
        }
    }

    Write-Log "Windows Update installation command completed." 'OK'
}

function Test-WURebootRequired {
    try {
        if (Get-Command -Name Get-WURebootStatus -ErrorAction SilentlyContinue) {
            $status = Get-WURebootStatus -Silent
            return [bool]$status
        }
    }
    catch {
        Write-Log "Get-WURebootStatus check failed: $($_.Exception.Message)" 'WARN'
    }

    try {
        $sysInfo = New-Object -ComObject Microsoft.Update.SystemInfo
        return [bool]$sysInfo.RebootRequired
    }
    catch {
        Write-Log "Microsoft.Update.SystemInfo reboot check failed: $($_.Exception.Message)" 'WARN'
    }

    return $false
}

function Invoke-ExplicitReboot {
    param(
        [int]$Delay = 30
    )

    $comment = 'Restarting to complete Windows Update installation.'

    $arguments = @(
        '/r'
        '/t', $Delay.ToString()
        '/d', 'p:2:17'
        '/c', "`"$comment`""
        '/f'
    )

    Write-Log "Issuing reboot command: shutdown.exe $($arguments -join ' ')" 'INFO'

    & "$env:SystemRoot\System32\shutdown.exe" @arguments

    $exitCode = $LASTEXITCODE
    if ($exitCode -ne 0) {
        throw "shutdown.exe returned exit code $exitCode"
    }

    Write-Log "Reboot command issued successfully." 'OK'
}

# Main
if (-not (Test-IsAdministrator)) {
    Write-Error "Please run this script as Administrator."
    exit 1
}

Initialize-YamlLog
Write-Log "Initializing weekend Windows update script..." 'INFO'

try {
    Ensure-PSWindowsUpdate

    if ($ResetWUComponentsFirst) {
        Reset-WUComponentsSafe
    }

    Install-AvailableWindowsUpdates

    $script:RebootRequired = Test-WURebootRequired

    if ($script:RebootRequired) {
        $script:OverallResult = 'SucceededWithRebootRequired'
        Write-Log "Windows reports that a reboot is required." 'OK'
        Write-YamlLog
        Invoke-ExplicitReboot -Delay $RebootDelaySeconds
        exit 3010
    }
    else {
        $script:OverallResult = 'Succeeded'
        Write-Log "No reboot is currently required." 'OK'
        Write-YamlLog
        exit 0
    }
}
catch {
    $script:OverallResult = 'Failed'
    $script:FailureMessage = $_.Exception.Message
    Write-Log "Script failed: $($_.Exception.Message)" 'ERROR'
    Write-YamlLog
    exit 2
}
