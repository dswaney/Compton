# ScriptVersion: 1.1
# LastUpdated: 2026-03-31

[CmdletBinding()]
param(
    [string]$DownloadFolder = "$env:SystemDrive\Temp",
    [string]$LogPath = "$env:SystemDrive\Temp\Update-Edge-Silent.log",
    [switch]$KillEdgeProcesses = $false,
    [switch]$ForceReinstall = $false
)

$ErrorActionPreference = 'Stop'

# Microsoft Edge Stable Enterprise x64 MSI (latest stable via Microsoft redirect)
$EdgeMsiUrl = 'https://go.microsoft.com/fwlink/?LinkID=2093437'
$EdgeMsiPath = Join-Path $DownloadFolder 'MicrosoftEdgeEnterpriseX64.msi'
$MsiLogPath  = Join-Path $DownloadFolder 'MicrosoftEdgeEnterpriseX64-msi.log'

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
        $identity  = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object Security.Principal.WindowsPrincipal($identity)
        return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    }
    catch {
        return $false
    }
}

function ConvertTo-Version {
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [string]$Value
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $null
    }

    $match = [regex]::Match($Value, '\d+(?:\.\d+){1,3}')
    if (-not $match.Success) {
        return $null
    }

    try {
        return [version]$match.Value
    }
    catch {
        return $null
    }
}

function Get-InstalledEdgeVersion {
    $registryCandidates = @(
        'HKLM:\SOFTWARE\Microsoft\Edge\BLBeacon',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Edge\BLBeacon',
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Microsoft Edge',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Microsoft Edge'
    )

    foreach ($key in $registryCandidates) {
        try {
            if (Test-Path -LiteralPath $key) {
                $item = Get-ItemProperty -LiteralPath $key -ErrorAction Stop
                foreach ($propertyName in @('version', 'pv', 'DisplayVersion')) {
                    $candidate = ConvertTo-Version -Value ($item.$propertyName)
                    if ($candidate) {
                        return $candidate
                    }
                }
            }
        }
        catch {
        }
    }

    $fileCandidates = @(
        "$env:ProgramFiles(x86)\Microsoft\Edge\Application\msedge.exe",
        "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
        "$env:ProgramFiles(x86)\Microsoft\Edge\Application\msedge.dll",
        "$env:ProgramFiles\Microsoft\Edge\Application\msedge.dll"
    )

    foreach ($path in $fileCandidates) {
        if (Test-Path -LiteralPath $path) {
            try {
                $item = Get-Item -LiteralPath $path -ErrorAction Stop
                $candidate = ConvertTo-Version -Value $item.VersionInfo.ProductVersion
                if ($candidate) {
                    return $candidate
                }
            }
            catch {
            }
        }
    }

    return $null
}

function Get-MsiProductVersion {
    param(
        [Parameter(Mandatory)][string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        return $null
    }

    $installer = $null
    $database = $null
    $view = $null
    $record = $null

    try {
        $installer = New-Object -ComObject WindowsInstaller.Installer
        $database = $installer.GetType().InvokeMember('OpenDatabase', 'InvokeMethod', $null, $installer, @($Path, 0))
        $view = $database.GetType().InvokeMember('OpenView', 'InvokeMethod', $null, $database, @("SELECT `Value` FROM `Property` WHERE `Property`='ProductVersion'"))
        $view.GetType().InvokeMember('Execute', 'InvokeMethod', $null, $view, $null) | Out-Null
        $record = $view.GetType().InvokeMember('Fetch', 'InvokeMethod', $null, $view, $null)

        if ($record) {
            $value = $record.GetType().InvokeMember('StringData', 'GetProperty', $null, $record, 1)
            return (ConvertTo-Version -Value $value)
        }
    }
    catch {
        Write-Log "Could not read ProductVersion from MSI '$Path': $($_.Exception.Message)" 'WARN'
    }
    finally {
        foreach ($comObject in @($record, $view, $database, $installer)) {
            if ($null -ne $comObject) {
                try {
                    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($comObject)
                }
                catch {
                }
            }
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }

    return $null
}

function Ensure-Folder {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -Path $Path -ItemType Directory -Force | Out-Null
    }
}

function Download-File {
    param(
        [Parameter(Mandatory)][string]$Url,
        [Parameter(Mandatory)][string]$Destination
    )

    Write-Log "Downloading latest Microsoft Edge MSI from Microsoft..." 'INFO'

    try {
        Invoke-WebRequest -Uri $Url -OutFile $Destination -UseBasicParsing
    }
    catch {
        throw "Failed to download Edge MSI. $($_.Exception.Message)"
    }

    if (-not (Test-Path -LiteralPath $Destination)) {
        throw "Download completed but MSI file was not found at $Destination"
    }

    $size = (Get-Item -LiteralPath $Destination).Length
    if ($size -le 0) {
        throw "Downloaded MSI file is empty."
    }

    Write-Log "Downloaded MSI to $Destination ($size bytes)." 'OK'
}

function Stop-EdgeProcesses {
    Write-Log "Stopping Microsoft Edge processes..." 'WARN'

    $names = @(
        'msedge',
        'MicrosoftEdgeUpdate',
        'setup'
    )

    foreach ($name in $names) {
        Get-Process -Name $name -ErrorAction SilentlyContinue | ForEach-Object {
            try {
                Stop-Process -Id $_.Id -Force -ErrorAction Stop
                Write-Log "Stopped process $($_.ProcessName) (PID $($_.Id))." 'INFO'
            }
            catch {
                Write-Log "Could not stop process $($_.ProcessName) (PID $($_.Id)): $($_.Exception.Message)" 'WARN'
            }
        }
    }
}

function Install-EdgeMsi {
    param(
        [Parameter(Mandatory)][string]$MsiPath,
        [Parameter(Mandatory)][string]$MsiLog
    )

    $arguments = @(
        '/i'
        "`"$MsiPath`""
        '/qn'
        '/norestart'
        '/L*v'
        "`"$MsiLog`""
    )

    if ($ForceReinstall) {
        $arguments += @('REINSTALL=ALL', 'REINSTALLMODE=vomus')
    }

    Write-Log "Installing Microsoft Edge silently..." 'INFO'
    Write-Log "Running: msiexec.exe $($arguments -join ' ')" 'INFO'

    $proc = Start-Process -FilePath 'msiexec.exe' -ArgumentList $arguments -Wait -PassThru -WindowStyle Hidden
    $exitCode = $proc.ExitCode

    switch ($exitCode) {
        0 {
            Write-Log "Microsoft Edge MSI install completed successfully." 'OK'
        }
        1641 {
            Write-Log "Microsoft Edge install succeeded and initiated a restart." 'WARN'
        }
        3010 {
            Write-Log "Microsoft Edge install succeeded and requires a reboot." 'WARN'
        }
        default {
            throw "msiexec.exe returned exit code $exitCode. Review $MsiLog"
        }
    }

    return $exitCode
}

if (-not (Test-IsAdministrator)) {
    Write-Error "This script must be run as Administrator."
    exit 1
}

Ensure-Folder -Path $DownloadFolder
Write-Log "Initializing Microsoft Edge silent update script..." 'INFO'

$beforeVersion = Get-InstalledEdgeVersion
if ($beforeVersion) {
    Write-Log "Installed Edge version before update: $beforeVersion" 'INFO'
}
else {
    Write-Log "Microsoft Edge does not appear to be installed, or version could not be detected." 'WARN'
}

if ($KillEdgeProcesses) {
    Stop-EdgeProcesses
}

try {
    Download-File -Url $EdgeMsiUrl -Destination $EdgeMsiPath
}
catch {
    Write-Log "Script failed during download: $($_.Exception.Message)" 'ERROR'
    exit 2
}

$downloadedVersion = Get-MsiProductVersion -Path $EdgeMsiPath
if ($downloadedVersion) {
    Write-Log "Downloaded Edge MSI version: $downloadedVersion" 'INFO'
}
else {
    Write-Log "Could not determine downloaded Edge MSI version before install." 'WARN'
}

if (-not $ForceReinstall -and $beforeVersion -and $downloadedVersion) {
    if ($beforeVersion -ge $downloadedVersion) {
        Write-Log "Installed Edge version ($beforeVersion) is already the same as or newer than the downloaded MSI version ($downloadedVersion). Skipping install." 'OK'
        exit 0
    }
}

try {
    $installExitCode = Install-EdgeMsi -MsiPath $EdgeMsiPath -MsiLog $MsiLogPath
}
catch {
    Write-Log "Script failed during install: $($_.Exception.Message)" 'ERROR'
    exit 3
}

Start-Sleep -Seconds 5

$afterVersion = Get-InstalledEdgeVersion
if ($afterVersion) {
    Write-Log "Installed Edge version after update: $afterVersion" 'INFO'
}
else {
    Write-Log "Could not detect Edge version after install." 'WARN'
}

if ($beforeVersion -and $downloadedVersion) {
    if ($beforeVersion -lt $downloadedVersion) {
        Write-Log "Installed version was older than the downloaded version and an update was needed." 'INFO'
    }
    elseif ($beforeVersion -eq $downloadedVersion) {
        Write-Log "Installed version already matched the downloaded MSI version before install." 'INFO'
    }
    else {
        Write-Log "Installed version before update was newer than the downloaded MSI version. Review whether the download source is behind your installed build." 'WARN'
    }
}

if ($afterVersion -and $downloadedVersion) {
    if ($afterVersion -eq $downloadedVersion) {
        Write-Log "Installed Edge version now matches the downloaded MSI version." 'OK'
    }
    elseif ($afterVersion -gt $downloadedVersion) {
        Write-Log "Installed Edge version after update is newer than the downloaded MSI version. Edge Update may have advanced the build beyond the MSI version." 'WARN'
    }
    else {
        Write-Log "Installed Edge version after update is still lower than the downloaded MSI version. Review MSI log: $MsiLogPath" 'WARN'
    }
}
elseif ($beforeVersion -and $afterVersion) {
    if ($afterVersion -gt $beforeVersion) {
        Write-Log "Edge was upgraded successfully." 'OK'
    }
    elseif ($afterVersion -eq $beforeVersion) {
        Write-Log "Edge version did not change. This usually means the same version was already installed, Edge was in use, or the installer did not replace the existing build." 'WARN'
    }
    else {
        Write-Log "Edge version after install appears lower than before install. Review MSI log: $MsiLogPath" 'WARN'
    }
}

switch ($installExitCode) {
    0    { exit 0 }
    1641 { exit 3010 }
    3010 { exit 3010 }
    default { exit 0 }
}
