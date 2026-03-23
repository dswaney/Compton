# ScriptVersion: 1.0
# LastUpdated: 2026-03-23

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

function Get-InstalledEdgeVersion {
    $candidates = @(
        "$env:ProgramFiles(x86)\Microsoft\Edge\Application\msedge.exe",
        "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe"
    )

    foreach ($path in $candidates) {
        if (Test-Path -LiteralPath $path) {
            try {
                return [version](Get-Item -LiteralPath $path).VersionInfo.ProductVersion
            }
            catch {
            }
        }
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
    $installExitCode = Install-EdgeMsi -MsiPath $EdgeMsiPath -MsiLog $MsiLogPath
}
catch {
    Write-Log "Script failed: $($_.Exception.Message)" 'ERROR'
    exit 2
}

Start-Sleep -Seconds 5

$afterVersion = Get-InstalledEdgeVersion
if ($afterVersion) {
    Write-Log "Installed Edge version after update: $afterVersion" 'INFO'
}
else {
    Write-Log "Could not detect Edge version after install." 'WARN'
}

if ($beforeVersion -and $afterVersion) {
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