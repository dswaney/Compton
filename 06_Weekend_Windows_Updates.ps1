# ScriptVersion: 1.0
# LastUpdated: 2026-03-23

<#
.SYNOPSIS
    Installs Windows Updates using PSWindowsUpdate and explicitly reboots
    if Windows reports that a reboot is required.

.DESCRIPTION
    This script is intended for scheduled/task-based execution where relying
    on -AutoReboot alone may not be consistent enough.

    It:
      - Ensures it is running as Administrator
      - Verifies/imports PSWindowsUpdate
      - Optionally resets Windows Update components
      - Installs available updates
      - Explicitly checks whether a reboot is required
      - Explicitly issues the reboot command

.NOTES
    Recommended to run as SYSTEM or a local administrator in Task Scheduler.
#>

[CmdletBinding()]
param(
    [switch]$ResetWUComponentsFirst = $false,
    [int]$OperationTimeoutSeconds = 1800,
    [int]$RebootDelaySeconds = 30,
    [string]$LogPath = "$env:SystemDrive\Temp\Weekend-Windows-Updates.log"
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

function Invoke-WithTimeout {
    param(
        [Parameter(Mandatory)]
        [scriptblock]$ScriptBlock,

        [int]$TimeoutSeconds = 1800,

        [string]$ActivityName = 'Operation'
    )

    Write-Log "Starting: $ActivityName" 'INFO'

    $job = Start-Job -ScriptBlock $ScriptBlock

    try {
        if (Wait-Job -Job $job -Timeout $TimeoutSeconds) {
            $output = Receive-Job -Job $job -Keep
            if ($output) {
                $output | ForEach-Object { Write-Log "$_" 'INFO' }
            }

            if ($job.State -eq 'Failed') {
                throw "Job failed for activity: $ActivityName"
            }

            Write-Log "Completed: $ActivityName" 'OK'
        }
        else {
            Stop-Job -Job $job -Force -ErrorAction SilentlyContinue | Out-Null
            throw "Timed out after $TimeoutSeconds seconds: $ActivityName"
        }
    }
    finally {
        Remove-Job -Job $job -Force -ErrorAction SilentlyContinue | Out-Null
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
        $results | ForEach-Object {
            Write-Log "$_" 'INFO'
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

Write-Log "Initializing weekend Windows update script..." 'INFO'

try {
    Ensure-PSWindowsUpdate

    if ($ResetWUComponentsFirst) {
        Reset-WUComponentsSafe
    }

    Install-AvailableWindowsUpdates

    $rebootRequired = Test-WURebootRequired

    if ($rebootRequired) {
        Write-Log "Windows reports that a reboot is required." 'OK'
        Invoke-ExplicitReboot -Delay $RebootDelaySeconds
        exit 3010
    }
    else {
        Write-Log "No reboot is currently required." 'OK'
        exit 0
    }
}
catch {
    Write-Log "Script failed: $($_.Exception.Message)" 'ERROR'
    exit 2
}