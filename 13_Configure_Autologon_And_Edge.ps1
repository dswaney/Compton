# =====================================================================
# ScriptName: 13_Configure_Autologon_And_Edge.ps1
# ScriptVersion: 1.1.0
# LastUpdated: 2026-07-23
# =====================================================================

[CmdletBinding()]
param(
    # Add or remove computer-name patterns here.
    # Wildcards are supported:
    #   'SSB-122-*' matches names beginning with SSB-122-
    #   'SSB-114*'  matches names beginning with SSB-114
    #   'LAB-205-01' matches one exact computer name
    [string[]]$ComputerNamePatterns = @(
        'SSB-122-*',
        'SSB-114*'
    ),

    [string]$DefaultUserName = 'CC-Student',

    [string]$DefaultPassword = '',

    [string]$DefaultDomainName = 'Compton.edu',

    [string]$EdgeUrl = 'https://www.compton.edu',

    [string]$LogDirectory = 'C:\Logs'
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$ScriptVersion = '1.1.0'
$ComputerName = $env:COMPUTERNAME
$LogPath = Join-Path $LogDirectory 'Configure-Autologon-And-Edge.log'
$WinlogonPath = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon'
$RunKeyPath = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run'
$EdgeRunValueName = 'LaunchComptonEdge'
$AllUsersStartupPath = Join-Path $env:ProgramData 'Microsoft\Windows\Start Menu\Programs\Startup'
$LegacyChromeShortcutPath = Join-Path $AllUsersStartupPath 'Google Chrome.lnk'

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','OK','WARN','ERROR')][string]$Level = 'INFO'
    )

    $line = '[{0}] [{1,-5}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message

    switch ($Level) {
        'OK'    { Write-Host $line -ForegroundColor Green }
        'WARN'  { Write-Host $line -ForegroundColor Yellow }
        'ERROR' { Write-Host $line -ForegroundColor Red }
        default { Write-Host $line }
    }

    try {
        if (-not (Test-Path -LiteralPath $LogDirectory)) {
            New-Item -Path $LogDirectory -ItemType Directory -Force | Out-Null
        }
        Add-Content -LiteralPath $LogPath -Value $line -Encoding UTF8
    }
    catch {
        Write-Warning "Unable to write to the log file: $($_.Exception.Message)"
    }
}

function Test-IsAdministrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Test-ComputerNameMatch {
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string[]]$Patterns
    )

    foreach ($pattern in $Patterns) {
        if (-not [string]::IsNullOrWhiteSpace($pattern) -and $Name -like $pattern) {
            return $true
        }
    }

    return $false
}

function Set-RegistryString {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][AllowEmptyString()][string]$Value
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -Path $Path -Force | Out-Null
    }

    $current = $null
    try {
        $current = (Get-ItemProperty -LiteralPath $Path -Name $Name -ErrorAction Stop).$Name
    }
    catch {
        $current = $null
    }

    if ([string]$current -ne $Value) {
        New-ItemProperty -LiteralPath $Path -Name $Name -Value $Value -PropertyType String -Force | Out-Null
        Write-Log "Updated registry value: $Path\$Name" 'OK'
    }
    else {
        Write-Log "Registry value is already correct: $Path\$Name"
    }
}

function Set-AutologonConfiguration {
    Write-Log "Ensuring Windows autologon is configured for user '$DefaultUserName'."

    Set-RegistryString -Path $WinlogonPath -Name 'AutoAdminLogon' -Value '1'
    Set-RegistryString -Path $WinlogonPath -Name 'DefaultUserName' -Value $DefaultUserName
    Set-RegistryString -Path $WinlogonPath -Name 'DefaultPassword' -Value $DefaultPassword

    Set-RegistryString -Path $WinlogonPath -Name 'DefaultDomainName' -Value $DefaultDomainName

    Set-RegistryString -Path $WinlogonPath -Name 'ForceAutoLogon' -Value '1'
    Write-Log "Autologon configuration is enabled for '$DefaultUserName'." 'OK'
}

function Remove-LegacyChromeStartupShortcut {
    Write-Log 'Checking for the legacy All Users Google Chrome startup shortcut.'

    if (Test-Path -LiteralPath $LegacyChromeShortcutPath -PathType Leaf) {
        Remove-Item -LiteralPath $LegacyChromeShortcutPath -Force -ErrorAction Stop

        if (Test-Path -LiteralPath $LegacyChromeShortcutPath -PathType Leaf) {
            throw "The legacy Chrome startup shortcut could not be removed: $LegacyChromeShortcutPath"
        }

        Write-Log "Removed legacy Chrome startup shortcut: $LegacyChromeShortcutPath" 'OK'
    }
    else {
        Write-Log 'Legacy Google Chrome startup shortcut is not present.'
    }
}

function Get-EdgePath {
    $candidatePaths = @(
        (Join-Path ${env:ProgramFiles(x86)} 'Microsoft\Edge\Application\msedge.exe'),
        (Join-Path ${env:ProgramFiles} 'Microsoft\Edge\Application\msedge.exe')
    )

    return $candidatePaths |
        Where-Object { Test-Path -LiteralPath $_ -PathType Leaf } |
        Select-Object -First 1
}

function Set-EdgeAutoLaunch {
    $edgePath = Get-EdgePath

    if (-not $edgePath) {
        throw 'Microsoft Edge was not found in either Program Files location.'
    }

    $command = '"{0}" --inprivate --new-window --start-maximized "{1}"' -f $edgePath, $EdgeUrl
    Set-RegistryString -Path $RunKeyPath -Name $EdgeRunValueName -Value $command

    Write-Log "Configured Edge to launch for every user at logon: $EdgeUrl" 'OK'
}

if (-not (Test-IsAdministrator)) {
    Write-Error 'Please run this script as Administrator or through a SYSTEM scheduled task.'
    exit 1
}

Write-Log "===== Autologon and Edge configuration v$ScriptVersion started ====="
Write-Log "Computer name: $ComputerName"
Write-Log "Configured computer-name patterns: $($ComputerNamePatterns -join ', ')"

try {
    if (-not (Test-ComputerNameMatch -Name $ComputerName -Patterns $ComputerNamePatterns)) {
        Write-Log "Computer '$ComputerName' does not match the configured pattern list. No changes were made." 'INFO'
        Write-Log '===== Script completed: computer not targeted =====' 'OK'
        exit 0
    }

    Write-Log "Computer '$ComputerName' matches the configured pattern list." 'OK'

    Set-AutologonConfiguration
    Remove-LegacyChromeStartupShortcut
    Set-EdgeAutoLaunch

    Write-Log 'Autologon, legacy Chrome startup cleanup, and Edge startup settings were verified successfully.' 'OK'
    Write-Log '===== Script completed successfully =====' 'OK'
    exit 0
}
catch {
    Write-Log "Configuration failed: $($_.Exception.Message)" 'ERROR'
    Write-Log '===== Script completed with errors =====' 'ERROR'
    exit 1
}
