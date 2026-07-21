<#
.SYNOPSIS
    Installs and registers a SHARP printer driver, then silently installs the
    PaperCut Print Deploy client.

.DESCRIPTION
    1. Compares the fileshare SHARP driver package with the last successfully
       deployed package by using a deterministic SHA-256 package fingerprint.
    2. Updates and registers the printer driver only when it is missing or the
       fileshare package has changed.
    3. Ensures the PaperCut Print Deploy client is installed.
    4. Ensures the StudentSecurePrint shared printer connection exists.
    5. Supports interactive execution and SYSTEM scheduled-task execution.
    6. Writes a transcript-style log to C:\Logs.

.VERSION
    1.1.0

.DATE
    2026-07-21

.CHANGELOG
    1.1.0
    - Added scheduled-task-safe, idempotent maintenance logic.
    - Added a deterministic SHA-256 fingerprint for the source driver package.
    - Driver deployment runs only when the driver is missing or the source
      package differs from the last successfully deployed package.
    - Added a local deployment-state file under C:\ProgramData\Compton.
    - Added SYSTEM detection and a per-computer printer connection using
      PrintUIEntry /ga when running as SYSTEM.
    - Interactive runs continue to create and verify a per-user connection.
    - PaperCut installation and printer connection are checked on every run.

    1.0.4
    - Added centralized colored console output.
    - Added a startup banner for easier identification.
    - Added ACTION log level for active installation steps.
    - Log files remain plain text for compatibility.

    1.0.3
    - Added automatic connection to the shared printer
      \\papercut\StudentSecurePrint.
    - Verifies the shared printer connection after installation.
    - Does not change the user's default printer.

    1.0.2
    - Added REBOOT=ReallySuppress to the PaperCut MSI installation.
    - Retained /norestart as an additional restart-suppression safeguard.
    - Exit code 1641 is no longer accepted as a successful result.
    - Exit code 3010 is logged as success with a pending restart, without
      restarting the computer.

    1.0.1
    - Fixed PaperCut uninstall-registry detection under Set-StrictMode.
    - Safely handles registry entries that do not contain DisplayName or
      DisplayVersion properties.
    - Corrected the final failure message so it identifies the active step.

.NOTES
    Run from an elevated 64-bit Windows PowerShell session.

    IMPORTANT:
    Set $DriverSourcePath and $PrinterDriverName in the Configuration section.
    The printer driver name must exactly match the model name contained in the
    SHARP INF package and shown in Print Server Properties > Drivers.
#>

[CmdletBinding()]
param()

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

# =========================
# Configuration
# =========================

# Folder containing the fully extracted SHARP driver package.
# Example:
# $DriverSourcePath = '\\filesvr\labscripts\Sharp\MX-Series-PCL6'
$DriverSourcePath = '\\papercut\Printer Drivers\MFP\win\SH_D31_PCL6_PS_2410a_EnglishUS_64bit'

# Exact Windows printer-driver model name from the SHARP INF.
# Example only:
# $PrinterDriverName = 'SHARP MX-6071 PCL6'
$PrinterDriverName = 'Sharp BP-70C31 PCL6'

$PaperCutMsiPath = '\\papercut\Print Deploy Clients\win\pc-print-deploy-client[10.2.3.44].msi'

# Shared printer connection to create after the driver and PaperCut client
# installation steps complete.
$PrinterSharePath = '\\papercut\StudentSecurePrint'

$LocalDriverStage = 'C:\ProgramData\Compton\Drivers\Sharp'
$StateDirectory   = 'C:\ProgramData\Compton\State'
$StatePath        = Join-Path $StateDirectory 'SharpDriver-PaperCut-State.json'
$LogDirectory     = 'C:\Logs'
$LogPath          = Join-Path $LogDirectory 'Install-SharpDriver-And-PaperCut.log'

# =========================
# Helper functions
# =========================

function Write-Log {
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [ValidateSet('INFO', 'ACTION', 'WARN', 'ERROR', 'SUCCESS')]
        [string]$Level = 'INFO'
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[{0}] [{1}] {2}" -f $timestamp, $Level, $Message

    $foregroundColor = switch ($Level) {
        'ACTION'  { 'Yellow' }
        'SUCCESS' { 'Green' }
        'WARN'    { 'DarkYellow' }
        'ERROR'   { 'Red' }
        default   { 'White' }
    }

    Write-Host $line -ForegroundColor $foregroundColor
    Add-Content -LiteralPath $LogPath -Value $line -Encoding UTF8
}

function Write-Banner {
    $border = '=' * 68

    Write-Host ''
    Write-Host $border -ForegroundColor Cyan
    Write-Host '     SHARP Driver and PaperCut Print Deploy Installation' -ForegroundColor Cyan
    Write-Host $border -ForegroundColor Cyan
    Write-Host ''
}

function Test-IsAdministrator {
    $identity  = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole(
        [Security.Principal.WindowsBuiltInRole]::Administrator
    )
}

function Invoke-ExternalProcess {
    param(
        [Parameter(Mandatory)]
        [string]$FilePath,

        [Parameter(Mandatory)]
        [string[]]$ArgumentList,

        [int[]]$SuccessExitCodes = @(0)
    )

    Write-Log "Running: $FilePath $($ArgumentList -join ' ')" 'ACTION'

    $process = Start-Process `
        -FilePath $FilePath `
        -ArgumentList $ArgumentList `
        -Wait `
        -PassThru `
        -WindowStyle Hidden

    Write-Log "Process exit code: $($process.ExitCode)"

    if ($process.ExitCode -notin $SuccessExitCodes) {
        throw "Command failed with exit code $($process.ExitCode): $FilePath"
    }

    return $process.ExitCode
}

function Get-PrinterInfFiles {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    $allInfFiles = Get-ChildItem `
        -LiteralPath $Path `
        -Filter '*.inf' `
        -File `
        -Recurse `
        -ErrorAction Stop

    if (-not $allInfFiles) {
        throw "No INF files were found under: $Path"
    }

    # Prefer INF files that declare the Printer class. If the package does not
    # expose that text plainly, return all INF files and let PnPUtil process them.
    $printerInfFiles = foreach ($inf in $allInfFiles) {
        $content = Get-Content -LiteralPath $inf.FullName -ErrorAction SilentlyContinue
        if ($content -match '^\s*Class\s*=\s*Printer\s*$') {
            $inf
        }
    }

    if ($printerInfFiles) {
        return @($printerInfFiles)
    }

    Write-Log 'No INF explicitly declared Class=Printer; all INF files will be staged.' 'WARN'
    return @($allInfFiles)
}


function Test-IsSystemAccount {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    return $identity.User.Value -eq 'S-1-5-18'
}

function Get-DriverPackageFingerprint {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    $files = @(
        Get-ChildItem `
            -LiteralPath $Path `
            -File `
            -Recurse `
            -ErrorAction Stop |
        Sort-Object FullName
    )

    if (-not $files) {
        throw "No files were found in the driver source package: $Path"
    }

    $manifestLines = foreach ($file in $files) {
        $relativePath = $file.FullName.Substring($Path.TrimEnd('\').Length).TrimStart('\')
        $hash = Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256 -ErrorAction Stop

        '{0}|{1}|{2}' -f $relativePath.ToLowerInvariant(), $file.Length, $hash.Hash
    }

    $manifestText = $manifestLines -join "`n"
    $sha256 = [Security.Cryptography.SHA256]::Create()

    try {
        $bytes = [Text.Encoding]::UTF8.GetBytes($manifestText)
        $fingerprintBytes = $sha256.ComputeHash($bytes)
        return ([BitConverter]::ToString($fingerprintBytes)).Replace('-', '')
    }
    finally {
        $sha256.Dispose()
    }
}

function Get-DeploymentState {
    if (-not (Test-Path -LiteralPath $StatePath -PathType Leaf)) {
        return $null
    }

    try {
        return Get-Content -LiteralPath $StatePath -Raw -ErrorAction Stop |
            ConvertFrom-Json -ErrorAction Stop
    }
    catch {
        Write-Log "The deployment-state file could not be read and will be rebuilt. $($_.Exception.Message)" 'WARN'
        return $null
    }
}

function Save-DeploymentState {
    param(
        [Parameter(Mandatory)]
        [string]$Fingerprint
    )

    New-Item -Path $StateDirectory -ItemType Directory -Force | Out-Null

    $state = [ordered]@{
        DriverName              = $PrinterDriverName
        SourcePath              = $DriverSourcePath
        SourceFingerprintSHA256 = $Fingerprint
        LastSuccessfulUpdate    = (Get-Date).ToString('o')
        ComputerName            = $env:COMPUTERNAME
    }

    $state |
        ConvertTo-Json -Depth 3 |
        Set-Content -LiteralPath $StatePath -Encoding UTF8 -Force
}

function Test-PerMachinePrinterConnection {
    param(
        [Parameter(Mandatory)]
        [string]$ConnectionPath
    )

    $connectionsPath = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Connections'

    if (-not (Test-Path -LiteralPath $connectionsPath)) {
        return $false
    }

    $serverAndShare = $ConnectionPath.TrimStart('\') -split '\\', 2

    if ($serverAndShare.Count -ne 2) {
        return $false
    }

    $expectedKeyName = ',,' + $serverAndShare[0] + ',' + $serverAndShare[1]

    return $null -ne (
        Get-ChildItem -LiteralPath $connectionsPath -ErrorAction SilentlyContinue |
        Where-Object { $_.PSChildName -ieq $expectedKeyName } |
        Select-Object -First 1
    )
}

function Ensure-SharedPrinterConnection {
    param(
        [Parameter(Mandatory)]
        [string]$ConnectionPath
    )

    if (Test-IsSystemAccount) {
        Write-Log 'Running as SYSTEM; checking the per-computer printer connection.'

        if (Test-PerMachinePrinterConnection -ConnectionPath $ConnectionPath) {
            Write-Log "Per-computer printer connection is already configured: $ConnectionPath" 'SUCCESS'
            return
        }

        Write-Log "Creating per-computer printer connection: $ConnectionPath" 'ACTION'

        Invoke-ExternalProcess `
            -FilePath "$env:SystemRoot\System32\rundll32.exe" `
            -ArgumentList @(
                'printui.dll,PrintUIEntry',
                '/ga',
                "/n`"$ConnectionPath`""
            ) `
            -SuccessExitCodes @(0) | Out-Null

        if (-not (Test-PerMachinePrinterConnection -ConnectionPath $ConnectionPath)) {
            throw "The per-computer printer connection could not be verified: $ConnectionPath"
        }

        Write-Log 'Per-computer printer connection created. It becomes available to users at their next sign-in.' 'SUCCESS'
        return
    }

    Write-Log "Checking the current user's printer connection: $ConnectionPath"

    $serverAndShare = $ConnectionPath.TrimStart('\') -split '\\', 2
    $serverName = $serverAndShare[0]
    $shareName = $serverAndShare[1]

    $existingConnection = Get-Printer -ErrorAction SilentlyContinue |
        Where-Object {
            $_.Name -eq $ConnectionPath -or
            ($_.ComputerName -ieq $serverName -and $_.ShareName -ieq $shareName)
        } |
        Select-Object -First 1

    if ($existingConnection) {
        Write-Log "Shared printer is already connected: $ConnectionPath" 'SUCCESS'
        return
    }

    Write-Log "Connecting shared printer for the current user: $ConnectionPath" 'ACTION'

    try {
        Add-Printer -ConnectionName $ConnectionPath -ErrorAction Stop
    }
    catch {
        Write-Log "Add-Printer failed; trying PrintUIEntry fallback. $($_.Exception.Message)" 'WARN'

        Invoke-ExternalProcess `
            -FilePath "$env:SystemRoot\System32\rundll32.exe" `
            -ArgumentList @(
                'printui.dll,PrintUIEntry',
                '/in',
                "/n`"$ConnectionPath`""
            ) `
            -SuccessExitCodes @(0) | Out-Null
    }

    Start-Sleep -Seconds 2

    $connectedPrinter = Get-Printer -ErrorAction SilentlyContinue |
        Where-Object {
            $_.Name -eq $ConnectionPath -or
            ($_.ComputerName -ieq $serverName -and $_.ShareName -ieq $shareName)
        } |
        Select-Object -First 1

    if (-not $connectedPrinter) {
        throw "The shared printer connection could not be verified: $ConnectionPath"
    }

    Write-Log "Shared printer connected successfully: $ConnectionPath" 'SUCCESS'
}

# =========================
# Main
# =========================

try {
    New-Item -Path $LogDirectory -ItemType Directory -Force | Out-Null
    Write-Banner
    Write-Log 'Starting SHARP driver and PaperCut Print Deploy installation.'

    if (-not (Test-IsAdministrator)) {
        throw 'This script must be run as an administrator.'
    }

    if ($PrinterDriverName -like 'CHANGE ME*' -or
        [string]::IsNullOrWhiteSpace($PrinterDriverName)) {
        throw 'Set $PrinterDriverName to the exact SHARP driver model name before running the script.'
    }

    if (-not (Test-Path -LiteralPath $DriverSourcePath -PathType Container)) {
        throw "The SHARP driver source folder is unavailable: $DriverSourcePath"
    }

    if (-not (Test-Path -LiteralPath $PaperCutMsiPath -PathType Leaf)) {
        throw "The PaperCut Print Deploy MSI is unavailable: $PaperCutMsiPath"
    }

    Import-Module PrintManagement -ErrorAction Stop

    $sourceFingerprint = Get-DriverPackageFingerprint -Path $DriverSourcePath
    $deploymentState = Get-DeploymentState
    $existingDriver = Get-PrinterDriver -Name $PrinterDriverName -ErrorAction SilentlyContinue

    $storedFingerprint = $null
    if ($null -ne $deploymentState) {
        $fingerprintProperty = $deploymentState.PSObject.Properties['SourceFingerprintSHA256']
        if ($null -ne $fingerprintProperty) {
            $storedFingerprint = [string]$fingerprintProperty.Value
        }
    }

    $driverUpdateRequired = $false

    if (-not $existingDriver) {
        Write-Log "Printer driver is not registered: $PrinterDriverName" 'WARN'
        $driverUpdateRequired = $true
    }
    elseif ([string]::IsNullOrWhiteSpace($storedFingerprint)) {
        Write-Log 'No prior deployment fingerprint exists; the driver package will be validated and deployed once.' 'WARN'
        $driverUpdateRequired = $true
    }
    elseif ($storedFingerprint -ne $sourceFingerprint) {
        Write-Log 'The fileshare driver package differs from the last successfully deployed package.' 'WARN'
        $driverUpdateRequired = $true
    }
    else {
        Write-Log "The registered driver matches the last successfully deployed fileshare package: $PrinterDriverName" 'SUCCESS'
    }

    if ($driverUpdateRequired) {
        if (Test-Path -LiteralPath $LocalDriverStage) {
            Write-Log "Clearing the local driver staging folder: $LocalDriverStage" 'ACTION'
            Remove-Item -LiteralPath $LocalDriverStage -Recurse -Force -ErrorAction Stop
        }

        New-Item -Path $LocalDriverStage -ItemType Directory -Force | Out-Null

        Write-Log "Copying SHARP driver package from '$DriverSourcePath' to '$LocalDriverStage'." 'ACTION'
        Copy-Item `
            -Path (Join-Path $DriverSourcePath '*') `
            -Destination $LocalDriverStage `
            -Recurse `
            -Force `
            -ErrorAction Stop

        $infFiles = @(Get-PrinterInfFiles -Path $LocalDriverStage)
        Write-Log "Found $($infFiles.Count) applicable INF file(s)."

        foreach ($inf in $infFiles) {
            Invoke-ExternalProcess `
                -FilePath "$env:SystemRoot\System32\pnputil.exe" `
                -ArgumentList @('/add-driver', "`"$($inf.FullName)`"") `
                -SuccessExitCodes @(0, 3010) | Out-Null
        }

        Write-Log 'Driver package staging completed.' 'SUCCESS'
        Write-Log "Registering or refreshing printer driver: $PrinterDriverName" 'ACTION'

        Add-PrinterDriver -Name $PrinterDriverName -ErrorAction Stop

        $registeredDriver = Get-PrinterDriver `
            -Name $PrinterDriverName `
            -ErrorAction SilentlyContinue

        if (-not $registeredDriver) {
            throw "The driver package was staged, but '$PrinterDriverName' was not registered with the Print Spooler."
        }

        Save-DeploymentState -Fingerprint $sourceFingerprint
        Write-Log "Printer driver deployment completed successfully: $PrinterDriverName" 'SUCCESS'
    }

    # Install PaperCut only after the printer driver has been verified.
    $uninstallRoots = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )

    $paperCutProduct = Get-ItemProperty `
        -Path $uninstallRoots `
        -ErrorAction SilentlyContinue |
        Where-Object {
            $displayNameProperty = $_.PSObject.Properties['DisplayName']

            $null -ne $displayNameProperty -and
            -not [string]::IsNullOrWhiteSpace([string]$displayNameProperty.Value) -and
            [string]$displayNameProperty.Value -match 'PaperCut.*Print Deploy|Print Deploy Client'
        } |
        Select-Object -First 1

    if ($paperCutProduct) {
        $paperCutDisplayName = [string]$paperCutProduct.PSObject.Properties['DisplayName'].Value
        $versionProperty = $paperCutProduct.PSObject.Properties['DisplayVersion']

        if ($null -ne $versionProperty -and
            -not [string]::IsNullOrWhiteSpace([string]$versionProperty.Value)) {
            $paperCutDisplayVersion = [string]$versionProperty.Value
            Write-Log "PaperCut Print Deploy is already installed: $paperCutDisplayName $paperCutDisplayVersion" 'SUCCESS'
        }
        else {
            Write-Log "PaperCut Print Deploy is already installed: $paperCutDisplayName" 'SUCCESS'
        }
    }
    else {
        Write-Log 'Installing PaperCut Print Deploy client silently.' 'ACTION'

        $msiLogPath = Join-Path $LogDirectory 'PaperCut-Print-Deploy-MSI.log'

        $exitCode = Invoke-ExternalProcess `
            -FilePath "$env:SystemRoot\System32\msiexec.exe" `
            -ArgumentList @(
                '/i',
                "`"$PaperCutMsiPath`"",
                '/qn',
                '/norestart',
                'REBOOT=ReallySuppress',
                '/L*v',
                "`"$msiLogPath`""
            ) `
            -SuccessExitCodes @(0, 3010)

        switch ($exitCode) {
            0 {
                Write-Log 'PaperCut Print Deploy installed successfully.' 'SUCCESS'
            }

            3010 {
                Write-Log 'PaperCut installed successfully. A restart is pending, but no automatic restart was allowed.' 'WARN'
            }

            default {
                throw "PaperCut installation returned an unexpected exit code: $exitCode"
            }
        }
    }

    # Ensure the printer connection exists. SYSTEM creates a per-computer
    # connection; an interactive run creates a current-user connection.
    Ensure-SharedPrinterConnection -ConnectionPath $PrinterSharePath

    Write-Log 'All requested installation steps completed successfully.' 'SUCCESS'
    exit 0
}
catch {
    Write-Log $_.Exception.Message 'ERROR'
    Write-Log 'Installation stopped before all requested steps completed.' 'ERROR'
    exit 1
}
