# ─────────────────────────────────────────────────────────────────────────────
# Ensure script runs as Administrator
# ─────────────────────────────────────────────────────────────────────────────
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Error "Please run this script as Administrator."
    exit
}

# ─────────────────────────────────────────────────────────────────────────────
# Set Execution Policy AFTER verifying admin rights
# ─────────────────────────────────────────────────────────────────────────────
Set-ExecutionPolicy RemoteSigned -Scope LocalMachine -Force

$ErrorActionPreference = "Stop"
trap { Write-Error "Unhandled Error: $_"; exit 1 }

# ——————————————————————————————————————————————
# Install HP Drivers and clean up install files (runs on all hardware)
# Does not update the Flash BIOS
# ——————————————————————————————————————————————
Write-Host "Installing HP drivers and cleaning up…" -ForegroundColor Cyan
Get-HPDrivers -NoPrompt -DeleteInstallationFiles
