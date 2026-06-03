# Run as Administrator

$ProviderSource = "\\filesvr\Labscripts\DellCommandPowerShellProvider_2.10.0_153"
$LogPath = "C:\Logs\Dell-BIOS-Power-WOL.log"

if (-not (Test-Path "C:\Logs")) {
    New-Item -Path "C:\Logs" -ItemType Directory -Force | Out-Null
}

function Write-Log {
    param([string]$Message)
    $line = "[{0}] {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Message
    Write-Host $line
    Add-Content -Path $LogPath -Value $line
}

Write-Log "Starting Dell BIOS Auto-On and Wake-on-LAN configuration."

# Import Dell Command PowerShell Provider
$module = Get-Module -ListAvailable -Name DellBIOSProvider

if (-not $module) {
    Write-Log "DellBIOSProvider module not found locally. Attempting import from file share."

    $psd1 = Get-ChildItem -Path $ProviderSource -Filter "DellBIOSProvider.psd1" -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1

    if (-not $psd1) {
        throw "Could not find DellBIOSProvider.psd1 under $ProviderSource"
    }

    Import-Module $psd1.FullName -Force
}
else {
    Import-Module DellBIOSProvider -Force
}

if (-not (Get-PSDrive -Name DellSmbios -ErrorAction SilentlyContinue)) {
    throw "DellSmbios PSDrive was not loaded. Confirm Dell Command PowerShell Provider is installed correctly."
}

Write-Log "DellBIOSProvider loaded successfully."

# Optional: show current values
Write-Log "Current AutoOn: $((Get-Item DellSmbios:\PowerManagement\AutoOn).CurrentValue)"
Write-Log "Current AutoOnSun: $((Get-Item DellSmbios:\PowerManagement\AutoOnSun).CurrentValue)"
Write-Log "Current AutoOnHr: $((Get-Item DellSmbios:\PowerManagement\AutoOnHr).CurrentValue)"
Write-Log "Current AutoOnMn: $((Get-Item DellSmbios:\PowerManagement\AutoOnMn).CurrentValue)"
Write-Log "Current WakeOnLan: $((Get-Item DellSmbios:\PowerManagement\WakeOnLan).CurrentValue)"

# Configure Sunday midnight power-on
Set-Item -Path DellSmbios:\PowerManagement\AutoOn -Value SelectDays
Set-Item -Path DellSmbios:\PowerManagement\AutoOnSun -Value Enabled
Set-Item -Path DellSmbios:\PowerManagement\AutoOnHr -Value 0
Set-Item -Path DellSmbios:\PowerManagement\AutoOnMn -Value 0

# Optional: disable other days so only Sunday powers on
Set-Item -Path DellSmbios:\PowerManagement\AutoOnMon -Value Disabled -ErrorAction SilentlyContinue
Set-Item -Path DellSmbios:\PowerManagement\AutoOnTue -Value Disabled -ErrorAction SilentlyContinue
Set-Item -Path DellSmbios:\PowerManagement\AutoOnWed -Value Disabled -ErrorAction SilentlyContinue
Set-Item -Path DellSmbios:\PowerManagement\AutoOnThur -Value Disabled -ErrorAction SilentlyContinue
Set-Item -Path DellSmbios:\PowerManagement\AutoOnFri -Value Disabled -ErrorAction SilentlyContinue
Set-Item -Path DellSmbios:\PowerManagement\AutoOnSat -Value Disabled -ErrorAction SilentlyContinue

# Enable Wake on LAN
# Use LanOnly for wired LAN only.
# Use LanWlan if you want LAN or WLAN where supported.
Set-Item -Path DellSmbios:\PowerManagement\WakeOnLan -Value LanOnly

Write-Log "Updated BIOS settings:"
Write-Log "AutoOn: $((Get-Item DellSmbios:\PowerManagement\AutoOn).CurrentValue)"
Write-Log "AutoOnSun: $((Get-Item DellSmbios:\PowerManagement\AutoOnSun).CurrentValue)"
Write-Log "AutoOnHr: $((Get-Item DellSmbios:\PowerManagement\AutoOnHr).CurrentValue)"
Write-Log "AutoOnMn: $((Get-Item DellSmbios:\PowerManagement\AutoOnMn).CurrentValue)"
Write-Log "WakeOnLan: $((Get-Item DellSmbios:\PowerManagement\WakeOnLan).CurrentValue)"

Write-Log "Dell BIOS configuration completed."