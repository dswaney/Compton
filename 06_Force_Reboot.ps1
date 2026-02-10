<#
.SYNOPSIS
  Clears common Windows 10/11 "pending reboot" flags, then performs a reboot.

.DESCRIPTION
  Force-clears well-known flags BEFORE rebooting to avoid phantom follow-up restarts:
    - CBS:
        HKLM\...\Component Based Servicing\RebootPending       (delete key)
        HKLM\...\Component Based Servicing\RebootInProgress    (delete key)
    - Windows Update:
        HKLM\...\WindowsUpdate\Auto Update\RebootRequired      (delete key)
    - Session Manager:
        HKLM\...\Session Manager\PendingFileRenameOperations   (delete value)
        HKLM\...\Session Manager\PendingFileRenameOperations2  (delete value, if present)
    - Legacy:
        HKLM\SOFTWARE\Microsoft\Updates\UpdateExeVolatile      (set to 0 if present)

  After clearing, it logs a verification snapshot and reboots (visible countdown).

.PARAMETER GracePeriodSeconds
  Seconds to wait before the reboot (via shutdown.exe). Default: 30

.PARAMETER Force
  Use -Force with Restart-Computer if we fall back to it. Default: $true

.NOTES
  - Run as Administrator.
  - Clearing these flags is not officially supported by Microsoft; use with care.
#>

[CmdletBinding(SupportsShouldProcess)]
param(
  [int]$GracePeriodSeconds = 30,
  [bool]$Force = $true
)

#--------------------------- Logging ---------------------------#
function Write-Log {
  param(
    [Parameter(Mandatory)][ValidateSet('INFO','WARN','ERROR','SUCCESS')][string]$Level,
    [Parameter(Mandatory)][string]$Message
  )
  $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
  $prefix = @{
    'INFO'    = '[INFO]   '
    'WARN'    = '[WARN]   '
    'ERROR'   = '[ERROR]  '
    'SUCCESS' = '[SUCCESS]'
  }[$Level]
  Write-Host "$ts $prefix $Message"
}

#--------------------------- Helpers ---------------------------#
function Test-IsElevated {
  $id = [Security.Principal.WindowsIdentity]::GetCurrent()
  $p  = New-Object Security.Principal.WindowsPrincipal($id)
  return $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Test-RegistryPathExists {
  param([Parameter(Mandatory)][string]$Path)
  try { return (Get-Item -LiteralPath $Path -ErrorAction Stop) -ne $null } catch { return $false }
}

function Remove-RegistryKeySafe {
  [CmdletBinding(SupportsShouldProcess)]
  param([Parameter(Mandatory)][string]$Path)

  if (Test-RegistryPathExists $Path) {
    try {
      if ($PSCmdlet.ShouldProcess($Path,'Remove-Item (registry key)')) {
        Remove-Item -LiteralPath $Path -Recurse -Force -ErrorAction Stop
        Write-Log -Level SUCCESS -Message "Removed key: ${Path}"
        return $true
      }
    } catch {
      Write-Log -Level ERROR -Message "Failed to remove key ${Path}: $($_.Exception.Message)"
    }
  } else {
    Write-Log -Level INFO -Message "Key not present: ${Path}"
  }
  return $false
}

function Remove-RegistryValueSafe {
  [CmdletBinding(SupportsShouldProcess)]
  param(
    [Parameter(Mandatory)][string]$Path,
    [Parameter(Mandatory)][string]$Name
  )
  try {
    $item = Get-Item -LiteralPath $Path -ErrorAction Stop
    $has  = ($item.GetValue($Name, $null) -ne $null)
    if ($has) {
      if ($PSCmdlet.ShouldProcess("$Path\$Name",'Remove-ItemProperty (registry value)')) {
        Remove-ItemProperty -LiteralPath $Path -Name $Name -Force -ErrorAction Stop
        Write-Log -Level SUCCESS -Message "Removed value: ${Path}\${Name}"
        return $true
      }
    } else {
      Write-Log -Level INFO -Message "Value not present: ${Path}\${Name}"
    }
  } catch {
    Write-Log -Level INFO -Message "Key not found (skipping): ${Path}"
  }
  return $false
}

#--------------------------- Clear flags ---------------------------#
function Clear-PendingRebootFlags {
  [CmdletBinding(SupportsShouldProcess)]
  param()

  Write-Log -Level INFO -Message "Clearing common reboot flags..."

  # CBS flags
  Remove-RegistryKeySafe 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending'    | Out-Null
  Remove-RegistryKeySafe 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootInProgress' | Out-Null

  # Windows Update RebootRequired
  Remove-RegistryKeySafe 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired'   | Out-Null

  # Session Manager: pending file rename arrays
  Remove-RegistryValueSafe 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' 'PendingFileRenameOperations'  | Out-Null
  Remove-RegistryValueSafe 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' 'PendingFileRenameOperations2' | Out-Null

  # Legacy UpdateExeVolatile (set to 0 if present)
  try {
    $path = 'HKLM:\SOFTWARE\Microsoft\Updates'
    $name = 'UpdateExeVolatile'
    $exists = $false
    try {
      $exists = (Get-ItemProperty -LiteralPath $path -ErrorAction Stop).PSObject.Properties.Name -contains $name
    } catch {}
    if ($exists) {
      if ($PSCmdlet.ShouldProcess("$path\$name", 'Set-ItemProperty (set to 0)')) {
        Set-ItemProperty -LiteralPath $path -Name $name -Value 0 -Type DWord -Force -ErrorAction Stop
        Write-Log -Level SUCCESS -Message "Set ${path}\${name} to 0"
      }
      # If you prefer deletion instead, swap for:
      # Remove-RegistryValueSafe $path $name | Out-Null
    } else {
      Write-Log -Level INFO -Message "Legacy value not present: ${path}\${name}"
    }
  } catch {
    Write-Log -Level ERROR -Message "UpdateExeVolatile handling failed: $($_.Exception.Message)"
  }

  Write-Log -Level SUCCESS -Message "Flag clearing pass complete."
}

#--------------------------- Verification snapshot ---------------------------#
function Get-PendingSignalsSnapshot {
  # Compute each signal first (no inline try/catch inside hashtable)
  $cbsPending   = Test-RegistryPathExists 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending'
  $cbsInProg    = Test-RegistryPathExists 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootInProgress'
  $wuReq        = Test-RegistryPathExists 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired'

  $smPend1 = $false
  try {
    $v = (Get-ItemProperty -LiteralPath 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -ErrorAction Stop).PendingFileRenameOperations
    if ($null -ne $v) { $smPend1 = @($v).Count -gt 0 }
  } catch {}

  $smPend2 = $false
  try {
    $v2 = (Get-ItemProperty -LiteralPath 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -ErrorAction Stop).PendingFileRenameOperations2
    if ($null -ne $v2) { $smPend2 = @($v2).Count -gt 0 }
  } catch {}

  $legacyVol = $false
  try {
    $val = (Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Updates' -ErrorAction Stop).UpdateExeVolatile
    if ($null -ne $val) { $legacyVol = ([int]$val) -gt 0 }
  } catch {}

  return [ordered]@{
    'CBS_RebootPending_Key' = $cbsPending
    'CBS_RebootInProgress'  = $cbsInProg
    'WU_RebootRequired_Key' = $wuReq
    'SM_PendingFileRename'  = $smPend1
    'SM_PendingFileRename2' = $smPend2
    'Legacy_UpdateVolatile' = $legacyVol
  }
}

#--------------------------- Main ---------------------------#
if (-not (Test-IsElevated)) {
  Write-Log -Level ERROR -Message "Please run this script as Administrator."
  exit 1
}

Write-Log -Level INFO -Message "Starting pre-reboot cleanup..."
Clear-PendingRebootFlags

$signals = Get-PendingSignalsSnapshot
$signals.GetEnumerator() | ForEach-Object {
  $state = if ($_.Value) { 'STILL PRESENT' } else { 'cleared' }
  Write-Log -Level INFO -Message ("Verify {0,-24}: {1}" -f $_.Key, $state)
}

# Always proceed to reboot after cleanup (per your requirement)
if ($PSCmdlet.ShouldProcess("Local computer","Reboot after clearing flags")) {
  if ($GracePeriodSeconds -gt 0) {
    Write-Log -Level WARN -Message "Rebooting in $GracePeriodSeconds seconds so users can save work..."
    try {
      Start-Process -FilePath "$env:SystemRoot\System32\shutdown.exe" `
                    -ArgumentList "/r /t $GracePeriodSeconds /c `"Maintenance reboot (flags cleared pre-restart)`"" `
                    -WindowStyle Hidden
      Start-Sleep -Seconds 2
      exit 0
    } catch {
      Write-Log -Level WARN -Message "Could not schedule with shutdown.exe: $($_.Exception.Message). Falling back to Restart-Computer after delay."
      Start-Sleep -Seconds $GracePeriodSeconds
    }
  }

  try {
    Write-Log -Level INFO -Message "Invoking Restart-Computer..."
    Restart-Computer -Force:$Force -Confirm:$false
  } catch {
    Write-Log -Level ERROR -Message "Restart-Computer failed: $($_.Exception.Message). Trying immediate shutdown.exe."
    try {
      Start-Process -FilePath "$env:SystemRoot\System32\shutdown.exe" -ArgumentList "/r /t 0" -WindowStyle Hidden
    } catch {
      Write-Log -Level ERROR -Message "shutdown.exe also failed: $($_.Exception.Message)"
      exit 1
    }
  }
}
