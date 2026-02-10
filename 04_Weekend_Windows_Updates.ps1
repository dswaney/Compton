# ─────────────────────────────────────────────────────────────────────────────
# Ensure script runs as Administrator
# ─────────────────────────────────────────────────────────────────────────────
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Error "Please run this script as Administrator."
    exit
}

# ─────────────────────────────────────────────────────────────────────────────
# Reset Windows Update Components with timeout and retry
# ─────────────────────────────────────────────────────────────────────────────
function Invoke-WithTimeoutAndRetry {
    param(
        [Parameter(Mandatory)][scriptblock]$ScriptBlock,
        [int]$TimeoutSeconds = 300,
        [int]$MaxRetries = 2
    )
    for ($i = 1; $i -le $MaxRetries; $i++) {
        Write-Output "Starting attempt $i of $MaxRetries..."
        $job = Start-Job -ScriptBlock $ScriptBlock
        if (Wait-Job $job -Timeout $TimeoutSeconds) {
            Receive-Job $job | Write-Output
            Remove-Job $job | Out-Null
            return
        } else {
            Write-Warning "Attempt $i timed out after $TimeoutSeconds seconds."
            Stop-Job $job | Out-Null; Remove-Job $job | Out-Null
        }
        catch { Write-Warning "Unexpected error: $_" }
    }
    Write-Error "Operation exceeded $MaxRetries retries and timed out."
    catch { Write-Warning "Unexpected error: $_" }
}

# ─────────────────────────────────────────────────────────────────────────────
# Execute Reset-WUComponents with automated timeout and retry
# ─────────────────────────────────────────────────────────────────────────────
Invoke-WithTimeoutAndRetry -ScriptBlock { Reset-WUComponents -Verbose } -TimeoutSeconds 300 -MaxRetries 2

# ─────────────────────────────────────────────────────────────────────────────
# Install Windows Updates
# ─────────────────────────────────────────────────────────────────────────────
Get-WindowsUpdate -AcceptAll -Install -AutoReboot
