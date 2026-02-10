# Suppress all errors and warnings
try {
    Set-ExecutionPolicy Bypass -Scope Process -Force -ErrorAction Stop -WarningAction SilentlyContinue
} catch {
    # Swallow any failure silently
}

function Invoke-SystemMaintenance {
    [CmdletBinding()]
    param (
        [switch]$ArchiveLogs,
        [string]$LogArchivePath = 'C:\Logs'
    )

    $ErrorActionPreference = 'Stop'

    if ($ArchiveLogs -and -not (Test-Path $LogArchivePath)) {
        try {
            New-Item -Path $LogArchivePath -ItemType Directory -Force | Out-Null
            Write-Verbose "Created log folder: $LogArchivePath"
        } catch {
            Write-Warning ("Could not create log folder: " + $_)
        }
    }

    try {
        Write-Host "▶ Repairing filesystem…" -ForegroundColor Cyan
        $sysDriveLetter = ($env:SystemDrive.Substring(0,1)).ToUpper()
        if (Get-Volume -DriveLetter $sysDriveLetter -ErrorAction SilentlyContinue) {
            try {
                Repair-Volume -DriveLetter $sysDriveLetter -Scan -ErrorAction Stop -WarningAction SilentlyContinue
            } catch {
                Write-Verbose "   • Scan on ${sysDriveLetter}: failed or no objects found."
            }
            try {
                Repair-Volume -DriveLetter $sysDriveLetter -OfflineScanAndFix -ErrorAction Stop -WarningAction SilentlyContinue
            } catch {
                Write-Verbose "   • OfflineScanAndFix on ${sysDriveLetter}: failed or no objects found."
            }
        }

        Write-Host "▶ Checking & repairing component store…" -ForegroundColor Cyan
        DISM.exe /Online /Cleanup-Image /CheckHealth
        DISM.exe /Online /Cleanup-Image /ScanHealth
        DISM.exe /Online /Cleanup-Image /RestoreHealth
        DISM.exe /Online /Cleanup-Image /StartComponentCleanup /ResetBase

        Write-Host "▶ Verifying WMI repository…" -ForegroundColor Cyan
        winmgmt /verifyrepository
        winmgmt /salvagerepository

        Write-Host "▶ Resetting network stack & flushing DNS…" -ForegroundColor Cyan
        netsh winsock reset
        netsh int ip reset
        ipconfig /flushdns

        Write-Host "▶ Cleaning up temp and OEM folders…" -ForegroundColor Cyan

        $pathsToDelete = @("C:\SWSetup", "C:\Temp", "C:\system.sav")
        foreach ($path in $pathsToDelete) {
            if (Test-Path $path) {
                try {
                    Remove-Item -Path $path -Recurse -Force -ErrorAction Stop
                    Write-Host "   ✔ Deleted: $path"
                } catch {
                    Write-Warning ("   ⚠ Failed to delete " + $path + ": " + $_)
                }
            }
        }

        $windowsTemp = "C:\Windows\Temp"
        if (Test-Path $windowsTemp) {
            try {
                Get-ChildItem -Path $windowsTemp -Recurse -Force -ErrorAction Stop | Remove-Item -Recurse -Force -ErrorAction Stop
                Write-Host "   ✔ Cleared contents of C:\Windows\Temp"
            } catch {
                Write-Warning ("   ⚠ Failed to clean C:\Windows\Temp: " + $_)
            }
        }

        $extraCleanPaths = @(
            "$env:TEMP",
            "C:\Windows\SoftwareDistribution\Download",
            "C:\Windows\Prefetch",
            "C:\Windows\Logs\CBS",
            "$env:LOCALAPPDATA\Microsoft\Windows\INetCache",
            "$env:LOCALAPPDATA\Microsoft\Windows\WebCache",
            "C:\ProgramData\Microsoft\Windows\WER\ReportQueue",
            "$env:LOCALAPPDATA\CrashDumps",
            "$env:LOCALAPPDATA\Microsoft\Windows\DeliveryOptimization\Cache"
        )

        foreach ($folder in $extraCleanPaths) {
            if (Test-Path $folder) {
                try {
                    Get-ChildItem -Path $folder -Recurse -Force -ErrorAction Stop |
                        Remove-Item -Recurse -Force -ErrorAction Stop
                    Write-Host "   ✔ Cleaned contents of $folder"
                } catch {
                    Write-Warning ("   ⚠ Failed to clean " + $folder + ": " + $_)
                }
            }
        }

        Write-Host "▶ Performing SSD trim (retrim) on system drive…" -ForegroundColor Cyan
        try {
            Optimize-Volume -DriveLetter $sysDriveLetter -ReTrim -Verbose -ErrorAction Stop
            Write-Host "   ✔ SSD Trim completed on drive $sysDriveLetter."
        } catch {
            Write-Warning ("   ⚠ SSD Trim failed: " + $_)
        }

        Write-Host "▶ Checking for corrupt scheduled tasks…" -ForegroundColor Cyan
        try {
            Get-ScheduledTask | Where-Object { $_.State -eq 'Unknown' } | ForEach-Object {
                Write-Warning ("⚠ Corrupt or orphaned scheduled task: " + $_.TaskName)
            }
        } catch {
            Write-Warning ("⚠ Failed to evaluate scheduled tasks: " + $_)
        }

        Write-Host "▶ Rebuilding icon cache…" -ForegroundColor Cyan
        try {
            Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
            Remove-Item "$env:LOCALAPPDATA\IconCache.db" -Force -ErrorAction SilentlyContinue
            Start-Process explorer
            Write-Host "   ✔ Icon cache rebuilt."
        } catch {
            Write-Warning ("⚠ Failed to rebuild icon cache: " + $_)
        }

        Write-Host "▶ Disabling Windows Search indexing…" -ForegroundColor Cyan
        try {
            Stop-Service WSearch -Force -ErrorAction SilentlyContinue
            Set-Service WSearch -StartupType Disabled
            Write-Host "   ✔ Windows Search service disabled."
        } catch {
            Write-Warning ("⚠ Failed to disable search indexing: " + $_)
        }

        Write-Host "▶ Clearing DNS Resolver Cache…" -ForegroundColor Cyan
        try {
            Clear-DnsClientCache
            Write-Host "   ✔ DNS cache cleared."
        } catch {
            Write-Warning ("⚠ Failed to clear DNS cache: " + $_)
        }

        Write-Host "▶ Enabling Storage Sense…" -ForegroundColor Cyan
        try {
            Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\StorageSense\Parameters\StoragePolicy" `
                             -Name "01" -Value 1 -Force
            Write-Host "   ✔ Storage Sense enabled."
        } catch {
            Write-Warning ("⚠ Failed to enable Storage Sense: " + $_)
        }

        Write-Host "▶ Resetting Windows Firewall rules…" -ForegroundColor Cyan
        try {
            netsh advfirewall reset
            Write-Host "   ✔ Firewall rules reset."
        } catch {
            Write-Warning ("⚠ Failed to reset firewall rules: " + $_)
        }

        Write-Host "▶ Re-enabling Windows Defender…" -ForegroundColor Cyan
        try {
            Set-MpPreference -DisableRealtimeMonitoring $false
            Write-Host "   ✔ Defender real-time protection enabled."
        } catch {
            Write-Warning ("⚠ Failed to enable Defender protection: " + $_)
        }

        Write-Host "▶ Removing provisioned bloatware apps…" -ForegroundColor Cyan
        try {
            Get-AppxProvisionedPackage -Online |
            Where-Object DisplayName -Match "XBox|Zune|Skype|Weather" |
            ForEach-Object {
                Remove-AppxProvisionedPackage -Online -PackageName $_.PackageName -ErrorAction SilentlyContinue
                Write-Host ("   ✔ Removed provisioned package: " + $_.DisplayName)
            }
        } catch {
            Write-Warning ("⚠ Failed to remove provisioned packages: " + $_)
        }

        Write-Host "▶ Clearing all event logs…" -ForegroundColor Cyan
        try {
            Get-WinEvent -ListLog * | ForEach-Object {
                try {
                    Clear-EventLog -LogName $_.LogName
                    Write-Host ("   ✔ Cleared log: " + $_.LogName)
                } catch {
                    Write-Warning ("⚠ Failed to clear log: " + $_.LogName)
                }
            }
        } catch {
            Write-Warning ("⚠ Failed to enumerate event logs: " + $_)
        }

    } catch {
        Write-Error ("Maintenance failed: " + $_)
    }
}

Invoke-SystemMaintenance -ArchiveLogs -Verbose