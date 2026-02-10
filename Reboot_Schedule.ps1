# Create/replace weekly reboot tasks: Sunday 8pm/9pm + Monday 4am
$tasks = @(
    @{ Name = 'Reboot-Sun-2000'; Day = 'Sunday' ; At = '8:00 PM' ; Desc = 'Weekly reboot every Sunday at 8:00 PM'  },
    @{ Name = 'Reboot-Sun-2100'; Day = 'Sunday' ; At = '9:00 PM' ; Desc = 'Weekly reboot every Sunday at 9:00 PM'  },
    @{ Name = 'Reboot-Mon-0400'; Day = 'Monday' ; At = '4:00 AM' ; Desc = 'Weekly reboot every Monday at 4:00 AM'  }
)

$action    = New-ScheduledTaskAction -Execute 'shutdown.exe' -Argument '/r /t 0 /f'
$principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
$settings  = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
                                          -StartWhenAvailable -WakeToRun:$false -Compatibility Win8

foreach ($t in $tasks) {
    $trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek $t.Day -At $t.At

    # Replace if it already exists
    $existing = Get-ScheduledTask -TaskName $t.Name -ErrorAction SilentlyContinue
    if ($existing) {
        Unregister-ScheduledTask -TaskName $t.Name -Confirm:$false
    }

    Register-ScheduledTask -TaskName $t.Name -Action $action -Trigger $trigger `
        -Principal $principal -Settings $settings -Description $t.Desc | Out-Null

    Write-Host "✔ Created task $($t.Name): $($t.Day) at $($t.At)"
}
