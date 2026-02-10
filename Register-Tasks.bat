@echo off
setlocal

REM Set credentials
set "USERNAME=MISAdmin"
set "PASSWORD="

REM Register Weekly Tasks (Every Sunday)
call :CreateTask "Remove User Profiles Weekly"        "C:\Windows\Scripts\01_Remove-UserProfiles.ps1"       "01:30"
call :CreateTask "Weekend Apps Updates"               "C:\Windows\Scripts\02_Weekend_Apps_Update.ps1"       "02:00"
call :CreateTask "Weekend HP Drivers Update"          "C:\Windows\Scripts\03_Weekend_HP_Drivers_Update.ps1" "02:45"
call :CreateTask "Weekend Windows Updates"            "C:\Windows\Scripts\04_Weekend_Windows_Updates.ps1"   "04:00"
call :CreateTask "System Repair"                      "C:\Windows\Scripts\05_SystemRepair.ps1"              "08:00"
call :CreateTask "Force Reboot 0900"                  "C:\Windows\Scripts\06_Force_Reboot.ps1"              "09:00"
call :CreateTask "Force Reboot 1900"                  "C:\Windows\Scripts\06_Force_Reboot.ps1"              "19:00"

goto :EOF

:CreateTask
set "TASKNAME=%~1"
set "SCRIPT=%~2"
set "TIME=%~3"

schtasks /Create ^
 /TN "%TASKNAME%" ^
 /TR "powershell.exe -NoProfile -ExecutionPolicy Bypass -File \"%SCRIPT%\"" ^
 /SC WEEKLY /D SUN ^
 /ST %TIME% ^
 /RL HIGHEST ^
 /RU %USERNAME% /RP %PASSWORD% ^
 /F

if %errorlevel% equ 0 (
    echo ✅ Task "%TASKNAME%" created successfully for %TIME%.
) else (
    echo ❌ Failed to create task "%TASKNAME%".
)

goto :EOF
