@echo off
setlocal EnableExtensions

REM Register Weekly Tasks (Every Tuesday) to run as Local System
REM Adjust the times below if you want a different order or spacing.

call :CreateTask "01. Check for Updated Scripts"           "C:\Scripts\00_Update-Scripts-FromGitHub.ps1"         "01:15"
call :CreateTask "02. Enable Windows Update Services"      "C:\Scripts\01_Enable_Windows_Update_Services.ps1"    "01:20"
call :CreateTask "03. Remove User Profiles Weekly"         "C:\Scripts\02_Remove_User_Profiles.ps1"              "01:30"
call :CreateTask "04. Weekend Apps Update"                 "C:\Scripts\03_Weekend_Apps_Update.ps1"               "02:00"
call :CreateTask "05. Update Edge Silent"                  "C:\Scripts\04_Update_Edge_Silent.ps1"                "02:45" "-KillEdgeProcesses"
call :CreateTask "06. Weekend HP Drivers Update"           "C:\Scripts\05_Weekend_HP_Drivers_Update.ps1"         "03:00"
call :CreateTask "07. Weekend Windows Updates - 1st Pass"  "C:\Scripts\06_Weekend_Windows_Updates.ps1"           "04:00"
call :CreateTask "08. Force Reboot Install Updates"        "C:\Scripts\07_Force_Reboot_Install_Updates.ps1"      "05:00"
call :CreateTask "09. Weekend Windows Updates - 2nd Pass"  "C:\Scripts\06_Weekend_Windows_Updates.ps1"           "05:30"
call :CreateTask "10. Disable Windows Update Services"     "C:\Scripts\09_Disable_Windows_Update_Services.ps1"   "06:00"
call :CreateTask "11. Force Reboot Install Updates 2"      "C:\Scripts\07_Force_Reboot_Install_Updates.ps1"      "06:05"
call :CreateTask "12. System Repair"                       "C:\Scripts\08_System_Repair.ps1"                     "06:15"
call :CreateTask "13. Force Reboot Install Updates 3"      "C:\Scripts\07_Force_Reboot_Install_Updates.ps1"      "07:00"

echo.
echo All task registration commands have completed.
goto :EOF

:CreateTask
set "TASKNAME=%~1"
set "SCRIPT=%~2"
set "TIME=%~3"
set "ARGS=%~4"
set "TASKCMD=powershell.exe -NoProfile -ExecutionPolicy Bypass -File \"%SCRIPT%\""

if not "%ARGS%"=="" set "TASKCMD=%TASKCMD% %ARGS%"

schtasks /Create /TN "%TASKNAME%" /TR "%TASKCMD%" /SC WEEKLY /D SUN /ST %TIME% /RL HIGHEST /RU SYSTEM /F

if %errorlevel% equ 0 (
    echo [OK] Task "%TASKNAME%" created successfully for %TIME% as SYSTEM.
) else (
    echo [FAIL] Failed to create task "%TASKNAME%".
)

goto :EOF