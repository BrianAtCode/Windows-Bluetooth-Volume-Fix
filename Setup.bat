@echo off
echo Windows Bluetooth Volume Fix
echo ====================================================
echo Auto‑fix via Task Scheduler (Win10)
echo.
echo This will set up automated Windows Update monitoring,
echo Bluetooth AVRCP registry management, and event-based triggering.
echo.
pause

echo Checking for Administrator privileges...
net session >nul 2>&1
if %errorLevel% == 0 (
    echo Administrator privileges confirmed.
    echo.
) else (
    echo ERROR: This script must be run as Administrator!
    echo Please right-click and select "Run as administrator"
    echo.
    pause
    exit /b 1
)

echo Starting PowerShell setup script...

cd /d %~dp0

setlocal enabledelayedexpansion

:loop_start
echo Which PowerShell version do you prefer?
echo 1. PowerShell 5.1
echo 2. PowerShell 7 or later
echo Please enter the number corresponding to your choice:

set /p  "powershell_version=" 

if "%powershell_version%" == "1" (
	echo. 
	
	PowerShell.exe -ExecutionPolicy Bypass -File "%~dp0SetupTaskScheduler.ps1"
	
	echo.
	
) else if "%powershell_version%" == "2" (
	echo. 
	
	pwsh.exe -ExecutionPolicy Bypass -File "%~dp0SetupTaskScheduler.ps1"
	
	echo.
) else (
	echo ERROR: The input option is not included, not support or incorrect.

    GOTO :loop_start
)
endlocal

echo Setup completed.
echo Your Windows 10 Bluetooth Policy is now active!
echo Check the output above for any errors.
pause
