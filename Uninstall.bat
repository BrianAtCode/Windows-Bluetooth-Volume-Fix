@echo off

echo Windows Bluetooth Volume Fix Uninstaller
echo =======================================================
echo Auto‑fix via Task Scheduler (Win10) Script - Removal
echo.
echo This will remove all scheduled tasks related to the Bluetooth Policy automation.
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

echo Starting PowerShell uninstall script...
cd /d %~dp0
setlocal enabledelayedexpansion

:run_uninstaller
:loop_start
echo Which PowerShell version do you prefer?
echo 1. PowerShell 5.1
echo 2. PowerShell 7 or later
echo Please enter the number corresponding to your choice:
set /p "powershell_version="

if "%powershell_version%" == "1" (
echo.
PowerShell.exe -ExecutionPolicy Bypass -File "%~dp0UninstallTaskScheduler.ps1"
echo.
) else if "%powershell_version%" == "2" (
echo.
pwsh.exe -ExecutionPolicy Bypass -File "%~dp0UninstallTaskScheduler.ps1"
echo.
) else (
echo ERROR: The input option is not included, not support or incorrect.
GOTO :loop_start
)

:end_script
endlocal
echo Uninstall process completed.
echo Check the output above for any errors.
pause
