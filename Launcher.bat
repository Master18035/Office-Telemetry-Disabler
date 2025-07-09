@echo off
setlocal enabledelayedexpansion

:: ====================================================
:: Simplified PowerShell Script Launcher
:: ====================================================

title Office Privacy and Telemetry Disabler Launcher

:: Set script directory
set "SCRIPT_DIR=%~dp0"

:: Define PowerShell paths
set "PS7_PATH=C:\Program Files\PowerShell\7\pwsh.exe"
set "PS7_PREVIEW_PATH=C:\Program Files\PowerShell\7-preview\pwsh.exe"
set "PS5_PATH=C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"

:: Initialize variables
set "PS_EXE="
set "PS_SCRIPT="
set "PS_VERSION="

:: ====================================================
:: Find PowerShell Executable
:: ====================================================

:: Check for PowerShell 7 first (preferred)
if exist "%PS7_PATH%" (
    set "PS_EXE=%PS7_PATH%"
    set "PS_VERSION=PowerShell 7"
    goto :found_powershell
)

:: Check for PowerShell 7 Preview
if exist "%PS7_PREVIEW_PATH%" (
    set "PS_EXE=%PS7_PREVIEW_PATH%"
    set "PS_VERSION=PowerShell 7 Preview"
    goto :found_powershell
)

:: Check for PowerShell 5
if exist "%PS5_PATH%" (
    set "PS_EXE=%PS5_PATH%"
    set "PS_VERSION=PowerShell 5"
    goto :found_powershell
)

:: No PowerShell found
echo [ERROR] No compatible PowerShell version found!
echo.
echo Please install either:
echo  - PowerShell 7 (recommended)
echo  - PowerShell 5 (Windows PowerShell)
echo.
pause
exit /b 1

:found_powershell

:: ====================================================
:: Find PowerShell Script
:: ====================================================

:: Try to find the appropriate script file
set "SCRIPT_FOUND="

:: Check for Windows 10/11 script first
set "TEST_SCRIPT=%SCRIPT_DIR%office_privacy_telemetry_disabler.ps1"
if exist "%TEST_SCRIPT%" (
    set "PS_SCRIPT=%TEST_SCRIPT%"
    set "SCRIPT_FOUND=YES"
    set "SCRIPT_TYPE=Windows 10/11"
    goto :script_found
)

:: Check in script subdirectory
set "TEST_SCRIPT=%SCRIPT_DIR%script\office_privacy_telemetry_disabler.ps1"
if exist "%TEST_SCRIPT%" (
    set "PS_SCRIPT=%TEST_SCRIPT%"
    set "SCRIPT_FOUND=YES"
    set "SCRIPT_TYPE=Windows 10/11 (from script folder)"
    goto :script_found
)

:: Check for Windows 7+ script
set "TEST_SCRIPT=%SCRIPT_DIR%office_privacy_telemetry_disabler_win7+.ps1"
if exist "%TEST_SCRIPT%" (
    set "PS_SCRIPT=%TEST_SCRIPT%"
    set "SCRIPT_FOUND=YES"
    set "SCRIPT_TYPE=Windows 7/8/8.1"
    goto :script_found
)

:: Check in script subdirectory
set "TEST_SCRIPT=%SCRIPT_DIR%script\office_privacy_telemetry_disabler_win7+.ps1"
if exist "%TEST_SCRIPT%" (
    set "PS_SCRIPT=%TEST_SCRIPT%"
    set "SCRIPT_FOUND=YES"
    set "SCRIPT_TYPE=Windows 7/8/8.1 (from script folder)"
    goto :script_found
)

:: No script found
echo [ERROR] No PowerShell script found!
echo.
echo Please make sure one of these files exists:
echo  - office_privacy_telemetry_disabler.ps1 (for Windows 10/11)
echo  - office_privacy_telemetry_disabler_win7+.ps1 (for Windows 7/8/8.1)
echo.
echo Either in the same directory as this launcher or in a 'script' subdirectory.
echo.
pause
exit /b 1

:script_found

:: ====================================================
:: Display Information and Confirmation
:: ====================================================

echo.
echo ====================================================
echo    Office Privacy and Telemetry Disabler Launcher
echo.
echo                      by EXLOUD
echo              https://github.com/EXLOUD
echo.
echo ====================================================
echo.
echo System Information:
echo  - PowerShell: %PS_VERSION%
echo  - Script: !SCRIPT_TYPE!
echo  - Location: !PS_SCRIPT!
echo.
echo This will disable telemetry and privacy features for:
echo  - Microsoft Office 2010-2024
echo  - Office logging and telemetry
echo  - Customer Experience Improvement Program
echo  - Connected Experiences
echo  - Automatic updates and notifications
echo  - Scheduled telemetry tasks
echo.

:confirmation
set /p "CONFIRM=Do you want to continue? (Y/N): "

if /i "!CONFIRM!"=="y" goto :proceed
if /i "!CONFIRM!"=="yes" goto :proceed
if /i "!CONFIRM!"=="n" goto :cancel
if /i "!CONFIRM!"=="no" goto :cancel

echo Invalid input. Please enter Y or N.
goto :confirmation

:cancel
echo.
echo Operation cancelled by user.
pause
exit /b 0

:proceed
cls

echo.
echo [INFO] Launching Office Privacy Disabler...
echo [INFO] PowerShell: %PS_VERSION%
echo [INFO] Script: !SCRIPT_TYPE!
echo.
echo [WARNING] Administrator rights may be required for some registry changes.
echo.

:: Change directory to script location
cd /d "%SCRIPT_DIR%"

:: Launch PowerShell script with execution policy bypass
"%PS_EXE%" -ExecutionPolicy Bypass -NoProfile -File "%PS_SCRIPT%"

:: Check exit code
if %errorLevel% == 0 (
    echo.
    echo [SUCCESS] Script completed successfully!
    echo.
    echo Office privacy and telemetry settings have been disabled.
    echo Some changes may require restarting Office applications.
) else (
    echo.
    echo [ERROR] Script encountered errors. Exit code: %errorLevel%
    echo.
    echo This may happen if:
    echo  - Office is not installed
    echo  - Administrator rights are required
    echo  - Registry access is restricted
)

echo.
echo Press any key to exit...
pause >nul
exit /b 0
