@echo off
setlocal enabledelayedexpansion

:: ====================================================
:: PowerShell Script Launcher for Office Privacy & Telemetry Disabler
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

:: Check for PowerShell 7 first (preferred)
if exist "%PS7_PATH%" (
    set "PS_EXE=%PS7_PATH%"
    set "PS_SCRIPT=%SCRIPT_DIR%script\office_privacy_telemetry_disabler.ps1"
    set "PS_VERSION=PowerShell 7"
    goto :found_powershell
)

:: Check for PowerShell 7 Preview
if exist "%PS7_PREVIEW_PATH%" (
    set "PS_EXE=%PS7_PREVIEW_PATH%"
    set "PS_SCRIPT=%SCRIPT_DIR%script\office_privacy_telemetry_disabler.ps1"
    set "PS_VERSION=PowerShell 7 Preview"
    goto :found_powershell
)

:: Check for PowerShell 5
if exist "%PS5_PATH%" (
    set "PS_EXE=%PS5_PATH%"
    set "PS_SCRIPT=%SCRIPT_DIR%script\office_privacy_telemetry_disabler.ps1"
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
:: Check if the PowerShell script exists
if not exist "%PS_SCRIPT%" (
    echo [ERROR] PowerShell script not found: %PS_SCRIPT%
    echo.
    echo Make sure office_privacy_telemetry_disabler.ps1 is in the 'script' subdirectory.
    echo.
    pause
    exit /b 1
)

echo.
echo ====================================================
echo    Office Privacy and Telemetry Disabler Launcher
echo.
echo                      by EXLOUD
echo              https://github.com/EXLOUD
echo.
echo ====================================================
echo.
echo Using: %PS_VERSION%
echo Script location: %PS_SCRIPT%
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
echo [INFO] Launching Office Privacy Disabler on %PS_VERSION% ...
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