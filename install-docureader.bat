@echo off
REM DocuReader Installer Batch Wrapper
REM Runs the PowerShell installation script

setlocal enabledelayedexpansion

echo.
echo === DocuReader Installer ===
echo.

REM Check for admin privileges
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo Error: This script requires Administrator privileges.
    echo Please right-click and select "Run as Administrator"
    pause
    exit /b 1
)

REM Get script directory
set "SCRIPT_DIR=%~dp0"

REM Run PowerShell installer with elevated privileges
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%install-docureader.ps1"

if %errorlevel% neq 0 (
    echo.
    echo Installation failed. Please check the error messages above.
    pause
    exit /b 1
)

echo.
echo Installation complete! Press any key to exit.
pause
