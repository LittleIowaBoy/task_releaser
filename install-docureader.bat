@echo off
REM DocuReader Installer Batch Wrapper
REM Per-user installation - no Administrator rights required.

setlocal enabledelayedexpansion

echo.
echo === DocuReader Installer (per-user) ===
echo.

REM Get script directory
set "SCRIPT_DIR=%~dp0"

REM Run PowerShell installer (no elevation needed)
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
