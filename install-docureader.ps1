#!/usr/bin/env powershell
<#
.SYNOPSIS
    DocuReader installer script - extracts and registers the application for the current user.

.DESCRIPTION
    This script installs DocuReader to %LOCALAPPDATA%\Programs\DocuReader (no admin
    rights required), creates Start Menu shortcuts, and optionally creates a Desktop
    shortcut. Installing per-user lets the in-app auto-updater overwrite files without
    triggering UAC.

.EXAMPLE
    .\install-docureader.ps1

.NOTES
    No Administrator privileges required.
#>

[CmdletBinding()]
param(
    [switch]$DesktopShortcut = $true,
    [switch]$NoStartMenu = $false
)

$ErrorActionPreference = "Stop"

# Installation paths
$AppName = "DocuReader"
$InstallDir = Join-Path $env:LOCALAPPDATA "Programs\$AppName"
$LegacyInstallDir = Join-Path $env:ProgramFiles $AppName
$SourceDir = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "freeze_build\cx_freeze"

Write-Host "=== DocuReader Installer ===" -ForegroundColor Cyan
Write-Host "Installing to: $InstallDir" -ForegroundColor Yellow

# Check source exists
if (-not (Test-Path $SourceDir)) {
    Write-Host "Error: Source directory not found: $SourceDir" -ForegroundColor Red
    exit 1
}

try {
    # Warn about (but don't touch) any old admin install at Program Files.
    if (Test-Path $LegacyInstallDir) {
        Write-Host "Note: An older install at '$LegacyInstallDir' was detected." -ForegroundColor Yellow
        Write-Host "      The new per-user install will not remove it; delete it manually if desired." -ForegroundColor Yellow
    }

    # Create installation directory
    if (Test-Path $InstallDir) {
        Write-Host "Removing previous per-user installation..." -ForegroundColor Yellow
        Remove-Item -Recurse -Force $InstallDir
    }

    Write-Host "Creating installation directory..." -ForegroundColor Yellow
    New-Item -ItemType Directory -Path $InstallDir -Force | Out-Null

    # Copy application files
    Write-Host "Copying application files..." -ForegroundColor Yellow
    Copy-Item -Path "$SourceDir\*" -Destination $InstallDir -Recurse -Force

    # Create Start Menu shortcuts
    if (-not $NoStartMenu) {
        $StartMenuDir = Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\Programs\$AppName"
        Write-Host "Creating Start Menu shortcuts..." -ForegroundColor Yellow

        if (Test-Path $StartMenuDir) {
            Remove-Item -Recurse -Force $StartMenuDir
        }
        New-Item -ItemType Directory -Path $StartMenuDir -Force | Out-Null

        $Shell = New-Object -ComObject WScript.Shell
        $Shortcut = $Shell.CreateShortcut("$StartMenuDir\$AppName.lnk")
        $Shortcut.TargetPath = Join-Path $InstallDir "DocuReader.exe"
        $Shortcut.WorkingDirectory = $InstallDir
        $Shortcut.IconLocation = Join-Path $InstallDir "DocuReader.exe"
        $Shortcut.Save()

        $UninstallShortcut = $Shell.CreateShortcut("$StartMenuDir\Uninstall.lnk")
        $UninstallShortcut.TargetPath = "powershell.exe"
        $UninstallShortcut.Arguments = "-NoProfile -ExecutionPolicy Bypass -Command ""& {Remove-Item -Recurse -Force '$InstallDir'; Remove-Item -Recurse -Force '$StartMenuDir';}"""
        $UninstallShortcut.IconLocation = "systemroot\System32\shell32.dll,131"
        $UninstallShortcut.Save()

        Write-Host "Start Menu shortcuts created." -ForegroundColor Green
    }

    # Create Desktop shortcut
    if ($DesktopShortcut) {
        Write-Host "Creating Desktop shortcut..." -ForegroundColor Yellow
        $DesktopDir = [System.IO.Path]::Combine([Environment]::GetFolderPath("Desktop"))
        $Shell = New-Object -ComObject WScript.Shell
        $Shortcut = $Shell.CreateShortcut("$DesktopDir\$AppName.lnk")
        $Shortcut.TargetPath = Join-Path $InstallDir "DocuReader.exe"
        $Shortcut.WorkingDirectory = $InstallDir
        $Shortcut.IconLocation = Join-Path $InstallDir "DocuReader.exe"
        $Shortcut.Save()
        Write-Host "Desktop shortcut created." -ForegroundColor Green
    }

    Write-Host ""
    Write-Host "=== Installation Complete ===" -ForegroundColor Green
    Write-Host "DocuReader has been installed to: $InstallDir" -ForegroundColor Green
    Write-Host "You can launch it from the Start Menu or Desktop shortcut." -ForegroundColor Green
    Write-Host "In-app updates will install automatically (no admin prompt needed)." -ForegroundColor Green

} catch {
    Write-Host "Error during installation: $_" -ForegroundColor Red
    exit 1
}
