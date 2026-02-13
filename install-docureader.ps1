#!/usr/bin/env powershell
<#
.SYNOPSIS
    DocuReader installer script - extracts and registers the application
    
.DESCRIPTION
    This script installs DocuReader to Program Files, creates Start Menu shortcuts,
    and optionally creates a Desktop shortcut.
    
.EXAMPLE
    .\install-docureader.ps1
    
.NOTES
    Requires Administrator privileges
#>

[CmdletBinding()]
param(
    [switch]$DesktopShortcut = $true,
    [switch]$NoStartMenu = $false
)

# Check for admin privileges
$Principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $Principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "Error: This script requires Administrator privileges." -ForegroundColor Red
    Write-Host "Please run as Administrator and try again." -ForegroundColor Red
    exit 1
}

$ErrorActionPreference = "Stop"

# Installation paths
$AppName = "DocuReader"
$ProgramFilesDir = Join-Path $env:ProgramFiles $AppName
$SourceDir = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "freeze_build\cx_freeze"

Write-Host "=== DocuReader Installer ===" -ForegroundColor Cyan
Write-Host "Installing to: $ProgramFilesDir" -ForegroundColor Yellow

# Check source exists
if (-not (Test-Path $SourceDir)) {
    Write-Host "Error: Source directory not found: $SourceDir" -ForegroundColor Red
    exit 1
}

try {
    # Create installation directory
    if (Test-Path $ProgramFilesDir) {
        Write-Host "Removing previous installation..." -ForegroundColor Yellow
        Remove-Item -Recurse -Force $ProgramFilesDir
    }
    
    Write-Host "Creating installation directory..." -ForegroundColor Yellow
    New-Item -ItemType Directory -Path $ProgramFilesDir -Force | Out-Null
    
    # Copy application files
    Write-Host "Copying application files..." -ForegroundColor Yellow
    Copy-Item -Path "$SourceDir\*" -Destination $ProgramFilesDir -Recurse -Force
    
    # Create Start Menu shortcuts
    if (-not $NoStartMenu) {
        $StartMenuDir = Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\Programs\$AppName"
        Write-Host "Creating Start Menu shortcuts..." -ForegroundColor Yellow
        
        if (Test-Path $StartMenuDir) {
            Remove-Item -Recurse -Force $StartMenuDir
        }
        New-Item -ItemType Directory -Path $StartMenuDir -Force | Out-Null
        
        # Create shortcuts using WScript.Shell COM
        $Shell = New-Object -ComObject WScript.Shell
        $Shortcut = $Shell.CreateShortcut("$StartMenuDir\$AppName.lnk")
        $Shortcut.TargetPath = Join-Path $ProgramFilesDir "DocuReader.exe"
        $Shortcut.WorkingDirectory = $ProgramFilesDir
        $Shortcut.IconLocation = Join-Path $ProgramFilesDir "DocuReader.exe"
        $Shortcut.Save()
        
        # Create uninstall shortcut
        $UninstallShortcut = $Shell.CreateShortcut("$StartMenuDir\Uninstall.lnk")
        $UninstallShortcut.TargetPath = "powershell.exe"
        $UninstallShortcut.Arguments = "-NoProfile -ExecutionPolicy Bypass -Command ""& {Remove-Item -Recurse -Force '$ProgramFilesDir'; Remove-Item -Recurse -Force '$StartMenuDir';}"""
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
        $Shortcut.TargetPath = Join-Path $ProgramFilesDir "DocuReader.exe"
        $Shortcut.WorkingDirectory = $ProgramFilesDir
        $Shortcut.IconLocation = Join-Path $ProgramFilesDir "DocuReader.exe"
        $Shortcut.Save()
        Write-Host "Desktop shortcut created." -ForegroundColor Green
    }
    
    Write-Host ""
    Write-Host "=== Installation Complete ===" -ForegroundColor Green
    Write-Host "DocuReader has been installed to: $ProgramFilesDir" -ForegroundColor Green
    Write-Host "You can launch it from the Start Menu or Desktop shortcut." -ForegroundColor Green
    
} catch {
    Write-Host "Error during installation: $_" -ForegroundColor Red
    exit 1
}
