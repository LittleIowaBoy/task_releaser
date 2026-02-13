# DocuReader

Inventory DocuReader is a small PyQt6 GUI and Excel/CSV parser that helps analyze inventory data and highlight replenishment timing.

## Features
- Load CSV/XLSX/XLS files from Downloads.
- View data in a sortable table.
- Highlight replenishment timing against short time.
- Copy Task IDs to clipboard.

## Installation

```bash
pip install docureader
```

## Usage

```bash
docureader
```

## Requirements
- Python 3.10+

## Windows Installation

### Option 1: Use the Installer Script (Recommended)

#### On Windows with Batch File:
1. Right-click `install-docureader.bat` and select **"Run as Administrator"**
2. Follow the on-screen prompts
3. The app will be installed to `C:\Program Files\DocuReader\` with Start Menu + Desktop shortcuts

#### On Windows with PowerShell:
```powershell
# Run as Administrator:
.\install-docureader.ps1
```

### Option 2: Portable Executable
No installation needed! Run directly:
```
freeze_build/cx_freeze/DocuReader.exe
```

### Option 3: Build a New Bundled Executable

Install build dependency:
```bash
pip install cx_Freeze
```

Create a bundled executable (includes Python interpreter + dependencies):
```bash
python freeze_setup.py build_exe
```

Output: `freeze_build/cx_freeze/DocuReader.exe` (portable, ~50MB)
