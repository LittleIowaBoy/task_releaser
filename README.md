# DocuReader

**Version: 0.2.0**

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

### Option 1: Portable ZIP (Recommended)

1. Download `DocuReader-0.2.0-portable.zip`
2. Extract anywhere (e.g., `C:\Users\YourName\DocuReader` or `C:\Program Files\DocuReader`)
3. Run `DocuReader.exe` directly — no installation or admin privileges required
4. (Optional) Create a shortcut to `DocuReader.exe` and place on Desktop or Start Menu for quick access

### Option 2: Portable Executable
No installation needed! Run directly from the build output:
```
freeze_build/DocuReader/DocuReader.exe
```

### Option 3: Legacy Installer Scripts (Deprecated)

The PowerShell and Batch installer scripts are no longer recommended. Please use the Portable ZIP option instead.

If you still want to use them for system-wide installation:
```powershell
# Run as Administrator:
.\install-docureader.ps1
```

### Option 4: Build a New Bundled Executable

For development or to create a custom build:

Install build dependency:
```bash
pip install cx_Freeze
```

Create a bundled executable:
```bash
python freeze_setup.py build
```

Output: `freeze_build/DocuReader/` containing portable executable (~100MB)
## Update & Maintenance

### Checking for Updates

The application includes an automated update system that checks against the GitHub repository ('task_releaser'). 

**Check for available updates:**
\\\ash
python update.py --check-only
\\\

This will display the current version and latest available version without making any changes.

### Installing Updates

**Perform automatic update:**
\\\ash
python update.py
\\\

The script will:
1. Check if updates are available on GitHub
2. Backup your current version (in 'backup/' directory)
3. Pull changes from the repository
4. Automatically rebuild the executable if source code or dependencies changed
5. Validate the new build
6. Create a new portable ZIP package

You'll see progress messages and a completion status. After the update completes, restart the application to use the new version.

**Force rebuild (useful for development):**
\\\ash
python update.py --force-rebuild
\\\

This rebuilds the executable even if no source code changes are detected.

### Rollback (Reverting Updates)

If an update causes issues, you can rollback to the previous version:
\\\ash
python update.py --rollback
\\\

This restores the previously backed-up executable and dependencies.

### Update Log

All update operations are logged to 'update.log' for troubleshooting and audit purposes.

### Version Information

- Current version is displayed in the application window title
- Check 'tr_gui.py' for the '\__version__\' constant
- Versions follow semantic versioning (MAJOR.MINOR.PATCH, e.g., 0.2.0)
- GitHub releases are tagged with version numbers (v0.2.0, v0.2.1, etc.)

### For Developers

When releasing a new version:
1. Update the '\__version__\' constant in 'tr_gui.py' (e.g., from "0.2.0" to "0.2.1")
2. Test the application thoroughly
3. Commit changes to the repository
4. Tag the release: \git tag -a v0.2.1 -m "Release 0.2.1"\
5. Push tag to remote: \git push task_releaser v0.2.1\

The update script will automatically detect the new tag and make it available to users.
