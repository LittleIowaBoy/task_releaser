# DocuReader

**Version: 0.2.3**

Inventory DocuReader is a PyQt6 GUI and Excel/CSV parser for inventory analysis and replenishment workflows.

## Features
- Load CSV/XLSX/XLS files from Downloads.
- Display parsed data in a table with location-aware ordering.
- Highlight replenishment timing against short time.
- Copy matched Task IDs to clipboard.

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
1. Download `DocuReader-0.2.3-portable.zip`
2. Extract anywhere (for example `C:\Users\YourName\DocuReader`)
3. Run `DocuReader.exe`

### Option 2: Run Directly from Build Output
```text
freeze_build/cx_freeze/DocuReader.exe
```

### Option 3: Build Locally
```bash
pip install cx_Freeze
python freeze_setup.py build
```

Build output:
- `freeze_build/cx_freeze/DocuReader.exe`
- `freeze_build/cx_freeze/update.exe`

## Update & Maintenance

### Check for Updates
```bash
python update.py --check-only
```

### Install Updates
```bash
python update.py
```

The updater may reset the repository to a release tag. If local uncommitted changes are present, update is blocked by default.

Optional flags:
- `--yes` skip interactive confirmation prompt.
- `--allow-dirty` allow updates with uncommitted local changes.
- `--force-rebuild` rebuild executable even when no source changes are detected.
- `--rollback` restore latest backup.

### Pip Entry Points
- `docureader` launches the GUI.
- `docureader-update` runs the updater.

### Update Log
All update operations are written to `update.log`.

## Developer Release Notes
When releasing:
1. Update version in `tr_gui.py` and `pyproject.toml`.
2. Validate GUI and updater behavior.
3. Commit and tag release (example: `git tag -a v0.2.3 -m "Release 0.2.3"`).
4. Push tag to remote.
