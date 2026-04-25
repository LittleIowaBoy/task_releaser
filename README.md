# DocuReader

**Version: 0.3.0**

Inventory DocuReader is a PyQt6 GUI and Excel/CSV parser for inventory analysis
and replenishment workflows.

## Features
- Load CSV/XLSX/XLS files from Downloads.
- **Template-driven views**: column drop / rename / reorder / sort and
  conditional cell highlighting are all defined per file category in JSON.
  Built-in templates ship for Replenishment Audit, Chase Tasks, and
  Locked Full Container reports; users can author their own without
  changing any code.
- Auto-detection of the right template based on filename pattern + column
  signature (longest-required-columns wins; tie-break by `priority`).
- Multi-sheet Excel workbooks prompt for the sheet to load (last choice
  remembered per file).
- **Export view...** writes the currently displayed table to CSV or XLSX,
  preserving template highlights as openpyxl cell fills.
- **Auto-update from GitHub Releases**: when running the frozen build, the
  in-app "Check & Install Updates" button downloads the latest signed ZIP,
  verifies its SHA-256, stages it under `%LOCALAPPDATA%\DocuReader\updates\`,
  and swaps the install on next start. No admin / UAC prompt.

## Installation

### From source
```bash
pip install docureader
```

### Windows portable
1. Download `DocuReader-<version>-portable.zip` from the
   [Releases page](https://github.com/LittleIowaBoy/task_releaser/releases).
2. Extract anywhere (or use `install-docureader.bat` for a per-user install).
3. Run `DocuReader.exe`.

`install-docureader.bat` / `install-docureader.ps1` install to
`%LOCALAPPDATA%\Programs\DocuReader` (no admin required) and create Start
Menu + Desktop shortcuts. Older admin installs at `%ProgramFiles%\DocuReader`
are auto-migrated to the new path on first launch.

## Usage

```bash
docureader
```

## Requirements
- Python 3.10+
- PyQt6 >= 6.5
- pandas >= 2.0, openpyxl >= 3.1, xlrd >= 2.0

## Building Windows binaries

```bash
pip install -e ".[build]"
python rebuild_and_package.py
```

Outputs:
- `freeze_build/cx_freeze/DocuReader.exe`
- `freeze_build/cx_freeze/update.exe` (git-based dev updater)
- `freeze_build/cx_freeze/update_github.exe` (GitHub Releases updater)
- `DocuReader-<version>-portable.zip`
- `SHA256SUMS.txt`

## Templates

Bundled defaults live in [default_templates.json](default_templates.json).
Per-user overrides are stored at `~/.docureader/templates.json`. New
bundled templates added in future releases are merged in by name without
overwriting any user-edited template.

Edit templates in-app via the **Templates...** button (raw-JSON editor with
new/delete/import/export). A template is a JSON object like:

```jsonc
{
  "name": "My Report",
  "filename_patterns": ["*MyReport*.xlsx"],
  "required_columns": ["Task ID", "Item"],
  "drop_columns": ["Notes"],
  "rename_columns": {"Task ID": "TASK_ID"},
  "column_order": ["TASK_ID", "Item", "Quantity"],
  "sort_by": ["TASK_ID"],
  "location_columns": ["Bin", "Aisle"],
  "highlights": [
    { "color": "darkgreen", "column": "Status", "operator": "==", "value": "OK" }
  ]
}
```

`location_columns` enables location-aware sort + visual divider rows.

## Updating

### Frozen install (recommended for end users)
Click **Check & Install Updates** in the GUI, or run:
```bash
update_github.exe --check-only
update_github.exe --yes
update_github.exe --include-prereleases
```
Updates are SHA-256 verified against `SHA256SUMS.txt` published with each
release. An update with a missing or mismatched checksum is refused.

### Source checkout
```bash
python update.py --check-only
python update.py
```
The git-based updater may reset the repo to a release tag.
- `--yes` skip the confirmation prompt.
- `--allow-dirty` allow updates with uncommitted local changes.
- `--force-rebuild` rebuild even when no source changes are detected.
- `--rollback` restore latest backup.

### Console scripts
- `docureader` launches the GUI.
- `docureader-update` runs the source / git updater.
- `docureader-update-github` runs the GitHub Releases updater.

## Tests

```bash
pip install -e ".[dev]"
python -m pytest tests -q
```

## Developer Release Process
The version lives in a single file: [_version.py](_version.py). Bump it,
commit, tag, push:

```bash
git tag -a v0.3.0 -m "Release 0.3.0"
git push origin v0.3.0
```

The [.github/workflows/release.yml](.github/workflows/release.yml) workflow
runs on the tag push, builds the Windows portable ZIP via
`rebuild_and_package.py`, computes `SHA256SUMS.txt`, and attaches both to a
GitHub Release. Existing installs pick the new release up automatically the
next time a user clicks **Check & Install Updates**.
