# DocuReader

**Version: 0.6.0**

Inventory DocuReader is a PyQt6 GUI and Excel/CSV parser for inventory analysis
and replenishment workflows.

## Features

- **Load CSV/XLSX/XLS files** from your Downloads folder. The most recent files
  are listed first.
- **Multi-document parsing**: select one primary file and add as many additional
  files as you need using the **Add files...** button. All files are concatenated
  into a single dataset before analysis — useful when monitoring multiple towers
  (e.g. East and West) at the same time. Rows and task IDs from all files are
  combined automatically.
- **Two analysis modes**:
  - *Chase Tasks Needing Released* — compares Active OHB vs Allocated across all
    selected files and returns the combined list of task IDs ready to release,
    plus a summary of items that still need replenishment.
  - *The Whole Table* — displays every row across all selected files with
    template-driven column ordering, sorting, and conditional cell highlights.
- **Template-driven views**: column drop / rename / reorder / `columns_last`
  (pin specific columns to the far right) / sort and conditional cell
  highlighting are all defined per file category in JSON. Built-in templates
  ship for:
  - Replenishment Audit
  - Chase Tasks (Active OHB / Allocated)
  - Locked Full Container Chase Tasks
  - Export – Tasks Assigned to LOCKED *(new in 0.4.0 — moves the Task column
    to the last position before the Done? column)*
  - Users can author their own templates without changing any code.
- **Auto-detection** of the right template based on filename glob pattern and
  column signature (longest `required_columns` match wins).
- **Multi-sheet Excel workbooks** prompt for the sheet to load; the last choice
  is remembered per file.
- **Done? checkbox column** in The Whole Table view. Checking a row marks it
  with red strikethrough. For files that have a `location_columns` field set,
  checking any row in a location group marks the entire group at once. Works
  for all document types, including task exports that have no location column.
- **Export view...** writes the currently displayed table to CSV or XLSX,
  preserving template highlights as openpyxl cell fills.
- **Batch export** applies each file's matched template and writes one
  `.view.xlsx` per source file to a chosen output folder.
- **Auto-update from GitHub Releases**: the in-app **Tools → Check & Install Updates**
  menu item downloads the latest signed ZIP, verifies its SHA-256, and stages it
  for installation. No admin / UAC prompt required.

## Installation

### Windows portable (recommended for end users)

1. Download `DocuReader-<version>-portable.zip` from the
   [Releases page](https://github.com/LittleIowaBoy/task_releaser/releases).
2. Right-click the ZIP and extract it.
3. Open the extracted folder and run **`install-docureader.bat`** (or
   `install-docureader.ps1` in PowerShell).
   - Installs to `%LOCALAPPDATA%\Programs\DocuReader` — no admin rights needed.
   - Creates Start Menu and Desktop shortcuts automatically.
4. Launch **DocuReader** from the Start Menu or Desktop shortcut.

> Older installs located at `%ProgramFiles%\DocuReader` are auto-migrated to the
> per-user path on first launch so that automatic updates can run without UAC.

### From source
```bash
pip install docureader
docureader
```

## Requirements
- Python 3.10+
- PyQt6 >= 6.5
- pandas >= 2.0, openpyxl >= 3.1, xlrd >= 2.0

## Templates

Bundled defaults live in [default_templates.json](default_templates.json).
Per-user overrides are stored at `~/.docureader/templates.json`. New bundled
templates added in future releases are merged in by name without overwriting
any user-edited template.

Edit templates in-app via **Tools → Templates...** (raw-JSON editor with
new / delete / import / export). A template is a JSON object like:

```jsonc
{
  "name": "My Report",
  "filename_patterns": ["*MyReport*.xlsx"],
  "required_columns": ["Task ID", "Item"],
  "drop": ["Notes"],
  "rename": {"Task ID": "TASK_ID"},
  "order": ["TASK_ID", "Item", "Quantity"],
  "columns_last": ["Task"],
  "sort_by": [["TASK_ID", "asc"]],
  "location_columns": ["Bin", "Aisle"],
  "highlights": [
    {
      "name": "OK status",
      "when": "Status == 'OK'",
      "target_columns": ["Status"],
      "color": "darkgreen",
      "priority": 10
    }
  ]
}
```

| Field | Purpose |
|---|---|
| `filename_patterns` | Glob patterns matched against the loaded filename |
| `required_columns` | Columns that must be present for this template to match |
| `drop` | Columns to remove before display |
| `rename` | Mapping of old → new column names |
| `order` | Preferred column order (unspecified columns follow in original order) |
| `columns_last` | Columns pinned to the far right, after all other columns |
| `sort_by` | `[["column", "asc\|desc"], ...]` pairs |
| `location_columns` | Enables location-aware natural sort and divider rows |
| `highlights` | Conditional formatting rules (see below) |

`columns_last` and `order` can be combined. Columns in `columns_last` are
excluded from the `order` pass and appended after all remaining columns.

## Updating

### Frozen install — end users
Click **Check & Install Updates** in the app. The updater will:
1. Query the GitHub Releases API for the latest version.
2. Download `DocuReader-<version>-portable.zip` and `SHA256SUMS.txt`.
3. Verify the SHA-256 checksum (refuses the update if it doesn't match).
4. Stage the new files under `%LOCALAPPDATA%\DocuReader\updates\`.
5. On next launch, replace the running install with the new version.

To include pre-release / beta builds, tick **Tools → Include Pre-Releases**
before triggering the update.

Command-line equivalent (run from the install folder):
```
update_github.exe --check-only          # print available version only
update_github.exe --yes                 # update without confirmation prompt
update_github.exe --include-prereleases # include beta / RC releases
```

### Source checkout — developers
```bash
python update.py           # pull latest release tag and rebuild
python update.py --check-only
python update.py --yes
python update.py --allow-dirty
python update.py --force-rebuild
python update.py --rollback
```

## Building Windows binaries

```bash
pip install -e ".[build]"
python rebuild_and_package.py
```

Outputs:
- `freeze_build/DocuReader/DocuReader.exe`
- `freeze_build/DocuReader/update_github.exe`
- `DocuReader-<version>-portable.zip`
- `SHA256SUMS.txt`

## Tests

```bash
pip install -e ".[dev]"
python -m pytest tests -q
```

## Developer Release Process

The version lives in a single file: [_version.py](_version.py). Bump it,
commit, tag, and push:

```bash
# edit _version.py
git add _version.py
git commit -m "vX.Y.Z: <summary>"
git tag vX.Y.Z
git push
git push origin vX.Y.Z
```

The [.github/workflows/release.yml](.github/workflows/release.yml) workflow
fires on the tag push, builds the Windows portable ZIP via
`rebuild_and_package.py`, computes `SHA256SUMS.txt`, and attaches both to a
GitHub Release. Existing installs pick the new release up automatically the
next time a user clicks **Check & Install Updates**.

## Changelog

### 0.5.1
- **Dark-mode status text**: the status banner at the top of the window now
  detects the active color theme at startup and uses white text on dark themes,
  black text on light themes — instead of a hard-coded gray.
- **Auto-refresh on Add files**: clicking **Add files...** now silently
  refreshes the primary file list before the browser dialog opens, so
  newly-downloaded files are always indexed and the table is constructed
  from up-to-date file paths.

### 0.5.0
- **Tools menu**: moved *Check & Install Updates*, *Include Pre-Releases*,
  *Templates...*, *Export View...*, *Batch Export*, and *Terminate* off the
  main window and into a **Tools** drop-down on the menu bar. The button row
  is now much narrower.
- **Responsive window sizing**: the initial window is sized as a percentage of
  the user's available screen area and centered precisely after the window
  frame is established, so it never opens off-screen.
- **Collapsible log pane**: raw activity output is hidden by default. Click
  **Show Log** in the button row to expand a 150 px log pane beneath the
  table; click **Hide Log** to collapse it again.
- **Cleaner log output**: technical details (row/column counts, column name
  lists, template match reasons) are no longer written to the log. The log
  now shows only user-relevant messages: file names loaded, detected template,
  replenishment summary, and errors.
- **Status banner**: the status label is now centered, bold, and italic —
  acting as a top-of-window activity indicator.
- **Undo / Redo** (10 steps, Ctrl+Z / Ctrl+Y): tracks both checkbox toggles
  and cell-text edits made in the table.

### 0.4.0
- **Multi-document parsing**: add extra files alongside the primary selection;
  all files are concatenated before analysis. Enables east + west tower
  parsing in a single run.
- **Renamed** the checkbox column from *Counted?* to *Done?*.
- **Fixed** strikethrough not applying for document types without a location
  column (e.g. task exports).
- **New template field** `columns_last`: pins named columns to the rightmost
  position in the table, before the *Done?* column.
- **New built-in template** *Export – Tasks Assigned to LOCKED*: matches
  `Export-Tasks assigned to LOCKED` filenames and moves the Task column last.

### 0.3.0
- Template system with JSON-defined column ordering, sorting, highlights.
- GitHub Releases auto-updater with SHA-256 verification.
- Per-user install path; auto-migration from legacy `%ProgramFiles%` location.
- Multi-sheet workbook support with remembered sheet selection.
- Batch export mode.

### 0.2.x
- Locked Full Container Chase Tasks support.
- Location-aware natural sort with divider rows.
- Theme-aware strikethrough restore.
