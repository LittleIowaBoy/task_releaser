# DocuReader Improvement Roadmap

Status key: `[ ]` Not started · `[~]` In progress · `[x]` Complete

---

## High Impact / Low Effort

- [x] **1. Interactive column sort on header click** — Hook `QHeaderView.sectionClicked` to toggle ascending/descending sort on any column. Data is already a DataFrame; only UI wiring is needed.
- [x] **2. Row filter bar (Ctrl+F)** — Input above the table that hides non-matching rows. DataFrame is in memory; filtering is trivial. Replaces the need to scroll a 200-row table.
- [x] **3. Column width persistence per template** — Save/restore per-template column widths in `QSettings`. Prevents `resizeColumnsToContents()` from resetting manual sizing on each analysis run.
- [x] **4. Row count / status summary footer** — Status bar line showing `142 rows | 6 checked | 12 task IDs ready` so analysts have instant context without scrolling.

---

## High Impact / Medium Effort

- [x] **5. File system watcher (auto-refresh Downloads)** — `QFileSystemWatcher` on `~/Downloads` repopulates the file combo the moment a new CSV/XLSX appears. No more manual "Refresh Files" after every WMS export.
- [x] **6. Configurable source folder** — The Downloads path is hardcoded in `populate_downloads_files()`. A persistent folder picker lets users point at a network share, SharePoint sync, or WMS drop directory.
- [ ] **7. Session save / restore** — Serialize checked rows, strikethrough state, and cell edits to a small JSON file keyed by source filename so analysts can hand off mid-shift.
- [ ] **8. Keyboard shortcuts** — `Ctrl+F` filter focus, `Space` toggle Done? on selected row(s), `Ctrl+E` export, `Ctrl+R` refresh files. Currently the entire workflow is mouse-only.

---

## Medium Impact / Medium Effort

- [x] **9. Template condition builder UI** — Column picker + operator dropdown + value field in the Templates dialog. Makes highlight rule authoring accessible without knowing the `"col OP value"` syntax.
- [ ] **10. Template import / export** — One-click "Export template…" / "Import template…" in the Templates dialog so teams can share configs as files or email attachments without editing JSON manually.
- [ ] **11. Comparison / diff mode** — Accept two files of the same template and highlight new rows (green), resolved rows (strikethrough), changed cells (yellow). Replaces manual side-by-side comparison.
- [ ] **12. Summary statistics panel** — Collapsible panel showing total rows, rows by location prefix, items below replen threshold, % tasks ready to release.

---

## Longer-Term / Higher Effort

- [x] **13. Multi-tab interface** — `QTabWidget` wrapping the central widget so multiple analyses can be open simultaneously without relaunching.
- [x] **14. Scheduled / silent auto-analysis** — Combine file watcher (#5) with template auto-detection: when a matching file lands in the watched folder, run analysis in a background thread and badge the taskbar icon.
- [ ] **15. PDF / print export** — `QTextDocument` + `QPrinter` for paginated PDF output. Serves shift-report distribution without needing to open Excel.
- [ ] **16. System tray integration** — Run minimized to tray with notification bubbles when new matching files appear.

---

## Code Health

- [x] **17. Remove or redirect diagnostic `print()` calls in `tr.py`** — `read_excel()`, `filter_by_condition()`, `get_values()`, and `get_multiple_columns()` all call `print()`. In the frozen app these go nowhere. The `display()` method is effectively dead code in the GUI context. Options: remove the print calls entirely, or route errors through a logging signal. See notes below.
- [x] **18. Expand undo history limit and scope** — Raise the 10-step `UNDO_HISTORY_LIMIT` and extend undo to cover column sorts and filter changes.
