import sys
import pandas as pd
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QTextEdit, QLabel, QComboBox, QMessageBox, QSplitter,
    QTableWidget, QTableWidgetItem, QCheckBox, QDialog, QDialogButtonBox,
    QListWidget, QListWidgetItem, QLineEdit, QFormLayout, QPlainTextEdit,
    QFileDialog, QInputDialog
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QProcess, QSettings
from PyQt6.QtGui import QFont, QTextDocument, QTextCursor, QColor
from tr import ExcelParser
from _version import __version__
from templates import (
    TemplateRegistry,
    Template,
    HighlightRule,
    USER_TEMPLATES_PATH,
)
from view import ViewMeta, apply_template, parse_location_parts


import json as _json


class TemplatesDialog(QDialog):
    """Minimal CRUD dialog over the user template store.

    Editing happens as raw JSON for the selected template - this keeps the
    dialog small while still exposing every option in :class:`templates.Template`
    (drop, rename, order, sort_by, location_columns, highlights, etc.). Wiring
    up form widgets per field can come later without changing the storage
    format.
    """

    def __init__(self, registry: TemplateRegistry, parent=None):
        super().__init__(parent)
        self.registry = registry
        self.setWindowTitle("DocuReader Templates")
        self.resize(900, 600)
        self._dirty = False

        layout = QHBoxLayout(self)

        # Left: list of templates.
        left = QVBoxLayout()
        left.addWidget(QLabel("Templates"))
        self.list_widget = QListWidget()
        self.list_widget.currentRowChanged.connect(self._on_select)
        left.addWidget(self.list_widget)

        list_buttons = QHBoxLayout()
        self.new_button = QPushButton("New")
        self.new_button.clicked.connect(self._on_new)
        list_buttons.addWidget(self.new_button)
        self.delete_button = QPushButton("Delete")
        self.delete_button.clicked.connect(self._on_delete)
        list_buttons.addWidget(self.delete_button)
        left.addLayout(list_buttons)

        io_buttons = QHBoxLayout()
        self.import_button = QPushButton("Import...")
        self.import_button.clicked.connect(self._on_import)
        io_buttons.addWidget(self.import_button)
        self.export_button = QPushButton("Export...")
        self.export_button.clicked.connect(self._on_export)
        io_buttons.addWidget(self.export_button)
        left.addLayout(io_buttons)

        layout.addLayout(left, 1)

        # Right: JSON editor for the selected template.
        right = QVBoxLayout()
        right.addWidget(QLabel(f"Editing template (stored at {USER_TEMPLATES_PATH})"))
        self.editor = QPlainTextEdit()
        self.editor.setFont(QFont("Courier New", 10))
        self.editor.textChanged.connect(self._mark_dirty)
        right.addWidget(self.editor)

        self.error_label = QLabel("")
        self.error_label.setStyleSheet("QLabel { color: #b00; }")
        right.addWidget(self.error_label)

        button_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Save
            | QDialogButtonBox.StandardButton.Close
        )
        button_box.button(QDialogButtonBox.StandardButton.Save).clicked.connect(self._on_save)
        button_box.button(QDialogButtonBox.StandardButton.Close).clicked.connect(self.accept)
        right.addWidget(button_box)

        layout.addLayout(right, 2)

        self._reload_list()

    def _reload_list(self):
        self.list_widget.clear()
        for t in self.registry.templates:
            label = t.name + ("  (built-in)" if t.builtin else "")
            self.list_widget.addItem(QListWidgetItem(label))
        if self.registry.templates:
            self.list_widget.setCurrentRow(0)

    def _current_template(self) -> Optional[Template]:
        row = self.list_widget.currentRow()
        if row < 0 or row >= len(self.registry.templates):
            return None
        return self.registry.templates[row]

    def _on_select(self, _row: int):
        t = self._current_template()
        if t is None:
            self.editor.setPlainText("")
            return
        self.editor.blockSignals(True)
        self.editor.setPlainText(_json.dumps(t.to_dict(), indent=2))
        self.editor.blockSignals(False)
        self._dirty = False
        self.error_label.setText("")

    def _mark_dirty(self):
        self._dirty = True
        self.error_label.setText("Unsaved changes.")

    def _parse_editor(self) -> Optional[Template]:
        try:
            data = _json.loads(self.editor.toPlainText())
        except _json.JSONDecodeError as e:
            self.error_label.setText(f"Invalid JSON: {e}")
            return None
        try:
            return Template.from_dict(data)
        except (KeyError, TypeError, ValueError) as e:
            self.error_label.setText(f"Invalid template: {e}")
            return None

    def _on_save(self):
        t = self._parse_editor()
        if t is None:
            return
        original = self._current_template()
        if original is not None and original.name != t.name:
            self.registry.remove(original.name)
        self.registry.upsert(t)
        try:
            self.registry.save()
        except OSError as e:
            self.error_label.setText(f"Could not write user templates: {e}")
            return
        self._dirty = False
        self.error_label.setText("Saved.")
        current_row = self.list_widget.currentRow()
        self._reload_list()
        if current_row >= 0:
            self.list_widget.setCurrentRow(min(current_row, self.list_widget.count() - 1))

    def _on_new(self):
        new = Template(
            name=f"New Template {len(self.registry.templates) + 1}",
            description="Describe when this template applies.",
            filename_patterns=["*example*"],
            required_columns=[],
            drop=[],
            order=[],
            sort_by=[],
            location_columns=[],
            highlights=[],
        )
        self.registry.upsert(new)
        try:
            self.registry.save()
        except OSError as e:
            self.error_label.setText(f"Could not write user templates: {e}")
        self._reload_list()
        self.list_widget.setCurrentRow(len(self.registry.templates) - 1)

    def _on_delete(self):
        t = self._current_template()
        if t is None:
            return
        if t.builtin:
            QMessageBox.information(
                self,
                "Built-in Template",
                "Built-in templates cannot be deleted. Edit and rename to create a copy.",
            )
            return
        if QMessageBox.question(
            self,
            "Delete Template",
            f"Delete template '{t.name}'? This cannot be undone.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        ) != QMessageBox.StandardButton.Yes:
            return
        self.registry.remove(t.name)
        try:
            self.registry.save()
        except OSError as e:
            self.error_label.setText(f"Could not write user templates: {e}")
        self._reload_list()

    def _on_import(self):
        path, _ = QFileDialog.getOpenFileName(self, "Import template", "", "JSON (*.json)")
        if not path:
            return
        try:
            data = _json.loads(Path(path).read_text(encoding="utf-8"))
            t = Template.from_dict(data)
        except (OSError, _json.JSONDecodeError, KeyError, TypeError, ValueError) as e:
            QMessageBox.critical(self, "Import failed", str(e))
            return
        self.registry.upsert(t)
        try:
            self.registry.save()
        except OSError as e:
            QMessageBox.critical(self, "Import failed", str(e))
            return
        self._reload_list()

    def _on_export(self):
        t = self._current_template()
        if t is None:
            return
        path, _ = QFileDialog.getSaveFileName(self, "Export template", f"{t.name}.json", "JSON (*.json)")
        if not path:
            return
        try:
            Path(path).write_text(_json.dumps(t.to_dict(), indent=2), encoding="utf-8")
        except OSError as e:
            QMessageBox.critical(self, "Export failed", str(e))


class WorkerThread(QThread):
    """Worker thread to run the parser without blocking the GUI"""
    finished = pyqtSignal()
    error = pyqtSignal(str)
    output = pyqtSignal(str)
    table_ready = pyqtSignal(pd.DataFrame, object)  # df, ViewMeta
    data_loaded = pyqtSignal(object)  # Emit the dataframe
    task_ids_ready = pyqtSignal(list)  # Emit the task IDs for clipboard copy
    template_matched = pyqtSignal(str, str)  # template_name, match_reason
    
    def __init__(self, filepath: str, function_name: str, registry: TemplateRegistry, sheet_name: Optional[str] = None):
        super().__init__()
        self.filepath = filepath
        self.function_name = function_name
        self.registry = registry
        self.sheet_name = sheet_name
        self.parser = None
        self.matched_template: Optional[Template] = None

    def format_df(self, df: pd.DataFrame) -> str:
        """Format a DataFrame into an aligned text table for the log pane."""
        if df is None or df.empty:
            return "No data to display\n"

        cols = list(df.columns)
        rows: List[List[str]] = []
        col_widths: Dict[str, int] = {c: len(str(c)) for c in cols}

        for _, r in df.iterrows():
            row = ["" if pd.isna(r[c]) else str(r[c]) for c in cols]
            rows.append(row)
            for i, val in enumerate(row):
                col_name = cols[i]
                if len(val) > col_widths[col_name]:
                    col_widths[col_name] = len(val)

        header_parts = [str(c).ljust(col_widths[c]) for c in cols]
        sep_parts = ["-" * col_widths[c] for c in cols]
        header = " | ".join(header_parts)
        separator = "-+-".join(sep_parts)

        lines = [header, separator]
        for row in rows:
            row_parts = [row[i].ljust(col_widths[cols[i]]) for i in range(len(cols))]
            lines.append(" | ".join(row_parts))

        return "\n".join(lines) + "\n"

    def run(self):
        try:
            self.parser = ExcelParser(self.filepath)
            if self.sheet_name:
                self.parser.read_excel(sheet_name=self.sheet_name)
            else:
                self.parser.read_excel()

            if self.parser.df is None:
                self.error.emit("Failed to load data")
                return

            extension = Path(self.filepath).suffix.lower()
            file_type = "CSV" if extension == ".csv" else "Excel"
            self.output.emit(
                f"Successfully loaded {file_type} file with {len(self.parser.df)} rows "
                f"and {len(self.parser.df.columns)} columns\n"
            )
            self.output.emit(f"Columns: {list(self.parser.df.columns)}\n")

            # Match a template based on filename + columns.
            match = self.registry.select(self.parser.df.columns, self.filepath)
            self.matched_template = match.template
            self.output.emit(f"Detected template: {match.template.name}  ({match.reason})\n")
            self.template_matched.emit(match.template.name, match.reason)

            self.execute_selected_function()
            self.data_loaded.emit(self.parser.df)

        except Exception as e:
            self.error.emit(f"Error: {str(e)}")
        finally:
            self.finished.emit()
    
    def execute_selected_function(self):
        """Execute the selected function from the dropdown."""
        try:
            if self.parser is None or self.parser.df is None:
                self.output.emit("Parser is not initialized.\n")
                return

            template = self.matched_template
            template_name = template.name if template else ""

            if self.function_name == "get_task_ids_where_condition":
                # Special-case the Locked Full Container template (was an inline
                # column-set sniff in tr.py / tr_gui.py; now driven by template name).
                if template_name == "Locked Full Container Chase Tasks":
                    self.output.emit("\n--- Item Condition Summary ---\n")
                    self.output.emit("Not applicable for Locked Full Container Chase Tasks.\n")
                    task_ids = self.parser.get_unique_numeric_values("TASK_ID")
                    self.output.emit("\n--- All Task IDs (TASK_ID/Aisle input) ---\n")
                    self.output.emit(f"{task_ids}\n")
                    items_df = pd.DataFrame(columns=["Item", "Affected Task ID Count"])
                    self.table_ready.emit(items_df, ViewMeta(template_name=template_name))
                    self.task_ids_ready.emit(task_ids)
                    return

                task_ids, items_not_met = self.parser.get_task_ids_where_condition(
                    task_id_col="Task ID",
                    condition_col1="Active OHB",
                    condition_col2="Allocated",
                    comparison=">=",
                    item_col="Item",
                )

                self.output.emit("\n--- Items that need Replenishment Task (Active OHB < Allocated) ---\n")
                if items_not_met:
                    for item, task_id_count in sorted(items_not_met.items(), key=lambda x: x[1], reverse=True):
                        self.output.emit(f"  {item}: affects {task_id_count} Task ID(s)\n")
                else:
                    self.output.emit("No items found in Task IDs that don't meet the condition.\n")

                self.output.emit("\n--- Tasks that need Released (Active OHB >= Allocated) ---\n")
                self.output.emit(f"{task_ids}\n")

                if items_not_met:
                    sorted_items = sorted(items_not_met.items(), key=lambda x: x[1], reverse=True)
                    items_df = pd.DataFrame(sorted_items, columns=["Item", "Affected Task ID Count"])
                else:
                    items_df = pd.DataFrame(columns=["Item", "Affected Task ID Count"])

                self.table_ready.emit(items_df, ViewMeta(template_name=template_name))
                self.task_ids_ready.emit(task_ids)

            elif self.function_name == "display_all":
                self.output.emit("\n--- All Data ---\n")
                if template is None:
                    self.output.emit("No template matched; showing raw data.\n")
                    self.table_ready.emit(self.parser.df.reset_index(drop=True), ViewMeta())
                    return
                prepared_df, meta = apply_template(self.parser.df, template)
                self.table_ready.emit(prepared_df, meta)

            else:
                self.output.emit(f"Unknown function: {self.function_name}\n")

        except Exception as e:
            self.output.emit(f"Error executing function: {str(e)}\n")


class ExcelParserGUI(QMainWindow):
    """PyQt GUI for the Excel Parser"""
    
    def __init__(self):
        super().__init__()
        self.worker_thread = None
        self.current_df = None
        self.task_ids = None
        self.strikethrough_rows = set()  # Track which rows have strikethrough applied
        self.cell_highlights: Dict[Tuple[int, str], str] = {}  # (data_row, col) -> color
        self.table_row_map = None  # Map table row index to data row index
        self.table_row_location_values = None  # Map table row index to location value
        self.location_col = None  # Track location column for table grouping
        self.bulk_checkbox_update = False  # Prevent recursive checkbox handling
        self.update_process = None
        self.registry = TemplateRegistry.load()
        self.settings = QSettings("DocuReader", "DocuReader")
        self.view_df: Optional[pd.DataFrame] = None
        self.view_meta: Optional[ViewMeta] = None
        self.init_ui()
    
    def init_ui(self):
        """Initialize the user interface"""
        self.setWindowTitle("Lomar Inventory Control - DocuReader")
        self.setGeometry(100, 100, 1000, 600)
        
        # Create central widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Create main layout
        main_layout = QVBoxLayout()
        
        # Title
        title = QLabel("Inventory DocuReader")
        title_font = title.font()
        title_font.setPointSize(14)
        title_font.setBold(True)
        title.setFont(title_font)
        main_layout.addWidget(title)
        
        # Status label (initialize early before populate_downloads_files)
        self.status_label = QLabel("Ready")
        main_layout.addWidget(self.status_label)
        
        # File selection section
        file_layout = QHBoxLayout()
        file_layout.addWidget(QLabel("Which file do you want?"))
        
        self.file_combo = QComboBox()
        self.populate_downloads_files()
        file_layout.addWidget(self.file_combo)
        
        main_layout.addLayout(file_layout)
        
        # Function selection section
        function_layout = QHBoxLayout()
        function_layout.addWidget(QLabel("What do you need?"))
        
        self.function_combo = QComboBox()
        self.function_combo.addItem("Chase Tasks Needing Released", "get_task_ids_where_condition")
        #self.function_combo.addItem("filter_by_value", "filter_by_value")
        #self.function_combo.addItem("filter_by_range", "filter_by_range")
        #self.function_combo.addItem("filter_by_contains", "filter_by_contains")
        self.function_combo.addItem("The Whole Table", "display_all")
        
        function_layout.addWidget(self.function_combo)
        main_layout.addLayout(function_layout)

        # Detected template / category label (populated after analysis runs).
        self.detected_label = QLabel("Detected category: (none yet - run an analysis)")
        self.detected_label.setStyleSheet("QLabel { color: #555; font-style: italic; }")
        main_layout.addWidget(self.detected_label)

        # Buttons layout
        button_layout = QHBoxLayout()
        
        self.start_button = QPushButton("Analyze and Parse")
        self.start_button.clicked.connect(self.start_analysis)
        button_layout.addWidget(self.start_button)

        self.copy_button = QPushButton("Copy Task IDs to Clipboard")
        self.copy_button.clicked.connect(self.copy_to_clipboard)
        self.copy_button.setEnabled(False)
        button_layout.addWidget(self.copy_button)
        
        self.refresh_button = QPushButton("Refresh Files")
        self.refresh_button.clicked.connect(self.populate_downloads_files)
        button_layout.addWidget(self.refresh_button)
        
        self.clear_button = QPushButton("Clear Output")
        self.clear_button.clicked.connect(self.clear_output)
        button_layout.addWidget(self.clear_button)

        self.update_button = QPushButton("Check & Install Updates")
        self.update_button.clicked.connect(self.check_for_updates)
        self.update_button.setStyleSheet("QPushButton { background-color: #4caf50; color: white; font-weight: bold; }")
        button_layout.addWidget(self.update_button)

        # Channel toggle - persisted via QSettings.
        self.prerelease_checkbox = QCheckBox("Include pre-releases")
        self.prerelease_checkbox.setToolTip(
            "When checked, the updater will also offer release candidates / beta tags."
        )
        self.prerelease_checkbox.setChecked(
            self.settings.value("updater/include_prereleases", False, type=bool)
        )
        self.prerelease_checkbox.toggled.connect(
            lambda v: self.settings.setValue("updater/include_prereleases", bool(v))
        )
        button_layout.addWidget(self.prerelease_checkbox)

        self.templates_button = QPushButton("Templates...")
        self.templates_button.clicked.connect(self.open_templates_dialog)
        button_layout.addWidget(self.templates_button)

        self.export_button = QPushButton("Export view...")
        self.export_button.clicked.connect(self.export_view)
        self.export_button.setEnabled(False)
        button_layout.addWidget(self.export_button)

        self.batch_button = QPushButton("Batch export...")
        self.batch_button.clicked.connect(self.batch_export)
        self.batch_button.setToolTip(
            "Pick multiple CSV/XLSX files, apply each file's matched template, "
            "and write one templated .xlsx per source file to a chosen folder."
        )
        button_layout.addWidget(self.batch_button)
        
        self.terminate_button = QPushButton("Terminate")
        self.terminate_button.clicked.connect(self.terminate_program)
        self.terminate_button.setStyleSheet("QPushButton { background-color: #ff6b6b; color: white; font-weight: bold; }")
        button_layout.addWidget(self.terminate_button)
        
        main_layout.addLayout(button_layout)
        
        # Output display
        output_label = QLabel("Analysis Output:")
        main_layout.addWidget(output_label)
        
        # Create a splitter for logs and table
        splitter = QSplitter(Qt.Orientation.Vertical)
        
        # Text output for logs
        self.output_text = QTextEdit()
        self.output_text.setReadOnly(True)
        monospace_font = QFont("Courier New", 9)
        monospace_font.setFixedPitch(True)
        self.output_text.setFont(monospace_font)
        self.output_text.setMaximumHeight(100)
        splitter.addWidget(self.output_text)
        
        # Table widget for data display
        self.table_widget = QTableWidget()
        self.table_widget.setColumnCount(0)
        self.table_widget.setRowCount(0)
        splitter.addWidget(self.table_widget)
        
        main_layout.addWidget(splitter)
        
        # Set layout
        central_widget.setLayout(main_layout)
    
    def populate_downloads_files(self):
        """Populate the file combo box with CSV and Excel files from Downloads"""
        self.file_combo.clear()
        downloads_path = Path.home() / "Downloads"
        
        files = []
        for ext in ['*.csv', '*.xlsx', '*.xls']:
            files.extend(downloads_path.glob(ext))
        
        if files:
            # Sort by modification time (most recent first)
            files.sort(key=lambda f: f.stat().st_mtime, reverse=True)
            
            for file in files:
                self.file_combo.addItem(file.name, str(file))
            
            self.status_label.setText(f"Found {len(files)} file(s) in Downloads")
        else:
            self.file_combo.addItem("No files found", None)
            self.status_label.setText("No CSV or Excel files found in Downloads folder")
    
    def start_analysis(self):
        """Start the analysis in a worker thread"""
        if self.file_combo.currentData() is None:
            QMessageBox.warning(self, "No File", "No file selected. Please select a file from Downloads.")
            return
        
        filepath = self.file_combo.currentData()
        selected_function = self.function_combo.currentData()

        # If this is a multi-sheet workbook, prompt the user (with the last
        # used sheet for this file remembered via QSettings).
        sheet_name = self._pick_sheet(filepath)
        if sheet_name is False:  # user cancelled the dialog
            return

        # Disable buttons while processing
        self.start_button.setEnabled(False)
        self.file_combo.setEnabled(False)
        self.function_combo.setEnabled(False)
        self.refresh_button.setEnabled(False)
        
        # Clear the previous table
        self.table_widget.setRowCount(0)
        self.table_widget.setColumnCount(0)
        self.strikethrough_rows.clear()
        self.task_ids = None
        self.copy_button.setEnabled(False)
        
        self.status_label.setText("Processing... please wait")
        self.output_text.append("Starting analysis...\n")
        
        # Create and start worker thread
        self.worker_thread = WorkerThread(filepath, selected_function, self.registry, sheet_name=sheet_name or None)
        self.worker_thread.output.connect(self.append_output)
        self.worker_thread.error.connect(self.on_error)
        self.worker_thread.table_ready.connect(self.on_table_ready)
        self.worker_thread.task_ids_ready.connect(self.on_task_ids_ready)
        self.worker_thread.data_loaded.connect(self.on_data_loaded)
        self.worker_thread.template_matched.connect(self.on_template_matched)
        self.worker_thread.finished.connect(self.on_finished)
        self.worker_thread.start()
    
    def append_output(self, text: str):
        """Append text to the output display"""
        self.output_text.append(text)

    def on_template_matched(self, template_name: str, reason: str):
        """Update the 'Detected category' label after template selection."""
        self.detected_label.setText(f"Detected category: {template_name}")
        self.detected_label.setToolTip(f"Matched by: {reason}")

    def _is_dark_theme(self) -> bool:
        """Determine whether the active application theme is dark."""
        palette = self.table_widget.palette()
        window_color = palette.color(palette.ColorRole.Window)
        return window_color.lightness() < 128

    def _default_table_text_color(self) -> QColor:
        """Return standard table text color based on active theme."""
        return QColor("white") if self._is_dark_theme() else QColor("black")

    def _is_highlighted_cell(self, data_row_idx: int, col_name: str) -> bool:
        """Return True when a cell has an active background highlight."""
        return (data_row_idx, col_name) in self.cell_highlights

    @staticmethod
    def _color_for_name(name: str) -> Optional[QColor]:
        """Map template highlight colour names to QColor swatches."""
        return {
            "darkgreen": QColor(130, 200, 150),
            "darkyellow": QColor(230, 200, 90),
            "red": QColor(220, 120, 120),
            "blue": QColor(140, 180, 230),
        }.get(name)

    def on_table_ready(self, df: pd.DataFrame, meta: ViewMeta):
        """Populate the table with data and add the checkbox column."""
        if df is None or df.empty:
            self.output_text.append("No data to display in table.\n")
            return

        # Remember the view-shaped df + highlights so "Export view..." can
        # write exactly what the user is looking at.
        self.view_df = df.copy()
        self.view_meta = meta
        self.export_button.setEnabled(True)

        self.strikethrough_rows.clear()

        # Build cell-highlight lookup from the template-driven ViewMeta.
        self.cell_highlights = {(int(r), c): color for (r, c, color) in meta.highlights}

        # Use the template-detected location column for divider rows when present.
        location_col = meta.location_column if meta.location_column in df.columns else None
        self.location_col = location_col

        render_rows = []
        self.table_row_map = []
        self.table_row_location_values = []
        prev_prefix = None
        for df_idx, (_, row_data) in enumerate(df.iterrows()):
            prefix = None
            location_text = None
            if location_col:
                location_value = row_data.get(location_col)
                location_text = "" if pd.isna(location_value) else str(location_value)
                parsed_prefix, parsed_number, _ = parse_location_parts(location_text)
                prefix = (parsed_prefix, parsed_number // 100000 if parsed_number != float("inf") else float("inf"))

            if location_col and prev_prefix is not None and prefix != prev_prefix:
                render_rows.append({"type": "divider"})
                self.table_row_map.append(None)
                self.table_row_location_values.append(None)

            render_rows.append({"type": "data", "df_idx": df_idx, "row_data": row_data})
            self.table_row_map.append(df_idx)
            self.table_row_location_values.append(location_text if location_col else None)
            if location_col:
                prev_prefix = prefix

        self.table_widget.setRowCount(len(render_rows))
        self.table_widget.setColumnCount(len(df.columns) + 1)

        headers = list(df.columns) + ["Counted?"]
        self.table_widget.setHorizontalHeaderLabels(headers)

        default_text_color = self._default_table_text_color()

        for row_idx, render_row in enumerate(render_rows):
            if render_row["type"] == "divider":
                for col_idx in range(len(df.columns)):
                    item = QTableWidgetItem("-----")
                    item.setFlags(Qt.ItemFlag.NoItemFlags)
                    item.setForeground(default_text_color)
                    self.table_widget.setItem(row_idx, col_idx, item)
                continue

            row_data = render_row["row_data"]
            df_idx = render_row["df_idx"]

            for col_idx, col_name in enumerate(df.columns):
                value = row_data[col_name]
                cell_text = "" if pd.isna(value) else str(value)
                item = QTableWidgetItem(cell_text)
                color_name = self.cell_highlights.get((df_idx, col_name))
                if color_name:
                    qcolor = self._color_for_name(color_name)
                    if qcolor is not None:
                        item.setBackground(qcolor)
                        item.setForeground(QColor("black"))
                self.table_widget.setItem(row_idx, col_idx, item)

            checkbox = QCheckBox()
            checkbox.stateChanged.connect(lambda checked, r=row_idx: self.on_checkbox_changed(r, checked))
            self.table_widget.setCellWidget(row_idx, len(df.columns), checkbox)

        self.table_widget.resizeColumnsToContents()
        header = self.table_widget.horizontalHeader()
        if header is not None:
            header.setStretchLastSection(True)
        self.output_text.append(f"Table loaded with {len(df)} rows and {len(df.columns)} columns.\n")
    
    def on_checkbox_changed(self, row_idx: int, state):
        """Handle checkbox state change - apply/remove strikethrough on row"""
        if self.bulk_checkbox_update:
            return

        is_checked = state == 2  # Qt.CheckState.Checked is 2

        data_row_idx = None
        if isinstance(self.table_row_map, list) and row_idx < len(self.table_row_map):
            data_row_idx = self.table_row_map[row_idx]
        if data_row_idx is None:
            return

        location_value = None
        if isinstance(self.table_row_location_values, list) and row_idx < len(self.table_row_location_values):
            location_value = self.table_row_location_values[row_idx]
        if location_value is None:
            return
        if not isinstance(self.table_row_location_values, list):
            return

        self.bulk_checkbox_update = True
        try:
            for table_row_idx, row_location in enumerate(self.table_row_location_values):
                if row_location != location_value:
                    continue

                checkbox = self.table_widget.cellWidget(table_row_idx, self.table_widget.columnCount() - 1)
                if isinstance(checkbox, QCheckBox) and checkbox.isChecked() != is_checked:
                    checkbox.setChecked(is_checked)

                self.apply_row_strikethrough(table_row_idx, is_checked)
        finally:
            self.bulk_checkbox_update = False

    def apply_row_strikethrough(self, row_idx: int, is_checked: bool):
        """Apply or remove strikethrough for a table row."""
        data_row_idx = None
        if isinstance(self.table_row_map, list) and row_idx < len(self.table_row_map):
            data_row_idx = self.table_row_map[row_idx]
        if data_row_idx is None:
            return
        
        if is_checked and row_idx not in self.strikethrough_rows:
            # Apply strikethrough and red color
            self.strikethrough_rows.add(row_idx)
            for col_idx in range(self.table_widget.columnCount() - 1):  # Skip checkbox column
                item = self.table_widget.item(row_idx, col_idx)
                if item:
                    font = item.font()
                    font.setStrikeOut(True)
                    item.setFont(font)
                    item.setForeground(QColor("red"))
        elif not is_checked and row_idx in self.strikethrough_rows:
            # Remove strikethrough and restore original formatting
            self.strikethrough_rows.discard(row_idx)
            default_text_color = self._default_table_text_color()
            
            for col_idx in range(self.table_widget.columnCount() - 1):  # Skip checkbox column
                item = self.table_widget.item(row_idx, col_idx)
                if item:
                    font = item.font()
                    font.setStrikeOut(False)
                    item.setFont(font)
                    
                    # Get column name
                    header_item = self.table_widget.horizontalHeaderItem(col_idx)
                    if header_item is None:
                        continue
                    col_name = header_item.text()
                    
                    # Restore text color: black for highlighted cells, theme-default for others
                    if self._is_highlighted_cell(data_row_idx, col_name):
                        item.setForeground(QColor("black"))
                    else:
                        item.setForeground(default_text_color)
    
    def on_error(self, error_msg: str):
        """Handle errors from worker thread"""
        self.output_text.append(f"\nERROR: {error_msg}\n")
        self.status_label.setText("Error occurred")
    
    def on_task_ids_ready(self, task_ids: list):
        """Handle task IDs from worker thread"""
        self.task_ids = task_ids
        self.copy_button.setEnabled(bool(task_ids))
    
    def on_data_loaded(self, df: pd.DataFrame):
        """Handle data loaded from worker thread"""
        self.current_df = df
        self.export_button.setEnabled(df is not None and not df.empty)
    
    def on_finished(self):
        """Called when worker thread finishes"""
        # Re-enable buttons
        self.start_button.setEnabled(True)
        self.file_combo.setEnabled(True)
        self.function_combo.setEnabled(True)
        self.refresh_button.setEnabled(True)
        
        self.status_label.setText("Analysis complete")
        self.output_text.append("\nAnalysis complete. You can start a new analysis or terminate the program.")
    
    def clear_output(self):
        """Clear the output display and table"""
        self.output_text.clear()
        self.table_widget.setRowCount(0)
        self.table_widget.setColumnCount(0)
        self.strikethrough_rows.clear()
        self.task_ids = None
        self.copy_button.setEnabled(False)
        self.status_label.setText("Output cleared")

    def open_templates_dialog(self):
        """Show the templates CRUD dialog and reload the registry on close."""
        dialog = TemplatesDialog(self.registry, self)
        dialog.exec()
        # Re-load from disk so any external edits are picked up too.
        self.registry = TemplateRegistry.load()

    def _pick_sheet(self, filepath: str):
        """Resolve which sheet to read for an Excel file.

        Returns:
            - The sheet name (str) when one is chosen,
            - "" when the file is a CSV or has at most one sheet (caller will
              pass ``None`` to ``WorkerThread`` and let pandas default),
            - ``False`` when the user cancels the dialog (caller should abort).
        """
        if filepath.lower().endswith(".csv"):
            return ""
        try:
            sheets = ExcelParser(filepath).list_sheets()
        except Exception:
            return ""
        if len(sheets) <= 1:
            return ""

        key = f"sheets/{filepath}"
        last = self.settings.value(key, type=str) or sheets[0]
        try:
            current_index = sheets.index(last)
        except ValueError:
            current_index = 0

        chosen, ok = QInputDialog.getItem(
            self,
            "Select sheet",
            f"This workbook contains {len(sheets)} sheets. Pick one:",
            sheets,
            current_index,
            False,
        )
        if not ok:
            return False
        self.settings.setValue(key, chosen)
        return chosen

    def export_view(self):
        """Export the current view (template-applied dataframe) to CSV or XLSX.

        For .xlsx, cell highlights from the active ViewMeta are preserved
        using openpyxl pattern fills.
        """
        if self.view_df is None or self.view_df.empty:
            QMessageBox.information(self, "Nothing to export", "Run an analysis first.")
            return

        default_name = "docureader_view.xlsx"
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Export current view",
            default_name,
            "Excel Workbook (*.xlsx);;CSV (*.csv)",
        )
        if not path:
            return

        try:
            if path.lower().endswith(".csv"):
                self.view_df.to_csv(path, index=False)
            else:
                self._export_xlsx_with_highlights(path)
        except Exception as e:
            QMessageBox.critical(self, "Export failed", str(e))
            return

        self.status_label.setText(f"Exported view to {path}")
        self.output_text.append(f"\nExported view to {path}\n")

    def _export_xlsx_with_highlights(self, path: str) -> None:
        """Write ``self.view_df`` to ``path`` and apply highlight fills."""
        from openpyxl import Workbook
        from openpyxl.styles import PatternFill

        df = self.view_df
        meta = self.view_meta
        wb = Workbook()
        ws = wb.active
        ws.title = "View"

        cols = list(df.columns)
        ws.append(cols)
        for _, row in df.iterrows():
            ws.append(["" if pd.isna(v) else v for v in row.tolist()])

        # Map our colour names to hex fills.
        hex_for = {
            "darkgreen": "82C896",
            "darkyellow": "E6C85A",
            "red": "DC7878",
            "blue": "8CB4E6",
        }
        if meta is not None and meta.highlights:
            col_index = {c: i + 1 for i, c in enumerate(cols)}  # openpyxl is 1-based
            for (data_row, col_name, color_name) in meta.highlights:
                hex_code = hex_for.get(color_name)
                col_idx = col_index.get(col_name)
                if not hex_code or col_idx is None:
                    continue
                # +2: header row offset (row 1) + 0-based data_row.
                ws.cell(row=int(data_row) + 2, column=col_idx).fill = PatternFill(
                    start_color=hex_code, end_color=hex_code, fill_type="solid"
                )

        wb.save(path)

    def batch_export(self):
        """Pick N source files, apply each one's matched template, and write
        one ``<source>.view.xlsx`` per file into a chosen output folder.

        Runs synchronously on the GUI thread - intended for short batches
        from Downloads. Highlights are preserved (same export path as the
        single-file export).
        """
        downloads = str(Path.home() / "Downloads")
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "Select files to batch process",
            downloads,
            "Data files (*.csv *.xlsx *.xls);;All files (*)",
        )
        if not files:
            return

        out_dir = QFileDialog.getExistingDirectory(
            self, "Select output folder", downloads
        )
        if not out_dir:
            return

        out_path = Path(out_dir)
        successes = 0
        failures: list[tuple[str, str]] = []

        self.output_text.append("\n" + "=" * 60 + "\n")
        self.output_text.append(f"Batch processing {len(files)} file(s) -> {out_dir}\n")
        self.output_text.append("=" * 60 + "\n")

        prev_view_df = self.view_df
        prev_view_meta = self.view_meta

        for src in files:
            src_path = Path(src)
            try:
                parser = ExcelParser(str(src_path))
                parser.read_excel()  # pandas defaults; sheet selector is interactive only.
                if parser.df is None:
                    raise RuntimeError("failed to load")

                match = self.registry.select(parser.df.columns, str(src_path))
                view_df, meta = apply_template(parser.df, match.template)

                # Reuse the existing exporter by stashing the view temporarily.
                self.view_df = view_df
                self.view_meta = meta
                dst = out_path / f"{src_path.stem}.view.xlsx"
                self._export_xlsx_with_highlights(str(dst))

                self.output_text.append(
                    f"  OK  {src_path.name}  -> {dst.name}  "
                    f"[template: {match.template.name}, {len(view_df)} rows]\n"
                )
                successes += 1
            except Exception as e:
                failures.append((src_path.name, str(e)))
                self.output_text.append(f"  FAIL  {src_path.name}: {e}\n")

        # Restore the on-screen view's state.
        self.view_df = prev_view_df
        self.view_meta = prev_view_meta

        self.output_text.append(
            f"\nBatch complete: {successes} succeeded, {len(failures)} failed.\n"
        )
        self.status_label.setText(
            f"Batch export: {successes}/{len(files)} succeeded"
        )
        if failures:
            QMessageBox.warning(
                self,
                "Batch export finished with errors",
                f"{successes} of {len(files)} files exported.\n\n"
                + "\n".join(f"- {n}: {e}" for n, e in failures[:10])
                + ("" if len(failures) <= 10 else f"\n... and {len(failures) - 10} more"),
            )

    def copy_to_clipboard(self):
        """Copy task IDs to clipboard"""
        if self.task_ids:
            clipboard_text = ", ".join(map(str, self.task_ids))
            clipboard = QApplication.clipboard()
            if clipboard is not None:
                clipboard.setText(clipboard_text)
            self.status_label.setText(f"Copied {len(self.task_ids)} Task ID(s) to clipboard")
            self.output_text.append(f"\nCopied to clipboard: {clipboard_text}\n")

    def check_for_updates(self):
        """Check for and automatically install application updates."""
        reply = QMessageBox.question(
            self,
            "Update Application",
            "This will run the updater and may reset local files to a release version.\n\n"
            "Any uncommitted local changes can be lost.\n\n"
            "The application will need to be restarted after updating.\n\n"
            "Do you want to continue?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )

        if reply != QMessageBox.StandardButton.Yes:
            return

        self.output_text.append("\n" + "=" * 60 + "\n")
        self.output_text.append("Checking for updates and installing if available...\n")
        self.output_text.append("=" * 60 + "\n")

        self.update_button.setEnabled(False)
        self.update_button.setText("Updating...")

        self.update_process = QProcess(self)
        self.update_process.readyReadStandardOutput.connect(self.handle_update_output)
        self.update_process.readyReadStandardError.connect(self.handle_update_error)
        self.update_process.finished.connect(self.update_finished)

        update_command = self.resolve_update_command()
        if update_command is None:
            self.output_text.append("Updater is not available in this installation.\n")
            self.status_label.setText("Updater unavailable")
            self.update_button.setEnabled(True)
            self.update_button.setText("Check & Install Updates")
            self.update_process = None
            return

        update_args = update_command[1:] + ["--yes"]
        if self.prerelease_checkbox.isChecked() and self._update_command_supports_prereleases(update_command):
            update_args.append("--include-prereleases")
        self.update_process.start(update_command[0], update_args)

    @staticmethod
    def _update_command_supports_prereleases(update_command: list) -> bool:
        """The ``--include-prereleases`` flag is only known to the GitHub
        Releases updater, not the legacy git-based ``update.py``/``update.exe``."""
        joined = " ".join(update_command).lower()
        return "update_github" in joined or "updater_github" in joined

    def resolve_update_command(self) -> Optional[list]:
        """Resolve updater command.

        Selection order (matches the plan, Phase 3.4):
        1. Frozen exe with ``update_github.exe`` next to it -> the new GitHub
           Releases updater (works with no Python / no git on the client).
        2. Frozen exe with ``update.exe`` next to it -> legacy git-based path
           (developer/test installs only).
        3. Source checkout with a ``.git`` folder -> ``update.py`` (git-based).
        4. Source checkout without ``.git`` -> ``updater_github.py`` (network).
        """
        frozen = getattr(sys, "frozen", False)
        if frozen:
            exe_dir = Path(sys.executable).resolve().parent
            github = exe_dir / "update_github.exe"
            if github.exists():
                return [str(github)]
            legacy = exe_dir / "update.exe"
            if legacy.exists():
                return [str(legacy)]
            return None

        script_dir = Path(__file__).resolve().parent
        if (script_dir / ".git").exists():
            update_script = script_dir / "update.py"
            if update_script.exists():
                return [sys.executable, str(update_script)]
        github_script = script_dir / "updater_github.py"
        if github_script.exists():
            return [sys.executable, str(github_script)]
        return [sys.executable, "-m", "updater_github"]

    def handle_update_output(self):
        """Handle stdout from update process"""
        if self.update_process:
            data = self.update_process.readAllStandardOutput()
            output = bytes(data.data()).decode("utf-8", errors="replace")
            self.output_text.append(output)

    def handle_update_error(self):
        """Handle stderr from update process"""
        if self.update_process:
            data = self.update_process.readAllStandardError()
            output = bytes(data.data()).decode("utf-8", errors="replace")
            if output.strip():
                self.output_text.append(f"[ERROR] {output}")

    def update_finished(self, exit_code, exit_status):
        """Handle update process completion"""
        self.output_text.append("\n" + "=" * 60 + "\n")
        if exit_code == 0:
            self.output_text.append("Update process completed.\n")
            self.status_label.setText("Update check complete")
        elif exit_code == 2:
            self.output_text.append("Already on the latest version.\n")
            self.status_label.setText("No updates available")
        else:
            self.output_text.append(f"Update process completed with exit code {exit_code}.\n")
            if exit_status != QProcess.ExitStatus.NormalExit:
                self.output_text.append("Update process ended unexpectedly.\n")
            self.status_label.setText("Update failed")

        self.update_button.setEnabled(True)
        self.update_button.setText("Check & Install Updates")
        self.update_process = None
    
    def terminate_program(self):
        """Terminate the application"""
        reply = QMessageBox.question(
            self,
            "Terminate Program",
            "Are you sure you want to close the application?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )
        
        if reply == QMessageBox.StandardButton.Yes:
            # Stop worker thread if running
            if self.worker_thread and self.worker_thread.isRunning():
                self.worker_thread.quit()
                self.worker_thread.wait()
            
            self.close()
    
    def closeEvent(self, event):
        """Handle window close event"""
        # Stop worker thread if running
        if self.worker_thread and self.worker_thread.isRunning():
            self.worker_thread.quit()
            self.worker_thread.wait()
        
        event.accept()


def main():
    # First-run migration: if launched from %ProgramFiles% (the legacy admin
    # install location), copy ourselves to %LOCALAPPDATA% and relaunch so
    # subsequent auto-updates can write without UAC.
    try:
        from migrate import maybe_migrate_install
        if maybe_migrate_install():
            return
    except Exception:
        pass

    app = QApplication(sys.argv)
    gui = ExcelParserGUI()
    gui.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
