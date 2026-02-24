import sys
import re
import pandas as pd
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QTextEdit, QLabel, QComboBox, QMessageBox, QSplitter,
    QTableWidget, QTableWidgetItem, QCheckBox
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QProcess
from PyQt6.QtGui import QFont, QTextDocument, QTextCursor, QColor
from tr import ExcelParser

__version__ = "0.2.3"

LOCATION_PATTERN = re.compile(r'^\s*([A-Za-z-]*?)(\d+)?([A-Za-z]*)\s*$')


def parse_location_parts(value: Any) -> Tuple[str, float, str]:
    """Parse location into prefix, numeric body, and suffix for stable grouping/sorting."""
    text = "" if pd.isna(value) else str(value).strip().upper()
    match = LOCATION_PATTERN.match(text)
    if not match:
        return (text, float("inf"), "")

    prefix, number_part, suffix = match.groups()
    if number_part:
        if suffix and len(number_part) < 7:
            number = int(number_part.ljust(7, "0"))
        else:
            number = int(number_part)
    else:
        number = float("inf")
    return (prefix or "", number, suffix or "")


class WorkerThread(QThread):
    """Worker thread to run the parser without blocking the GUI"""
    finished = pyqtSignal()
    error = pyqtSignal(str)
    output = pyqtSignal(str)
    table_ready = pyqtSignal(pd.DataFrame)  # Emit processed DataFrame for table display
    data_loaded = pyqtSignal(object)  # Emit the dataframe
    task_ids_ready = pyqtSignal(list)  # Emit the task IDs for clipboard copy
    
    def __init__(self, filepath: str, function_name: str):
        super().__init__()
        self.filepath = filepath
        self.function_name = function_name
        self.parser = None

    def format_df(self, df: pd.DataFrame) -> str:
        """Format a DataFrame into an aligned table string similar to ExcelParser.display().

        This version keeps track of each column's maximum width while iterating
        through rows, updating widths as wider values are encountered. Each column
        is padded to match the widest value in that column.
        
        Columns matching certain names are automatically dropped before formatting.
        """
        if df is None or df.empty:
            return "No data to display\n"

        # Drop columns with specific names
        columns_to_drop = ['Rsn Code', 'Tie']
        df = df.drop(columns=[col for col in columns_to_drop if col in df.columns])

        # Sort by 'Location' column if present (numeric-aware sort ascending)
        if 'Location' in df.columns:
            try:
                # Strip any non-digit prefix (e.g., 'L-') and convert the remainder to numeric for sorting
                df = df.sort_values(
                    by='Location',
                    key=lambda s: pd.to_numeric(s.astype(str).str.replace(r'^\D+', '', regex=True), errors='coerce')
                )
            except Exception:
                # Fallback to default lexicographic sort if unexpected errors occur
                df = df.sort_values(by='Location')

        cols = list(df.columns)
        rows = []

        # Initialize col widths with header widths
        col_widths: Dict[str, int] = {c: len(str(c)) for c in cols}

        # Build rows and update column widths on the fly
        for _, r in df.iterrows():
            row = ["" if pd.isna(r[c]) else str(r[c]) for c in cols]
            rows.append(row)
            for i, val in enumerate(row):
                col_name = cols[i]
                if len(val) > col_widths[col_name]:
                    col_widths[col_name] = len(val)

        # Build header and separator using the computed max widths
        # Each header and separator must be exactly the same width as its column
        header_parts = []
        sep_parts = []
        for c in cols:
            width = col_widths[c]
            header_parts.append(str(c).ljust(width))
            sep_parts.append("-" * width)
        
        header = " | ".join(header_parts)
        separator = "-+-".join(sep_parts)

        lines = [header, separator]
        for row in rows:
            row_parts = []
            for i, val in enumerate(row):
                width = col_widths[cols[i]]
                row_parts.append(val.ljust(width))
            lines.append(" | ".join(row_parts))

        return "\n".join(lines) + "\n"
    
    def prepare_df_for_table(self, df: pd.DataFrame) -> pd.DataFrame:
        """Prepare DataFrame for table display: drop unwanted columns, sort by location column."""
        if df is None or df.empty:
            return df

        df = df.copy()

        # Drop columns with specific names
        columns_to_drop = ['Rsn Code', 'Tie', 'LOCN_CLASS', 'CYCLE_CNT_PENDING']
        df = df.drop(columns=[col for col in columns_to_drop if col in df.columns])

        # Check if all expected columns are present and reorder them
        expected_columns = ['Location', 'Item', 'Description', 'Current OHB', 'Last Replen', 'Last Replen Qty', 'Replen Location', 'Short Time', 'User']
        if all(col in df.columns for col in expected_columns):
            # Reorder columns in the specified order
            desired_order = ['Location', 'Replen Location', 'Item', 'Description', 'Current OHB', 'Last Replen Qty', 'Last Replen', 'Short Time', 'User']
            df = df[desired_order]

        # Sort by location column if present (numeric-aware sort ascending)
        # Check for different location column names
        location_columns = ['Location', 'Locn', 'DSP_LOCN', 'Destination Location']
        location_col = None
        for col in location_columns:
            if col in df.columns:
                location_col = col
                break
        
        if location_col:
            try:
                df = df.sort_values(
                    by=location_col,
                    key=lambda s: s.map(parse_location_parts)
                )
            except Exception:
                # Fallback to default lexicographic sort if unexpected errors occur
                df = df.sort_values(by=location_col)

        df = df.reset_index(drop=True)

        row_highlights: List[Optional[Dict[str, str]]] = [None] * len(df)
        if 'Last Replen' in df.columns and 'Short Time' in df.columns:
            last_replen = pd.to_datetime(df['Last Replen'], errors='coerce')
            short_time = pd.to_datetime(df['Short Time'], errors='coerce')

            after_mask = last_replen > short_time
            before_mask = last_replen < short_time

            for idx, is_after in enumerate(after_mask):
                if is_after:
                    row_highlights[idx] = {
                        'Last Replen': 'darkgreen',
                        'Short Time': 'darkgreen'
                    }
            for idx, is_before in enumerate(before_mask):
                if is_before:
                    highlight = {
                        'Last Replen': 'darkyellow',
                        'Short Time': 'darkyellow'
                    }
                    if 'Current OHB' in df.columns:
                        try:
                            numeric_ohb = pd.to_numeric(df.at[idx, 'Current OHB'], errors='coerce')
                            current_ohb = float(numeric_ohb) if pd.notna(numeric_ohb) else None
                        except Exception:
                            current_ohb = None
                        if current_ohb is not None:
                            if current_ohb <= 5:
                                highlight['Current OHB'] = 'darkgreen'
                            else:
                                highlight['Current OHB'] = 'darkyellow'
                    row_highlights[idx] = highlight

        df.attrs['row_highlights'] = row_highlights
        
        return df
    
    def run(self):
        try:
            #self.output.emit(f"Loading file: {Path(self.filepath).name}\n")
            self.parser = ExcelParser(self.filepath)
            self.parser.read_excel()
            
            if self.parser.df is not None:
                extension = Path(self.filepath).suffix.lower()
                file_type = "CSV" if extension == ".csv" else "Excel"
                self.output.emit(f"Successfully loaded {file_type} file with {len(self.parser.df)} rows and {len(self.parser.df.columns)} columns\n")
                self.output.emit(f"Columns: {list(self.parser.df.columns)}\n")
                #self.output.emit("\nFirst few rows:\n")
                #self.output.emit(self.format_df(self.parser.df.head()))
                
                # Call the selected function
                self.execute_selected_function()
                
                self.data_loaded.emit(self.parser.df)
            else:
                self.error.emit("Failed to load data")
        
        except Exception as e:
            self.error.emit(f"Error: {str(e)}")
        finally:
            self.finished.emit()
    
    def execute_selected_function(self):
        """Execute the selected function from the dropdown"""
        try:
            if self.parser is None:
                self.output.emit("Parser is not initialized.\n")
                return

            if self.function_name == "get_task_ids_where_condition":
                # Get Task IDs where Active OHB < Allocated
                task_ids, items_not_met = self.parser.get_task_ids_where_condition(
                    task_id_col="Task ID",
                    condition_col1="Active OHB",
                    condition_col2="Allocated",
                    comparison=">=",
                    item_col="Item"
                )
                
                # Display items from Task IDs that don't meet the condition first
                self.output.emit(f"\n--- Items that need Replenishment Task (Active OHB < Allocated) ---\n")
                if items_not_met:
                    for item, task_id_count in sorted(items_not_met.items(), key=lambda x: x[1], reverse=True):
                        self.output.emit(f"  {item}: affects {task_id_count} Task ID(s)\n")
                else:
                    self.output.emit("No items found in Task IDs that don't meet the condition.\n")
                
                # Then display task IDs that meet the condition
                self.output.emit(f"\n--- Tasks that need Released (Active OHB >= Allocated) ---\n")
                self.output.emit(f"{task_ids}\n")
                
                self.task_ids_ready.emit(task_ids)
            
            elif self.function_name == "display_all":
                self.output.emit(f"\n--- All Data ---\n")
                if self.parser.df is None:
                    self.output.emit("No data loaded.\n")
                    return
                prepared_df = self.prepare_df_for_table(self.parser.df)
                self.table_ready.emit(prepared_df)
            
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
        self.row_highlights = None  # Store row highlights for use in checkbox handler
        self.table_row_map = None  # Map table row index to data row index
        self.table_row_location_values = None  # Map table row index to location value
        self.location_col = None  # Track location column for table grouping
        self.bulk_checkbox_update = False  # Prevent recursive checkbox handling
        self.update_process = None
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
        self.worker_thread = WorkerThread(filepath, selected_function)
        self.worker_thread.output.connect(self.append_output)
        self.worker_thread.error.connect(self.on_error)
        self.worker_thread.table_ready.connect(self.on_table_ready)
        self.worker_thread.task_ids_ready.connect(self.on_task_ids_ready)
        self.worker_thread.data_loaded.connect(self.on_data_loaded)
        self.worker_thread.finished.connect(self.on_finished)
        self.worker_thread.start()
    
    def append_output(self, text: str):
        """Append text to the output display"""
        self.output_text.append(text)

    def _is_dark_theme(self) -> bool:
        """Determine whether the active application theme is dark."""
        palette = self.table_widget.palette()
        window_color = palette.color(palette.ColorRole.Window)
        return window_color.lightness() < 128

    def _default_table_text_color(self) -> QColor:
        """Return standard table text color based on active theme."""
        return QColor("white") if self._is_dark_theme() else QColor("black")

    def _is_highlighted_cell(self, data_row_idx: int, col_name: str) -> bool:
        """Return True when a cell has active background highlight."""
        if not isinstance(self.row_highlights, list) or data_row_idx >= len(self.row_highlights):
            return False

        highlight = self.row_highlights[data_row_idx]
        if isinstance(highlight, dict):
            return col_name in highlight

        return highlight in ("darkgreen", "darkyellow") and col_name in ("Last Replen", "Short Time")
    
    def on_table_ready(self, df: pd.DataFrame):
        """Populate the table with data and add checkbox column"""
        if df is None or df.empty:
            self.output_text.append("No data to display in table.\n")
            return
        
        # Clear any strikethrough tracking and reset table
        self.strikethrough_rows.clear()
        
        # Determine which location column to use for divider rows
        location_columns = ['Location', 'Locn', 'DSP_LOCN']
        location_col = None
        for col in location_columns:
            if col in df.columns:
                location_col = col
                break
        self.location_col = location_col

        # Build render rows with divider rows when location prefix changes
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

        # Set table dimensions (add 1 for checkbox column)
        self.table_widget.setRowCount(len(render_rows))
        self.table_widget.setColumnCount(len(df.columns) + 1)
        
        # Set column headers (including "Counted?" checkbox column)
        headers = list(df.columns) + ["Counted?"]
        self.table_widget.setHorizontalHeaderLabels(headers)
        
        # Populate table with data
        self.row_highlights = df.attrs.get('row_highlights')
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
            highlight = None
            if isinstance(self.row_highlights, list) and df_idx < len(self.row_highlights):
                highlight = self.row_highlights[df_idx]

            if highlight == 'darkgreen':
                row_color = QColor(130, 200, 150)
            elif highlight == 'darkyellow':
                row_color = QColor(230, 200, 90)
            else:
                row_color = None

            for col_idx, col_name in enumerate(df.columns):
                value = row_data[col_name]
                cell_text = "" if pd.isna(value) else str(value)
                item = QTableWidgetItem(cell_text)
                if isinstance(highlight, dict) and col_name in highlight:
                    if highlight[col_name] == 'darkgreen':
                        item.setBackground(QColor(130, 200, 150))
                        item.setForeground(QColor("black"))
                    elif highlight[col_name] == 'darkyellow':
                        item.setBackground(QColor(230, 200, 90))
                        item.setForeground(QColor("black"))
                elif row_color is not None and col_name in ("Last Replen", "Short Time"):
                    item.setBackground(row_color)
                    item.setForeground(QColor("black"))
                self.table_widget.setItem(row_idx, col_idx, item)

            # Add checkbox in the last column
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
        self.update_process.start(update_command[0], update_args)

    def resolve_update_command(self) -> Optional[list]:
        """Resolve updater command for source checkout, installed package, or frozen app."""
        frozen = getattr(sys, "frozen", False)

        if frozen:
            exe_dir = Path(sys.executable).resolve().parent
            candidate = exe_dir / "update.exe"
            if candidate.exists():
                return [str(candidate)]
            return None

        update_script = Path(__file__).resolve().parent / "update.py"
        if update_script.exists():
            return [sys.executable, str(update_script)]

        return [sys.executable, "-m", "update"]

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
    app = QApplication(sys.argv)
    gui = ExcelParserGUI()
    gui.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
