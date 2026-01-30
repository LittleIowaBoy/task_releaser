import sys
import pandas as pd
from pathlib import Path
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QTextEdit, QLabel, QComboBox, QMessageBox, QSplitter
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from tr import ExcelParser


class WorkerThread(QThread):
    """Worker thread to run the parser without blocking the GUI"""
    finished = pyqtSignal()
    error = pyqtSignal(str)
    output = pyqtSignal(str)
    data_loaded = pyqtSignal(object)  # Emit the dataframe
    
    def __init__(self, filepath: str):
        super().__init__()
        self.filepath = filepath
        self.parser = None
    
    def run(self):
        try:
            self.output.emit(f"Loading file: {Path(self.filepath).name}\n")
            self.parser = ExcelParser(self.filepath)
            self.parser.read_excel()
            
            if self.parser.df is not None:
                self.output.emit(f"Successfully loaded CSV file with {len(self.parser.df)} rows and {len(self.parser.df.columns)} columns\n")
                self.output.emit(f"Columns: {list(self.parser.df.columns)}\n")
                self.output.emit(f"\nFirst few rows:\n{self.parser.df.head().to_string()}\n")
                
                # Get Task IDs where Active OHB > Allocated
                task_ids = self.parser.get_task_ids_where_condition(
                    task_id_col="Task ID",
                    condition_col1="Active OHB",
                    condition_col2="Allocated",
                    comparison=">"
                )
                
                self.output.emit(f"\n--- Task IDs where Active OHB > Allocated ---\n")
                self.output.emit(f"{task_ids}\n")
                
                self.data_loaded.emit(self.parser.df)
            else:
                self.error.emit("Failed to load data")
        
        except Exception as e:
            self.error.emit(f"Error: {str(e)}")
        finally:
            self.finished.emit()


class ExcelParserGUI(QMainWindow):
    """PyQt GUI for the Excel Parser"""
    
    def __init__(self):
        super().__init__()
        self.worker_thread = None
        self.current_df = None
        self.init_ui()
    
    def init_ui(self):
        """Initialize the user interface"""
        self.setWindowTitle("Excel Parser GUI")
        self.setGeometry(100, 100, 1000, 700)
        
        # Create central widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Create main layout
        main_layout = QVBoxLayout()
        
        # Title
        title = QLabel("Excel Parser Application")
        title_font = title.font()
        title_font.setPointSize(14)
        title_font.setBold(True)
        title.setFont(title_font)
        main_layout.addWidget(title)
        
        # Status label (initialize early before populate_downloads_files)
        self.status_label = QLabel("Ready")
        
        # File selection section
        file_layout = QHBoxLayout()
        file_layout.addWidget(QLabel("Select file from Downloads:"))
        
        self.file_combo = QComboBox()
        self.populate_downloads_files()
        file_layout.addWidget(self.file_combo)
        
        main_layout.addLayout(file_layout)
        
        # Buttons layout
        button_layout = QHBoxLayout()
        
        self.start_button = QPushButton("Start Analysis")
        self.start_button.clicked.connect(self.start_analysis)
        button_layout.addWidget(self.start_button)
        
        self.refresh_button = QPushButton("Refresh Files")
        self.refresh_button.clicked.connect(self.populate_downloads_files)
        button_layout.addWidget(self.refresh_button)
        
        self.clear_button = QPushButton("Clear Output")
        self.clear_button.clicked.connect(self.clear_output)
        button_layout.addWidget(self.clear_button)
        
        self.terminate_button = QPushButton("Terminate")
        self.terminate_button.clicked.connect(self.terminate_program)
        self.terminate_button.setStyleSheet("QPushButton { background-color: #ff6b6b; color: white; font-weight: bold; }")
        button_layout.addWidget(self.terminate_button)
        
        main_layout.addLayout(button_layout)
        
        # Output display
        output_label = QLabel("Analysis Output:")
        main_layout.addWidget(output_label)
        
        self.output_text = QTextEdit()
        self.output_text.setReadOnly(True)
        main_layout.addWidget(self.output_text)
        
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
        
        # Disable buttons while processing
        self.start_button.setEnabled(False)
        self.file_combo.setEnabled(False)
        self.refresh_button.setEnabled(False)
        
        self.status_label.setText("Processing... please wait")
        self.output_text.append("Starting analysis...\n")
        
        # Create and start worker thread
        self.worker_thread = WorkerThread(filepath)
        self.worker_thread.output.connect(self.append_output)
        self.worker_thread.error.connect(self.on_error)
        self.worker_thread.data_loaded.connect(self.on_data_loaded)
        self.worker_thread.finished.connect(self.on_finished)
        self.worker_thread.start()
    
    def append_output(self, text: str):
        """Append text to the output display"""
        self.output_text.append(text)
    
    def on_error(self, error_msg: str):
        """Handle errors from worker thread"""
        self.output_text.append(f"\nERROR: {error_msg}\n")
        self.status_label.setText("Error occurred")
    
    def on_data_loaded(self, df: pd.DataFrame):
        """Handle data loaded from worker thread"""
        self.current_df = df
    
    def on_finished(self):
        """Called when worker thread finishes"""
        # Re-enable buttons
        self.start_button.setEnabled(True)
        self.file_combo.setEnabled(True)
        self.refresh_button.setEnabled(True)
        
        self.status_label.setText("Analysis complete")
        self.output_text.append("\nAnalysis complete. You can start a new analysis or terminate the program.")
    
    def clear_output(self):
        """Clear the output display"""
        self.output_text.clear()
        self.status_label.setText("Output cleared")
    
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
