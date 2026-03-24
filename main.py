import sys
import os
import re
import pandas as pd
from pathlib import Path
import openpyxl

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QFileDialog, QMessageBox
)
from PyQt6.QtCore import QThread, pyqtSignal
from PyQt6.QtGui import QIcon

# Import the UI class you generated
from lokalise_xlsx_converter import Ui_MainWindow

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


class ConversionWorker(QThread):
    """
    Runs the Excel conversion in the background to prevent the UI from freezing.
    Emits signals to update the progress bar and notify when finished.
    """
    progress_update = pyqtSignal(int)
    finished_success = pyqtSignal()
    finished_error = pyqtSignal(str)

    def __init__(self, input_dir, output_dir):
        super().__init__()
        self.input_dir = Path(input_dir)
        self.output_dir = Path(output_dir)

    def run(self):
        try:
            file_paths = list(self.input_dir.rglob("*.xlsx"))
            total_files = len(file_paths)

            if total_files == 0:
                self.finished_error.emit("No .xlsx files found in the input folder.")
                return

            for index, file_path in enumerate(file_paths):
                # Folder Mirroring Logic
                relative_path = file_path.relative_to(self.input_dir)
                output_subfolder = self.output_dir / relative_path.parent
                output_subfolder.mkdir(parents=True, exist_ok=True)

                # Data Extraction
                try:
                    df_raw = pd.read_excel(file_path, header=None)
                except Exception:
                    continue # Skip unreadable files

                task_name = ""
                total_words = 0
                repetitions = 0
                table_start_idx = -1

                for idx, row in df_raw.iterrows():
                    row_str = str(row.values)
                    col0_str = str(row[0]).strip() if pd.notna(row[0]) else ""

                    if "Task name" in col0_str:
                        task_name = str(row[1]).strip() if pd.notna(row[1]) else col0_str.replace("Task name", "").strip()
                    elif "Source words" in col0_str:
                        total_words = int(row[1]) if pd.notna(row[1]) else int(re.search(r'\d+', col0_str).group())
                    elif "Repetitions" in col0_str:
                        repetitions = int(row[1]) if pd.notna(row[1]) else int(re.search(r'\d+', col0_str).group())

                    if "TM 100%" in row_str:
                        table_start_idx = idx
                        break

                if table_start_idx == -1:
                    continue # Skip if table header isn't found

                # Table Formatting
                df_table = pd.read_excel(file_path, skiprows=table_start_idx)
                target_rows = []

                for _, row in df_table.iterrows():
                    lang_col = str(row.iloc[0])
                    if "→" not in lang_col:
                        continue

                    match = re.search(r'→\s*.*?\((.*?)\)', lang_col)
                    if not match:
                        continue

                    target_code = match.group(1).strip()
                    
                    new_row = {
                        "Target Locale": f"en-US to {target_code}",
                        "Asset": target_code,
                        "Total": total_words,
                        "ICE Match": 0,
                        "100%": row.get("TM 100%", 0),
                        "100-95%": row.get("TM 95-99%", 0),
                        "95-85%": row.get("TM 85-94%", 0),
                        "85-75%": row.get("TM 75-84%", 0),
                        "75-50%": row.get("TM 50-74%", 0),
                        "50-0%": row.get("TM 0-49%", 0),
                        "Repetition": repetitions,
                        "MT Fuzzy Words": 0
                    }
                    target_rows.append(new_row)

                    # Save new file with auto-fitting columns
                    if target_rows:
                        df_output = pd.DataFrame(target_rows)
                        safe_filename = task_name.replace("::", "_") + ".xlsx" if task_name else file_path.name
                        output_file_path = output_subfolder / safe_filename

                        # Use ExcelWriter to access the underlying openpyxl workbook
                        with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
                            df_output.to_excel(writer, index=False, sheet_name='Sheet1')
                            worksheet = writer.sheets['Sheet1']

                            # Iterate through each column to find the widest text
                            for idx, col in enumerate(df_output.columns):
                                # Get max length of data in the column or the header itself
                                max_len = max(
                                    df_output[col].astype(str).map(len).max(),
                                    len(str(col))
                                ) + 2  # Add 2 spaces of padding so it's not squished

                                # Convert column index (0, 1, 2...) to Excel letter (A, B, C...)
                                col_letter = openpyxl.utils.get_column_letter(idx + 1)
                                worksheet.column_dimensions[col_letter].width = max_len

                # Update Progress Bar
                progress_percentage = int(((index + 1) / total_files) * 100)
                self.progress_update.emit(progress_percentage)

            self.finished_success.emit()

        except Exception as e:
            self.finished_error.emit(str(e))


class AppWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        
        # Load the UI from your generated file
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)

        # Lock down the text boxes so users must use the Browse buttons
        self.ui.inputFolderLineEdit.setReadOnly(True)
        self.ui.outputFolderLineEdit.setReadOnly(True)

        # Set the window icon using our dynamic PyInstaller-safe path
        icon_path = resource_path(os.path.join("resources", "app_icon.ico"))
        self.setWindowIcon(QIcon(icon_path))

        # Reset UI defaults
        self.ui.progressBar.setValue(0)
        
        # Connect Buttons to functions
        self.ui.inputFolderPushButton.clicked.connect(self.browse_input_folder)
        self.ui.outputFolderPushButton.clicked.connect(self.browse_output_folder)
        self.ui.cancelPushButton.clicked.connect(self.close)
        self.ui.runPushButton.clicked.connect(self.run_conversion)

    def browse_input_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Input Folder")
        if folder:
            self.ui.inputFolderLineEdit.setText(folder)

    def browse_output_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Output Folder")
        if folder:
            self.ui.outputFolderLineEdit.setText(folder)

    def run_conversion(self):
        input_dir = self.ui.inputFolderLineEdit.text().strip()
        output_dir = self.ui.outputFolderLineEdit.text().strip()

        # Basic Validation
        if not input_dir or not output_dir:
            QMessageBox.warning(self, "Missing Folders", "Please select both input and output folders.")
            return
        if not os.path.isdir(input_dir):
            QMessageBox.warning(self, "Invalid Folder", "The input folder does not exist.")
            return

        # Prepare UI for processing
        self.ui.runPushButton.setEnabled(False)
        self.ui.progressBar.setValue(0)

        # Start the background worker
        self.worker = ConversionWorker(input_dir, output_dir)
        self.worker.progress_update.connect(self.update_progress)
        self.worker.finished_success.connect(self.on_success)
        self.worker.finished_error.connect(self.on_error)
        self.worker.start()

    def update_progress(self, value):
        self.ui.progressBar.setValue(value)

    def on_success(self):
        self.ui.runPushButton.setEnabled(True)
        QMessageBox.information(self, "Complete", "All files have been converted successfully!")
        self.ui.progressBar.setValue(0)

    def on_error(self, error_message):
        self.ui.runPushButton.setEnabled(True)
        self.ui.progressBar.setValue(0)
        QMessageBox.critical(self, "Error", f"An error occurred:\n{error_message}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = AppWindow()
    window.show()
    sys.exit(app.exec())
