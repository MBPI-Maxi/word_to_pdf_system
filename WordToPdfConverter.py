from PyQt6.QtWidgets import (
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QLabel,
    QHBoxLayout,
    QPushButton,
    QListWidget,
    QLineEdit,
    QProgressBar,
    QMessageBox,
    QFileDialog,
    QInputDialog,
)

from PyQt6.QtCore import Qt, QThread, QObject
from PyQt6.QtGui import QColor
from typing import Type

import qtawesome as qta
import os

class WordToPdfConverter(QMainWindow):
    def __init__(self, converter_worker: Type[QObject]):
        super().__init__()
        self.converter_worker = converter_worker
        self.setWindowIcon(qta.icon("fa5s.file-pdf", color="red"))
        self.setWindowTitle("Batch Word to PDF Converter")
        self.setGeometry(100, 100, 600, 500)
        self.passwords = {}  # Store passwords for files
        self.default_password = None
        

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.main_layout = QVBoxLayout(self.central_widget)

        # --- UI Widgets ---
        self.main_layout.addWidget(QLabel("1. Add Word files to the list."))

        button_layout = QHBoxLayout()
        self.add_button = QPushButton("Add Files...")
        self.clear_button = QPushButton("Clear List")
        self.open_destination_loc = QPushButton("Open Destination Folder")
        button_layout.addWidget(self.add_button)
        button_layout.addWidget(self.clear_button)
        button_layout.addStretch()
        self.main_layout.addLayout(button_layout)
        
        
        self.file_list_widget = QListWidget()
        self.main_layout.addWidget(self.file_list_widget)

        # Output directory selection
        self.main_layout.addWidget(
            QLabel("2. (Optional) Select an output folder. If empty, PDFs save next to originals."))
        output_dir_layout = QHBoxLayout()
        self.output_dir_edit = QLineEdit()
        self.output_dir_edit.setPlaceholderText("Select an output folder...")
        self.output_dir_edit.setReadOnly(True)
        self.output_dir_button = QPushButton("Browse...")
        output_dir_layout.addWidget(self.output_dir_edit)
        
        output_dir_layout.addWidget(self.output_dir_button)
        self.main_layout.addWidget(self.open_destination_loc)
        self.main_layout.addLayout(output_dir_layout)

        # Default password option
        self.main_layout.addWidget(QLabel("3. (Optional) Set a default password for protected files:"))
        password_layout = QHBoxLayout()
        self.password_edit = QLineEdit()
        self.password_edit.setPlaceholderText("Enter default password...")
        self.password_edit.setEchoMode(QLineEdit.EchoMode.Password)
        self.set_password_button = QPushButton("Set Password")
        password_layout.addWidget(self.password_edit)
        password_layout.addWidget(self.set_password_button)
        self.main_layout.addLayout(password_layout)

        self.main_layout.addWidget(QLabel("4. Start the conversion."))
        self.convert_button = QPushButton("Convert All to PDF")
        self.convert_button.setEnabled(False)
        self.main_layout.addWidget(self.convert_button)

        self.current_file_label = QLabel("Ready to start.")
        self.main_layout.addWidget(self.current_file_label)

        self.progress_bar = QProgressBar()
        self.main_layout.addWidget(self.progress_bar)

        self.statusBar().showMessage("Ready")

        # --- Connect Signals ---
        self.add_button.clicked.connect(self.add_files)
        self.clear_button.clicked.connect(self.clear_list)
        self.output_dir_button.clicked.connect(self.select_output_folder)
        self.set_password_button.clicked.connect(self.set_default_password)
        self.convert_button.clicked.connect(self.start_conversion)
        self.open_destination_loc.clicked.connect(self.open_destination_folder)

        self.thread = None
        self.worker = None

    def open_destination_folder(self):
        path = self.output_dir_edit.text()

        if path:
            os.startfile(os.path.normpath(path))
        else:
            QMessageBox.information(
                self,
                "No destination folder selected.",
                "Please select a destination folder."
            )

    def add_files(self):
        file_paths, _ = QFileDialog.getOpenFileNames(self, "Select Word Documents", "", "Word Files (*.doc *.docx)")
        if file_paths:
            for path in file_paths:
                if not self.file_list_widget.findItems(path, Qt.MatchFlag.MatchExactly):
                    self.file_list_widget.addItem(path)
            self.update_ui_state()

    def select_output_folder(self):
        folder_path = QFileDialog.getExistingDirectory(self, "Select Output Folder")
        if folder_path:
            self.output_dir_edit.setText(folder_path)

    def set_default_password(self):
        password = self.password_edit.text()
        self.default_password = password if password else None
        self.statusBar().showMessage(
            "Default password set." if password else "Default password cleared.", 
            3000
        )

    def start_conversion(self):
        paths_to_process = [self.file_list_widget.item(i).text() for i in range(self.file_list_widget.count())]
        output_dir = self.output_dir_edit.text()

        if not paths_to_process:
            return

        self.set_ui_for_processing(True)

        self.thread = QThread()
        self.worker = self.converter_worker(paths_to_process, output_dir, self.default_password)
        self.worker.moveToThread(self.thread)

        # Connect all signals
        self.thread.started.connect(self.worker.run)
        self.worker.overall_progress.connect(self.progress_bar.setValue)
        self.worker.current_file_progress.connect(self.current_file_label.setText)
        self.worker.file_finished.connect(self.on_file_finished)
        self.worker.batch_finished.connect(self.on_batch_finished)
        self.worker.fatal_error.connect(self.on_fatal_error)
        self.worker.password_required.connect(self.handle_password_required)
        # --- ADD THIS NEW CONNECTION ---
        self.worker.overwrite_request.connect(self.handle_overwrite_request) 
        
        self.thread.finished.connect(self.thread.deleteLater)
        self.worker.batch_finished.connect(self.thread.quit)
        self.worker.fatal_error.connect(self.thread.quit)

        self.thread.start()

    def handle_password_required(self, file_index, file_path):
        item = self.file_list_widget.item(file_index)
        original_text = item.text().split(" --- ")[0]
        filename = os.path.basename(original_text)
        
        password, ok = QInputDialog.getText(
            self, 
            f"Password Required for: {filename}",
            f"The document '{filename}' is password protected.\nPlease enter the password:",
            QLineEdit.EchoMode.Password
        )
        
        if ok and password:
            self.worker.provide_password(password)
        else:
            # User cancelled - skip this file
            item.setText(f"{original_text} --- ❌ Skipped (password required)")
            item.setBackground(QColor("#FFCCCB"))
            self.worker.provide_password(None)

    # --- ADD THIS NEW SLOT FOR OVERWRITE HANDLING ---
    def handle_overwrite_request(self, file_index: int, full_output_path: str, pdf_name_only: str):
        item = self.file_list_widget.item(file_index)
        original_text = item.text().split(" --- ")[0]

        reply = QMessageBox.question(
            self, # Parent is the main window
            "Overwrite Existing File",
            f"The file '{pdf_name_only}' already exists at '{os.path.dirname(full_output_path)}'.\nDo you want to overwrite it?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No # Default button is 'No'
        )

        if reply == QMessageBox.StandardButton.Yes:
            self.worker.set_overwrite_action('yes')
        else: # QMessageBox.StandardButton.No or dialog closed
            self.worker.set_overwrite_action('no')
            # Update the list item immediately to reflect the skip
            item.setText(f"{original_text} --- ➖ Skipped (user chose not to overwrite)")
            item.setBackground(QColor("lightgray"))


    def clear_list(self):
        self.file_list_widget.clear()
        self.reset_list_visuals()
        self.update_ui_state()

    def update_ui_state(self):
        has_items = self.file_list_widget.count() > 0
        self.convert_button.setEnabled(has_items)
        self.clear_button.setEnabled(has_items)
        if not has_items:
            self.current_file_label.setText("Ready to start.")
            self.progress_bar.setValue(0)
            self.statusBar().showMessage("Ready")

    def set_ui_for_processing(self, processing: bool):
        self.add_button.setEnabled(not processing)
        self.clear_button.setEnabled(not processing)
        self.convert_button.setEnabled(not processing)
        self.output_dir_button.setEnabled(not processing)
        self.file_list_widget.setEnabled(not processing)
        self.set_password_button.setEnabled(not processing)
        self.password_edit.setEnabled(not processing)

    def on_file_finished(self, index: int, message: str, success: bool):
        item = self.file_list_widget.item(index)
        if item:
            original_text = item.text().split(" --- ")[0]
            item.setText(f"{original_text} --- {message}")
            color = QColor("lightgreen") if success else QColor("#FFCCCB")
            item.setBackground(color)

    def on_batch_finished(self, message: str):
        self.statusBar().showMessage("Batch conversion complete.", 5000)
        self.current_file_label.setText(message)
        self.set_ui_for_processing(False)

    def on_fatal_error(self, message: str):
        QMessageBox.critical(self, "Fatal Error", message)
        self.statusBar().showMessage("A fatal error occurred!", 5000)
        self.current_file_label.setText(message)
        self.progress_bar.setValue(0)
        self.set_ui_for_processing(False)

    def reset_list_visuals(self):
        for i in range(self.file_list_widget.count()):
            item = self.file_list_widget.item(i)
            if item:
                original_path = item.text().split(" --- ")[0]
                item.setText(original_path)
                item.setBackground(QColor("white"))