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
    QFrame,
)

from PyQt6.QtCore import Qt, QThread, QObject
from PyQt6.QtGui import QColor, QFont
from typing import Type

import qtawesome as qta
import os

class WordToPdfConverter(QMainWindow):
    def __init__(self, converter_worker: Type[QObject]):
        super().__init__()
        self.converter_worker = converter_worker
        self.setWindowIcon(qta.icon("fa5s.file-pdf", color="#f44336"))
        self.setWindowTitle("Batch Word to PDF Converter")
        self.setGeometry(100, 100, 700, 600)
        self.passwords = {}  # Store passwords for files
        self.default_password = None
        
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.main_layout = QVBoxLayout(self.central_widget)
        self.main_layout.setContentsMargins(20, 20, 20, 20)
        self.main_layout.setSpacing(15)

        # Header
        header = QLabel("Batch Word to PDF Converter")
        header_font = QFont()
        header_font.setPointSize(18)
        header_font.setWeight(600)
        header.setFont(header_font)
        header.setStyleSheet("color: #212121;")
        self.main_layout.addWidget(header)

        # Step 1: Add files
        step1_frame = QFrame()
        step1_frame.setFrameShape(QFrame.Shape.StyledPanel)
        step1_layout = QVBoxLayout(step1_frame)
        step1_layout.setContentsMargins(15, 15, 15, 15)
        step1_layout.setSpacing(10)
        
        step1_label = QLabel("1. Add Word files to convert")
        step1_label.setStyleSheet("font-weight: bold;")
        step1_layout.addWidget(step1_label)

        button_layout = QHBoxLayout()
        self.add_button = QPushButton("Add Files")
        self.add_button.setIcon(qta.icon("fa5s.folder-open", color="white"))
        self.clear_button = QPushButton("Clear List")
        self.clear_button.setIcon(qta.icon("fa5s.trash", color="white"))
        self.clear_button.setObjectName("danger")
        button_layout.addWidget(self.add_button)
        button_layout.addWidget(self.clear_button)
        button_layout.addStretch()
        step1_layout.addLayout(button_layout)
        
        self.file_list_widget = QListWidget()
        self.file_list_widget.setSelectionMode(QListWidget.SelectionMode.ExtendedSelection)
        step1_layout.addWidget(self.file_list_widget)
        
        self.main_layout.addWidget(step1_frame)

        # Step 2: Output directory
        step2_frame = QFrame()
        step2_frame.setFrameShape(QFrame.Shape.StyledPanel)
        step2_layout = QVBoxLayout(step2_frame)
        step2_layout.setContentsMargins(15, 15, 15, 15)
        step2_layout.setSpacing(10)
        
        step2_label = QLabel("2. Select output folder (optional)")
        step2_label.setStyleSheet("font-weight: bold;")
        step2_layout.addWidget(step2_label)
        
        step2_help = QLabel("If empty, PDFs will be saved next to the original files.")
        step2_help.setStyleSheet("color: #757575; font-size: 13px;")
        step2_layout.addWidget(step2_help)

        output_dir_layout = QHBoxLayout()
        self.output_dir_edit = QLineEdit()
        self.output_dir_edit.setPlaceholderText("No folder selected...")
        self.output_dir_edit.setReadOnly(True)
        self.output_dir_button = QPushButton("Browse...")
        self.output_dir_button.setIcon(qta.icon("fa5s.folder", color="white"))
        output_dir_layout.addWidget(self.output_dir_edit)
        output_dir_layout.addWidget(self.output_dir_button)
        step2_layout.addLayout(output_dir_layout)
        
        self.open_destination_loc = QPushButton("Open Destination Folder")
        self.open_destination_loc.setIcon(qta.icon("fa5s.external-link-alt", color="white"))
        step2_layout.addWidget(self.open_destination_loc)
        
        self.main_layout.addWidget(step2_frame)

        # Step 3: Password protection
        step3_frame = QFrame()
        step3_frame.setFrameShape(QFrame.Shape.StyledPanel)
        step3_layout = QVBoxLayout(step3_frame)
        step3_layout.setContentsMargins(15, 15, 15, 15)
        step3_layout.setSpacing(10)
        
        step3_label = QLabel("3. Set default password for protected files (optional)")
        step3_label.setStyleSheet("font-weight: bold;")
        step3_layout.addWidget(step3_label)
        
        password_layout = QHBoxLayout()
        self.password_edit = QLineEdit()
        self.password_edit.setPlaceholderText("Enter default password...")
        self.password_edit.setEchoMode(QLineEdit.EchoMode.Password)
        self.set_password_button = QPushButton("Set Password")
        self.set_password_button.setIcon(qta.icon("fa5s.key", color="white"))
        password_layout.addWidget(self.password_edit)
        password_layout.addWidget(self.set_password_button)
        step3_layout.addLayout(password_layout)
        
        self.main_layout.addWidget(step3_frame)

        # Step 4: Convert
        step4_frame = QFrame()
        step4_frame.setFrameShape(QFrame.Shape.StyledPanel)
        step4_layout = QVBoxLayout(step4_frame)
        step4_layout.setContentsMargins(15, 15, 15, 15)
        step4_layout.setSpacing(15)
        
        step4_label = QLabel("4. Start conversion")
        step4_label.setStyleSheet("font-weight: bold;")
        step4_layout.addWidget(step4_label)
        
        self.convert_button = QPushButton("Convert All to PDF")
        self.convert_button.setIcon(qta.icon("fa5s.file-export", color="white"))
        self.convert_button.setEnabled(False)
        self.convert_button.setObjectName("convert-btn")
        
        step4_layout.addWidget(self.convert_button)

        # Progress area
        progress_frame = QFrame()
        progress_layout = QVBoxLayout(progress_frame)
        progress_layout.setContentsMargins(0, 0, 0, 0)
        progress_layout.setSpacing(8)
        
        self.current_file_label = QLabel("Ready to start conversion.")
        self.current_file_label.setStyleSheet("color: #616161;")
        progress_layout.addWidget(self.current_file_label)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setTextVisible(True)
        progress_layout.addWidget(self.progress_bar)
        
        step4_layout.addWidget(progress_frame)
        self.main_layout.addWidget(step4_frame)

        # Status bar
        self.statusBar().setObjectName("status-bar-qstatus")
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

        self.apply_styles()

    def apply_styles(self):
        try:
            qss_path = os.path.join(os.getcwd(), "styles.css")

            with open(qss_path, "r") as f:
                self.setStyleSheet(f.read())
        except FileNotFoundError:
            print("Warning styles is not loaded due to 'styles.css' is missing")

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
