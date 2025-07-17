from PyQt6.QtCore import (
    QObject,
    pyqtSignal,
    QEventLoop
)

import os

class ConverterWorker(QObject):
    overall_progress = pyqtSignal(int)
    current_file_progress = pyqtSignal(str)
    file_finished = pyqtSignal(int, str, bool)
    batch_finished = pyqtSignal(str)
    fatal_error = pyqtSignal(str)
    password_required = pyqtSignal(int, str)  # (file_index, file_path)
    overwrite_request = pyqtSignal(int, str, str) # (file_index, full_output_path, pdf_name_only)

    def __init__(self, input_paths, output_dir=None, default_password=None):
        super().__init__()
        self.input_paths = input_paths
        self.output_dir = output_dir
        self.default_password = default_password
        self.passwords = {}  # Store passwords for specific files
        self._is_running = True
        self.current_password = None
        self.password_event_loop = None
        self.overwrite_event_loop = None 
        self.overwrite_choice = None     

    def run(self):
        try:
            import win32com.client
        except ImportError:
            self.fatal_error.emit("Fatal Error: pywin32 is not installed. Please run: pip install pywin32")
            return

        word_app = None
        total_files = len(self.input_paths)
        if total_files == 0:
            self.batch_finished.emit("No files were selected to process.")
            return

        success_count = 0

        try:
            self.current_file_progress.emit("Starting Microsoft Word...")
            word_app = win32com.client.Dispatch("Word.Application")
            word_app.Visible = False  # Ensure Word stays hidden

            for i, input_path in enumerate(self.input_paths):
                if not self._is_running: 
                    break

                doc = None
                file_name = os.path.basename(input_path)
                self.current_file_progress.emit(f"Processing ({i + 1}/{total_files}): {file_name}")
                base_progress = int((i / total_files) * 100)
                self.overall_progress.emit(base_progress)

                try:
                    absolute_input_path = os.path.abspath(input_path)
                    
                    # Determine output path
                    if self.output_dir:
                        original_base_name = os.path.splitext(file_name)[0]
                        original_base_name = original_base_name.replace(" ", "_")
                        pdf_name = original_base_name + "_converted.pdf"
                        output_path = os.path.join(self.output_dir, pdf_name)
                    else:
                        original_base_name = os.path.splitext(input_path)[0]
                        pdf_name = os.path.basename(original_base_name) + "_converted.pdf" # get only filename part
                        output_path = original_base_name + "_converted.pdf"

                    self.overall_progress.emit(base_progress + 50)

                    wdFormatPDF = 17 
                    
                    # --- NEW OVERWRITE LOGIC HERE ---
                    if os.path.exists(output_path):
                        self.overwrite_choice = None # Reset for each file
                        # Emit signal to main thread to ask user
                        self.overwrite_request.emit(i, output_path, pdf_name) 
                        
                        # Create and run an event loop to pause this thread
                        self.overwrite_event_loop = QEventLoop()
                        self.overwrite_event_loop.exec() 

                        # Check the user's choice after the event loop quits
                        if self.overwrite_choice == 'no':
                            self.file_finished.emit(i, "➖ Skipped (user chose not to overwrite)", False)
                            continue # Skip to the next file
                        elif self.overwrite_choice == 'yes':
                            # --- ADD THIS LINE TO DELETE THE EXISTING FILE ---
                            os.remove(output_path) 
                            # If 'yes', proceed with saving (no explicit action here)
                    # --- END NEW OVERWRITE LOGIC ---

                    # Try with no password first, then with default password if provided
                    passwords_to_try = [None]
                    if self.default_password:
                        passwords_to_try.append(self.default_password)
                    if absolute_input_path in self.passwords:
                        passwords_to_try.append(self.passwords[absolute_input_path])

                    last_error = None
                    doc = None
                    
                    for password in passwords_to_try:
                        try:
                            doc = word_app.Documents.Open(
                                FileName=absolute_input_path,
                                ConfirmConversions=False,
                                ReadOnly=True,
                                AddToRecentFiles=False,
                                PasswordDocument="" if password is None else str(password),
                                Visible=False,  # Important: keep document hidden
                                OpenAndRepair=True
                            )
                            break  # Successfully opened
                        except Exception as e:
                            last_error = e
                            continue

                    if doc is None:
                        # If we still couldn't open, check if it's password protected
                        error_str = str(last_error).lower()
                        if "password" in error_str:
                            # Request password from main thread
                            self.password_required.emit(i, absolute_input_path)
                            
                            # Wait for password response
                            self.password_event_loop = QEventLoop()
                            self.password_event_loop.exec()
                            
                            if not self.current_password:
                                self.file_finished.emit(i, "❌ Skipped (password required)", False)
                                continue
                                
                            # Try with the new password
                            try:
                                doc = word_app.Documents.Open(
                                    FileName=absolute_input_path,
                                    ConfirmConversions=False,
                                    ReadOnly=True,
                                    AddToRecentFiles=False,
                                    PasswordDocument=str(self.current_password),
                                    Visible=False,
                                    OpenAndRepair=True
                                )
                                self.passwords[absolute_input_path] = self.current_password
                                self.current_password = None
                            except Exception as e:
                                self.file_finished.emit(i, f"❌ Error: {str(e)}", False)
                                continue
                        else:
                            self.file_finished.emit(i, f"❌ Error: {str(last_error)}", False)
                            continue

                    # Document successfully opened
                    self.overall_progress.emit(base_progress + 25)

                    # Accept all tracked changes for a clean final version
                    if doc.Revisions.Count > 0:
                        doc.AcceptAllRevisions()

                    self.overall_progress.emit(base_progress + 50)

                    # --- PERFORM THE SAVE HERE ---
                    doc.SaveAs(output_path, FileFormat=wdFormatPDF)
                    self.overall_progress.emit(base_progress + 90)

                    self.file_finished.emit(i, "✅ Converted", True)
                    success_count += 1

                except Exception as e:
                    self.file_finished.emit(i, f"❌ Error: {str(e)}", False)

                finally:
                    if doc:
                        doc.Close(0)  # 0 = wdDoNotSaveChanges

            self.overall_progress.emit(100)
            final_message = f"Batch complete. {success_count} of {total_files} files converted successfully."
            self.batch_finished.emit(final_message)

        except Exception as e:
            self.fatal_error.emit(f"A fatal error occurred: {e}\nEnsure MS Word is installed and not blocked.")

        finally:
            if word_app:
                word_app.Quit()

    def stop(self):
        self._is_running = False
        if self.password_event_loop and self.password_event_loop.isRunning():
            self.password_event_loop.quit()
        if self.overwrite_event_loop and self.overwrite_event_loop.isRunning():
            self.overwrite_event_loop.quit()

    def provide_password(self, password):
        self.current_password = password
        if self.password_event_loop and self.password_event_loop.isRunning():
            self.password_event_loop.quit()

    def set_overwrite_action(self, choice: str):
        self.overwrite_choice = choice
        if self.overwrite_event_loop and self.overwrite_event_loop.isRunning():
            self.overwrite_event_loop.quit()