import os
import shutil
import subprocess
import sys
import time
import win32com.client
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                            QPushButton, QLineEdit, QLabel, QProgressBar, QTextEdit, QFrame)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QColor

class StatusBox(QFrame):
    def __init__(self, title):
        super().__init__()
        self.setFrameStyle(QFrame.Shape.NoFrame)
        layout = QVBoxLayout(self)
        layout.setSpacing(5)
        layout.setContentsMargins(10, 10, 10, 10)
        
        self.title = QLabel(title)
        self.title.setStyleSheet("font-weight: bold;")
        self.status = QLabel("Pending...")
        self.status.setStyleSheet("color: #888888;")
        
        layout.addWidget(self.title)
        layout.addWidget(self.status)
    
    def set_status(self, status, color="white"):
        self.status.setText(status)
        self.status.setStyleSheet(f"color: {color};")

class WorkerThread(QThread):
    progress = pyqtSignal(int)
    log = pyqtSignal(str)
    finished = pyqtSignal()
    error = pyqtSignal(str)
    status_update = pyqtSignal(str, str, str)  # step_id, status, color

    def __init__(self, patch_number):
        super().__init__()
        self.patch_number = patch_number

    def find_patch_folder(self, patch_number):
        base_path = "C:\\SmtDB\\Install"
        parts = str(patch_number).split('.')
        main_patch = parts[0]
        addon = parts[1] if len(parts) > 1 else None
        
        for folder in os.listdir(base_path):
            if "fixes" in folder.lower():
                continue
                
            if f"Patch {main_patch}" in folder:
                patch_folder = os.path.join(base_path, folder)
                
                cisystems_path = os.path.join(patch_folder, "CISystems")
                if os.path.exists(cisystems_path):
                    return patch_folder
                    
                for subdir in os.listdir(patch_folder):
                    subdir_path = os.path.join(patch_folder, subdir)
                    if os.path.isdir(subdir_path):
                        cisystems_path = os.path.join(subdir_path, "CISystems")
                        if os.path.exists(cisystems_path):
                            return subdir_path
                
                raise ValueError(f"Found patch folder but no CISystems directory in: {patch_folder}")
        
        raise ValueError(f"Could not find folder for Patch {patch_number}")

    def copy_with_progress(self, source, destination):
        total_size = 0
        for dirpath, dirnames, filenames in os.walk(source):
            for f in filenames:
                fp = os.path.join(dirpath, f)
                total_size += os.path.getsize(fp)
        
        copied_size = 0
        for dirpath, dirnames, filenames in os.walk(source):
            structure = os.path.join(destination, os.path.relpath(dirpath, source))
            os.makedirs(structure, exist_ok=True)
            
            for f in filenames:
                src_file = os.path.join(dirpath, f)
                dst_file = os.path.join(structure, f)
                shutil.copy2(src_file, dst_file)
                copied_size += os.path.getsize(src_file)
                progress = int((copied_size / total_size) * 100)
                self.progress.emit(progress)
                self.log.emit(f"Copying: {f}")

    def run(self):
        try:
            program_files_path = "C:\\Program Files (x86)\\CISystems"
            
            # Find source folder
            self.status_update.emit("find", "Finding patch folder...", "blue")
            source_path = self.find_patch_folder(self.patch_number)
            source_path = os.path.join(source_path, "CISystems")
            self.log.emit(f"Found source path: {source_path}")
            self.status_update.emit("find", "✓ Found patch folder", "green")
            
            # Stop services and processes
            self.status_update.emit("delete", "Stopping services...", "blue")
            try:
                # Stop CISystems services
                services_to_stop = ["StartupServ", "CISystemsService"]
                processes_to_kill = ["StartupServ.exe", "CISystemsService.exe", "dllhost.exe"]
                
                # Stop services first
                for service in services_to_stop:
                    self.log.emit(f"Attempting to stop service {service}...")
                    subprocess.run(["net", "stop", service], 
                                 capture_output=True, 
                                 text=True, 
                                 shell=True)
                
                # Force kill any remaining processes
                for process in processes_to_kill:
                    self.log.emit(f"Force killing process {process}...")
                    subprocess.run(["taskkill", "/F", "/IM", process], 
                                 capture_output=True,
                                 text=True,
                                 shell=True)
                
                # Give processes time to fully stop
                time.sleep(3)
                self.log.emit("Services and processes stopped successfully")
                
            except Exception as e:
                self.log.emit(f"Warning when stopping services/processes: {str(e)}")
            
            # Delete existing directory
            self.status_update.emit("delete", "Deleting existing directory...", "blue")
            if os.path.exists(program_files_path):
                self.log.emit("Deleting existing CISystems directory...")
                
                # Try multiple times with longer waits
                max_attempts = 5  # Increased attempts
                for attempt in range(max_attempts):
                    try:
                        shutil.rmtree(program_files_path)
                        break
                    except Exception as e:
                        if attempt == max_attempts - 1:
                            raise
                        self.log.emit(f"Deletion attempt {attempt + 1} failed, retrying...")
                        time.sleep(3)  # Longer wait between attempts
                
                self.progress.emit(25)
            self.status_update.emit("delete", "✓ Deleted existing directory", "green")
            
            # Copy files
            self.status_update.emit("copy", "Copying files...", "blue")
            self.log.emit("Starting file copy...")
            self.copy_with_progress(source_path, program_files_path)
            self.status_update.emit("copy", "✓ Files copied", "green")
            self.progress.emit(50)
            
            # Run batch file
            self.status_update.emit("batch", "Running RegAllDll's.bat...", "blue")
            batch_path = os.path.join(program_files_path, "RegAllDll's.bat")
            self.log.emit("Running RegAllDll's.bat...")
            process = subprocess.Popen(
                batch_path,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                universal_newlines=True,
                shell=True,
                cwd=os.path.dirname(batch_path)
            )
            
            while True:
                line = process.stdout.readline()
                if not line and process.poll() is not None:
                    break
                if line:
                    self.log.emit(line.strip())
            
            self.status_update.emit("batch", "✓ Batch file completed", "green")

            # Add COM+ Applications cleanup and creation
            self.status_update.emit("complus", "Setting up COM+ Applications...", "blue")
            self.log.emit("Starting COM+ Applications setup...")

            try:
                catalog = win32com.client.Dispatch("COMAdmin.COMAdminCatalog")
                applications = catalog.GetCollection("Applications")
                applications.Populate()

                # List of applications to manage
                com_app_names = [
                    "SMT Customers",
                    "SMT General",
                    "SMT Printing",
                    "SMT Procedures",
                    "SMT Products",
                    "SMT Reports",
                    "SMT Sales",
                    "SMT Settings",
                    "SMT Suppliers",
                    "SMT Sync"
                ]

                # First, delete existing applications
                for i in range(applications.Count - 1, -1, -1):
                    app = applications.Item(i)
                    if app.Value("Name") in com_app_names:
                        self.log.emit(f"Removing existing COM+ application: {app.Value('Name')}")
                        applications.Remove(i)
                        applications.SaveChanges()

                # Wait a moment for COM+ to process deletions
                time.sleep(2)

                # Calculate progress increment per application
                progress_per_app = 25 / len(com_app_names)  # Assuming we want this phase to be 25% of total progress
                current_progress = 50  # Starting from 50% (after previous operations)

                # Add this before the COM+ setup code
                com_app_dirs = {
                    "SMT Customers": os.path.join(program_files_path, "ComPlus Applications", "Customers"),
                    "SMT General": os.path.join(program_files_path, "ComPlus Applications", "General"),
                    "SMT Printing": os.path.join(program_files_path, "ComPlus Applications", "Printing"),
                    "SMT Procedures": os.path.join(program_files_path, "ComPlus Applications", "Procedures"),
                    "SMT Products": os.path.join(program_files_path, "ComPlus Applications", "Products"),
                    "SMT Reports": os.path.join(program_files_path, "ComPlus Applications", "Reports"),
                    "SMT Sales": os.path.join(program_files_path, "ComPlus Applications", "Sales"),
                    "SMT Settings": os.path.join(program_files_path, "ComPlus Applications", "Settings"),
                    "SMT Suppliers": os.path.join(program_files_path, "ComPlus Applications", "Suppliers"),
                    "SMT Sync": os.path.join(program_files_path, "ComPlus Applications", "Sync")
                }

                # Now create new applications and register components
                for app_name in com_app_names:
                    try:
                        self.log.emit(f"Creating COM+ application: {app_name}")
                        self.status_update.emit("complus", f"Setting up {app_name}...", "blue")
                        
                        new_app = applications.Add()
                        new_app.SetValue("Name", app_name)
                        new_app.SetValue("ApplicationAccessChecksEnabled", True)
                        new_app.SetValue("Authentication", 2)
                        new_app.SetValue("AuthenticationCapability", 2)
                        applications.SaveChanges()

                        # Wait for application to be fully created
                        time.sleep(1)

                        if app_name in com_app_dirs:
                            component_dir = com_app_dirs[app_name]
                            if os.path.exists(component_dir):
                                for dll in os.listdir(component_dir):
                                    if dll.lower().endswith('.dll'):
                                        dll_path = os.path.join(component_dir, dll)
                                        try:
                                            catalog.InstallComponent(app_name, dll_path, "", "")
                                            self.log.emit(f"Registered component: {dll}")
                                        except Exception as e:
                                            self.log.emit(f"Error registering {dll}: {str(e)}")
                                            # Continue with next DLL even if one fails

                        # Add roles and permissions
                        roles = applications.GetCollection("Roles", new_app.Key)
                        roles.Populate()
                        
                        role = roles.Add()
                        role.SetValue("Name", "CreatorOwner")
                        roles.SaveChanges()

                        users_in_role = roles.GetCollection("UsersInRole", role.Key)
                        users_in_role.Populate()
                        
                        user = users_in_role.Add()
                        user.SetValue("User", "Everyone")
                        users_in_role.SaveChanges()

                        # After setting up roles and permissions
                        new_app.SetValue("RunForever", True)
                        applications.SaveChanges()

                        # Update progress
                        current_progress += progress_per_app
                        self.progress.emit(int(current_progress))

                    except Exception as e:
                        self.log.emit(f"Error setting up {app_name}: {str(e)}")
                        # Continue with next application even if one fails

                self.log.emit("COM+ Applications setup completed successfully")
                self.status_update.emit("complus", "✓ COM+ Setup completed", "green")
                try:
                    # Force close any open COM+ windows
                    os.system('taskkill /f /im mmc.exe /fi "windowtitle eq Component Services"')
                except:
                    pass
                
            except Exception as e:
                self.log.emit(f"Error in COM+ setup: {str(e)}")
                self.error.emit(str(e))

            self.progress.emit(100)
            self.finished.emit()
            
        except Exception as e:
            self.error.emit(str(e))
            self.status_update.emit("find", "✗ Error", "red")
            self.status_update.emit("delete", "✗ Error", "red")
            self.status_update.emit("copy", "✗ Error", "red")
            self.status_update.emit("batch", "✗ Error", "red")

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Patch Switcher Pro")
        self.setMinimumSize(900, 700)
        
        # Set the window background color
        self.setStyleSheet("""
            QMainWindow {
                background-color: #1e1e1e;
            }
            QWidget {
                background-color: #1e1e1e;
                color: #ffffff;
            }
        """)
        
        # Create central widget and layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        layout.setSpacing(20)
        layout.setContentsMargins(30, 30, 30, 30)
        
        # Create title section
        title_label = QLabel("Patch Switcher Pro")
        title_label.setStyleSheet("""
            QLabel {
                color: #4dabf7;
                font-size: 24px;
                font-weight: bold;
                padding: 10px;
                border-bottom: 2px solid #4dabf7;
            }
        """)
        layout.addWidget(title_label)
        
        # Create input section with modern styling
        input_layout = QHBoxLayout()
        input_layout.setSpacing(15)
        
        self.label = QLabel("Enter Patch Number:")
        self.label.setStyleSheet("color: #ffffff; font-size: 14px;")
        
        self.input = QLineEdit()
        self.input.setPlaceholderText("e.g., 180 or 180.1")
        self.input.setStyleSheet("""
            QLineEdit {
                padding: 10px;
                border: 2px solid #333333;
                border-radius: 5px;
                background-color: #2d2d2d;
                color: #ffffff;
                font-size: 14px;
            }
            QLineEdit:focus {
                border: 2px solid #4dabf7;
            }
        """)
        
        self.button = QPushButton("Switch Patch")
        self.button.setStyleSheet("""
            QPushButton {
                padding: 10px 20px;
                background-color: #4dabf7;
                color: white;
                border: none;
                border-radius: 5px;
                font-size: 14px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #3793dd;
            }
            QPushButton:pressed {
                background-color: #2d7ab8;
            }
            QPushButton:disabled {
                background-color: #666666;
            }
        """)
        
        input_layout.addWidget(self.label)
        input_layout.addWidget(self.input, 1)
        input_layout.addWidget(self.button)
        layout.addLayout(input_layout)
        
        # Create status boxes with modern design
        status_layout = QHBoxLayout()
        status_layout.setSpacing(10)
        self.status_boxes = {
            "find": StatusBox("Find Patch"),
            "delete": StatusBox("Delete Existing"),
            "copy": StatusBox("Copy Files"),
            "batch": StatusBox("Run Batch File"),
            "complus": StatusBox("COM+ Setup")
        }
        
        # Update StatusBox styling
        for box in self.status_boxes.values():
            box.setStyleSheet("""
                QFrame {
                    background-color: #2d2d2d;
                    border-radius: 8px;
                    padding: 10px;
                }
                QLabel {
                    background-color: transparent;
                    color: #ffffff;
                    font-size: 13px;
                }
            """)
            status_layout.addWidget(box)
        layout.addLayout(status_layout)
        
        # Create progress bar with modern styling
        self.progress = QProgressBar()
        self.progress.setStyleSheet("""
            QProgressBar {
                border: none;
                border-radius: 5px;
                background-color: #2d2d2d;
                height: 20px;
                text-align: center;
            }
            QProgressBar::chunk {
                background-color: #4dabf7;
                border-radius: 5px;
            }
        """)
        layout.addWidget(self.progress)
        
        # Create log window with improved styling
        self.log_window = QTextEdit()
        self.log_window.setReadOnly(True)
        self.log_window.setStyleSheet("""
            QTextEdit {
                background-color: #2d2d2d;
                color: #ffffff;
                border: none;
                border-radius: 8px;
                padding: 15px;
                font-family: 'Consolas', monospace;
                font-size: 12px;
                line-height: 1.5;
            }
            QScrollBar:vertical {
                border: none;
                background-color: #2d2d2d;
                width: 14px;
                margin: 15px 0 15px 0;
                border-radius: 0px;
            }
            QScrollBar::handle:vertical {
                background-color: #4dabf7;
                min-height: 30px;
                border-radius: 7px;
            }
            QScrollBar::handle:vertical:hover {
                background-color: #3793dd;
            }
            QScrollBar::sub-line:vertical, QScrollBar::add-line:vertical {
                border: none;
                background: none;
            }
        """)
        layout.addWidget(self.log_window)
        
        # Create credits section
        credits_layout = QHBoxLayout()
        credits_layout.setSpacing(5)

        credits_label = QLabel("Created by Cassie & Janco")
        credits_label.setStyleSheet("""
            QLabel {
                color: #666666;
                font-size: 12px;
                font-style: italic;
                padding: 5px;
            }
        """)
        credits_layout.addStretch()
        credits_layout.addWidget(credits_label)
        credits_layout.addStretch()
        layout.addLayout(credits_layout)
        
        # Connect button
        self.button.clicked.connect(self.start_process)

    def log_message(self, message, color="white"):
        self.log_window.append(f'<span style="color: {color};">{message}</span>')

    def start_process(self):
        self.button.setEnabled(False)
        self.log_window.clear()
        self.progress.setValue(0)
        
        # Reset status boxes
        for box in self.status_boxes.values():
            box.set_status("Pending...", "gray")
        
        # Create and start worker thread
        self.worker = WorkerThread(self.input.text().strip())
        self.worker.progress.connect(self.progress.setValue)
        self.worker.log.connect(lambda msg: self.handle_log(msg))
        self.worker.finished.connect(self.process_finished)
        self.worker.error.connect(self.handle_error)
        self.worker.status_update.connect(self.update_status)
        self.worker.start()
    
    def handle_log(self, message):
        if "Error:" in message or "failed" in message.lower():
            self.log_message(message, "#ff6b6b")  # Red for errors
        elif "success" in message.lower() or "completed" in message.lower():
            self.log_message(message, "#69db7c")  # Green for success
        elif "attempting" in message.lower() or "starting" in message.lower():
            self.log_message(message, "#4dabf7")  # Blue for actions
        elif "copying:" in message.lower():
            self.log_message(message, "#ffd43b")  # Yellow for file operations
        else:
            self.log_message(message, "white")  # Default color
    
    def update_status(self, step_id, status, color):
        self.status_boxes[step_id].set_status(status, color)
    
    def process_finished(self):
        self.button.setEnabled(True)
        self.log_message("Process completed successfully!", "#69db7c")
    
    def handle_error(self, error_msg):
        self.button.setEnabled(True)
        self.log_message(f"Error: {error_msg}", "#ff6b6b")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
