import sys
import os
import ctypes
import shutil
import subprocess
import time
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                            QLabel, QLineEdit, QPushButton, QProgressBar, QMessageBox, QFrame)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QFont, QIcon

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def run_as_admin():
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)

class SwitchWorker(QThread):
    progress = pyqtSignal(str)
    dll_progress = pyqtSignal(int, int, int)  # current, success, failed
    finished = pyqtSignal(bool, str)

    def __init__(self, patch_number):
        super().__init__()
        self.patch_number = str(patch_number)  # Convert to string to handle decimal points

    def parse_patch_number(self):
        # Split patch number into base and addon if it exists (e.g., "178.3" -> "178", "3")
        parts = self.patch_number.split('.')
        base_number = parts[0]
        addon_number = parts[1] if len(parts) > 1 else None
        return base_number, addon_number

    def find_matching_patch_folder(self, install_path):
        base_number, addon_number = self.parse_patch_number()
        
        def check_folder(folder_path):
            all_folders = [f for f in os.listdir(folder_path) 
                          if os.path.isdir(os.path.join(folder_path, f))]
            
            # Filter out folders containing "fix" or "fixes" (case insensitive)
            valid_folders = [f for f in all_folders 
                            if not any(fix in f.lower() 
                                     for fix in ['fix', 'fixes', 'hotfix'])]
            
            # First, find folders matching the base patch number
            base_matches = [f for f in valid_folders 
                           if f"Patch {base_number}".lower() in f.lower()]
            
            if not base_matches:
                return None
            
            if not addon_number:
                # If no addon number, return the first matching base patch
                return base_matches[0]
            
            # Look for addon number in the matching base patches with more flexible matching
            addon_patterns = [
                f"- Add On {addon_number}",
                f"- Add on {addon_number}",
                f"- Addon {addon_number}",
                f"_Add On {addon_number}",
                f"_Add on {addon_number}",
                f"_Addon {addon_number}"
            ]
            addon_matches = [f for f in base_matches 
                            if any(pattern.lower() in f.lower() for pattern in addon_patterns)]
            
            return addon_matches[0] if addon_matches else None

        # Check main folder
        matching_folder = check_folder(install_path)
        if matching_folder:
            folder_path = os.path.join(install_path, matching_folder)
            # Check if CISystems exists in this path
            if os.path.exists(os.path.join(folder_path, "CISystems")):
                return matching_folder
        
        # Check one level deeper for nested folders
        for folder in os.listdir(install_path):
            nested_path = os.path.join(install_path, folder)
            if os.path.isdir(nested_path):
                matching_folder = check_folder(nested_path)
                if matching_folder:
                    nested_full_path = os.path.join(nested_path, matching_folder)
                    if os.path.exists(os.path.join(nested_full_path, "CISystems")):
                        return os.path.join(folder, matching_folder)
        
        return None

    def register_dlls(self, folder_path):
        os.chdir(folder_path)
        
        # Show CMD window for regsvr32 commands
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags &= ~subprocess.STARTF_USESHOWWINDOW
        
        dll_files = [f for f in os.listdir() if f.lower().endswith('.dll')]
        total_dlls = len(dll_files)
        success_count = 0
        failed_count = 0

        for idx, dll in enumerate(dll_files, 1):
            try:
                self.progress.emit(f"Registering: {dll}")
                
                result = subprocess.run(
                    ["regsvr32", dll],
                    capture_output=True,
                    text=True,
                    startupinfo=startupinfo
                )
                
                if result.returncode == 0:
                    success_count += 1
                else:
                    failed_count += 1
                
                self.dll_progress.emit(idx, success_count, failed_count)
                
            except Exception as e:
                failed_count += 1
                self.dll_progress.emit(idx, success_count, failed_count)

        return success_count, failed_count

    def run(self):
        try:
            # Look for patch folder
            self.progress.emit("Searching for patch folder...")
            install_path = r"C:\SmtDB\Install"
            
            matching_folder = self.find_matching_patch_folder(install_path)
            
            if not matching_folder:
                base_number, addon_number = self.parse_patch_number()
                error_msg = f"No patch folder found for Patch {base_number}"
                if addon_number:
                    error_msg += f" Add-on {addon_number}"
                self.finished.emit(False, error_msg)
                return

            source_folder = os.path.join(install_path, matching_folder, "CISystems")
            destination_folder = r"C:\Program Files (x86)\CISystems"

            # Verify source folder exists and contains RegAllDll's.bat
            if not os.path.exists(source_folder):
                self.finished.emit(False, f"CISystems folder not found in patch: {source_folder}")
                return

            if not os.path.exists(os.path.join(source_folder, "RegAllDll's.bat")):
                self.finished.emit(False, f"RegAllDll's.bat not found in source folder: {source_folder}")
                return

            # Force remove existing destination if it exists
            if os.path.exists(destination_folder):
                self.progress.emit("Removing existing CISystems folder...")
                
                # Create a batch file with aggressive removal commands
                remove_batch = '''@echo off
taskkill /F /IM SmartTrade.exe 2>nul
taskkill /F /IM SmartTradeServer.exe 2>nul
taskkill /F /IM dllhost.exe 2>nul
taskkill /F /IM mmc.exe 2>nul

REM Stop COM+ services
net stop "COM+ System Application" /y
net stop "System Event Notification Service" /y
net stop COMSysApp /y

cd /d "C:\Program Files (x86)"

REM Take ownership and set permissions
takeown /F CISystems /R /D Y
icacls CISystems /grant administrators:F /T /C
icacls CISystems /reset /T
icacls CISystems /grant Everyone:F /T

REM Wait a moment
timeout /t 2 /nobreak

REM Aggressive delete attempts
rd /s /q CISystems
if exist CISystems (
    del /F /S /Q CISystems\*.*
    rd /s /q CISystems
)

REM Restart COM+ services
net start "COM+ System Application"
net start "System Event Notification Service"
net start COMSysApp

REM Create a flag file when done
echo Done > "%TEMP%\deletion_complete.flag"
exit
'''
                batch_path = os.path.join(os.environ['TEMP'], 'remove_folder.bat')
                flag_path = os.path.join(os.environ['TEMP'], 'deletion_complete.flag')
                
                # Remove flag file if it exists from previous run
                if os.path.exists(flag_path):
                    os.remove(flag_path)
                
                with open(batch_path, 'w') as f:
                    f.write(remove_batch)
                
                # Execute with highest privileges
                ctypes.windll.shell32.ShellExecuteW(
                    None,
                    "runas",
                    "cmd.exe",
                    f'/c "{batch_path}"',
                    None,
                    0  # SW_HIDE
                )
                
                # Wait for deletion to complete (max 30 seconds)
                max_wait = 30
                while max_wait > 0 and not os.path.exists(flag_path):
                    time.sleep(1)
                    max_wait -= 1
                    self.progress.emit(f"Waiting for folder deletion... {max_wait}s")
                
                # Clean up flag file
                if os.path.exists(flag_path):
                    os.remove(flag_path)
                
                # Verify deletion
                if os.path.exists(destination_folder):
                    self.finished.emit(False, "Failed to remove existing CISystems folder. Please close Component Services manually.")
                    return

                # Extra wait to ensure everything is cleaned up
                time.sleep(2)

            # Continue with the copy and registration
            self.progress.emit("Fast copying CISystems folder...")
            try:
                # Use robocopy for better copying
                copy_command = f'robocopy "{source_folder}" "{destination_folder}" /E /NFL /NDL /NJH /NJS /NC /NS /MT:8'
                subprocess.run(copy_command, shell=True, check=False)
                
                # Verify copy was successful and RegAllDll's.bat exists
                if not os.path.exists(destination_folder):
                    self.finished.emit(False, "Failed to copy CISystems folder")
                    return
                    
                if not os.path.exists(os.path.join(destination_folder, "RegAllDll's.bat")):
                    self.finished.emit(False, "RegAllDll's.bat was not copied successfully")
                    return

                self.progress.emit("Running RegAllDll's.bat...")
                
                # Create a batch file to run the commands
                batch_content = '''@echo off
cd /d "C:\Program Files (x86)\CISystems"
RegAllDll's.bat
pause
'''
                batch_path = os.path.join(os.environ['TEMP'], 'register_dlls.bat')
                with open(batch_path, 'w') as f:
                    f.write(batch_content)
                
                # Launch the batch file with ShellExecute to trigger UAC
                ctypes.windll.shell32.ShellExecuteW(
                    None, 
                    "runas",
                    "cmd.exe",
                    f'/k "{batch_path}"',
                    None,
                    1  # SW_SHOWNORMAL
                )
                
                self.finished.emit(True, "Patch switched successfully! Please wait for DLL registration to complete in the CMD window.\n\nPlease restart the application to switch patches again.")
                
            except Exception as e:
                self.finished.emit(False, f"Error during operation: {str(e)}")
                return

        except Exception as e:
            self.finished.emit(False, f"Error: {str(e)}")

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        # Window setup
        self.setWindowTitle('SmartTrade Patch Switcher')
        self.setFixedSize(400, 300)
        
        # Create central widget and layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        layout.setSpacing(20)
        layout.setContentsMargins(20, 20, 20, 20)

        # Title label
        title_label = QLabel('SmartTrade Patch Switcher')
        title_label.setFont(QFont('Arial', 16, QFont.Weight.Bold))
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title_label)

        # Patch number input
        self.patch_input = QLineEdit()
        self.patch_input.setPlaceholderText('Enter patch number...')
        self.patch_input.setFont(QFont('Arial', 12))
        layout.addWidget(self.patch_input)

        # Switch button
        self.switch_button = QPushButton('Switch Patch')
        self.switch_button.setFont(QFont('Arial', 12))
        self.switch_button.clicked.connect(self.start_switch)
        layout.addWidget(self.switch_button)

        # Add a horizontal spacer after the switch button
        layout.addSpacing(10)

        # Register DLLs only button
        self.register_button = QPushButton('Register DLLs Only')
        self.register_button.setFont(QFont('Arial', 12))
        self.register_button.clicked.connect(self.register_dlls)
        layout.addWidget(self.register_button)

        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setTextVisible(False)
        layout.addWidget(self.progress_bar)

        # Status label
        self.status_label = QLabel('')
        self.status_label.setFont(QFont('Arial', 10))
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.status_label.setWordWrap(True)
        layout.addWidget(self.status_label)

        # Add DLL progress labels
        self.dll_progress_frame = QFrame()
        self.dll_progress_frame.setStyleSheet("background-color: white; border-radius: 10px;")
        dll_progress_layout = QVBoxLayout(self.dll_progress_frame)
        dll_progress_layout.setContentsMargins(20, 20, 20, 20)

        self.dll_count_label = QLabel('DLL Registration Progress: 0/0')
        self.dll_count_label.setFont(QFont('Arial', 10))
        self.dll_count_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        dll_progress_layout.addWidget(self.dll_count_label)

        self.dll_status_label = QLabel('Successful: 0 | Failed: 0')
        self.dll_status_label.setFont(QFont('Arial', 10))
        self.dll_status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        dll_progress_layout.addWidget(self.dll_status_label)

        layout.addWidget(self.dll_progress_frame)
        self.dll_progress_frame.hide()  # Hidden by default

    def start_switch(self):
        patch_number = self.patch_input.text().strip()
        
        if not patch_number:
            QMessageBox.warning(self, 'Error', 'Please enter a patch number.')
            return

        # Disable input during processing
        self.patch_input.setEnabled(False)
        self.switch_button.setEnabled(False)
        self.progress_bar.setMaximum(0)  # Show indeterminate progress
        
        # Create and start worker thread
        self.worker = SwitchWorker(patch_number)
        self.worker.progress.connect(self.update_status)
        self.worker.finished.connect(self.switch_completed)
        self.worker.dll_progress.connect(self.update_dll_progress)
        self.worker.start()

        self.dll_progress_frame.show()
        self.dll_count_label.setText('DLL Registration Progress: 0/0')
        self.dll_status_label.setText('Successful: 0 | Failed: 0')

    def update_status(self, message):
        self.status_label.setText(message)

    def switch_completed(self, success, message):
        self.patch_input.setEnabled(True)
        self.switch_button.setEnabled(True)
        self.progress_bar.setMaximum(100)
        self.progress_bar.setValue(100 if success else 0)
        
        QMessageBox.information(self, 'Success' if success else 'Error', message)
        
        self.status_label.setText('')
        self.progress_bar.setValue(0)
        self.dll_progress_frame.hide()

    def register_dlls(self):
        self.register_button.setEnabled(False)
        self.switch_button.setEnabled(False)
        self.progress_bar.setMaximum(0)

        try:
            self.status_label.setText("Registering DLLs...")
            destination_folder = r"C:\Program Files (x86)\CISystems"
            
            if not os.path.exists(destination_folder):
                QMessageBox.critical(self, 'Error', 'CISystems folder not found!')
                return

            os.chdir(destination_folder)
            
            # Show CMD window
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags &= ~subprocess.STARTF_USESHOWWINDOW
            
            result = subprocess.run(
                ["cmd.exe", "/c", "regalldll's"], 
                capture_output=True,
                text=True,
                startupinfo=startupinfo
            )
            
            if result.returncode == 0:
                QMessageBox.information(self, 'Success', 'DLLs registered successfully!\n\nPlease restart the application to switch patches again.')
            else:
                QMessageBox.critical(self, 'Error', f'Failed to register DLLs: {result.stderr}')

        except Exception as e:
            QMessageBox.critical(self, 'Error', f'Error registering DLLs: {str(e)}')

        finally:
            # Re-enable buttons and reset progress
            self.register_button.setEnabled(True)
            self.switch_button.setEnabled(True)
            self.progress_bar.setMaximum(100)
            self.progress_bar.setValue(0)
            self.status_label.setText('')

        self.dll_progress_frame.show()
        self.dll_count_label.setText('DLL Registration Progress: 0/0')
        self.dll_status_label.setText('Successful: 0 | Failed: 0')

    def update_dll_progress(self, current, success, failed):
        total = success + failed
        self.dll_count_label.setText(f'DLL Registration Progress: {current}/{total}')
        self.dll_status_label.setText(f'Successful: {success} | Failed: {failed}')

def main():
    if not is_admin():
        # Re-run the program with admin rights
        run_as_admin()
        sys.exit()

    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    window = MainWindow()
    window.show()
    sys.exit(app.exec())

if __name__ == '__main__':
    main()
