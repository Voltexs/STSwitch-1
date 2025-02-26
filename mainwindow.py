from PyQt6.QtWidgets import QMainWindow, QLabel, QVBoxLayout, QWidget

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        
        # Create main layout
        layout = QVBoxLayout()
        central_widget = QWidget()
        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)
        
        # Add status labels
        self.complus_status = QLabel("COM+ Applications: Waiting...")
        self.complus_status.setStyleSheet("color: gray;")
        layout.addWidget(self.complus_status)

        # ... rest of your initialization code ...

    def update_status(self, operation, message, color):
        if operation == "delete":
            self.delete_status.setText(f"Delete: {message}")
            self.delete_status.setStyleSheet(f"color: {color};")
        elif operation == "copy":
            self.copy_status.setText(f"Copy: {message}")
            self.copy_status.setStyleSheet(f"color: {color};")
        elif operation == "batch":
            self.batch_status.setText(f"Batch: {message}")
            self.batch_status.setStyleSheet(f"color: {color};")
        elif operation == "complus":  # Add this condition
            self.complus_status.setText(f"COM+ Applications: {message}")
            self.complus_status.setStyleSheet(f"color: {color};")

        # ... rest of your update_status method ... 