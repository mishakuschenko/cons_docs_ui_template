from PySide6.QtWidgets import (
    QApplication,
    QDialog,
    QLabel,
    QLineEdit,
    QPushButton,
    QVBoxLayout,
    QFormLayout,
    QWidget
)

class ErrWindow(QDialog):
    def __init__(self, message="Неизвестная ошибка"):
        super().__init__()

        self.setWindowTitle("Ошибка")
        self.setFixedSize(300, 150)

        self.label = QLabel(message)
        self.label.setWordWrap(True)

        self.btn = QPushButton("ОК")
        self.btn.clicked.connect(self.accept)

        layout = QVBoxLayout()
        layout.addWidget(self.label)
        layout.addWidget(self.btn)

        self.setLayout(layout)

        self.setStyleSheet("font-size: 14pt; font-weight: bold;")
        
        




