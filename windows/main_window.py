from PySide6.QtWidgets import (
    QDialog,
    QLabel,
    QLineEdit,
    QPushButton,
    QVBoxLayout,
    QHBoxLayout,
    QFormLayout,
    QWidget,
    QFileDialog,
    QMessageBox
)
from utils import format_input, validate
from .err_window import *

import pandas as pd
import os

WIDTH = 1000
HEIGHT = 300
MAIN_WINDOW_NAME="СтройДок"

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.a = ""
        self.b = ""
        self.c = ""

        self.setFixedSize(WIDTH, HEIGHT)
        self.setWindowTitle(MAIN_WINDOW_NAME)

        self._margin = ""


        self.label1 = QLabel("Цена за м^2 руб", self)
        self.label2 = QLabel("Себестоимость м^2, руб", self)
        self.label3 = QLabel("Количество м^2", self)
        self._margin_field = QLabel("Маржинальность:", self)
        self.margin_label = QLabel(self._margin, self)


        self.margin_label.setStyleSheet("font-weight: bold; color: #2ecc71;")

        self.input_a = QLineEdit(self)
        self.input_b = QLineEdit(self)
        self.input_c = QLineEdit(self)


        self.input_a.returnPressed.connect(self._handle_enter)
        self.input_b.returnPressed.connect(self._handle_enter)
        self.input_c.returnPressed.connect(self._handle_enter)

        self.save_button = QPushButton("Сохранить", self)
        self.export_button = QPushButton("Загрузить отчет (.xlsx)", self)

        main_layout = QVBoxLayout()
        form_layout = QFormLayout()

        # 1. Заполняем форму ввода
        form_layout.addRow(self.label1, self.input_a)
        form_layout.addRow(self.label2, self.input_b)
        form_layout.addRow(self.label3, self.input_c)
        form_layout.addRow(QLabel(""))
        form_layout.addRow(self._margin_field, self.margin_label)

        main_layout.addLayout(form_layout)
        main_layout.addStretch()

        main_layout.addWidget(self.save_button)
        main_layout.addWidget(self.export_button)

        self.setLayout(main_layout)

        self.save_button.clicked.connect(self.save_data)
        self.export_button.clicked.connect(self.export_into_xlsx)

        self.setStyleSheet("font-size: 14pt;")

    def _handle_enter(self):
        from utils import format_input
        line_edit = self.sender()

        text = line_edit.text()
        if text == "":
            return
        formatted = format_input(float(text))

        line_edit.setText(formatted)
    def export_into_xlsx(self):
        if not all([self.input_a.text(), self.input_b.text(), self.input_c.text()]):
            ErrWindow("Нет данных! Сначала заполните поля.").exec()
            return

        if not all([self.a, self.b, self.c]):
            reply = QMessageBox.warning(
                self,
                "Предупреждение",
                "Данные не были сохранены (кнопка 'Сохранить').\n\n"
                "Экспортировать текущие данные из полей ввода?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No
            )

            if reply != QMessageBox.StandardButton.Yes:
                return

        file_path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить файл", "data.xlsx", "Excel Files (*.xlsx)"
        )
        if not file_path:
            return

        if not file_path.endswith(".xlsx"):
            file_path += ".xlsx"

        try:
            # Берем данные из полей ввода (актуальные)
            a = float(self.input_a.text().replace(" ", ""))
            b = float(self.input_b.text().replace(" ", ""))
            c = float(self.input_c.text().replace(" ", ""))
            margin = a * c - b * c

            df = pd.DataFrame({
                "Стоимость за м² (руб.)": [a],
                "Себестоимость за м² (руб.)": [b],
                "Количество (м²)": [c],
                "Маржинальность (руб.)": [margin]
            })

            df.to_excel(file_path, index=False)
            QMessageBox.information(self, "Успех", f"Данные сохранены в:\n{file_path}")

        except Exception as e:
            ErrWindow(f"Ошибка: {str(e)}").exec()


    def save_data(self):
        self.a = self.input_a.text()
        self.b = self.input_b.text()
        self.c = self.input_c.text()

        try:
            self.a = validate(self.a)
            self.b = validate(self.b)
            self.c = validate(self.c)

            self._ma = self.a * self.c - self.b * self.c

            if self._ma > 0:

                self.margin_label.setStyleSheet("font-weight: bold; color: #2ecc71;")

            else:

                self.margin_label.setStyleSheet("font-weight: bold; color: #c71824;")

            self._m = f"{(self._ma):_}".replace("_", " ")

            self.margin_label.setText(self._m + " " + "Rub")

        except (TypeError, ValueError):
            err_window = ErrWindow("Ошибка Ввода")
            err_window.exec()
