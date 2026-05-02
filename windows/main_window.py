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

from pptx import Presentation
from pptx.util import Inches, Pt

from utils import format_input, validate
from .err_window import ErrWindow
from .spread.spreadsheet_manager import GoogleSheetsManager
from .config import Config

import pandas as pd
import os

WIDTH = 1000
HEIGHT = 400
MAIN_WINDOW_NAME = "СтройДок"

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.is_saved = False

        self.a = ""
        self.b = ""
        self.c = ""

        self.setFixedSize(WIDTH, HEIGHT)
        self.setWindowTitle(MAIN_WINDOW_NAME)

        self._margin = ""
        
        # Загружаем конфигурацию
        self.config = Config()
        self.google_manager = None
        
        # Пытаемся подключиться к Google Sheets
        if self.config.settings["google_sheets"]["enabled"]:
            try:
                credentials_file = self.config.settings["google_sheets"]["credentials_file"]
                sheet_name = self.config.settings["google_sheets"]["sheet_name"]
                
                if os.path.exists(credentials_file):
                    self.google_manager = GoogleSheetsManager(credentials_file, sheet_name)
                    print("✅ Подключение к Google Sheets установлено")
                else:
                    print(f"⚠️ Файл учетных данных не найден: {credentials_file}")
                    print("Создайте сервисный аккаунт в Google Cloud и сохраните ключ как credentials.json")
            except Exception as e:
                print(f"❌ Ошибка подключения к Google Sheets: {str(e)}")

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
        self.google_sync_button = QPushButton("Синхронизировать с Google Таблицу", self)
        self.export_to_presentation = QPushButton("Экспортировать в Google Презентацию", self)

        #self.load_history_button = QPushButton("Загрузить историю из Google Sheets", self)

        main_layout = QVBoxLayout()
        form_layout = QFormLayout()

        # Заполняем форму ввода
        form_layout.addRow(self.label1, self.input_a)
        form_layout.addRow(self.label2, self.input_b)
        form_layout.addRow(self.label3, self.input_c)
        form_layout.addRow(QLabel(""))
        form_layout.addRow(self._margin_field, self.margin_label)

        main_layout.addLayout(form_layout)
        main_layout.addStretch()

        main_layout.addWidget(self.save_button)
        main_layout.addWidget(self.export_button)
        main_layout.addWidget(self.export_to_presentation)

        # Добавляем кнопки Google Sheets только если менеджер инициализирован
        if self.google_manager:
            main_layout.addWidget(self.google_sync_button)
            #main_layout.addWidget(self.load_history_button)

        self.setLayout(main_layout)

        self.save_button.clicked.connect(self.save_data)
        self.export_button.clicked.connect(self.export_into_xlsx)
        self.export_to_presentation.clicked.connect(self.export_to_google_presentation)
        
        if self.google_manager:
            self.google_sync_button.clicked.connect(self.sync_to_google_sheets)
            #self.load_history_button.clicked.connect(self.load_from_google_sheets)

        self.setStyleSheet("font-size: 14pt;")

    def _handle_enter(self):
        from utils import format_input
        line_edit = self.sender()

        text = line_edit.text()
        if text == "":
            return
        formatted = format_input(float(text))

        line_edit.setText(formatted)


    def export_to_google_presentation(self):
        """Генерирует презентацию на основе текущих данных и сохраняет ее в выбранное место"""

        if not self.is_saved:
            QMessageBox.warning(
                self,
                "Предупреждение",
                "Данные не были сохранены (кнопка 'Сохранить').\n"
            )
            return

        prs: Presentation = Presentation()

        # --- СЛАЙД 1: Титульный ---
        first_page_layout = prs.slide_layouts[0]
        first_slide = prs.slides.add_slide(first_page_layout)
        first_slide_title = first_slide.shapes.title
        first_slide_title.text = "Отчет о маржинальности проекта"

        try:
            _a = float(self.input_a.text().replace(" ", ""))
            _b = float(self.input_b.text().replace(" ", ""))
            _c = float(self.input_c.text().replace(" ", ""))
            _margin = _a * _c - _b * _c
        except ValueError:
            QMessageBox.warning(self, "Ошибка", "Пожалуйста, введите корректные числовые значения")
            return

        # --- СЛАЙД 2: Основные показатели (Список) ---
        main_page_layout = prs.slide_layouts[1] 
        main_slide = prs.slides.add_slide(main_page_layout)

        # Заголовок второго слайда
        main_slide_title = main_slide.shapes.title
        main_slide_title.text = "Основные показатели"

        # Обычно placeholder[1] на макете slide_layouts[1] — это блок для текста/списка
        body_shape = main_slide.placeholders[1]
        tf = body_shape.text_frame
        tf.text = f"Цена за м² (руб.): {_a}"

        # Добавляем новые пункты списка (параграфы)
        p = tf.add_paragraph()
        p.text = f"Себестоимость за м² (руб.): {_b}"
        p.level = 0 # Уровень вложенности списка

        p = tf.add_paragraph()
        p.text = f"Количество (м²): {_c}"
        p.level = 0

        # Добавляем пустую строку для отступа перед маржой
        p = tf.add_paragraph()
        p.text = ""

        # Добавляем итоговый показатель маржи
        p = tf.add_paragraph()
        p.text = f"Итоговая маржа: {_margin:.2f}" # .2f округлит до 2 знаков
        p.font.bold = True # Делаем жирным для акцента

        # --- СЛАЙД 3: Финальный ---
        # Можно использовать пустой макет (layout 6) или центрированный (layout 1 или 2)
        final_page_layout = prs.slide_layouts[1] 
        final_slide = prs.slides.add_slide(final_page_layout)

        # Убираем заголовок или используем его как центральный текст
        final_slide_title = final_slide.shapes.title
        final_slide_title.text = "Спасибо за внимание!"

        # Очищаем второй текстовый блок, если он не нужен на этом слайде
        if len(final_slide.placeholders) > 1:
            final_slide.placeholders[1].text = ""

        # --- СОХРАНЕНИЕ ---
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить файл", "presentation.pptx", "PowerPoint Files (*.pptx)"
        )

        if not file_path:
            return

        if not file_path.endswith(".pptx"):
            file_path += ".pptx"

        prs.save(file_path)


    def sync_to_google_sheets(self):
        """Синхронизирует текущие данные с Google Sheets"""
        if not self.google_manager:
            QMessageBox.warning(self, "Ошибка", "Google Sheets не настроен")
            return
        
        # Проверяем, есть ли данные для сохранения
        if not all([self.input_a.text(), self.input_b.text(), self.input_c.text()]):
            ErrWindow("Нет данных для синхронизации! Сначала заполните поля.").exec()
            return
        
        try:
            # Получаем текущие данные
            a = float(self.input_a.text().replace(" ", ""))
            b = float(self.input_b.text().replace(" ", ""))
            c = float(self.input_c.text().replace(" ", ""))
            margin = a * c - b * c
            
            # Подготавливаем данные для сохранения
            data = {
                "Цена за м² (руб.)": a,
                "Себестоимость за м² (руб.)": b,
                "Количество (м²)": c,
                "Маржинальность (руб.)": margin,
                "Дата": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")
            }
            
            # Сохраняем в Google Sheets
            if self.google_manager.save_data(data):
                QMessageBox.information(
                    self, 
                    "Успех", 
                    "Данные успешно синхронизированы с Google Sheets!"
                )
            else:
                QMessageBox.warning(self, "Ошибка", "Не удалось синхронизировать данные")
                
        except Exception as e:
            ErrWindow(f"Ошибка синхронизации: {str(e)}").exec()
    
    def load_from_google_sheets(self):
        """Загружает историю данных из Google Sheets"""
        if not self.google_manager:
            QMessageBox.warning(self, "Ошибка", "Google Sheets не настроен")
            return
        
        try:
            records = self.google_manager.get_all_records()
            
            if not records:
                QMessageBox.information(self, "Информация", "В Google Sheets нет сохраненных данных")
                return
            
            # Показываем последнюю запись в диалоговом окне
            last_record = records[-1]
            
            message = "Последние данные из Google Sheets:\n\n"
            message += f"💰 Цена за м²: {last_record.get('Цена за м² (руб.)', 'N/A')} руб.\n"
            message += f"🏭 Себестоимость за м²: {last_record.get('Себестоимость за м² (руб.)', 'N/A')} руб.\n"
            message += f"📐 Количество: {last_record.get('Количество (м²)', 'N/A')} м²\n"
            message += f"📈 Маржинальность: {last_record.get('Маржинальность (руб.)', 'N/A')} руб.\n"
            message += f"📅 Дата: {last_record.get('Дата', 'N/A')}\n"
            
            if len(records) > 1:
                message += f"\n📊 Всего записей в истории: {len(records)}"
            
            QMessageBox.information(self, "История из Google Sheets", message)
            
        except Exception as e:
            ErrWindow(f"Ошибка загрузки из Google Sheets: {str(e)}").exec()

    def export_into_xlsx(self):
        # if not all([self.input_a.text(), self.input_b.text(), self.input_c.text()]):
        #     ErrWindow("Нет данных! Сначала заполните поля.").exec()
        #     return

        if not self.is_saved:
            QMessageBox.warning(
                self,
                "Предупреждение",
                "Данные не были сохранены (кнопка 'Сохранить').\n"
            )
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

            self.is_saved = True
            
            # Автоматическая синхронизация при сохранении (опционально)
            if self.google_manager and hasattr(self, '_auto_sync') and self._auto_sync:
                self.sync_to_google_sheets()

        except (TypeError, ValueError):
            QMessageBox.warning(
                self,
                "Предупреждение",
                "Введите данные и повторите попытку.\n"
            )
            return
            #err_window = ErrWindow("Ошибка Ввода")
            #err_window.exec()
