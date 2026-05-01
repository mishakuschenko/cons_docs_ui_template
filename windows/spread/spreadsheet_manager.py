import gspread
from oauth2client.service_account import ServiceAccountCredentials
from typing import List, Dict, Any, Optional
import pandas as pd

class GoogleSheetsManager:
    def __init__(self, credentials_file: str, sheet_name: str = "СтройДок"):
        """
        Инициализация менеджера Google Sheets
        
        Args:
            credentials_file: путь к JSON файлу с учетными данными сервисного аккаунта
            sheet_name: название таблицы или ID таблицы
        """
        # Настройка доступа
        scope = [
            'https://spreadsheets.google.com/feeds',
            'https://www.googleapis.com/auth/drive'
        ]
        
        try:
            creds = ServiceAccountCredentials.from_json_keyfile_name(credentials_file, scope)
            self.client = gspread.authorize(creds)
            
            # Пытаемся открыть таблицу по имени или ID
            try:
                self.sheet = self.client.open_by_key(sheet_name).sheet1
            except gspread.SpreadsheetNotFound:
                # Если таблица не найдена, создаем новую
                self.sheet = self.client.create(sheet_name).sheet1
                print(f"Создана новая таблица: {sheet_name}")
                
        except Exception as e:
            raise Exception(f"Ошибка подключения к Google Sheets: {str(e)}")
    
    def save_data(self, data: Dict[str, Any]) -> bool:
        """
        Сохраняет данные в Google таблицу
        
        Args:
            data: словарь с данными для сохранения
        """
        try:
            # Проверяем, пустой ли лист
            existing_data = self.sheet.get_all_values()
            
            if not existing_data or len(existing_data) == 0:
                # Если лист пустой, добавляем заголовки
                headers = list(data.keys())
                self.sheet.append_row(headers)
            
            # Добавляем строку с данными
            row_values = list(data.values())
            self.sheet.append_row(row_values)
            
            return True
            
        except Exception as e:
            print(f"Ошибка сохранения в Google Sheets: {str(e)}")
            return False
    
    def get_all_records(self) -> List[Dict[str, Any]]:
        """Получает все записи из таблицы"""
        try:
            records = self.sheet.get_all_records()
            return records
        except Exception as e:
            print(f"Ошибка чтения из Google Sheets: {str(e)}")
            return []
    
    def get_last_record(self) -> Optional[Dict[str, Any]]:
        """Получает последнюю запись из таблицы"""
        try:
            all_records = self.get_all_records()
            if all_records:
                return all_records[-1]
            return None
        except Exception as e:
            print(f"Ошибка получения последней записи: {str(e)}")
            return None
    
    def export_to_excel(self, filename: str) -> bool:
        """Экспортирует данные из Google Sheets в Excel файл"""
        try:
            records = self.get_all_records()
            if records:
                df = pd.DataFrame(records)
                df.to_excel(filename, index=False)
                return True
            return False
        except Exception as e:
            print(f"Ошибка экспорта в Excel: {str(e)}")
            return False
