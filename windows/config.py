import json
import os

class Config:
    def __init__(self, config_file="settings.json"):
        self.config_file = config_file
        self.settings = self.load_config()
    
    def load_config(self):
        """Загружает конфигурацию из файла"""
        default_config = {
            "google_sheets": {
                "credentials_file": "credentials.json",  # Файл с ключами от Google API
                "sheet_name": "СтройДок_Данные",
                "enabled": True  # Включить/выключить синхронизацию с Google Sheets
            }
        }
        
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r') as f:
                    config = json.load(f)
                    # Мержим с дефолтными настройками
                    for key in default_config:
                        if key not in config:
                            config[key] = default_config[key]
                    return config
            except json.JSONDecodeError:
                print("Ошибка чтения конфигурации, использую настройки по умолчанию")
                return default_config
        else:
            # Создаем файл конфигурации с настройками по умолчанию
            with open(self.config_file, 'w') as f:
                json.dump(default_config, f, indent=4)
            print(f"Создан файл конфигурации: {self.config_file}")
            return default_config
    
    def save_config(self):
        """Сохраняет конфигурацию в файл"""
        with open(self.config_file, 'w') as f:
            json.dump(self.settings, f, indent=4)
