import sqlite3
import easyocr
import re
import pandas as pd
from PyQt5.QtWidgets import QApplication, QFileDialog
from config import search_company, search_contract, year, search_company_short


class DatabaseManager:
    def __init__(self, db_name):
        self.conn = sqlite3.connect(db_name)
        self.cursor = self.conn.cursor()
        self.create_table()

    def create_table(self):
        self.cursor.execute('''CREATE TABLE IF NOT EXISTS results
                               (id INTEGER NOT NULL,
                               iin INTEGER,
                               bin INTEGER,
                               at_company TEXT,
                               to_company TEXT,
                               contract TEXT,
                               PRIMARY KEY(id AUTOINCREMENT))''')

    def insert_data(self, iin, bin, contract, at_company, to_company):
        self.cursor.execute("INSERT INTO results (iin, bin, contract, at_company, to_company) VALUES (?,?,?,?,?)",
                            (iin, bin, contract, at_company, to_company))
        print("Данные записаны в БД")

    def export_to_excel(self, filename):
        query = "SELECT * FROM results"
        df = pd.read_sql_query(query, self.conn)
        df.to_excel(filename, index=False)

    def commit(self):
        self.conn.commit()

    def close(self):
        self.conn.commit()
        self.conn.close()


class OCRProcessor:
    def __init__(self, lang):
        self.reader = easyocr.Reader(lang)

    def process_image(self, image_path, db_manager):
        global contract
        result = self.reader.readtext(image_path)
        for detection in result:
            print(detection)
            text = detection[1]

            # найти iin и bin
            if text.isdigit() and len(text) > 6:
                iin_bin.append(text)  # добавляем в массив иин и бин
                print("Найденный элемент:", text)

            # найти название компании
            if re.search(search_company, text, re.IGNORECASE):
                company.append(text)  # добавляем в массив иин и бин
                print("Найдена компания:", text)

            if re.search(year, text):
                year_in_list.append(text)
                print("Найден аргумент похожий на дату:", text)

            # Контракт/Договор
            for text_split in text.split():
                count = sum(1 for a, b in zip(text_split, search_contract) if a == b)
                if count >= 5:
                    contract = text
                    print(contract)

        iin_bin_1 = int(iin_bin[0]) if len(iin_bin) > 0 else None
        iin_bin_2 = int(iin_bin[1]) if len(iin_bin) > 1 else None
        contract = contract or None
        company_1 = company[0] if len(company) > 0 else None
        company_2 = company[1] if len(company) > 1 else None
        db_manager.insert_data(iin_bin_1, iin_bin_2, contract, company_1, company_2)


class FileDialog:
    def __init__(self):
        self.app = QApplication([])

    def open_file_dialog(self):
        dialog = QFileDialog()
        file_path = dialog.getOpenFileName()[0]
        return file_path


print(f"Добавь свой файл:")

file_dialog = FileDialog()
file_path = file_dialog.open_file_dialog()
print(f"Выбранный файл: {file_path}")

iin_bin = []
company = []
year_in_list = []

db_manager = DatabaseManager('ocr_results.db')
ocr_processor = OCRProcessor(['en', 'ru'])
ocr_processor.process_image(file_path, db_manager)
db_manager.commit()
db_manager.export_to_excel('result.xlsx')
db_manager.close()

print(f"Работа завершина!")
input("Нажми Enter")
