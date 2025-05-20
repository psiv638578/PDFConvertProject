# gui/dialogs_excel.py (проверенная версия, работающая с текущим проектом)

from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QListWidget, QPushButton,
    QLabel, QListWidgetItem, QMessageBox
)
from PyQt5.QtCore import Qt
import os
import configparser
import openpyxl


class ExcelSheetsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Выбор листов Excel")
        self.setMinimumSize(600, 400)

        self.ini_path = os.path.join(os.path.dirname(__file__), "..", "setup.ini")
        self.config = configparser.ConfigParser()
        self.config.optionxform = str
        self.config.read(self.ini_path, encoding='utf-8')

        self.project_name = self.config.get("global", "current_project", fallback=None)
        if not self.project_name or not self.config.has_section(self.project_name):
            QMessageBox.critical(self, "Ошибка", "Не выбран или не найден текущий проект в setup.ini")
            self.reject()
            return

        self.xlsx_entries = self.find_excel_entries()
        self.current_key = None

        self.file_list = QListWidget()
        self.sheet_list = QListWidget()
        self.label_info = QLabel("Выберите файл слева, чтобы увидеть листы")

        for key, (path, _, _, _) in self.xlsx_entries.items():
            item = QListWidgetItem(os.path.basename(path))
            self.file_list.addItem(item)

        self.file_list.currentRowChanged.connect(self.load_sheets)

        btn_ok = QPushButton("OK")
        btn_cancel = QPushButton("Отмена")
        btn_ok.clicked.connect(self.save_and_close)
        btn_cancel.clicked.connect(self.reject)

        layout = QVBoxLayout(self)
        layout.addWidget(self.label_info)
        hbox = QHBoxLayout()
        hbox.addWidget(self.file_list, 1)
        hbox.addWidget(self.sheet_list, 2)
        layout.addLayout(hbox)

        buttons = QHBoxLayout()
        buttons.addStretch()
        buttons.addWidget(btn_ok)
        buttons.addWidget(btn_cancel)
        layout.addLayout(buttons)

    def find_excel_entries(self):
        entries = {}
        if not self.config.has_section(self.project_name):
            return entries

        for key, line in self.config.items(self.project_name):
            if not key.startswith("source_files_"):
                continue
            parts = [p.strip() for p in line.split("|")]
            if len(parts) < 4:
                continue
            path, sheets, enabled, merge = parts
            if path.lower().endswith(".xlsx") and os.path.exists(path):
                entries[key] = (path, sheets, enabled, merge)
        return entries

    def load_sheets(self, row_index):
        self.sheet_list.clear()
        keys = list(self.xlsx_entries.keys())
        if row_index < 0 or row_index >= len(keys):
            return

        key = keys[row_index]
        path, saved_sheets, *_ = self.xlsx_entries[key]
        self.current_key = key

        try:
            wb = openpyxl.load_workbook(path, read_only=True)
            available_sheets = wb.sheetnames
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось открыть Excel: {e}")
            return

        saved_set = set(s.strip() for s in saved_sheets.split(",")) if saved_sheets != "-" else set()

        for name in available_sheets:
            item = QListWidgetItem(name)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Checked if name in saved_set else Qt.Unchecked)
            self.sheet_list.addItem(item)

    def save_and_close(self):
        keys = list(self.xlsx_entries.keys())

        for i, key in enumerate(keys):
            path, _, enabled, merge = self.xlsx_entries[key]
            selected_sheets = []

            if i == self.file_list.currentRow():
                for j in range(self.sheet_list.count()):
                    item = self.sheet_list.item(j)
                    if item.checkState() == Qt.Checked:
                        selected_sheets.append(item.text())
            else:
                selected_sheets = self.xlsx_entries[key][1].split(",")

            sheet_str = ", ".join(selected_sheets) if selected_sheets else "-"
            line = f"{path} | {sheet_str} | {enabled} | {merge}"
            self.config.set(self.project_name, key, line)

        with open(self.ini_path, "w", encoding="utf-8") as f:
            self.config.write(f)

        self.accept()
