from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QLabel, QPushButton, QHBoxLayout, QCheckBox,
    QTableWidget, QTableWidgetItem, QAbstractItemView, QHeaderView
)
from PyQt5.QtCore import Qt
import configparser
import os

class TaskListDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Список заданий")
        self.setMinimumSize(600, 400)

        self.result_numbering = False
        self.result_skip_first_two = False
        self.source_paths = []

        layout = QVBoxLayout()

        self.table = QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(["Файл", "Параметры", "Статус", "Объединять"])
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        layout.addWidget(self.table)

        btn_layout = QHBoxLayout()

        self.btn_up = QPushButton("Вверх")
        self.btn_down = QPushButton("Вниз")
        self.btn_delete = QPushButton("Удалить")
        self.btn_ok = QPushButton("OK")
        self.btn_cancel = QPushButton("Отмена")

        btn_layout.addWidget(self.btn_up)
        btn_layout.addWidget(self.btn_down)
        btn_layout.addWidget(self.btn_delete)

        self.cb_numbering = QCheckBox("Нумеровать")
        self.cb_from_third = QCheckBox("С 3 листа")
        btn_layout.addWidget(self.cb_numbering)
        btn_layout.addWidget(self.cb_from_third)

        btn_layout.addStretch()
        btn_layout.addWidget(self.btn_ok)
        btn_layout.addWidget(self.btn_cancel)

        layout.addLayout(btn_layout)
        self.setLayout(layout)

        self.load_from_ini()
        self.update_checkbox_state()
        self.table.itemChanged.connect(self.update_checkbox_state)
        self.cb_numbering.stateChanged.connect(self.update_checkbox_state)

        self.btn_ok.clicked.connect(self.handle_accept)
        self.btn_cancel.clicked.connect(self.reject)

    def update_checkbox_state(self):
        has_merge = False
        for row in range(self.table.rowCount()):
            item = self.table.item(row, 3)
            if item and item.checkState() == Qt.Checked:
                has_merge = True
                break

        self.cb_numbering.setEnabled(has_merge)
        self.cb_from_third.setEnabled(self.cb_numbering.isChecked() and self.cb_numbering.isEnabled())

    def load_from_ini(self):
        ini_path = os.path.join(os.getcwd(), "setup.ini")
        if not os.path.exists(ini_path):
            return

        config = configparser.ConfigParser()
        config.optionxform = str
        config.read(ini_path, encoding='utf-8')

        section = config.get("global", "current_project", fallback="files")

        self.cb_numbering.setChecked(config.get(section, "add_page_numbers", fallback="no").lower() == "yes")
        self.cb_from_third.setChecked(config.get(section, "start_from_page3", fallback="no").lower() == "yes")

        if not config.has_section(section):
            return

        self.source_paths.clear()
        rows = [(k, v) for k, v in config.items(section) if k.startswith("source_files_")]
        rows.sort(key=lambda x: int(x[0].split("_")[-1]))

        for key, line in rows:
            parts = [p.strip() for p in line.split("|")]
            if len(parts) < 4:
                continue

            path, param, status, merge = parts
            row = self.table.rowCount()
            self.table.insertRow(row)
            self.source_paths.append(path)
            self.table.setItem(row, 0, QTableWidgetItem(os.path.basename(path)))
            self.table.setItem(row, 1, QTableWidgetItem(param))

            chk_status = QTableWidgetItem()
            chk_status.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
            chk_status.setCheckState(Qt.Checked if status.lower() == "enabled" else Qt.Unchecked)
            self.table.setItem(row, 2, chk_status)

            chk_merge = QTableWidgetItem()
            chk_merge.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
            chk_merge.setCheckState(Qt.Checked if merge.lower() == "merge" else Qt.Unchecked)
            self.table.setItem(row, 3, chk_merge)

        self.update_checkbox_state()

    def handle_accept(self):
        self.result_numbering = self.cb_numbering.isChecked()
        self.result_skip_first_two = self.cb_from_third.isChecked()

        ini_path = os.path.join(os.getcwd(), "setup.ini")
        config = configparser.ConfigParser()
        config.optionxform = str
        config.read(ini_path, encoding='utf-8')

        section = config.get("global", "current_project", fallback="files")
        if not config.has_section(section):
            config.add_section(section)

        config.set(section, "add_page_numbers", "yes" if self.result_numbering else "no")
        config.set(section, "start_from_page3", "yes" if self.result_skip_first_two else "no")

        for row in range(self.table.rowCount()):
            path = self.source_paths[row] if row < len(self.source_paths) else ""
            param = self.table.item(row, 1).text()
            status_item = self.table.item(row, 2)
            merge_item = self.table.item(row, 3)

            status = "enabled" if status_item.checkState() == Qt.Checked else "disabled"
            merge = "merge" if merge_item.checkState() == Qt.Checked else "merge not"

            config.set(section, f"source_files_{row + 1}", f"{path} | {param} | {status} | {merge}")

        with open(ini_path, "w", encoding='utf-8') as f:
            config.write(f)

        self.accept()
