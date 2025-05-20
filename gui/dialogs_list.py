# gui/dialogs_list.py (проверенная версия, работающая с текущим проектом)

from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QTableWidget, QTableWidgetItem,
    QPushButton, QCheckBox, QWidget, QHeaderView, QAbstractItemView,
    QMessageBox
)
from PyQt5.QtCore import Qt
import configparser
import os


class TaskListDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Список заданий")
        self.setMinimumSize(800, 400)

        self.ini_path = os.path.join(os.path.dirname(__file__), "..", "setup.ini")
        self.config = configparser.ConfigParser()
        self.config.optionxform = str
        self.config.read(self.ini_path, encoding="utf-8")

        self.project_name = self.config.get("global", "current_project", fallback=None)
        if not self.project_name or not self.config.has_section(self.project_name):
            QMessageBox.critical(self, "Ошибка", "Не выбран или не найден текущий проект в setup.ini")
            self.reject()
            return

        self.table = QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(["Файл", "Листы", "Конвертировать", "Объединять"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)

        btn_up = QPushButton("Вверх")
        btn_down = QPushButton("Вниз")
        btn_delete = QPushButton("Удалить")
        btn_ok = QPushButton("OK")
        btn_cancel = QPushButton("Отмена")

        btn_up.clicked.connect(self.move_up)
        btn_down.clicked.connect(self.move_down)
        btn_delete.clicked.connect(self.delete_row)
        btn_ok.clicked.connect(self.save_and_close)
        btn_cancel.clicked.connect(self.reject)

        hbox = QHBoxLayout()
        hbox.addWidget(btn_up)
        hbox.addWidget(btn_down)
        hbox.addWidget(btn_delete)
        hbox.addStretch()
        hbox.addWidget(btn_ok)
        hbox.addWidget(btn_cancel)

        layout = QVBoxLayout(self)
        layout.addWidget(self.table)
        layout.addLayout(hbox)

        self.load_data()

    def load_data(self):
        try:
            self.table.setRowCount(0)
            files = self.config.items(self.project_name)
            source_items = [(k, v) for k, v in files if k.startswith("source_files_")]
            sorted_files = sorted(source_items, key=lambda x: int(x[0].split("_")[-1]))

            for key, line in sorted_files:
                parts = [p.strip() for p in line.split("|")]
                file = parts[0] if len(parts) > 0 else ""
                sheet = parts[1] if len(parts) > 1 else "-"
                enabled = parts[2].lower() if len(parts) > 2 else "enabled"
                merge = parts[3].lower() if len(parts) > 3 else "merge"

                row = self.table.rowCount()
                self.table.insertRow(row)
                self.table.setItem(row, 0, QTableWidgetItem(file))
                self.table.setItem(row, 1, QTableWidgetItem(sheet))

                chk_enabled = QCheckBox()
                chk_enabled.setChecked(enabled == "enabled")
                self._set_checkbox(row, 2, chk_enabled)

                chk_merge = QCheckBox()
                chk_merge.setChecked(merge == "merge")
                self._set_checkbox(row, 3, chk_merge)

        except Exception as e:
            QMessageBox.critical(self, "Ошибка в load_data", str(e))

    def _set_checkbox(self, row, col, checkbox):
        container = QWidget()
        layout = QHBoxLayout(container)
        layout.addWidget(checkbox)
        layout.setAlignment(Qt.AlignCenter)
        layout.setContentsMargins(0, 0, 0, 0)
        container.setLayout(layout)
        self.table.setCellWidget(row, col, container)

    def move_up(self):
        row = self.table.currentRow()
        if row > 0:
            self._swap_rows(row, row - 1)
            self.table.selectRow(row - 1)

    def move_down(self):
        row = self.table.currentRow()
        if row < self.table.rowCount() - 1:
            self._swap_rows(row, row + 1)
            self.table.selectRow(row + 1)

    def _swap_rows(self, i, j):
        for col in range(2):
            text_i = self.table.item(i, col).text()
            text_j = self.table.item(j, col).text()
            self.table.item(i, col).setText(text_j)
            self.table.item(j, col).setText(text_i)
        for col in [2, 3]:
            chk_i = self.table.cellWidget(i, col).findChild(QCheckBox).isChecked()
            chk_j = self.table.cellWidget(j, col).findChild(QCheckBox).isChecked()
            self.table.cellWidget(i, col).findChild(QCheckBox).setChecked(chk_j)
            self.table.cellWidget(j, col).findChild(QCheckBox).setChecked(chk_i)

    def delete_row(self):
        row = self.table.currentRow()
        if row >= 0:
            self.table.removeRow(row)

    def save_and_close(self):
        self.config.remove_section(self.project_name)
        self.config.add_section(self.project_name)

        for i in range(self.table.rowCount()):
            file = self.table.item(i, 0).text()
            sheets = self.table.item(i, 1).text()
            enabled = "enabled" if self.table.cellWidget(i, 2).findChild(QCheckBox).isChecked() else "disabled"
            merge = "merge" if self.table.cellWidget(i, 3).findChild(QCheckBox).isChecked() else "merge not"
            line = f"{file} | {sheets} | {enabled} | {merge}"
            self.config.set(self.project_name, f"source_files_{i+1}", line)

        with open(self.ini_path, "w", encoding="utf-8") as f:
            self.config.write(f)
        self.accept()
