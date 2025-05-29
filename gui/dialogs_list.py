from PyQt5.QtWidgets import (
    QWidget, QDialog, QVBoxLayout, QLabel, QPushButton, QHBoxLayout, QCheckBox,
    QTableWidget, QTableWidgetItem, QAbstractItemView, QHeaderView
)
from PyQt5.QtCore import Qt
import configparser
import os
from gui.dialogs_excel import ExcelSheetsDialog

class TaskListDialog(QDialog):
    MERGE_COLUMN_INDEX = 3
    PROCESS_COLUMN_INDEX = 2

    def __init__(self, parent=None, config=None):
        super().__init__(parent)
        self.config = config
        self.ini_path = os.path.join(os.path.dirname(__file__), "..", "setup.ini")
        self.project_name = self.config.get("global", "current_project", fallback=None)

        self.setWindowTitle("Список заданий")
        self.setMinimumSize(600, 400)

        self.result_numbering = False
        self.result_skip_first_two = False
        self.source_paths = []

        layout = QVBoxLayout()

        self.table = QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(["Файл", "Листы", "Обрабатывать", "Объединять"])
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

        self.btn_up.clicked.connect(self.move_row_up)
        self.btn_down.clicked.connect(self.move_row_down)
        self.btn_delete.clicked.connect(self.delete_selected_row)


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
        self.table.blockSignals(True)  # 🔒 Отключаем сигналы

        has_merge = False
        for row in range(self.table.rowCount()):
            process_item = self.table.item(row, self.PROCESS_COLUMN_INDEX)
            merge_item = self.table.item(row, self.MERGE_COLUMN_INDEX)

            if process_item and merge_item:
                # 🔄 Обновление доступности чекбокса "Объединять"
                if process_item.checkState() == Qt.Checked:
                    merge_item.setFlags(merge_item.flags() | Qt.ItemIsEnabled)
                else:
                    merge_item.setFlags(merge_item.flags() & ~Qt.ItemIsEnabled)

                # 🔍 Проверка хотя бы одного активного merge-флага
                if merge_item.flags() & Qt.ItemIsEnabled and merge_item.checkState() == Qt.Checked:
                    has_merge = True

        self.table.blockSignals(False)  # 🔓 Включаем обратно
        self.table.viewport().update()

        # 🔁 Обновляем доступность внешних чекбоксов
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
            file_item = QTableWidgetItem(os.path.basename(path))
            file_item.setData(Qt.UserRole, path)  # сохраняем путь
            self.table.setItem(row, 0, file_item)

            display_param = "Все" if param == "-" or param.lower() == "all" else param
        

            if path.lower().endswith(".xlsx"):
                from PyQt5.QtWidgets import QWidget, QLabel, QHBoxLayout

                # Преобразуем параметр
                param = parts[1].strip()
                display_param = "Все" if param in ["-", "all"] else "Выборочно"

                # Сохраняем текст в скрытую структуру
                hidden = QTableWidgetItem(display_param)
                hidden.setFlags(Qt.ItemIsEnabled)
                self.table.setItem(row, 1, hidden)
            else:
                self.table.setItem(row, 1, QTableWidgetItem(display_param))

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

        # Записываем флаги нумерации
        config.set(section, "add_page_numbers", "yes" if self.result_numbering else "no")
        config.set(section, "start_from_page3", "yes" if self.result_skip_first_two else "no")

        # Удаляем старые записи source_files_
        keys_to_remove = [k for k in config[section] if k.startswith("source_files_")]
        for key in keys_to_remove:
            config.remove_option(section, key)

        # Сохраняем только видимые строки из таблицы
        file_index = 1
        for row in range(self.table.rowCount()):
            file_item = self.table.item(row, 0)
            param_item = self.table.item(row, 1)
            status_item = self.table.item(row, 2)
            merge_item = self.table.item(row, 3)

            # Проверка на наличие обязательных данных
            if not file_item:
                continue

            # path = file_item.text().strip()
            path = file_item.data(Qt.UserRole) or file_item.text().strip()

            if not path:
                continue  # Пропускаем пустые строки

            param_raw = param_item.text().strip() if param_item else "-"
            param = "all" if param_raw.lower() == "все" else param_raw

            status = "enabled" if status_item and status_item.checkState() == Qt.Checked else "disabled"
            merge = "merge" if merge_item and merge_item.checkState() == Qt.Checked else "merge not"

            config.set(section, f"source_files_{file_index}", f"{path} | {param} | {status} | {merge}")
            file_index += 1

        # Сохраняем результат
        with open(ini_path, "w", encoding='utf-8') as f:
            config.write(f)

        self.accept()
        
    # Обработчик изменения чекбокса "Обрабатывать"
    def on_process_checkbox_changed(self, state, row):
        merge_item = self.table.item(row, self.MERGE_COLUMN_INDEX)
        if merge_item:
            flags = merge_item.flags()
            if state == Qt.Checked:
                merge_item.setFlags(flags | Qt.ItemIsEnabled)
            else:
                merge_item.setFlags(flags & ~Qt.ItemIsEnabled)
            self.table.viewport().update()  # Принудительно перерисовать таблицу

    def delete_selected_row(self):
        row = self.table.currentRow()
        if row >= 0:
            self.table.removeRow(row)

            # Удалить путь из списка source_paths
            if row < len(self.source_paths):
                del self.source_paths[row]

            # Принудительно обновить состояние чекбоксов нумерации
            self.update_checkbox_state()

    def move_row_up(self):
        row = self.table.currentRow()
        if row > 0:
            self.swap_rows(row, row - 1)
            self.table.selectRow(row - 1)

    def move_row_down(self):
        row = self.table.currentRow()
        if row < self.table.rowCount() - 1:
            self.swap_rows(row, row + 1)
            self.table.selectRow(row + 1)

    def swap_rows(self, row1, row2):
        for col in range(self.table.columnCount()):
            item1 = self.table.item(row1, col)
            item2 = self.table.item(row2, col)

            # Создание новых ячеек
            new_item1 = QTableWidgetItem(item2.text() if item2 else "")
            new_item2 = QTableWidgetItem(item1.text() if item1 else "")

            # Если столбец содержит чекбокс (Обрабатывать или Объединять)
            if col in (self.PROCESS_COLUMN_INDEX, self.MERGE_COLUMN_INDEX):
                new_item1.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
                new_item1.setCheckState(item2.checkState() if item2 else Qt.Unchecked)

                new_item2.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
                new_item2.setCheckState(item1.checkState() if item1 else Qt.Unchecked)
            else:
                new_item1.setFlags(Qt.ItemIsEnabled)
                new_item2.setFlags(Qt.ItemIsEnabled)

            self.table.setItem(row1, col, new_item1)
            self.table.setItem(row2, col, new_item2)

        # Переместить соответствующие пути файлов
        if row1 < len(self.source_paths) and row2 < len(self.source_paths):
            self.source_paths[row1], self.source_paths[row2] = self.source_paths[row2], self.source_paths[row1]
