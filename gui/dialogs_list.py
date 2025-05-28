from PyQt5.QtWidgets import (
    QWidget, QDialog, QVBoxLayout, QLabel, QPushButton, QHBoxLayout, QCheckBox,
    QTableWidget, QTableWidgetItem, QAbstractItemView, QHeaderView
)
from PyQt5.QtCore import Qt
import configparser
import os
from gui.dialogs_excel import ExcelSheetsDialog

class TaskListDialog(QDialog):
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

            display_param = "Все" if param == "-" or param.lower() == "all" else param
        

            if path.lower().endswith(".xlsx"):
                from PyQt5.QtWidgets import QWidget, QLabel, QHBoxLayout

                # Преобразуем параметр
                param = parts[1].strip()
                display_param = "Все" if param in ["-", "all"] else "Выборочно"

                # # Виджет с лейблом и кнопкой
                # cell_widget = QWidget()
                # layout = QHBoxLayout(cell_widget)
                # layout.setContentsMargins(0, 0, 0, 0)

                # label = QLabel(display_param)
                # label.setAlignment(Qt.AlignCenter)
                # layout.addWidget(label)

                # # Кнопка "..."
                # btn = QPushButton("...")
                # btn.setFixedSize(25, 22)
                # btn.setProperty("row", row)

                # # ВНИМАНИЕ: используем parts[0].strip() вместо file_path
                # actual_file_path = parts[0].strip()
                # btn.clicked.connect(lambda _, f=actual_file_path, item=label: self.open_excel_sheet_dialog(f, item))
                # layout.addWidget(btn)

                # layout.addStretch()
                # self.table.setCellWidget(row, 1, cell_widget)

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

        config.set(section, "add_page_numbers", "yes" if self.result_numbering else "no")
        config.set(section, "start_from_page3", "yes" if self.result_skip_first_two else "no")

        for row in range(self.table.rowCount()):
            path = self.source_paths[row] if row < len(self.source_paths) else ""
            param_raw = self.table.item(row, 1).text().strip()
            param = "all" if param_raw.lower() == "все" else param_raw
            status_item = self.table.item(row, 2)
            merge_item = self.table.item(row, 3)

            status = "enabled" if status_item.checkState() == Qt.Checked else "disabled"
            merge = "merge" if merge_item.checkState() == Qt.Checked else "merge not"

            config.set(section, f"source_files_{row + 1}", f"{path} | {param} | {status} | {merge}")

        with open(ini_path, "w", encoding='utf-8') as f:
            config.write(f)

        self.accept()
        
    # def open_excel_sheet_dialog(self):
    #     sender = self.sender()
    #     if not sender:
    #         return

    #     row = sender.property("row")
    #     if row is None:
    #         return

    #     file_path = self.source_paths[row]
    #     if not os.path.exists(file_path):
    #         QMessageBox.warning(self, "Файл не найден", f"Файл не найден:\n{file_path}")
    #         return

    #     dlg = ExcelSheetsDialog(self)

    #     if dlg.exec_():
    #         selected_sheets = dlg.get_selected_sheets()  # Предполагается, что метод реализован
    #         if selected_sheets:
    #             text = ", ".join(str(i + 1) for i in range(len(selected_sheets)))
    #         else:
    #             text = "Все"

    #         self.table.item(row, 1).setText(text)

    #         # Обновляем текст в ячейке QLabel, если она есть
    #         container = self.table.cellWidget(row, 1)
    #         if container:
    #             label = container.findChild(QLabel)
    #             if label:
    #                 label.setText(text)
    #                 label.setToolTip(", ".join(selected_sheets))  # ← можно оставить как подсказку с оригинальными именами

    # def open_excel_sheet_dialog(self, file_path, table_item):
    #     # Открываем существующий диалог (все .xlsx сразу)
    #     dlg = ExcelSheetsDialog(self.config, self)
    #     dlg.exec_()

    #     # После закрытия — получаем новое значение из конфига
    #     project = self.project_name
    #     if not self.config.has_section(project):
    #         return

    #     # Найти нужную строку в setup.ini по имени файла
    #     for key in self.config.options(project):
    #         if key.startswith("source_files_"):
    #             value = self.config.get(project, key)
    #             parts = value.split("|")
    #             if parts and parts[0].strip() == file_path:
    #                 param = parts[1].strip()
    #                 display = "Все" if param in ["-", "all"] else "Выборочно"
    #                 table_item.setText(display)
    #                 break
