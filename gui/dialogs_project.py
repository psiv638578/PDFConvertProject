# gui/dialogs_project.py (обновлённый с улучшенным выводом и гарантией сохранения текущего проекта)

from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QListWidget, QPushButton, QLabel,
    QListWidgetItem, QMessageBox, QInputDialog
)
import configparser
import os
from PyQt5.QtCore import Qt

class ProjectSelectDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Выбор проекта")
        self.setMinimumSize(600, 400)

        self.ini_path = os.path.join(os.path.dirname(__file__), "..", "setup.ini")
        self.config = configparser.ConfigParser()
        self.config.optionxform = str
        self.config.read(self.ini_path, encoding='utf-8')

        self.project_list = QListWidget()
        self.details_label = QLabel("Выберите проект слева, чтобы увидеть настройки")
        self.details_label.setTextInteractionFlags(self.details_label.textInteractionFlags() | Qt.TextSelectableByMouse)
        self.current_selection = None

        self.project_list.currentItemChanged.connect(self.update_preview)
        self.project_list.itemDoubleClicked.connect(self.accept_project)

        self.load_projects()

        # Кнопки
        btn_create = QPushButton("Создать")
        btn_delete = QPushButton("Удалить")
        btn_select = QPushButton("Выбрать")
        btn_ok = QPushButton("OK")
        btn_cancel = QPushButton("Отмена")

        btn_create.clicked.connect(self.create_project)
        btn_delete.clicked.connect(self.delete_project)
        btn_select.clicked.connect(self.accept_project)
        btn_ok.clicked.connect(self.accept)
        btn_cancel.clicked.connect(self.reject)

        # Layout
        main_layout = QVBoxLayout(self)
        hbox = QHBoxLayout()
        hbox.addWidget(self.project_list, 1)
        hbox.addWidget(self.details_label, 2)
        main_layout.addLayout(hbox)

        btns = QHBoxLayout()
        btns.addWidget(btn_create)
        btns.addWidget(btn_delete)
        btns.addWidget(btn_select)
        btns.addStretch()
        btns.addWidget(btn_ok)
        btns.addWidget(btn_cancel)
        main_layout.addLayout(btns)

    def load_projects(self):
        self.project_list.clear()
        current = self.config.get("global", "current_project", fallback="")

        selected_item = None
        for section in self.config.sections():
            if section != "global":
                item = QListWidgetItem(section)
                self.project_list.addItem(item)
                if section == current:
                    selected_item = item

        if selected_item:
            self.project_list.setCurrentItem(selected_item)

    def update_preview(self):
        item = self.project_list.currentItem()
        if not item:
            self.details_label.setText("Нет проекта")
            self.current_selection = None
            return

        name = item.text()
        self.current_selection = name

        if not self.config.has_section(name):
            self.details_label.setText("Раздел не найден.")
            return

        data = self.config.items(name)
        summary = [f"Текущий проект: {name}", ""]
        file_counter = 0
        for k, v in data:
            if k == "output_folder":
                summary.append(f"Папка сохранения PDF: {v}")
            elif k == "merged_pdf_path":
                summary.append(f"Итоговый PDF: {v}")
            elif k.startswith("source_files_"):
                file_counter += 1
                parts = [p.strip() for p in v.split("|")]
                filepath = parts[0] if len(parts) > 0 else ""
                sheets = parts[1] if len(parts) > 1 else "-"
                enabled = parts[2].lower() == "enabled" if len(parts) > 2 else True
                merge = parts[3].lower() == "merge" if len(parts) > 3 else True
                summary.append(f"Файл: {os.path.basename(filepath)}")
        if file_counter == 0:
            summary.append("\nНет файлов в проекте.")
        self.details_label.setText("\n".join(summary))

    def create_project(self):
        name, ok = QInputDialog.getText(self, "Новый проект", "Введите имя проекта:")
        if not ok or not name.strip():
            return
        name = name.strip()
        if self.config.has_section(name):
            QMessageBox.warning(self, "Ошибка", "Проект с таким именем уже существует.")
            return
        self.config.add_section(name)
        self.project_list.addItem(name)

    def delete_project(self):
        item = self.project_list.currentItem()
        if not item:
            return
        name = item.text()
        confirm = QMessageBox.question(self, "Удалить проект",
                                       f"Удалить проект '{name}'?",
                                       QMessageBox.Yes | QMessageBox.No)
        if confirm == QMessageBox.Yes:
            self.config.remove_section(name)
            self.project_list.takeItem(self.project_list.row(item))
            self.details_label.setText("")

    def accept_project(self):
        if not self.current_selection:
            QMessageBox.warning(self, "Нет выбора", "Выберите проект.")
            return
        if not self.config.has_section("global"):
            self.config.add_section("global")
        self.config.set("global", "current_project", self.current_selection)
        with open(self.ini_path, "w", encoding="utf-8") as f:
            self.config.write(f)
        self.accept()

    def get_selected_project(self):
        return self.project_list.currentItem().text() if self.project_list.currentItem() else None

