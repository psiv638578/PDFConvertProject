# gui/main_gui.py

import os
import configparser
from PyQt5.QtWidgets import (
    QApplication,
	QMainWindow,
	QFileDialog, QInputDialog,
	QMessageBox,
    QMenuBar,
	QStatusBar, QProgressBar,
	QAction,
	QWidget,
	QVBoxLayout, QHBoxLayout,
	QPushButton,
	QMenuBar, 
    QLabel
)
from PyQt5.QtCore import Qt
from gui.dialogs_project import ProjectSelectDialog
from gui.dialogs_list import TaskListDialog
from gui.dialogs_excel import ExcelSheetsDialog
from core.converter_runner import ConvertWorker


class MainGui(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDFConvert GUI")
        self.setMinimumSize(300, 150)

        self.ini_path = os.path.join(os.path.dirname(__file__), "..", "setup.ini")
        self.config = configparser.ConfigParser()
        self.config.optionxform = str
        self.config.read(self.ini_path, encoding="utf-8")

        self.status = QStatusBar()
        self.setStatusBar(self.status)

        self.init_ui()

    def init_ui(self):
        self.create_menu()
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout()
        central_widget.setLayout(layout)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout()

        # Метка текущего проекта — ВВЕРХУ
        project_name = self.config.get("global", "current_project", fallback="не выбран")
        self.project_label = QLabel(f"Текущий проект: {project_name}")
        layout.addWidget(self.project_label)

        # Кнопки
        btn_layout = QHBoxLayout()
        btn_start = QPushButton("Конвертировать")
        btn_cancel = QPushButton("Отмена")
        btn_start.clicked.connect(self.start_conversion)
        btn_cancel.clicked.connect(self.close)

        btn_layout.addWidget(btn_start)
        btn_layout.addWidget(btn_cancel)
        layout.addLayout(btn_layout)

        central_widget.setLayout(layout)

        # Статус + прогресс
        self.status = QStatusBar()
        self.setStatusBar(self.status)
        self.progress = QProgressBar()
        self.progress.setVisible(False)
        self.status.addPermanentWidget(self.progress)

    def create_menu(self):
        menubar = self.menuBar()

        file_menu = menubar.addMenu("Файл")
        file_menu.addAction("Выбрать проект...", self.open_project_dialog)
        file_menu.addSeparator()
        file_menu.addAction("Выбрать папку с исходными файлами", self.select_source_folder)
        file_menu.addAction("Выбрать исходные файлы", self.select_source_files)
        file_menu.addAction("Папка сохранения PDF", self.select_output_folder)
        file_menu.addAction("Выбрать итоговый PDF", self.select_merged_pdf_path)
        file_menu.addAction("Изменить имя итогового PDF", self.change_merged_pdf_name)
        file_menu.addSeparator()
        file_menu.addAction("Выход", self.close)

        edit_menu = menubar.addMenu("Редактирование")
        edit_menu.addAction("Список заданий", self.open_task_list_dialog)
        edit_menu.addAction("Таблицы Excel", self.open_excel_sheets_dialog)
        
        help_menu = menubar.addMenu("Справка")
        help_menu.addAction("Руководство пользователя", self.open_manual)
        help_menu.addAction("О программе", self.open_about)


    def get_project_config(self):
        project_name = self.config.get("global", "current_project", fallback=None)
        if not project_name or not self.config.has_section(project_name):
            QMessageBox.warning(self, "Проект не выбран", "Выберите проект перед редактированием.")
            return None, None
        return project_name, self.config

    def open_project_dialog(self):
        dlg = ProjectSelectDialog(self)
        if dlg.exec_():
            selected_project = dlg.get_selected_project()
            if selected_project:
                self.config.set("global", "current_project", selected_project)
                with open(self.ini_path, "w", encoding="utf-8") as configfile:
                    self.config.write(configfile)
                self.status.showMessage(f"Проект переключён: {selected_project}", 5000)
                self.project_label.setText(f"Текущий проект: {selected_project}")
                self.config.read(self.ini_path, encoding="utf-8")

    def open_task_list_dialog(self):
        dlg = TaskListDialog(self)
        dlg.exec_()

    def open_excel_sheets_dialog(self):
        dlg = ExcelSheetsDialog(self)
        dlg.exec_()

    def select_source_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Выбрать папку с исходными файлами")
        if folder:
            exts = (".docx", ".xlsx", ".cdw", ".pdf")
            files = [f for f in os.listdir(folder) if f.lower().endswith(exts)]
            project_name, config = self.get_project_config()
            if not project_name:
                return

            for key in list(config[project_name].keys()):
                if key.startswith("source_files_"):
                    config.remove_option(project_name, key)

            for i, name in enumerate(files, start=1):
                full = os.path.join(folder, name).replace("\\", "/")
                line = f"{full} | - | enabled | merge"
                config.set(project_name, f"source_files_{i}", line)

            with open(self.ini_path, "w", encoding="utf-8") as f:
                config.write(f)
            self.status.showMessage(f"Добавлено файлов: {len(files)}")

    def select_source_files(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Выбрать файлы", "", "Документы и PDF (*.docx *.xlsx *.cdw *.pdf)")
        if files:
            project_name, config = self.get_project_config()
            if not project_name:
                return

            max_index = max([int(k.split("_")[-1]) for k in config[project_name] if k.startswith("source_files_")], default=0)
            for i, path in enumerate(files, start=max_index + 1):
                line = f"{path.replace('\\', '/')} | - | enabled | merge"
                config.set(project_name, f"source_files_{i}", line)

            with open(self.ini_path, "w", encoding="utf-8") as f:
                config.write(f)
            self.status.showMessage(f"Добавлено вручную: {len(files)} файлов")

    def select_output_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Выбрать папку для сохранения PDF")
        if folder:
            project_name, config = self.get_project_config()
            if not project_name:
                return
            config.set(project_name, "output_folder", folder.replace("\\", "/"))
            with open(self.ini_path, "w", encoding="utf-8") as f:
                config.write(f)
            self.status.showMessage("Папка сохранения обновлена")

    def select_merged_pdf_path(self):
        file, _ = QFileDialog.getSaveFileName(self, "Выбрать итоговый PDF", "", "PDF (*.pdf)")
        if file:
            project_name, config = self.get_project_config()
            if not project_name:
                return
            config.set(project_name, "merged_pdf_path", file.replace("\\", "/"))
            with open(self.ini_path, "w", encoding="utf-8") as f:
                config.write(f)
            self.status.showMessage("Путь итогового PDF обновлён")

    def change_merged_pdf_name(self):
        project_name, config = self.get_project_config()
        if not project_name:
            return

        old_path = config.get(project_name, "merged_pdf_path", fallback="")
        folder = os.path.dirname(old_path) if old_path else ""
        new_name, _ = QFileDialog.getSaveFileName(self, "Введите имя итогового PDF", folder, "PDF (*.pdf)")
        if new_name:
            config.set(project_name, "merged_pdf_path", new_name.replace("\\", "/"))
            with open(self.ini_path, "w", encoding="utf-8") as f:
                config.write(f)
            self.status.showMessage("Имя итогового PDF обновлено")

    def open_manual(self):
        try:
            path = os.path.join(os.path.dirname(__file__), "..", "manual.txt")
            if os.path.exists(path):
                with open(path, "r", encoding="utf-8") as f:
                    text = f.read()
                QMessageBox.information(self, "Руководство пользователя", text)
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", str(e))

    def open_about(self):
        try:
            path = os.path.join(os.path.dirname(__file__), "..", "about.txt")
            if os.path.exists(path):
                with open(path, "r", encoding="utf-8") as f:
                    text = f.read()
                QMessageBox.information(self, "О программе", text)
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", str(e))

    def start_conversion(self):
        self.progress.setVisible(True)
        self.progress.setValue(0)
        self.status.showMessage("Начинаем...")

        # Создание и запуск потока
        self.worker = ConvertWorker(
            ini_path=os.path.join(os.path.dirname(__file__), "..", "setup.ini")
        )

        # Подключения сигналов
        self.worker.update_status.connect(self.handle_status_message)
        self.worker.update_progress.connect(self.progress.setValue)
        self.worker.done.connect(self.conversion_finished)

        self.worker.start()

    def conversion_finished(self):
        self.progress.setVisible(False)
        self.status.showMessage("Конвертация завершена", 3000)

    def handle_status_message(self, text):
        if text.startswith("[BLOCKED]"):
            filename = text.replace("[BLOCKED] ", "")
            QMessageBox.critical(self, "Ошибка доступа", f"Файл «{filename}» заблокирован!\nВозможно, он открыт в другой программе.")
        else:
            self.status.showMessage(text, 5000)

def run_gui():
    app = QApplication([])
    gui = MainGui()
    gui.show()
    app.exec_()
