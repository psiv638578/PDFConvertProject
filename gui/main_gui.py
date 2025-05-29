# gui/main_gui.py

import os
import configparser
import webbrowser
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
from PyQt5.QtCore import Qt, pyqtSlot
from gui.dialogs_project import ProjectSelectDialog
from gui.dialogs_list import TaskListDialog
from gui.dialogs_excel import ExcelSheetsDialog
from core.converter_runner import ConvertWorker
from gui.dialogs_page_numbering import PageNumberingDialog

class MainGui(QMainWindow):
    def __init__(self):
        super().__init__()

        # Проверка setup.ini
        ini_path = os.path.join(os.getcwd(), "setup.ini")
        config = configparser.ConfigParser()
        config.optionxform = str  # сохранить регистр ключей

        ini_needs_creation = False
        ini_needs_fixing = False

        if not os.path.exists(ini_path):
            ini_needs_creation = True
        else:
            try:
                config.read(ini_path, encoding='utf-8')
                if "global" not in config or "current_project" not in config["global"]:
                    ini_needs_fixing = True
                else:
                    current_project = config.get("global", "current_project", fallback=None)
                    if current_project and current_project not in config:
                        ini_needs_fixing = True
            except Exception:
                ini_needs_fixing = True

        if ini_needs_creation:
            # Создаем новый setup.ini
            with open(ini_path, "w", encoding="utf-8") as f:
                f.write("[global]\ncurrent_project = Проект1\n\n[Проект1]\n")
            QMessageBox.information(self, "Отсутствует setup.ini",
                "Файл с настройками (setup.ini) не найден.\n"
                "Создан пустой файл-заготовка.\n\n"
                "Для дальнейшей работы необходимо:\n- выбрать исходные файлы;\n"
                "- указать папку для сохранения PDF;\n"
                "- задать имя объединённого файла (при необходимости).")

        elif ini_needs_fixing:
            with open(ini_path, "a", encoding="utf-8") as f:
                f.write("\n" + "-"*20 + "\n[global]\ncurrent_project = Проект1\n\n[Проект1]\n")
            QMessageBox.information(self, "Некорректный setup.ini",
                "Файл с настройками (setup.ini) имеет неправильную структуру.\n"
                "Он был дополнен необходимыми секциями.\n\n"
                "Для дальнейшей работы необходимо:\n- выбрать исходные файлы;\n"
                "- указать папку для сохранения PDF;\n"
                "- задать имя объединённого файла (при необходимости).")



        self.setWindowTitle("PDFConvert v1.2-beta")
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

        # Загружаем имя текущего проекта
        project_name = self.config.get("global", "current_project", fallback="(не задан)")

        # Метка и кнопка "Изменить..."
        self.project_label = QLabel(f"Текущий проект: {project_name}")
        self.project_button = QPushButton("Изменить...")
        self.project_button.clicked.connect(self.open_project_dialog)

        # Горизонтальное размещение project_label + кнопка
        project_layout = QHBoxLayout()
        project_layout.addWidget(self.project_label)
        project_layout.addWidget(self.project_button)
        project_layout.setContentsMargins(0, 0, 0, 20)  # left, top, right, bottom

        # Основной layout
        layout = QVBoxLayout()
        layout.addLayout(project_layout)
        
        # Кнопки
        btn_layout = QHBoxLayout()
        btn_start = QPushButton("Конвертировать")
        btn_cancel = QPushButton("Отмена")
        btn_start.clicked.connect(self.start_conversion)
        btn_cancel.clicked.connect(self.close)

        btn_layout.addWidget(btn_start)
        btn_layout.addWidget(btn_cancel)
        layout.addLayout(btn_layout)
        
        # Прогресс       
        self.progress = QProgressBar()
        self.progress.setValue(0)
        self.progress.setTextVisible(False)
        self.reset_progress_style_to_background()  # 👈 применить стили
        layout.addWidget(self.progress)

        # Статус
        self.status = QStatusBar()
        self.setStatusBar(self.status)
        self.status.setStyleSheet("""
            QStatusBar {
                border: 1px solid #aaa;
                margin: 2;
                min-height: 19px;
            }
        """) 
        
        # Завершаем: создаем центральный виджет и задаем layout
        central_widget = QWidget()
        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)

    def reset_progress_style_to_background(self):
        bg_color = self.palette().color(self.backgroundRole()).name()
        self.progress.setStyleSheet(f"""
            QProgressBar {{
                border: none;
                background-color: {bg_color};
            }}
        """)

    def set_progress_with_border(self):
        self.progress.setStyleSheet("""
            QProgressBar {
                border: 1px solid #999;
                border-radius: 3px;
                background-color: white;
            }
            QProgressBar::chunk {
                background-color: #3399ff;
                width: 10px;
            }
        """)


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
        # Пункт "Выбрать DOCX для нумерации"
        action_numbering = QAction("Выбрать файлы DOCX для нумерация страниц...", self)
        action_numbering.triggered.connect(self.open_page_numbering_dialog)
        file_menu.addAction(action_numbering)        
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
        selected_project = None  # ← инициализация по умолчанию
        if dlg.exec_():
            selected_project = dlg.get_selected_project()
        if selected_project:
            # 🔥 перечитываем setup.ini, чтобы сохранить свежие изменения (например, новую секцию)
            self.config.read(self.ini_path, encoding="utf-8")

            self.config.set("global", "current_project", selected_project)
            with open(self.ini_path, "w", encoding="utf-8") as configfile:
                self.config.write(configfile)

                self.status.showMessage(f"Проект переключён: {selected_project}", 1000)
                self.project_label.setText(f"Текущий проект: {selected_project}")
                self.config.read(self.ini_path, encoding="utf-8")

    def open_task_list_dialog(self):
        dlg = TaskListDialog(self, self.config)
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
        manual_path = os.path.abspath("manual.html")
        if os.path.exists(manual_path):
            webbrowser.open(f"file:///{manual_path}")
        else:
            QMessageBox.warning(self, "Ошибка", "Файл manual.html не найден.")

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
        self.set_progress_with_border()
        self.status.showMessage("Начинаем...")

        # Создание и запуск потока
        self.worker = ConvertWorker(
            ini_path=os.path.join(os.path.dirname(__file__), "..", "setup.ini")
        )

        # Подключения сигналов
        self.worker.update_status.connect(self.handle_status_message)
        self.worker.update_progress.connect(self.progress.setValue)
        self.worker.done.connect(self.conversion_finished)
        self.worker.show_info.connect(self.show_info_dialog)
        self.worker.show_blocking_dialog.connect(self.show_blocked_file_message)

        self.worker.start()

    def conversion_finished(self):
        self.progress.setValue(0)
        self.progress.setTextVisible(False)
        self.status.showMessage("Конвертация завершена", 2000)
        self.reset_progress_style_to_background()
        self.progress.setTextVisible(False)
        self.progress.setValue(0)


    def handle_status_message(self, text):
        if text.startswith("[BLOCKED]"):
            filename = text.replace("[BLOCKED] ", "")
            QMessageBox.critical(self, "Ошибка доступа", f"Файл «{filename}» заблокирован!\nВозможно, он открыт в другой программе.")
        else:
            self.status.showMessage(text, 3000)

    def open_page_numbering_dialog(self):
        from gui.dialogs_page_numbering import PageNumberingDialog
        dialog = PageNumberingDialog(self)
        dialog.exec_()

    def show_info_dialog(self, text):
        QMessageBox.information(self, "Информация", text)

    @pyqtSlot(str, int)
    def handle_status_message(self, text, duration):
        if text.startswith("[BLOCKED]"):
            filename = text.replace("[BLOCKED] ", "")
            QMessageBox.critical(self, "Ошибка доступа", f"Файл «{filename}» заблокирован!\nВозможно, он открыт в другой программе.")
        else:
            self.status.showMessage(text, duration)

    @pyqtSlot(str)
    def show_blocked_file_message(self, message):
        QMessageBox.critical(self, "Файлы заблокированы", message)
        QApplication.quit()

def run_gui():
    app = QApplication([])
    gui = MainGui()
    gui.show()
    app.exec_()

