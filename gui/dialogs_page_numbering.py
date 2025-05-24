from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QLabel, QPushButton,
    QListWidget, QFileDialog, QMessageBox, QHBoxLayout, QSpinBox, QWidget
)
from PyQt5.QtCore import Qt
import os
from core.page_numbering import number_docx_pages

class PageNumberingDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Нумерация страниц в DOCX")
        self.setMinimumSize(500, 400)

        self.selected_files = []

        layout = QVBoxLayout()

        self.label = QLabel("Выберите .docx файлы для нумерации:")
        layout.addWidget(self.label)

        self.list_widget = QListWidget()
        layout.addWidget(self.list_widget)

        self.btn_select = QPushButton("Выбрать файлы")
        self.btn_select.clicked.connect(self.select_files)
        layout.addWidget(self.btn_select)

        # Поле "Номер первого листа" с выравниванием по центру
        center_widget = QWidget()
        center_layout = QHBoxLayout()
        center_layout.setContentsMargins(0, 0, 0, 0)
        center_layout.setSpacing(10)
        center_layout.setAlignment(Qt.AlignCenter)

        start_label = QLabel("Номер первого листа:")
        self.start_number_input = QSpinBox()
        self.start_number_input.setMinimum(1)
        self.start_number_input.setMaximum(9999)
        self.start_number_input.setValue(3)
        self.start_number_input.setFixedWidth(70)

        center_layout.addWidget(start_label)
        center_layout.addWidget(self.start_number_input)
        center_widget.setLayout(center_layout)
        layout.addWidget(center_widget)

        self.btn_run = QPushButton("Нумеровать")
        self.btn_run.clicked.connect(self.run_numbering)
        layout.addWidget(self.btn_run)

        self.setLayout(layout)

    def select_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "Выберите файлы .docx",
            "",
            "Word Documents (*.docx)"
        )
        if files:
            self.selected_files = sorted(files)
            self.list_widget.clear()
            for f in self.selected_files:
                self.list_widget.addItem(os.path.basename(f))

    def run_numbering(self):
        if not self.selected_files:
            QMessageBox.warning(self, "Ошибка", "Сначала выберите хотя бы один файл.")
            return

        try:
            start_number = self.start_number_input.value()
            number_docx_pages(self.selected_files, start_number)
            QMessageBox.information(self, "Готово", "Нумерация завершена успешно.")
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Произошла ошибка:\n{str(e)}")

        self.accept()
