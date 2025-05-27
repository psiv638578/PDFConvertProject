import os
import configparser
import shutil
import pythoncom
from time import sleep
from PyQt5.QtCore import QThread, pyqtSignal
from PyQt5.QtWidgets import QMessageBox
from PyPDF2 import PdfMerger
from core.add_page_numbers import add_page_numbers
import win32com.client

class ConvertWorker(QThread):
    update_status = pyqtSignal(str)
    update_progress = pyqtSignal(int)
    done = pyqtSignal()
    show_info = pyqtSignal(str)     # определение сигнала

    def __init__(self, ini_path):
        super().__init__()
        self.ini_path = ini_path

    def run(self):
        # >>> Проверка параметров до старта
        config = configparser.ConfigParser()
        config.optionxform = str
        config.read(self.ini_path, encoding='utf-8')

        project_name = config.get("global", "current_project", fallback=None)
        if not project_name or not config.has_section(project_name):
            self.update_status.emit("Не выбран текущий проект или секция не найдена.")         
            sleep(1.5)
            self.done.emit()
            return

        section = config[project_name]

        # Проверка: исходные файлы заданы?
        items = [(k, v) for k, v in section.items() if k.startswith("source_files_")]
        if not items:
            self.update_status.emit("Исходные файлы не указаны.")
            sleep(1.5)
            self.done.emit()
            return

        # Проверка: папка вывода PDF
        output_folder = section.get("output_folder", "").strip()
        if not output_folder:
            self.update_status.emit("Папка сохранения ПДФ не указана.")
            sleep(1.5)
            self.done.emit()
            return

        if not os.path.exists(output_folder):
            os.makedirs(output_folder)

        # Проверка: путь объединенного PDF
        merged_pdf_path = section.get("merged_pdf_path", "").strip()
        has_merge = any("merge" in v.lower() for k, v in items)

        if not merged_pdf_path and has_merge:
            merged_pdf_path = os.path.join(output_folder, "Объединенный.pdf")
            merged_pdf_path = merged_pdf_path.replace("\\", "/")  # <-- это и есть нужное исправление
            config.set(project_name, "merged_pdf_path", merged_pdf_path)  # <== ВАЖНО
            self.show_info.emit(
                "Папка сохранения и имя объединенного ПДФ не указаны.\n"
                "Объединенный ПДФ-файл будет сохранен в папке вывода под именем 'Объединенный'."
            )

        with open(self.ini_path, "w", encoding="utf-8") as configfile:
            config.write(configfile)

        self.update_status.emit(">>> Запуск метода run()")
        self.update_status.emit(">>> Запуск метода run()")

        config = configparser.ConfigParser()
        config.optionxform = str
        config.read(self.ini_path, encoding='utf-8')

        project_name = config.get("global", "current_project", fallback=None)
        if not project_name or not config.has_section(project_name):
            self.update_status.emit("Не выбран текущий проект или секция не найдена.")
            self.done.emit()
            return

        output_folder = config.get(project_name, "output_folder", fallback=None)
        if not output_folder or not os.path.isdir(output_folder):
            self.update_status.emit("Папка вывода PDF не найдена.")
            self.done.emit()
            return

        items = sorted(
            [(k, v) for k, v in config.items(project_name) if k.startswith("source_files_")],
            key=lambda x: int(x[0].split('_')[-1])
        )

        self.update_status.emit(f"Файлов для обработки: {len(items)}")

        # Проверка существования всех исходных файлов ДО обработки
        for key, value in items:
            parts = [p.strip() for p in value.split("|")]
            if len(parts) < 4:
                self.update_status.emit(f"Недопустимая строка в setup.ini: {key} = {value}")
                self.done.emit()
                return
            path = parts[0]
            if not os.path.isfile(path):
                self.update_status.emit(f"Файл не найден: {path}")
                self.done.emit()
                return

        if not items:
            self.update_status.emit("Нет заданий для обработки.")
            self.done.emit()
            return

        merged_list = []
        processed = 0
        total = len(items)
        merged_created = False

        iConverter = None
        try:
            pythoncom.CoInitialize()
            from win32com.client import Dispatch, gencache
            kompas_api5 = gencache.EnsureModule("{0422828C-F174-495E-AC5D-D31014DBBE87}", 0, 1, 0)
            kompas_object = kompas_api5.KompasObject(
                Dispatch("Kompas.Application.5")._oleobj_.QueryInterface(
                    kompas_api5.KompasObject.CLSID, pythoncom.IID_IDispatch))

            kompas_api7 = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
            application = kompas_api7.IApplication(
                Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(
                    kompas_api7.IApplication.CLSID, pythoncom.IID_IDispatch))

            dll_path = kompas_object.ksSystemPath(5) + r"\Pdf2d.dll"
            iConverter = application.Converter(dll_path)

        except Exception as e:
            self.update_status.emit(f"Ошибка инициализации КОМПАС PDF2D: {e}")

        for key, value in items:
            parts = [p.strip() for p in value.split("|")]
            if len(parts) < 4:
                continue

            path, sheet, enabled, merge = parts
#            self.update_status.emit(f"Файл: {path}, enabled={enabled}, merge={merge}")

            if enabled.lower() != "enabled":
                continue

            if not os.path.isfile(path):
                self.update_status.emit(f"Файл не найден: {path}")
                continue

            ext = os.path.splitext(path)[1].lower()
            output_pdf = os.path.join(output_folder, os.path.splitext(os.path.basename(path))[0] + ".pdf")

            try:
                self.try_remove_existing(output_pdf)

                if ext == ".docx":
                    self.convert_docx(path, output_pdf)
                elif ext == ".xlsx":
                    self.convert_xlsx(path, output_pdf, sheet)
                elif ext == ".cdw":
                    if iConverter:
                        self.convert_cdw_pdf2d(iConverter, path, output_pdf)
                    else:
                        self.update_status.emit(f"Конвертер КОМПАС не инициализирован.")
                elif ext == ".pdf":
                    self.copy_pdf(path, output_pdf)
                else:
                    self.update_status.emit(f"Неизвестный формат: {path}")
                    continue

                if merge.lower() == "merge":
                    merged_list.append(output_pdf)

            except Exception as e:
                self.update_status.emit(f"Ошибка при обработке: {e}")

#            sleep(2)

            processed += 1
            percent = int((processed / total) * 100)
            self.update_progress.emit(percent)
            sleep(0.5)

            merged_pdf_path = section.get("merged_pdf_path", "").strip()

            self.update_status.emit(f"Файлы для объединения: {merged_list}")

        if merged_list and merged_pdf_path:
            try:
                merger = PdfMerger()
                for pdf_file in merged_list:
                    if os.path.exists(pdf_file):
                        merger.append(pdf_file)
                merger.write(merged_pdf_path)
                merger.close()
                self.update_status.emit(f"{os.path.basename(merged_pdf_path)} сохранён.")
                merged_created = True
            except Exception as e:
                self.update_status.emit(f"Ошибка при объединении PDF: {e}")
        else:
            self.update_status.emit("Объединение PDF не требуется.")

        if merged_created:
            sleep(1.5)

        if config.get(project_name, "add_page_numbers", fallback="no").lower() == "yes":
            start_page = 3 if config.get(project_name, "start_from_page3", fallback="no").lower() == "yes" else 1
            skip_pages = 2 if start_page == 3 else 0
            try:
                self.update_status.emit(f"Вызов нумерации: start={start_page}, skip={skip_pages}")
                add_page_numbers(merged_pdf_path, merged_pdf_path, start=start_page, skip=skip_pages)
                self.update_status.emit(f"Добавлены номера страниц (с {start_page}-го).")
            except Exception as e:
                self.update_status.emit(f"Ошибка нумерации PDF: {e}")

        # ✅ Завершающее сообщение
        self.update_status.emit("Конвертация завершена.")
        self.update_progress.emit(100)
        self.done.emit()

    def try_remove_existing(self, path):       # Удаляет файл, если он существует, и логирует ошибку при неудаче.

        if os.path.isfile(path):
            try:
                os.remove(path)
            except Exception:
                # Сообщаем о заблокированном файле
                self.update_status.emit(f"[BLOCKED] {os.path.basename(path)}")
                raise

    def convert_docx(self, input_path, output_path):
        self.try_remove_existing(output_path)
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(input_path, ReadOnly=1)
        doc.ExportAsFixedFormat(output_path, 17)
        doc.Close(False)
        word.Quit()
        self.update_status.emit(f"{os.path.basename(input_path)} конвертирован.")

    def convert_xlsx(self, input_path, output_path, sheet):
        self.try_remove_existing(output_path)
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        wb = excel.Workbooks.Open(input_path)
        if sheet != "-":
            try:
                ws = wb.Worksheets(sheet)
                ws.Select()
            except:
                pass
        wb.ExportAsFixedFormat(0, output_path)
        wb.Close(False)
        excel.Quit()
        self.update_status.emit(f"{os.path.basename(input_path)} конвертирован.")

    def convert_cdw_pdf2d(self, iConverter, input_path, output_path):
        self.try_remove_existing(output_path)
        result = iConverter.Convert(input_path, output_path, 0, False)
        if result:
            self.update_status.emit(f"{os.path.basename(input_path)} конвертирован.")
        else:
            self.update_status.emit(f"Ошибка при сохранении {os.path.basename(input_path)}.")

    def copy_pdf(self, input_path, output_path):
        self.try_remove_existing(output_path)
        shutil.copyfile(input_path, output_path)
