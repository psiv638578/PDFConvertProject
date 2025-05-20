# core/converter_runner.py (финальные версии всех convert_* с перезаписью и предупреждением)

import os
import configparser
import shutil
import pythoncom
from time import sleep
from PyQt5.QtCore import QThread, pyqtSignal
from PyQt5.QtWidgets import QMessageBox
import win32com.client
from PyPDF2 import PdfMerger

class ConvertWorker(QThread):
    update_status = pyqtSignal(str)
    update_progress = pyqtSignal(int)
    done = pyqtSignal()

    def __init__(self, ini_path):
        super().__init__()
        self.ini_path = ini_path

    def run(self):
        from PyPDF2 import PdfMerger

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

        if not items:
            self.update_status.emit("Нет заданий для обработки.")
            self.done.emit()
            return

        merged_list = []
        processed = 0
        total = len(items)
        merged_created = False

        # Инициализация PDF-конвертера КОМПАС
        iConverter = None
        try:
            import pythoncom
            from win32com.client import Dispatch, gencache
            pythoncom.CoInitialize()

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

            processed += 1
            percent = int((processed / total) * 100)
            self.update_progress.emit(percent)
            sleep(5)

        # Объединение PDF-файлов
        merged_created = False
        merged_pdf_path = config.get(project_name, "merged_pdf_path", fallback=None)

        if merged_list and merged_pdf_path:
            try:
                merger = PdfMerger()
                for pdf_file in merged_list:
                    if os.path.exists(pdf_file):
                        merger.append(pdf_file)
                merger.write(merged_pdf_path)
                merger.close()
                self.update_status.emit(f"{os.path.basename(merged_pdf_path)} успешно сохранён.")
                merged_created = True
            except Exception as e:
                self.update_status.emit(f"Ошибка при объединении PDF: {e}")
        else:
            self.update_status.emit("Объединение PDF не требуется.")

        # ⏱ Добавим задержку, чтобы сообщение об объединении отобразилось
        if merged_created:
            sleep(1.5)

        # ✅ Завершающее сообщение
        self.update_status.emit("Конвертация завершена.")
        self.update_progress.emit(100)
        self.done.emit()


    def try_remove_existing(self, path):
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

