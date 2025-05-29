from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
import io
import os

def add_page_numbers(input_pdf_path, output_pdf_path, start=1, skip=0):
    """
    Добавляет номера страниц в правый верхний угол каждой страницы PDF.
    :param input_pdf_path: путь к входному PDF
    :param output_pdf_path: путь для сохранения результата
    :param start: номер, с которого начинать нумерацию
    :param skip: количество первых страниц, которые нужно пропустить без нумерации
    """
    reader = PdfReader(input_pdf_path)
    writer = PdfWriter()

    for i, page in enumerate(reader.pages):

        print(f"i={i}, skip={skip}")

        packet = io.BytesIO()

        # Получаем фактические размеры страницы
        page_width = float(page.mediabox.width)
        page_height = float(page.mediabox.height)

        # Настраиваем холст с размерами текущей страницы
        can = canvas.Canvas(packet, pagesize=(page_width, page_height))

        # Настройки положения номера страницы
        offset_x_mm = 11  # отступ справа
        offset_y_mm = 10  # отступ сверху
        x = page_width - offset_x_mm * mm
        y = page_height - offset_y_mm * mm

        # Добавляем номер страницы, если не пропускаем
        if i >= skip:
            number = str(start + i - skip)
            can.setFont("Helvetica-Oblique", 10)
            can.drawString(x, y, number)

    # if i >= skip:
    #     number = str(start + i - skip)
    #     can.setFont("Helvetica-Bold", 12)
    #     can.setFillColorRGB(0, 0, 0)
    #     x = page_width - 15 * mm
    #     y = page_height - 15 * mm  # чуть ниже от верха
    #     can.drawString(x, y, number)

        # Завершаем работу с текущим холстом
        can.save()
        packet.seek(0)

        overlay = PdfReader(packet)
        page.merge_page(overlay.pages[0])
        writer.add_page(page)

    with open(output_pdf_path, "wb") as f:
        writer.write(f)