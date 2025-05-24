from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
import io
import os

def add_page_numbers(input_pdf_path, output_pdf_path, start=1, skip=0):
    """
    Добавляет номера страниц в правый верхний угол каждой страницы PDF.
    Страница A4, шрифт Helvetica 10pt.

    :param input_pdf_path: путь к входному PDF
    :param output_pdf_path: путь для сохранения результата
    :param start: номер, с которого начинать нумерацию
    :param skip: количество первых страниц, которые нужно пропустить без нумерации
    """
    reader = PdfReader(input_pdf_path)
    writer = PdfWriter()

    page_width, page_height = A4

    for i, page in enumerate(reader.pages):
        packet = io.BytesIO()
        can = canvas.Canvas(packet, pagesize=A4)

        offset_x_mm = 11
        offset_y_mm = 10
        x = page_width - offset_x_mm * mm
        y = page_height - offset_y_mm * mm

        if i >= skip:
            number = str(start + i - skip)
            can.setFont("Helvetica-Oblique", 10)
            can.drawString(x, y, number)

        can.save()
        packet.seek(0)

        overlay_pdf = PdfReader(packet)
        page.merge_page(overlay_pdf.pages[0])
        writer.add_page(page)

    with open(output_pdf_path, "wb") as f:
        writer.write(f)
        

