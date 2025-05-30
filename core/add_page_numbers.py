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

    print("=== add_page_numbers ===")
    print(f"input: {input_pdf_path}")
    print(f"output: {output_pdf_path}")
    print(f"start: {start}, skip: {skip}")
    print(f"pages: {len(reader.pages)}")


    for i, page in enumerate(reader.pages):

        print(f"Page {i}, skip={skip}")

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

            print(f"Рисуем номер: {number}")

            print(f"type(i)={type(i)}, type(skip)={type(skip)}")

            can.setFont("Helvetica-Oblique", 10)
            can.drawString(x, y, number)

            print(f"==> drawString at x={x}, y={y}")

        else:
            print("→ SKIPPED")


        # Завершаем работу с текущим холстом
        can.save()
        packet.seek(0)

        overlay = PdfReader(packet)
        page.merge_page(overlay.pages[0])
        writer.add_page(page)

    with open(output_pdf_path, "wb") as f:
        writer.write(f)