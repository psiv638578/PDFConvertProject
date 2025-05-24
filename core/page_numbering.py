import win32com.client
import os


def number_docx_pages(file_paths, start_number):
    """
    Выполняет сквозную нумерацию страниц во всех .docx файлов,
    обновляя поля { SEQ MyPage } внутри фигур в колонтитулах.
    Требует: Word должен быть закрыт перед запуском.
    """
    if _is_word_running():
        raise RuntimeError("Перед запуском закройте все окна Microsoft Word.")

    word = win32com.client.Dispatch("Word.Application")
    word.Visible = True
    word.DisplayAlerts = False

    first_field = True

    for file_path in sorted(file_paths):
        print(f"Обработка: {file_path}")
        doc = word.Documents.Open(file_path)

        for section in doc.Sections:
            for header in [section.Headers(1), section.Headers(2)]:  # Primary и FirstPage
                for shape in header.Shapes:
                    if shape.TextFrame.HasText:
                        for field in shape.TextFrame.TextRange.Fields:
                            try:
                                if "SEQ MyPage" in field.Code.Text:
                                    if first_field:
                                        field.Code.Text = f"SEQ MyPage \\r {start_number}"
                                        first_field = False
                            except Exception:
                                pass
                        shape.TextFrame.TextRange.Fields.Update()

        doc.Save()
        doc.Close(False)

    word.Quit()
    print("Готово.")


def _is_word_running():
    """Проверяет, запущен ли Word (перед запуском нужно закрыть)."""
    try:
        win32com.client.GetActiveObject("Word.Application")
        return True
    except Exception:
        return False
