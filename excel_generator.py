"""
Генератор акта приёма-передачи в формате Excel.
Создаёт новый файл на основе шаблона, подставляя данные из накладной.
"""

import shutil
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter


def generate_act(data: dict, template_path: str, output_path: str) -> str:
    """
    Заполняет акт на основе данных накладной.

    Args:
        data: словарь с полями doc_number, doc_date, supplier, recipient, items
        template_path: путь к шаблону .xlsm
        output_path: путь для сохранения результата .xlsx

    Returns:
        output_path
    """
    # Копируем шаблон, чтобы не портить оригинал
    shutil.copy(template_path, output_path)

    wb = load_workbook(output_path, keep_vba=False)
    ws = wb["Акт"]

    doc_number = data.get("doc_number", "")
    doc_date   = data.get("doc_date", "")
    supplier   = data.get("supplier", "")
    recipient  = data.get("recipient", "")
    items      = data.get("items", [])

    # ── Заголовочные поля ──────────────────────────────────────────────────

    # Дата передачи (A2)
    ws["A2"] = _parse_date(doc_date)
    ws["A2"].number_format = "DD.MM.YYYY"

    # Поставщик (C11) — перезаписываем формулу значением
    ws["C11"] = supplier

    # Дата доставки (C13)
    ws["C13"] = _parse_date(doc_date)
    ws["C13"].number_format = "DD.MM.YYYY"

    # Номер документа записываем в C12 (id доставки) как текст
    ws["C12"] = doc_number

    # Получатель: вставляем в G7 (используется в формуле A8)
    ws["G7"] = recipient

    # ── Таблица позиций (строки с 16) ─────────────────────────────────────
    START_ROW = 16

    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left   = Alignment(horizontal="left",   vertical="center", wrap_text=True)

    for i, item in enumerate(items):
        row = START_ROW + i

        # Растягиваем таблицу если строк больше, чем в шаблоне
        _ensure_row(ws, row)

        name     = item.get("name", "")
        article  = item.get("article", "")
        quantity = item.get("quantity", 0)
        unit     = item.get("unit", "")

        # Формируем наименование с артикулом
        full_name = name
        if article:
            full_name = f"{name}\nАрт.: {article}"

        # A — №
        ws.cell(row=row, column=1).value = i + 1
        ws.cell(row=row, column=1).alignment = center
        ws.cell(row=row, column=1).border = border

        # B-D — Наименование (объединённые колонки B:D)
        ws.cell(row=row, column=2).value = full_name
        ws.cell(row=row, column=2).alignment = left
        ws.cell(row=row, column=2).border = border
        ws.row_dimensions[row].height = 30 if article else 18

        # E — Количество
        ws.cell(row=row, column=5).value = quantity
        ws.cell(row=row, column=5).alignment = center
        ws.cell(row=row, column=5).border = border

        # F — Ед. изм.
        ws.cell(row=row, column=6).value = unit
        ws.cell(row=row, column=6).alignment = center
        ws.cell(row=row, column=6).border = border

    # Очищаем оставшиеся формульные строки шаблона ниже заполненных
    last_template_row = START_ROW + 100
    filled_rows = len(items)
    for row in range(START_ROW + filled_rows, last_template_row):
        for col in range(1, 7):
            cell = ws.cell(row=row, column=col)
            if isinstance(cell.value, str) and cell.value.startswith("="):
                cell.value = None

    wb.save(output_path)
    return output_path


def _parse_date(date_str: str):
    """Пробует разобрать строку даты в datetime, иначе возвращает строку."""
    for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(date_str.strip(), fmt)
        except (ValueError, AttributeError):
            continue
    return date_str  # вернём как есть если не распознали


def _ensure_row(ws, row: int):
    """Снимает формулы в строке если они есть (готовит к записи значений)."""
    for col in range(1, 7):
        cell = ws.cell(row=row, column=col)
        if isinstance(cell.value, str) and cell.value.startswith("="):
            cell.value = None
