"""Excel templates and append-only writing for a specific procurement workbook.

Do not use these mappings for arbitrary Excel files.  They are intentionally
based on the six worksheets in the user-provided workbook from 17 Nov 2026.
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from io import BytesIO
import math
import re
from typing import Any

from openpyxl import load_workbook
from openpyxl.utils.datetime import to_excel
from xlsx_patch import append_rows_preserving_original, last_physical_row


@dataclass(frozen=True)
class SheetMap:
    name_col: str
    unit_col: str
    quantity_col: str
    price_col: str
    total_col: str
    supplier_col: str | None = None
    date_col: str | None = None
    invoice_col: str | None = None
    purpose_col: str | None = None
    status_col: str | None = None
    ordinal_col: str | None = None
    caveat: str = ""


SHEET_MAPS: dict[str, SheetMap] = {
    "материалы_наличн": SheetMap(
        name_col="D", unit_col="E", quantity_col="F", price_col="G",
        total_col="H", supplier_col="K", date_col="L", purpose_col="N",
        caveat="В этом листе нет отдельной колонки номера счёта. Можно внести номер вручную в существующие примечания Excel при необходимости.",
    ),
    "УК ПК": SheetMap(
        name_col="E", unit_col="F", quantity_col="G", price_col="H",
        total_col="I", supplier_col="C", invoice_col="D", purpose_col="J",
        status_col="B",
        caveat="Внизу листа есть предварительно заполненные строки со статусом и номером счёта. Инструмент не перезаписывает их, а добавляет новые строки после всех существующих.",
    ),
    "ЦДС СС": SheetMap(
        name_col="C", unit_col="D", quantity_col="E", price_col="F",
        total_col="G", supplier_col="B", status_col="A",
        caveat="На листе смешаны два формата: верхний блок содержит НДС, а нижний — статус, поставщика и позиции. Новые строки создаются по нижнему формату (A–G); поля НДС не заполняются автоматически.",
    ),
    "ИМПЕРИАЛ": SheetMap(
        name_col="B", unit_col="C", quantity_col="D", price_col="E",
        total_col="F", supplier_col="I", date_col="J", purpose_col="K",
    ),
    "ЗЕФИР": SheetMap(
        name_col="C", unit_col="D", quantity_col="E", price_col="F",
        total_col="G", supplier_col="H", purpose_col="I", ordinal_col="B",
    ),
    "услуги": SheetMap(
        name_col="A", unit_col="B", quantity_col="C", price_col="D",
        total_col="E", date_col="F", supplier_col="G", purpose_col="H",
        caveat="В существующих строках G/H используются неодинаково (поставщик, назначение, плательщик). Проверьте значения этих столбцов перед использованием.",
    ),
}


def normalize(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "").casefold()).strip()


def number(value: Any) -> float | None:
    """Parse invoice-style numeric fields, preserving None for empty values."""
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        raise ValueError("Некорректное числовое значение")
    if isinstance(value, (int, float)):
        result = float(value)
        return result if math.isfinite(result) else None
    cleaned = str(value).replace("\u00a0", "").replace(" ", "").replace("₽", "").replace("р.", "")
    # 1.234,56 / 1,234.56 / 1234,56
    if "," in cleaned and "." in cleaned:
        if cleaned.rfind(",") > cleaned.rfind("."):
            cleaned = cleaned.replace(".", "").replace(",", ".")
        else:
            cleaned = cleaned.replace(",", "")
    elif "," in cleaned:
        cleaned = cleaned.replace(",", ".")
    try:
        result = float(cleaned)
        return result if math.isfinite(result) else None
    except ValueError as e:
        raise ValueError(f"Невозможно прочесть число: {value!r}") from e


def safe_text(value: Any) -> str:
    """Excel formula injection guard for text originating from invoices."""
    text = str(value or "").strip().replace("\x00", "")
    if text.startswith(("=", "+", "-", "@")):
        return "'" + text
    return text


def parse_date(value: Any) -> datetime | None:
    if isinstance(value, datetime):
        return value
    if value is None or not str(value).strip():
        return None
    text = str(value).strip()
    for fmt in ("%d.%m.%Y", "%d.%m.%y", "%Y-%m-%d", "%d/%m/%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    raise ValueError(f"Неверный формат даты: {value!r}. Используйте ДД.ММ.ГГГГ")


def read_workbook(raw: bytes):
    if not raw:
        raise ValueError("Excel-файл пуст")
    return load_workbook(BytesIO(raw), data_only=False)


def supported_sheets(wb) -> list[str]:
    return [name for name in wb.sheetnames if name in SHEET_MAPS]


def last_template_row(ws, mapping: SheetMap) -> int:
    """Find a real existing line item, not the trailing empty prefilled metadata."""
    for row in range(ws.max_row, 1, -1):
        if (ws[f"{mapping.name_col}{row}"].value is not None
                and ws[f"{mapping.quantity_col}{row}"].value is not None):
            return row
    return max(2, ws.max_row)


def workbook_sheet_preview(wb, sheet_name: str, limit: int = 6) -> list[dict[str, Any]]:
    ws, mapping = wb[sheet_name], SHEET_MAPS[sheet_name]
    rows = []
    for i in range(ws.max_row, 1, -1):
        cell = ws[f"{mapping.name_col}{i}"].value
        if cell is None or not str(cell).strip():
            continue
        row = {"Строка": i, "Наименование": str(cell)[:120]}
        for label, col in (("Количество", mapping.quantity_col), ("Стоимость", mapping.total_col)):
            row[label] = ws[f"{col}{i}"].value
        rows.append(row)
        if len(rows) >= limit:
            break
    return list(reversed(rows))


def existing_signatures(wb, sheet_name: str) -> set[tuple]:
    ws, m = wb[sheet_name], SHEET_MAPS[sheet_name]
    found = set()
    for i in range(2, ws.max_row + 1):
        name = normalize(ws[f"{m.name_col}{i}"].value)
        if not name:
            continue
        supplier = normalize(ws[f"{m.supplier_col}{i}"].value) if m.supplier_col else ""
        invoice = normalize(ws[f"{m.invoice_col}{i}"].value) if m.invoice_col else ""
        try:
            quantity = number(ws[f"{m.quantity_col}{i}"].value) if ws[f"{m.quantity_col}{i}"].data_type != 'f' else None
            price = number(ws[f"{m.price_col}{i}"].value) if ws[f"{m.price_col}{i}"].data_type != 'f' else None
        except ValueError:
            continue  # headings, totals and text cells are not invoice positions
        if invoice:
            found.add((name, supplier, invoice))
        elif quantity is not None and price is not None:
            found.add((name, supplier, round(quantity, 4), round(price, 2)))
    return found


def possible_duplicates(wb, sheet_name: str, lines: list[dict], supplier: str, invoice: str) -> list[bool]:
    m = SHEET_MAPS[sheet_name]
    existing = existing_signatures(wb, sheet_name)
    out = []
    for line in lines:
        name = normalize(line.get("name"))
        if not name:
            out.append(False)
            continue
        if m.invoice_col and invoice:
            # Excel contains human-entered prefixes ('№', 'счет', dates). Match invoice number inside cell.
            ws = wb[sheet_name]
            found = False
            for i in range(2, ws.max_row + 1):
                if normalize(ws[f"{m.name_col}{i}"].value) != name:
                    continue
                existing_invoice = normalize(ws[f"{m.invoice_col}{i}"].value)
                if normalize(invoice) in existing_invoice and normalize(supplier) == normalize(ws[f"{m.supplier_col}{i}"].value):
                    found = True
                    break
            out.append(found)
        else:
            q, p = number(line.get("quantity")), number(line.get("unit_price"))
            out.append(q is not None and p is not None and
                       (name, normalize(supplier), round(q, 4), round(p, 2)) in existing)
    return out


def validate_lines(lines: list[dict]) -> list[dict]:
    cleaned = []
    for row_num, line in enumerate(lines, start=1):
        if not bool(line.get("include", True)):
            continue
        name = safe_text(line.get("name"))
        if not name:
            raise ValueError(f"Позиция {row_num}: пустое наименование")
        q = number(line.get("quantity"))
        p = number(line.get("unit_price"))
        total = number(line.get("amount"))
        if q is None or q <= 0:
            raise ValueError(f"Позиция {row_num}: количество должно быть больше нуля")
        if p is not None and p < 0:
            raise ValueError(f"Позиция {row_num}: цена не может быть отрицательной")
        if total is not None and total < 0:
            raise ValueError(f"Позиция {row_num}: сумма не может быть отрицательной")
        if total is None and p is None:
            raise ValueError(f"Позиция {row_num}: укажите цену или сумму")
        if total is None:
            total = round(q * p, 2)
        if p is None:
            p = round(total / q, 6)
        cleaned.append({"name": name, "unit": safe_text(line.get("unit")),
                        "quantity": q, "unit_price": p, "amount": total})
    if not cleaned:
        raise ValueError("Не выбрана ни одна позиция для добавления")
    return cleaned


def add_items(raw: bytes, sheet_name: str, lines: list[dict], *, supplier: str = "",
              invoice: str = "", invoice_date: str = "", purpose: str = "",
              status: str = "не оплачено", pdf_filename: str = "") -> tuple[bytes, list[int]]:
    """Append rows without rewriting original XML, formula caches or print settings."""
    if sheet_name not in SHEET_MAPS:
        raise ValueError("Для выбранного листа нет безопасной схемы заполнения")
    wb = read_workbook(raw)
    if sheet_name not in wb.sheetnames:
        raise ValueError("В загруженном файле отсутствует выбранный лист")
    parsed = validate_lines(lines)
    m, ws = SHEET_MAPS[sheet_name], wb[sheet_name]
    dt = parse_date(invoice_date)
    supplier, invoice, purpose = (safe_text(x) for x in (supplier, invoice, purpose))
    start = last_physical_row(raw, sheet_name) + 1
    template = last_template_row(ws, m)
    ordinal = 0
    if m.ordinal_col:
        ordinal = max((int(ws[f"{m.ordinal_col}{r}"].value)
                       for r in range(1, ws.max_row + 1)
                       if isinstance(ws[f"{m.ordinal_col}{r}"].value, (int, float))), default=0)
    rows = []
    for offset, line in enumerate(parsed):
        cells = {
            m.name_col: line["name"], m.unit_col: line["unit"],
            m.quantity_col: line["quantity"], m.price_col: line["unit_price"],
            m.total_col: line["amount"],
        }
        if m.supplier_col:
            cells[m.supplier_col] = supplier or None
        if m.purpose_col:
            cells[m.purpose_col] = purpose or None
        if m.status_col:
            cells[m.status_col] = status or None
        if m.date_col:
            if dt and ws[f"{m.date_col}{template}"].is_date:
                cells[m.date_col] = to_excel(dt, wb.epoch)
            else:
                cells[m.date_col] = dt.strftime("%d.%m.%Y") if dt else None
        if m.invoice_col:
            cells[m.invoice_col] = (f"{invoice} от {dt:%d.%m.%Y}" if dt and invoice else invoice) or None
        if m.ordinal_col:
            cells[m.ordinal_col] = ordinal + offset + 1
        rows.append(cells)
    updated = append_rows_preserving_original(
        raw, sheet_name, template, start, rows,
        extend_status_validation=(sheet_name == "УК ПК"),
    )
    return updated, list(range(start, start + len(rows)))
