"""Combine multiple independently checked invoices into one updated workbook."""
from __future__ import annotations

from typing import Any

from excel_writer import add_items, parse_date, validate_lines


def append_invoice_batch(xlsx_bytes: bytes, invoices: list[dict[str, Any]]) -> tuple[bytes, list[dict[str, Any]]]:
    """Append invoice rows in order, possibly to different workbook sheets.

    Each item in ``invoices`` has ``sheet``, ``lines`` and optional invoice
    metadata. No output is returned if any document is invalid. The supplied
    XLSX bytes are never modified.
    """
    if not invoices:
        raise ValueError("Нет выбранных позиций для добавления")
    if len(invoices) > 8:
        raise ValueError("За одну операцию можно обработать не более 8 счетов")

    # Validate every document before creating a modified copy of the workbook.
    for index, document in enumerate(invoices, 1):
        label = str(document.get("pdf_filename") or f"Счёт {index}")
        try:
            validate_lines(document["lines"])
            parse_date(document.get("invoice_date", ""))
        except (KeyError, TypeError, ValueError) as exc:
            raise ValueError(f"{label}: {exc}") from exc

    current = xlsx_bytes
    report: list[dict[str, Any]] = []
    for document in invoices:
        label = str(document.get("pdf_filename") or "Счёт")
        try:
            current, rows = add_items(
                current, document["sheet"], document["lines"],
                supplier=document.get("supplier", ""),
                invoice=document.get("invoice", ""),
                invoice_date=document.get("invoice_date", ""),
                purpose=document.get("purpose", ""),
                status=document.get("status", "не оплачено"),
                pdf_filename=label,
            )
        except (KeyError, TypeError, ValueError) as exc:
            raise ValueError(f"{label}: {exc}") from exc
        report.append({"pdf_filename": label, "sheet": document["sheet"], "rows": rows})
    return current, report
