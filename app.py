"""Procurement invoice → existing Excel workbook. Start: streamlit run app.py"""
from __future__ import annotations

from datetime import datetime
from hashlib import sha256
import hmac
from io import BytesIO
import os
from pathlib import Path

import pandas as pd
import streamlit as st

from excel_writer import (
    SHEET_MAPS, add_items, number, possible_duplicates, read_workbook,
    supported_sheets, workbook_sheet_preview,
)
from ocr import DEFAULT_MODEL, parse_invoice_pdf
from xlsx_patch import last_physical_row


st.set_page_config(page_title="Счёт → Закупочная ведомость", page_icon="📑", layout="wide")
st.markdown("""
<style>
.block-container {max-width: 1250px; padding-top: 2rem;}
[data-testid="stMetric"] {border: 1px solid #e2e8f0; border-radius: 12px; padding: 16px;}
[data-testid="stFileUploader"] {border-radius: 14px;}
.step-note {color:#64748b; font-size: .93rem;}
</style>
""", unsafe_allow_html=True)


def secret(name: str, default: str = "") -> str:
    try:
        return str(st.secrets.get(name, os.environ.get(name, default)))
    except Exception:
        return os.environ.get(name, default)


def pdf_quality(pdf_bytes: bytes) -> tuple[int, str]:
    """Return page count and proposed OCR engine without sending data anywhere."""
    import fitz
    try:
        with fitz.open(stream=pdf_bytes, filetype="pdf") as document:
            if len(document) == 0:
                raise ValueError("PDF не содержит страниц")
            if len(document) > 8:
                raise ValueError("Для первой версии поддерживаются счета до 8 страниц")
            sample_text = "".join(page.get_text("text") for page in document)
            return len(document), ("mistral-ocr" if len(sample_text.strip()) < 100 else "cloudflare-ai")
    except ValueError:
        raise
    except Exception as exc:
        raise ValueError("Не удалось открыть PDF. Возможно, файл повреждён или защищён паролем.") from exc


def session_reset_if_changed(signature: str) -> None:
    if st.session_state.get("file_signature") != signature:
        st.session_state["file_signature"] = signature
        st.session_state.pop("draft", None)
        for key in ("meta_supplier", "meta_invoice", "meta_date", "meta_purpose"):
            st.session_state.pop(key, None)


def install_draft(draft: dict) -> None:
    st.session_state["draft"] = draft
    st.session_state["meta_supplier"] = draft.get("supplier", "")
    st.session_state["meta_invoice"] = draft.get("invoice_number", "")
    st.session_state["meta_date"] = draft.get("invoice_date", "")
    st.session_state["meta_purpose"] = ""


st.title("📑 Счёт → Закупочная ведомость")
app_password = secret("APP_PASSWORD")
if app_password:
    entered = st.text_input("Пароль доступа к приложению", type="password", key="access_password")
    if not hmac.compare_digest(entered, app_password):
        st.info("Для работы с ведомостью введите пароль доступа.")
        st.stop()
else:
    st.caption("Приложение не защищено паролем. Для общедоступного размещения установите APP_PASSWORD в Secrets.")
st.write("Добавляйте позиции из PDF-счёта в существующий Excel-файл, сохраняя другие листы и их данные.")

c1, c2 = st.columns(2, gap="large")
with c1:
    st.subheader("1. Excel-ведомость")
    xlsx_file = st.file_uploader("Загрузите текущий файл .xlsx", type=["xlsx"], key="xlsx")
with c2:
    st.subheader("2. PDF-счёт")
    pdf_file = st.file_uploader("Загрузите счёт .pdf", type=["pdf"], key="pdf")

xlsx_bytes = xlsx_file.getvalue() if xlsx_file else None
pdf_bytes = pdf_file.getvalue() if pdf_file else None
signature = sha256((xlsx_bytes or b"") + (pdf_bytes or b"")).hexdigest()
session_reset_if_changed(signature)

if xlsx_file and xlsx_bytes:
    try:
        wb = read_workbook(xlsx_bytes)
        sheets = supported_sheets(wb)
    except Exception:
        st.error("Не удалось прочитать Excel. Убедитесь, что это корректный файл .xlsx.")
        st.stop()
    if not sheets:
        st.error("Файл не содержит поддерживаемых листов. Приложение настроено на конкретный шаблон ведомости.")
        st.stop()
    st.success(f"Excel загружен · {len(wb.sheetnames)} листов · настроено для записи: {len(sheets)}")
    st.subheader("3. Куда добавлять позиции")
    selected_sheet = st.selectbox("Выберите лист ведомости", sheets)
    mapping = SHEET_MAPS[selected_sheet]
    if mapping.caveat:
        st.info(mapping.caveat)
    with st.expander("Показать последние позиции выбранного листа"):
        st.dataframe(workbook_sheet_preview(wb, selected_sheet), hide_index=True, use_container_width=True)
        st.caption(f"Новые позиции будут добавлены после текущей последней строки ({last_physical_row(xlsx_bytes, selected_sheet)}), без замены заполненных ячеек.")
else:
    st.info("Начните с загрузки Excel-файла. Можно сначала попробовать без PDF: позиции разрешается ввести вручную.")
    st.stop()

st.subheader("4. Распознать и проверить позиции")
key = secret("OPENROUTER_API_KEY")
model = secret("OPENROUTER_MODEL", DEFAULT_MODEL)
mode_map = {"Автоматически": "auto", "Текстовый PDF (без дополнительного OCR)": "cloudflare-ai", "Скан / фото в PDF (OCR)": "mistral-ocr"}

left, right = st.columns([3, 2])
with left:
    pdf_mode = st.selectbox("Режим распознавания", list(mode_map))
with right:
    st.caption("Автоматический разбор — через OpenRouter. Для сканов доступен OCR; он может оплачиваться отдельно.")

cc1, cc2 = st.columns([1, 1])
with cc1:
    parse_clicked = st.button("✨ Распознать PDF", type="primary", use_container_width=True, disabled=(not bool(pdf_bytes) or not bool(key)))
with cc2:
    manual_clicked = st.button("✏️ Ввести позиции вручную", use_container_width=True)

if not key:
    st.warning("Для автоматического разбора добавьте **OPENROUTER_API_KEY** в Secrets вашего Streamlit-приложения. Ручное заполнение работает без API.")
if pdf_file and len(pdf_bytes) > 18 * 1024 * 1024:
    st.error("PDF слишком большой (максимум 18 МБ).")

if manual_clicked:
    install_draft({
        "supplier": "", "invoice_number": "", "invoice_date": "", "invoice_total": None,
        "items": [{"include": True, "name": "", "unit": "шт", "quantity": 1.0,
                   "unit_price": 0.0, "amount": 0.0}],
        "warnings": ["Ручной ввод: проверьте значения перед скачиванием."],
    })

if parse_clicked:
    try:
        count, recommended = pdf_quality(pdf_bytes)
        engine = recommended if mode_map[pdf_mode] == "auto" else mode_map[pdf_mode]
        with st.spinner(f"Анализ PDF ({count} стр., режим: {engine})…"):
            parsed = parse_invoice_pdf(pdf_bytes, pdf_file.name, key, model=model, engine=engine)
        if not parsed["items"]:
            st.warning("В счёте не удалось уверенно найти позиции. Проверьте файл или добавьте строки вручную.")
        else:
            install_draft(parsed)
            st.success(f"Найдено {len(parsed['items'])} позиций. Проверьте данные ниже.")
    except Exception as exc:
        st.error(str(exc))

if "draft" not in st.session_state:
    st.caption("После распознавания появится редактируемая таблица. Вы сможете удалить ненужные строки, добавить свои и изменить числа.")
    st.stop()

draft = st.session_state["draft"]
if draft.get("warnings"):
    for warning in draft["warnings"][:5]:
        st.warning(warning)

with st.container(border=True):
    mc1, mc2 = st.columns(2)
    with mc1:
        supplier = st.text_input("Поставщик", key="meta_supplier")
        invoice = st.text_input("Номер счёта", key="meta_invoice")
    with mc2:
        invoice_date = st.text_input("Дата счёта (ДД.ММ.ГГГГ)", key="meta_date")
        purpose = st.text_input("Назначение / объект (если применимо)", key="meta_purpose")

    status = ""
    if mapping.status_col:
        status = st.selectbox("Статус оплаты", ["не оплачено", "оплачено", "часть аванс"], index=0)

    rows = []
    for item in draft["items"]:
        row = dict(item)
        for col in ("quantity", "unit_price", "amount"):
            try:
                row[col] = number(row.get(col))
            except ValueError:
                row[col] = None
        rows.append(row)
    editable = pd.DataFrame(rows, columns=["include", "name", "unit", "quantity", "unit_price", "amount"])
    editable["include"] = editable["include"].fillna(True).astype(bool)
    st.write("**Позиции счёта — все значения можно исправить**")
    edited = st.data_editor(
        editable, num_rows="dynamic", hide_index=True, use_container_width=True,
        column_config={
            "include": st.column_config.CheckboxColumn("Добавить", help="Снимите галочку, чтобы пропустить строку"),
            "name": st.column_config.TextColumn("Наименование", width="large", required=True),
            "unit": st.column_config.TextColumn("Ед. изм.", width="small"),
            "quantity": st.column_config.NumberColumn("Количество", min_value=0.0, format="%.4f"),
            "unit_price": st.column_config.NumberColumn("Цена", min_value=0.0, format="%.2f"),
            "amount": st.column_config.NumberColumn("Стоимость", min_value=0.0, format="%.2f"),
        },
        key=f"editor_{signature}_{selected_sheet}_{sha256(str(draft).encode()).hexdigest()[:10]}",
    )

lines = edited.where(pd.notnull(edited), None).to_dict("records")
lines = [dict(row, include=(False if row.get("include") is False else bool(row.get("include", True)))) for row in lines]
selected = [row for row in lines if row.get("include")]

if selected:
    total = 0.0
    for row in selected:
        try:
            val = number(row.get("amount"))
            if val is None:
                q, p = number(row.get("quantity")), number(row.get("unit_price"))
                val = (q or 0) * (p or 0)
            total += val
        except ValueError:
            pass
    a, b, c = st.columns(3)
    a.metric("Выбрано позиций", len(selected))
    b.metric("Сумма выбранных", f"{total:,.2f} ₽".replace(",", " "))
    if draft.get("invoice_total") is not None:
        c.metric("Итого в счёте", f"{draft['invoice_total']} ₽")

    mismatches = []
    for row in selected:
        try:
            q, p, s = [number(row.get(v)) for v in ("quantity", "unit_price", "amount")]
            if q is not None and p is not None and s is not None and abs(q * p - s) > max(1.0, .02 * abs(s)):
                mismatches.append(str(row.get("name", ""))[:65])
        except ValueError:
            continue
    if mismatches:
        st.warning("Стоимость отличается от количества × цены в позициях: " + ", ".join(mismatches[:4]) + ". Проверьте скидки, НДС и округление.")

    try:
        duplicates = possible_duplicates(wb, selected_sheet, selected, supplier, invoice)
        if any(duplicates):
            st.warning("В листе уже есть похожие позиции. Проверьте повторную загрузку: " + ", ".join(
                str(line.get("name"))[:70] for line, dup in zip(selected, duplicates) if dup)[:400])
    except Exception:
        st.caption("Не удалось выполнить проверку возможных дублей; проверьте строки вручную.")

try:
    output_bytes, new_rows = add_items(
        xlsx_bytes, selected_sheet, lines,
        supplier=supplier, invoice=invoice, invoice_date=invoice_date,
        purpose=purpose, status=status,
        pdf_filename=pdf_file.name if pdf_file else "",
    )
    stem = Path(xlsx_file.name).stem
    output_name = f"{stem}_обновлено.xlsx"
    st.success(f"Готово к скачиванию: добавится {len(new_rows)} позиций на лист «{selected_sheet}», строки {new_rows[0]}–{new_rows[-1]}.")
    st.download_button(
        "⬇️ Скачать обновлённый Excel", output_bytes,
        file_name=output_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary", use_container_width=True,
    )
    st.caption("Исходный файл не меняется. Скачайте новую копию и при следующем счёте загрузите именно её.")
except Exception as exc:
    st.error(f"Пока нельзя сформировать файл: {exc}")

st.divider()
st.caption("Конфиденциальность: Excel обрабатывается внутри приложения. При автоматическом распознавании PDF передаётся в OpenRouter/провайдер модели. В этой версии нет базы данных и автоматической отправки файлов в GitHub.")
