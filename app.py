"""Procurement invoices → existing Excel workbook. Start: streamlit run app.py"""
from __future__ import annotations

from hashlib import sha256
import hmac
import os
from pathlib import Path

import pandas as pd
import streamlit as st

from batch_writer import append_invoice_batch
from excel_writer import (
    SHEET_MAPS, normalize, number, possible_duplicates, read_workbook,
    supported_sheets, workbook_sheet_preview,
)
from ocr import DEFAULT_MODEL, parse_invoice_pdf
from xlsx_patch import last_physical_row

MAX_INVOICES = 8
MAX_PDF_SIZE = 18 * 1024 * 1024

st.set_page_config(page_title="Счета → Закупочная ведомость", page_icon="📑", layout="wide")
st.markdown("""
<style>
.block-container {max-width: 1250px; padding-top: 2rem;}
[data-testid="stMetric"] {border: 1px solid #e2e8f0; border-radius: 12px; padding: 16px;}
[data-testid="stFileUploader"] {border-radius: 14px;}
</style>
""", unsafe_allow_html=True)


def secret(name: str, default: str = "") -> str:
    try:
        return str(st.secrets.get(name, os.environ.get(name, default)))
    except Exception:
        return os.environ.get(name, default)


def pdf_quality(pdf_bytes: bytes) -> tuple[int, str]:
    """Check PDF locally and recommend the parser; no API billing."""
    import fitz
    if len(pdf_bytes) > MAX_PDF_SIZE:
        raise ValueError("PDF слишком большой (максимум 18 МБ на счёт)")
    try:
        with fitz.open(stream=pdf_bytes, filetype="pdf") as document:
            if len(document) == 0:
                raise ValueError("PDF не содержит страниц")
            if len(document) > 8:
                raise ValueError("Поддерживается максимум 8 страниц в одном PDF-счёте")
            sample_text = "".join(page.get_text("text") for page in document)
            return len(document), ("mistral-ocr" if len(sample_text.strip()) < 100 else "cloudflare-ai")
    except ValueError:
        raise
    except Exception as exc:
        raise ValueError("Не удалось открыть PDF. Возможно, он повреждён или защищён паролем.") from exc


def file_fingerprint(filename: str, content: bytes, position: int) -> str:
    return f"pdf_{position}_{sha256(content).hexdigest()[:16]}"


def reset_if_uploads_changed(xlsx_bytes: bytes | None, pdf_files: list) -> None:
    h = sha256(xlsx_bytes or b"")
    for file in pdf_files:
        h.update(file.name.encode("utf-8"))
        h.update(file.getvalue())
    signature = h.hexdigest()
    if st.session_state.get("uploads_signature") != signature:
        st.session_state["uploads_signature"] = signature
        st.session_state["documents"] = []
        st.session_state["parse_errors"] = []
        st.session_state["editor_revision"] = st.session_state.get("editor_revision", 0) + 1


def add_document(doc_id: str, filename: str, parsed: dict) -> dict:
    """Attach per-invoice editable metadata before Streamlit widgets are created."""
    for field, value in (
        ("supplier", parsed.get("supplier", "")),
        ("invoice", parsed.get("invoice_number", "")),
        ("date", parsed.get("invoice_date", "")),
        ("purpose", ""),
    ):
        st.session_state[f"{field}_{doc_id}"] = value or ""
    return {"id": doc_id, "filename": filename, "parsed": parsed}


st.title("📑 Счета → Закупочная ведомость")
app_password = secret("APP_PASSWORD")
if app_password:
    entered = st.text_input("Пароль доступа к приложению", type="password", key="access_password")
    if not hmac.compare_digest(entered, app_password):
        st.info("Для работы с ведомостью введите пароль доступа.")
        st.stop()
else:
    st.warning("Приложение не защищено паролем. Установите APP_PASSWORD в Secrets перед предоставлением ссылки коллегам.")

st.write("Загрузите текущий Excel и до 8 PDF-счетов. Проверьте позиции, выберите лист для каждого счёта и скачайте один обновлённый Excel.")

c1, c2 = st.columns(2, gap="large")
with c1:
    st.subheader("1. Excel-ведомость")
    xlsx_file = st.file_uploader("Загрузите текущий файл .xlsx", type=["xlsx"], key="xlsx")
with c2:
    st.subheader("2. PDF-счета — до 8 файлов")
    pdf_files = st.file_uploader(
        "Загрузите один или несколько счетов .pdf", type=["pdf"],
        accept_multiple_files=True, key="pdfs",
    ) or []
    st.caption(f"Выбрано файлов: {len(pdf_files)} / {MAX_INVOICES}. Каждый PDF — до 8 страниц и 18 МБ.")

if len(pdf_files) > MAX_INVOICES:
    st.error("Выбрано более 8 PDF. Удалите лишние файлы из загрузки.")
    st.stop()

xlsx_bytes = xlsx_file.getvalue() if xlsx_file else None
reset_if_uploads_changed(xlsx_bytes, pdf_files)

if not xlsx_bytes:
    st.info("Сначала загрузите Excel-файл. Если PDF нет, можно добавить позиции вручную.")
    st.stop()
try:
    wb = read_workbook(xlsx_bytes)
    sheets = supported_sheets(wb)
except Exception:
    st.error("Не удалось прочитать Excel. Проверьте формат .xlsx.")
    st.stop()
if not sheets:
    st.error("В файле нет поддерживаемых листов. Приложение настроено на определённый шаблон ведомости.")
    st.stop()
st.success(f"Excel загружен · всего листов: {len(wb.sheetnames)} · поддерживаются: {len(sheets)}")
with st.expander("Посмотреть последние записи и место добавления"):
    preview_sheet = st.selectbox("Лист для просмотра", sheets, key="preview_sheet")
    st.dataframe(workbook_sheet_preview(wb, preview_sheet), hide_index=True, use_container_width=True)
    st.caption(f"Добавление начинается после последней строки: {last_physical_row(xlsx_bytes, preview_sheet)}.")

st.subheader("3. Распознавание")
key = secret("OPENROUTER_API_KEY")
model = secret("OPENROUTER_MODEL", DEFAULT_MODEL)
mode_map = {"Автоматически": "auto", "Текстовый PDF": "cloudflare-ai", "Скан / фото в PDF": "mistral-ocr"}
left, right = st.columns([2, 3])
with left:
    pdf_mode = st.selectbox("Режим обработки PDF", list(mode_map))
with right:
    st.caption("Все выбранные PDF распознаются последовательно через OpenRouter. Каждый документ может расходовать баланс API. Повторное нажатие запускает распознавание повторно.")

button1, button2 = st.columns(2)
with button1:
    parse_clicked = st.button(
        "✨ Распознать выбранные PDF", type="primary", use_container_width=True,
        disabled=(not pdf_files or not key),
    )
with button2:
    manual_clicked = st.button("✏️ Добавить счёт вручную", use_container_width=True)
if not key:
    st.warning("Для распознавания добавьте OPENROUTER_API_KEY в Streamlit Secrets. Ручной ввод доступен без API.")

if manual_clicked:
    documents = st.session_state.get("documents", [])
    if len(documents) >= MAX_INVOICES:
        st.error("Можно подготовить не более 8 счетов за одну операцию.")
    else:
        next_id = f"manual_{st.session_state.get('manual_serial', 0) + 1}"
        st.session_state["manual_serial"] = st.session_state.get("manual_serial", 0) + 1
        documents.append(add_document(next_id, "Ручной ввод", {
            "supplier": "", "invoice_number": "", "invoice_date": "", "invoice_total": None,
            "items": [{"include": True, "name": "", "unit": "шт", "quantity": 1.0,
                       "unit_price": 0.0, "amount": 0.0}],
            "warnings": ["Ручной ввод: проверьте все значения."],
        }))
        st.session_state["documents"] = documents

if parse_clicked:
    # Validate the whole batch locally before any paid API requests.
    checks = []
    validation_errors = []
    seen_hashes = set()
    for index, file in enumerate(pdf_files):
        pdf_bytes = file.getvalue()
        digest = sha256(pdf_bytes).hexdigest()
        if digest in seen_hashes:
            validation_errors.append(f"{file.name}: одинаковый PDF загружен дважды")
            continue
        seen_hashes.add(digest)
        try:
            pages, engine = pdf_quality(pdf_bytes)
            checks.append((index, file, pdf_bytes, pages, engine))
        except ValueError as exc:
            validation_errors.append(f"{file.name}: {exc}")
    if validation_errors:
        st.session_state["parse_errors"] = validation_errors
        st.error("Проверьте выбранные PDF. API-запросы не выполнялись.")
    else:
        st.session_state["parse_errors"] = []
        new_docs = []
        progress = st.progress(0, text="Начало распознавания…")
        for completed, (index, file, pdf_bytes, pages, recommended) in enumerate(checks, 1):
            engine = recommended if mode_map[pdf_mode] == "auto" else mode_map[pdf_mode]
            try:
                with st.spinner(f"Счёт {completed}/{len(checks)}: {file.name} ({pages} стр.)…"):
                    parsed = parse_invoice_pdf(pdf_bytes, file.name, key, model=model, engine=engine)
                new_docs.append(add_document(file_fingerprint(file.name, pdf_bytes, index), file.name, parsed))
                if not parsed.get("items"):
                    st.session_state["parse_errors"].append(
                        f"{file.name}: позиции не найдены; при необходимости добавьте их вручную в таблицу.")
            except Exception as exc:
                st.session_state["parse_errors"].append(f"{file.name}: {exc}")
            progress.progress(completed / len(checks), text=f"Обработано: {completed}/{len(checks)}")
        st.session_state["documents"] = new_docs
        st.session_state["editor_revision"] = st.session_state.get("editor_revision", 0) + 1
        if new_docs:
            st.success(f"Распознано счетов: {len(new_docs)} из {len(checks)}. Проверьте данные каждого.")

for err in st.session_state.get("parse_errors", []):
    st.warning(err)

documents = st.session_state.get("documents", [])
if not documents:
    st.info("После распознавания появятся отдельные редактируемые таблицы для каждого счёта.")
    st.stop()

st.subheader("4. Проверка позиций и выбор листов")
st.caption("Для каждого счёта выберите лист отдельно. Можно исправить значения и исключить ненужные строки.")
prepared = []
batch_seen = set()
for index, doc in enumerate(documents, 1):
    doc_id, parsed = doc["id"], doc["parsed"]
    label = f"Счёт {index} из {len(documents)} · {doc['filename']} · {len(parsed.get('items', []))} позиций"
    with st.expander(label, expanded=(len(documents) == 1 or index == 1)):
        sh = st.selectbox("Записать на лист", sheets, key=f"sheet_{doc_id}")
        mapping = SHEET_MAPS[sh]
        if mapping.caveat:
            st.info(mapping.caveat)
        for warn in parsed.get("warnings", [])[:5]:
            st.warning(str(warn))
        m1, m2 = st.columns(2)
        with m1:
            supplier = st.text_input("Поставщик", key=f"supplier_{doc_id}")
            invoice = st.text_input("Номер счёта", key=f"invoice_{doc_id}")
        with m2:
            invoice_date = st.text_input("Дата счёта (ДД.ММ.ГГГГ)", key=f"date_{doc_id}")
            purpose = st.text_input("Назначение / объект (при необходимости)", key=f"purpose_{doc_id}")
        status = ""
        if mapping.status_col:
            status = st.selectbox("Статус оплаты", ["не оплачено", "оплачено", "часть аванс"], key=f"status_{doc_id}")

        rows = []
        for item in parsed.get("items", []):
            row = dict(item)
            for col in ("quantity", "unit_price", "amount"):
                try:
                    row[col] = number(row.get(col))
                except ValueError:
                    row[col] = None
            rows.append(row)
        editable = pd.DataFrame(rows, columns=["include", "name", "unit", "quantity", "unit_price", "amount"])
        if editable.empty:
            editable = pd.DataFrame([{"include": False, "name": "", "unit": "шт", "quantity": None,
                                      "unit_price": None, "amount": None}])
        editable["include"] = editable["include"].fillna(False).astype(bool)
        edited = st.data_editor(
            editable, num_rows="dynamic", hide_index=True, use_container_width=True,
            column_config={
                "include": st.column_config.CheckboxColumn("Добавить", help="Снимите галочку, чтобы исключить позицию"),
                "name": st.column_config.TextColumn("Наименование", width="large"),
                "unit": st.column_config.TextColumn("Ед. изм.", width="small"),
                "quantity": st.column_config.NumberColumn("Количество", min_value=0.0, format="%.4f"),
                "unit_price": st.column_config.NumberColumn("Цена", min_value=0.0, format="%.2f"),
                "amount": st.column_config.NumberColumn("Стоимость", min_value=0.0, format="%.2f"),
            },
            key=f"editor_{doc_id}_{st.session_state.get('editor_revision', 0)}",
        )
        lines = edited.where(pd.notnull(edited), None).to_dict("records")
        chosen = [dict(row) for row in lines if row.get("include") is True or row.get("include") == True]
        if not chosen:
            st.caption("Нет выбранных строк — этот счёт не попадёт в Excel.")
            continue
        total = 0.0
        bad_math = []
        for row in chosen:
            q, p, amount = (number(row.get(col)) for col in ("quantity", "unit_price", "amount"))
            total += (amount if amount is not None else (q or 0) * (p or 0))
            if all(v is not None for v in (q, p, amount)) and abs(q * p - amount) > max(1.0, 0.02 * abs(amount)):
                bad_math.append(str(row.get("name") or "")[:60])
            duplicate_key = (sh, normalize(invoice), normalize(supplier), normalize(row.get("name")), q, p)
            if duplicate_key in batch_seen:
                st.warning(f"Похожая позиция встречается в нескольких выбранных счетах: {row.get('name', '')}")
            batch_seen.add(duplicate_key)
        a, b = st.columns(2)
        a.metric("Позиций к добавлению", len(chosen))
        b.metric("Сумма выбранных", f"{total:,.2f} ₽".replace(",", " "))
        if bad_math:
            st.warning("Проверьте количество × цену: " + ", ".join(bad_math[:4]))
        try:
            duplicates = possible_duplicates(wb, sh, chosen, supplier, invoice)
            if any(duplicates):
                st.warning("Похожая позиция уже есть в Excel: " + ", ".join(
                    str(row.get("name"))[:65] for row, flag in zip(chosen, duplicates) if flag)[:300])
        except Exception:
            st.caption("Проверка дублей в Excel недоступна; проверьте позиции вручную.")
        prepared.append({"sheet": sh, "lines": chosen, "supplier": supplier, "invoice": invoice,
                         "invoice_date": invoice_date, "purpose": purpose, "status": status,
                         "pdf_filename": doc["filename"]})

st.subheader("5. Скачать результат")
if not prepared:
    st.info("Отметьте хотя бы одну позицию в счёте, чтобы сформировать Excel.")
    st.stop()
try:
    updated, report = append_invoice_batch(xlsx_bytes, prepared)
    lines_count = sum(len(entry["rows"]) for entry in report)
    st.success(f"Готово: {len(report)} счетов · {lines_count} позиций · {len(set(entry['sheet'] for entry in report))} листов. Один Excel-файл.")
    for entry in report:
        st.caption(f"{entry['pdf_filename']} → {entry['sheet']}: строки {entry['rows'][0]}–{entry['rows'][-1]}")
    st.download_button(
        "⬇️ Скачать обновлённый Excel", data=updated,
        file_name=f"{Path(xlsx_file.name).stem}_обновлено.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary", use_container_width=True,
    )
    st.caption("Исходный Excel не изменён. Для следующей партии загрузите скачанную обновлённую копию.")
except Exception as exc:
    st.error(f"Не удалось сформировать Excel: {exc}")

st.divider()
st.caption("Конфиденциальность: Excel обрабатывается в памяти приложения. Для AI-распознавания PDF передаются через OpenRouter. Файлы не отправляются в GitHub.")
