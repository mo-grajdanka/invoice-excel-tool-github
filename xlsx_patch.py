"""Surgical XLSX append that preserves all original archive members and formula caches.

Instead of reserializing the entire workbook, only the target sheet's XML is
modified. This keeps printer settings, cached formula values, existing comments,
shared strings, and other ancillary Excel components unchanged.
"""
from __future__ import annotations

from io import BytesIO
import re
from posixpath import normpath
from xml.sax.saxutils import escape
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from openpyxl.utils.datetime import to_excel

_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def target_sheet_xml(archive: ZipFile, name: str) -> str:
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    sheets = workbook.find(f"{{{_NS}}}sheets")
    sheet = next((s for s in sheets if s.get("name") == name), None)
    if sheet is None:
        raise ValueError("Лист не найден в XLSX")
    relationship_id = sheet.get(f"{{{_REL}}}id")
    relationships = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    relationship = next((r for r in relationships if r.get("Id") == relationship_id), None)
    if relationship is None:
        raise ValueError("Не найдена связь с листом Excel")
    target = relationship.get("Target", "")
    if target.startswith("/"):
        path = normpath(target.lstrip("/"))
    else:
        path = normpath("xl/" + target)
    if path not in archive.namelist() or not path.startswith("xl/worksheets/"):
        raise ValueError("Некорректный путь листа в XLSX")
    return path


def last_physical_row(raw: bytes, sheet_name: str) -> int:
    """Unlike openpyxl.max_row, also count explicitly formatted empty <row/> tags."""
    with ZipFile(BytesIO(raw), 'r') as archive:
        xml = archive.read(target_sheet_xml(archive, sheet_name)).decode('utf-8')
    match = re.search(r'<sheetData>(.*?)</sheetData>', xml, re.DOTALL)
    row_nums = [int(v) for v in re.findall(r'<row\b[^>]*\br="(\d+)"', match.group(1))] if match else []
    dimension = re.search(r'<dimension\s+ref="([^"]+)"', xml)
    if dimension:
        row_nums.extend(int(x) for x in re.findall(r'\d+', dimension.group(1)))
    return max(row_nums, default=1)


def extract_template(xml: str, template_row: int) -> tuple[str, dict[str, str]]:
    match = re.search(
        rf'<row\b(?P<attrs>[^>]*\br="{template_row}"[^>]*)>(?P<body>.*?)</row>',
        xml, re.DOTALL,
    )
    if match is None:
        raise ValueError(f"Нет исходной строки {template_row} для копирования оформления")
    style_map = {}
    for match_cell in re.finditer(r'<c\b[^>]*(?:/>|>.*?</c>)', match.group("body"), re.DOTALL):
        open_tag = match_cell.group(0).split(">", 1)[0]
        address = re.search(r'\br="([A-Z]+)\d+"', open_tag)
        style = re.search(r'\bs="(\d+)"', open_tag)
        if address and style:
            style_map[address.group(1)] = style.group(1)
    row_attrs = re.sub(r'\br="\d+"', 'r="{ROW}"', match.group("attrs"), count=1)
    return row_attrs, style_map


def xml_text(value: str) -> str:
    # XML 1.0 rejects these control characters even if properly escaped.
    return escape(re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', value))


def num_xml(value: int | float) -> str:
    return format(float(value), ".15g")


def cell_xml(col: str, row: int, value, styles: dict[str, str]) -> str:
    s = f' s="{styles[col]}"' if col in styles else ""
    if value is None:
        return f'<c r="{col}{row}"{s}/>'
    if isinstance(value, (int, float)):
        return f'<c r="{col}{row}"{s}><v>{num_xml(value)}</v></c>'
    return f'<c r="{col}{row}"{s} t="inlineStr"><is><t xml:space="preserve">{xml_text(str(value))}</t></is></c>'


def _column_sort(name: str) -> int:
    n = 0
    for ch in name:
        n = n * 26 + ord(ch) - ord('A') + 1
    return n


def _build_row(row: int, data: dict[str, object], row_attrs: str,
               style_map: dict[str, str]) -> str:
    all_cols = sorted(set(style_map) | set(data), key=_column_sort)
    cells = "".join(cell_xml(col, row, data.get(col), style_map) for col in all_cols)
    return f"<row{row_attrs.replace('{ROW}', str(row))}>{cells}</row>"


def _update_dimension(xml: str, last_row: int) -> str:
    def replace(match):
        address = match.group(2)
        if ':' in address:
            left, right = address.split(':', 1)
            new_right = re.sub(r'\d+$', str(last_row), right)
            return match.group(1) + left + ':' + new_right + match.group(3)
        return match.group(1) + address + ':' + re.sub(r'\d+$', str(last_row), address) + match.group(3)
    return re.sub(r'(<dimension\s+ref=")([A-Z]+\d+(?::[A-Z]+\d+)?)(")', replace, xml, count=1)


def _update_autofilter(xml: str, last_row: int) -> str:
    # Include all newly appended rows if the sheet already has a filter.
    def replace(match):
        address = match.group(2)
        if ':' not in address:
            return match.group(0)
        left, right = address.split(':', 1)
        return match.group(1) + left + ':' + re.sub(r'\d+$', str(last_row), right) + match.group(3)
    return re.sub(r'(<autoFilter\b[^>]*\bref=")([A-Z]+\d+:[A-Z]+\d+)(")', replace, xml, count=1)


def _extend_payment_validation(xml: str, start: int, end: int) -> str:
    # Preserve existing complex validation expressions; simply add a new range.
    def replace(match):
        old = match.group(2)
        if 'B314:B348' in old or 'B' in old:
            return match.group(1) + old + f' B{start}:B{end}' + match.group(3)
        return match.group(0)
    return re.sub(r'(<dataValidation\b[^>]*\bsqref=")([^"]+)(")', replace, xml, count=1)


def append_rows_preserving_original(raw: bytes, sheet_name: str, template_row: int,
                                    start: int, rows: list[dict[str, object]],
                                    *, extend_status_validation: bool = False) -> bytes:
    if not rows:
        raise ValueError("Нет строк для вставки")
    inbuf, outbuf = BytesIO(raw), BytesIO()
    with ZipFile(inbuf, 'r') as original:
        target_path = target_sheet_xml(original, sheet_name)
        target_xml = original.read(target_path).decode('utf-8')
        row_attrs, style_map = extract_template(target_xml, template_row)
        appended = "".join(_build_row(start + i, values, row_attrs, style_map)
                           for i, values in enumerate(rows))
        if '</sheetData>' not in target_xml:
            raise ValueError("Некорректный XML исходного листа")
        target_xml = target_xml.replace('</sheetData>', appended + '</sheetData>', 1)
        final_row = start + len(rows) - 1
        target_xml = _update_dimension(target_xml, final_row)
        target_xml = _update_autofilter(target_xml, final_row)
        if extend_status_validation:
            target_xml = _extend_payment_validation(target_xml, start, final_row)
        # Validate the only part being changed before returning the updated ZIP.
        ET.fromstring(target_xml)
        with ZipFile(outbuf, 'w') as updated:
            for part in original.infolist():
                data = (target_xml.encode('utf-8') if part.filename == target_path
                        else original.read(part.filename))
                updated.writestr(part, data)
    return outbuf.getvalue()
