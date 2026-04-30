from __future__ import annotations

import re
import xml.etree.ElementTree as ET
from pathlib import Path

from .models import CellData, RowData, SS, SheetData, WorkbookData, XML_NS


NS = {"ss": XML_NS}
FIELD_TOKEN_RE = re.compile(r"^[a-zA-Z0-9_]+$")
FORMULA_ATTRS = {f"{SS}Formula", f"{SS}ArrayRange"}


def _extract_cell_text(data_element: ET.Element | None) -> str:
    if data_element is None:
        return ""
    return "".join(data_element.itertext()).strip()


def parse_workbook(path: str | Path) -> WorkbookData:
    workbook_path = Path(path)
    return parse_workbook_bytes(
        workbook_path.read_bytes(),
        path=workbook_path,
        source_kind="local",
        source_label=workbook_path.name,
        snapshot_id=f"local:{workbook_path.resolve()}:{workbook_path.stat().st_mtime_ns}",
    )


def parse_workbook_bytes(
    xml_bytes: bytes,
    *,
    path: Path | None,
    source_kind: str,
    source_label: str,
    snapshot_id: str,
    source_meta: dict[str, str] | None = None,
) -> WorkbookData:
    root = ET.fromstring(xml_bytes)
    tree = ET.ElementTree(root)
    worksheet_nodes = {
        worksheet.attrib.get(f"{SS}Name", f"Sheet{index}"): worksheet
        for index, worksheet in enumerate(root.findall("ss:Worksheet", NS), start=1)
    }
    workbook = WorkbookData(
        path=path,
        tree=tree,
        sheet_names=list(worksheet_nodes.keys()),
        xml_bytes=xml_bytes,
        source_kind=source_kind,
        source_label=source_label,
        snapshot_id=snapshot_id,
        source_meta=source_meta or {},
    )
    workbook._sheet_loader = lambda sheet_name: _parse_sheet(worksheet_nodes.get(sheet_name), sheet_name)
    return workbook


def _parse_sheet(worksheet: ET.Element | None, sheet_name: str) -> SheetData | None:
    if worksheet is None:
        return None
    table = worksheet.find("ss:Table", NS)
    rows = _parse_rows(table) if table is not None else []
    _trim_trailing_empty_cells(rows)
    max_columns = _effective_max_columns(rows)
    sheet = SheetData(name=sheet_name, rows=rows, max_columns=max_columns)
    _infer_headers(sheet)
    _classify_rows(sheet)
    return sheet


def _parse_rows(table: ET.Element | None) -> list[RowData]:
    if table is None:
        return []

    parsed_rows: list[RowData] = []
    current_row_index = 1
    for row_element in table.findall("ss:Row", NS):
        explicit_index = row_element.attrib.get(f"{SS}Index")
        target_row_index = int(explicit_index) if explicit_index else current_row_index

        while current_row_index < target_row_index:
            parsed_rows.append(RowData(row_index=current_row_index, kind="blank"))
            current_row_index += 1

        row = RowData(
            row_index=target_row_index,
            attrs={key: value for key, value in row_element.attrib.items() if key != f"{SS}Index"},
        )

        current_column = 1
        for cell_element in row_element.findall("ss:Cell", NS):
            explicit_column = cell_element.attrib.get(f"{SS}Index")
            if explicit_column:
                current_column = int(explicit_column)

            data_element = cell_element.find("ss:Data", NS)
            value = _extract_cell_text(data_element)
            data_type = data_element.attrib.get(f"{SS}Type") if data_element is not None else None
            attrs = {key: value for key, value in cell_element.attrib.items() if key != f"{SS}Index"}
            row.cells.append(
                CellData(
                    column_index=current_column,
                    value=value,
                    data_type=data_type,
                    attrs=attrs,
                )
            )
            current_column += 1

        parsed_rows.append(row)
        current_row_index = target_row_index + 1

    return parsed_rows


def _trim_trailing_empty_cells(rows: list[RowData]) -> None:
    max_columns = _effective_max_columns(rows)
    for row in rows:
        row.cells = [
            cell
            for cell in row.cells
            if cell.column_index <= max_columns or _cell_has_meaningful_content(cell)
        ]


def _effective_max_columns(rows: list[RowData]) -> int:
    return max(
        (
            cell.column_index
            for row in rows
            for cell in row.cells
            if _cell_has_meaningful_content(cell)
        ),
        default=0,
    )


def _cell_has_meaningful_content(cell: CellData) -> bool:
    return bool(str(cell.value or "").strip()) or any(name in cell.attrs for name in FORMULA_ATTRS)


def _infer_headers(sheet: SheetData) -> None:
    candidates = [row for row in sheet.rows[: min(12, len(sheet.rows))] if row.non_empty_count]
    if not candidates:
        sheet.logical_headers = [f"col_{index}" for index in range(1, sheet.max_columns + 1)]
        sheet.display_headers = list(sheet.logical_headers)
        return

    field_row = None
    best_score = -1.0
    for row in candidates:
        values = [cell.value.strip() for cell in row.cells if cell.value.strip()]
        if not values:
            continue
        token_count = sum(1 for value in values if FIELD_TOKEN_RE.fullmatch(value))
        score = token_count / len(values)
        if score > best_score:
            best_score = score
            field_row = row

    if field_row is None:
        field_row = candidates[0]

    display_row = field_row
    for row in reversed(candidates):
        if row.row_index < field_row.row_index and row.non_empty_count >= max(2, field_row.non_empty_count // 2):
            display_row = row
            break

    sheet.field_header_row = field_row.row_index
    sheet.display_header_row = display_row.row_index
    sheet.logical_headers = [_header_value(field_row, index) for index in range(1, sheet.max_columns + 1)]
    sheet.display_headers = [_header_value(display_row, index) for index in range(1, sheet.max_columns + 1)]


def _header_value(row: RowData, index: int) -> str:
    value = row.value_at(index).strip()
    return value or f"col_{index}"


def _classify_rows(sheet: SheetData) -> None:
    field_row = sheet.field_header_row or 0
    display_row = sheet.display_header_row or 0
    key_like_indexes = [
        index
        for index, header in enumerate(sheet.logical_headers, start=1)
        if "id" in header.lower() or "key" in header.lower() or header.lower() == "name"
    ]

    for row in sheet.rows:
        if row.non_empty_count == 0:
            row.kind = "blank"
            continue
        if row.row_index in {field_row, display_row}:
            row.kind = "header"
            continue
        if row.non_empty_count == 1:
            row.kind = "group"
            continue
        if any(row.value_at(index).strip() for index in key_like_indexes):
            row.kind = "data"
            continue

        filled_ratio = row.non_empty_count / max(sheet.max_columns, 1)
        row.kind = "data" if filled_ratio >= 0.25 else "note"


def clone_row_with_values(base_row: RowData | None, values_by_column: dict[int, str]) -> RowData:
    row = base_row.clone() if base_row is not None else RowData(row_index=0)
    cell_map = {cell.column_index: cell.clone() for cell in row.cells}
    for column_index, value in values_by_column.items():
        cell = cell_map.get(column_index)
        if cell is None:
            cell = CellData(column_index=column_index, value=value)
            cell_map[column_index] = cell
        else:
            cell.value = value
            if cell.data_type is None:
                cell.data_type = "String"
    row.cells = [cell_map[index] for index in sorted(cell_map)]
    return row


def row_to_xml(row: RowData) -> ET.Element:
    row_element = ET.Element(f"{SS}Row")
    for key, value in row.attrs.items():
        row_element.set(key, value)

    previous_column = 0
    for cell in sorted(row.cells, key=lambda item: item.column_index):
        cell_element = ET.SubElement(row_element, f"{SS}Cell")
        if cell.column_index != previous_column + 1:
            cell_element.set(f"{SS}Index", str(cell.column_index))
        for key, value in cell.attrs.items():
            cell_element.set(key, value)

        data_element = ET.SubElement(cell_element, f"{SS}Data")
        data_element.set(f"{SS}Type", cell.data_type or _guess_type(cell.value))
        data_element.text = cell.value
        previous_column = cell.column_index

    return row_element


def _guess_type(value: str) -> str:
    if not value:
        return "String"
    if value.isdigit():
        return "Number"
    try:
        float(value)
    except ValueError:
        return "String"
    return "Number"
