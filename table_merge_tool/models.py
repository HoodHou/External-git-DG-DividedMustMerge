from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable


XML_NS = "urn:schemas-microsoft-com:office:spreadsheet"
SS = f"{{{XML_NS}}}"


@dataclass(slots=True)
class CellData:
    column_index: int
    value: str = ""
    data_type: str | None = None
    attrs: dict[str, str] = field(default_factory=dict)

    def clone(self) -> "CellData":
        return CellData(
            column_index=self.column_index,
            value=self.value,
            data_type=self.data_type,
            attrs=dict(self.attrs),
        )


@dataclass(slots=True)
class RowData:
    row_index: int
    cells: list[CellData] = field(default_factory=list)
    attrs: dict[str, str] = field(default_factory=dict)
    kind: str = "unknown"
    _cell_index_key: tuple[int, int] | None = field(default=None, repr=False)
    _cell_index: dict[int, CellData] = field(default_factory=dict, repr=False)

    def clone(self) -> "RowData":
        return RowData(
            row_index=self.row_index,
            cells=[cell.clone() for cell in self.cells],
            attrs=dict(self.attrs),
            kind=self.kind,
        )

    @property
    def non_empty_count(self) -> int:
        return sum(1 for cell in self.cells if cell.value.strip())

    def _ensure_cell_index(self) -> dict[int, CellData]:
        cells = self.cells
        key = (id(cells), len(cells))
        if self._cell_index_key != key:
            self._cell_index = {cell.column_index: cell for cell in cells}
            self._cell_index_key = key
        return self._cell_index

    def value_at(self, column_index: int) -> str:
        cell = self._ensure_cell_index().get(column_index)
        return cell.value if cell is not None else ""

    def cell_at(self, column_index: int) -> CellData | None:
        return self._ensure_cell_index().get(column_index)


@dataclass(slots=True)
class SheetData:
    name: str
    rows: list[RowData]
    max_columns: int
    display_header_row: int | None = None
    field_header_row: int | None = None
    logical_headers: list[str] = field(default_factory=list)
    display_headers: list[str] = field(default_factory=list)


@dataclass(slots=True)
class WorkbookData:
    path: Path | None
    tree: Any
    sheet_names: list[str]
    xml_bytes: bytes
    source_kind: str = "local"
    source_label: str = ""
    snapshot_id: str = ""
    source_meta: dict[str, str] = field(default_factory=dict)
    _sheet_cache: dict[str, SheetData] = field(default_factory=dict, repr=False)
    _sheet_loader: Callable[[str], SheetData | None] | None = field(default=None, repr=False)

    def get_sheet(self, sheet_name: str) -> SheetData | None:
        cached = self._sheet_cache.get(sheet_name)
        if cached is not None:
            return cached
        if self._sheet_loader is None:
            return None
        sheet = self._sheet_loader(sheet_name)
        if sheet is not None:
            self._sheet_cache[sheet_name] = sheet
        return sheet

    @property
    def sheets(self) -> list[SheetData]:
        return [sheet for name in self.sheet_names if (sheet := self.get_sheet(name)) is not None]

    @property
    def sheet_map(self) -> dict[str, SheetData]:
        return {name: sheet for name in self.sheet_names if (sheet := self.get_sheet(name)) is not None}


@dataclass(slots=True)
class ColumnBinding:
    key: str
    title: str
    left_index: int | None = None
    right_index: int | None = None
    base_index: int | None = None


CONFLICT_KIND_TWO_WAY = "two_way"
CONFLICT_KIND_THREE_WAY_DIVERGE = "three_way_diverge"
CONFLICT_KIND_THREE_WAY_LEFT_MODIFIED = "three_way_left_modified"
CONFLICT_KIND_THREE_WAY_RIGHT_MODIFIED = "three_way_right_modified"
CONFLICT_KIND_THREE_WAY_SAME_EDIT = "three_way_same_edit"


@dataclass(slots=True)
class AlignedRow:
    left_row: RowData | None
    right_row: RowData | None
    merged_row: RowData
    status: str
    reason: str
    score: float = 1.0
    conflict_columns: set[int] = field(default_factory=set)
    note: str = ""
    base_row: RowData | None = None
    conflict_kind: str = CONFLICT_KIND_TWO_WAY
    column_conflict_kinds: dict[int, str] = field(default_factory=dict)


@dataclass(slots=True)
class SheetAlignment:
    sheet_name: str
    columns: list[ColumnBinding]
    rows: list[AlignedRow]
    left_sheet: SheetData | None = None
    right_sheet: SheetData | None = None
    key_fields: list[str] = field(default_factory=list)
    base_sheet: SheetData | None = None

    @property
    def conflict_count(self) -> int:
        return sum(1 for row in self.rows if row.status == "conflict")

    @property
    def unresolved_count(self) -> int:
        return sum(1 for row in self.rows if row.status in {"conflict", "left_only", "right_only"})

    @property
    def is_three_way(self) -> bool:
        return self.base_sheet is not None
