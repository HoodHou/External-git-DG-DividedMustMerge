from __future__ import annotations

from PySide6.QtCore import Qt

from table_merge_tool.gui import MergeTableModel
from table_merge_tool.models import AlignedRow, CellData, ColumnBinding, RowData, SheetAlignment


def _alignment_with_row(status: str, left: RowData | None, right: RowData | None) -> SheetAlignment:
    merged = (left or right or RowData(row_index=1, cells=[])).clone()
    return SheetAlignment(
        sheet_name="Sheet",
        columns=[ColumnBinding(key="id", title="id", left_index=1, right_index=1)],
        rows=[AlignedRow(left_row=left, right_row=right, merged_row=merged, status=status, reason=status)],
    )


def _row(value: str) -> RowData:
    return RowData(row_index=1, cells=[CellData(column_index=1, value=value)])


def test_right_only_marker_only_appears_on_right_side_header():
    alignment = _alignment_with_row("right_only", None, _row("R"))

    assert MergeTableModel(alignment, "left").headerData(0, Qt.Vertical, Qt.DisplayRole) == "1"
    assert MergeTableModel(alignment, "right").headerData(0, Qt.Vertical, Qt.DisplayRole) == "1 +"


def test_left_only_marker_only_appears_on_left_side_header():
    alignment = _alignment_with_row("left_only", _row("L"), None)

    assert MergeTableModel(alignment, "left").headerData(0, Qt.Vertical, Qt.DisplayRole) == "1 -"
    assert MergeTableModel(alignment, "right").headerData(0, Qt.Vertical, Qt.DisplayRole) == "1"


def test_cell_tooltip_shows_char_level_diff_on_diff_cell():
    left_row = _row("今天天不好")
    right_row = _row("今天天挺好")
    alignment = _alignment_with_row("same", left_row, right_row)
    model = MergeTableModel(alignment, "middle")

    index = model.index(0, 0)
    tooltip = model.data(index, Qt.ToolTipRole)

    assert tooltip is not None
    # Inline-style spans so QToolTip renders them correctly
    assert "background-color:#ffebe9" in tooltip  # delete (不)
    assert "background-color:#dafbe1" in tooltip  # insert (挺)
    assert "<b>左:</b>" in tooltip
    assert "<b>右:</b>" in tooltip


def test_cell_tooltip_falls_back_to_plain_value_when_no_diff():
    row_left = _row("同值")
    row_right = _row("同值")
    alignment = _alignment_with_row("same", row_left, row_right)
    model = MergeTableModel(alignment, "middle")

    tooltip = model.data(model.index(0, 0), Qt.ToolTipRole)

    assert tooltip is not None
    assert "同值" in tooltip
    # No diff styling when values match
    assert "background-color:#ffebe9" not in tooltip
    assert "background-color:#dafbe1" not in tooltip


def test_cell_tooltip_highlights_conflict_marker_when_present():
    left_row = _row("A")
    right_row = _row("B")
    alignment = _alignment_with_row("conflict", left_row, right_row)
    alignment.rows[0].conflict_columns.add(1)
    model = MergeTableModel(alignment, "middle")

    tooltip = model.data(model.index(0, 0), Qt.ToolTipRole)

    assert tooltip is not None
    assert "冲突" in tooltip


def test_search_matches_left_only_content():
    alignment = _alignment_with_row("left_only", _row("只在左边"), None)
    model = MergeTableModel(alignment, "left")

    assert model.compute_visible_rows("all", "左边") == [0]


def test_search_matches_right_only_content():
    alignment = _alignment_with_row("right_only", None, _row("只在右边"))
    model = MergeTableModel(alignment, "left")

    assert model.compute_visible_rows("all", "右边") == [0]
