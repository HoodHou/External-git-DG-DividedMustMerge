from __future__ import annotations

from table_merge_tool.alignment import align_sheets_three_way
from table_merge_tool.merge_rules import get_merge_rule
from table_merge_tool.models import (
    CONFLICT_KIND_THREE_WAY_DIVERGE,
    CONFLICT_KIND_THREE_WAY_LEFT_MODIFIED,
    CONFLICT_KIND_THREE_WAY_RIGHT_MODIFIED,
    CONFLICT_KIND_THREE_WAY_SAME_EDIT,
    CONFLICT_KIND_TWO_WAY,
    CellData,
    RowData,
    SheetData,
)


def _row(row_index: int, values: list[str], kind: str) -> RowData:
    return RowData(
        row_index=row_index,
        cells=[CellData(column_index=index, value=value) for index, value in enumerate(values, start=1)],
        kind=kind,
    )


def _sheet(name: str, headers: list[str], data_rows: list[list[str]]) -> SheetData:
    rows = [_row(1, headers, "header"), _row(2, headers, "header")]
    for offset, values in enumerate(data_rows, start=3):
        rows.append(_row(offset, values, "data"))
    return SheetData(
        name=name,
        rows=rows,
        max_columns=len(headers),
        display_header_row=1,
        field_header_row=2,
        logical_headers=headers,
        display_headers=headers,
    )


def _row_by_key(alignment, key_value: str):
    for row in alignment.rows:
        candidate = row.left_row or row.right_row or row.base_row
        if candidate is None:
            continue
        if candidate.kind != "data":
            continue
        if candidate.value_at(1) == key_value:
            return row
    raise AssertionError(f"No row with id={key_value}")


def test_three_way_alignment_requires_base_or_falls_back_to_two_way():
    base = _sheet("Skill", ["id", "name"], [["1", "A"]])
    left = _sheet("Skill", ["id", "name"], [["1", "A"]])
    right = _sheet("Skill", ["id", "name"], [["1", "A"]])

    alignment = align_sheets_three_way(base, left, right, get_merge_rule("full_keep_left"))
    assert alignment.base_sheet is base
    assert alignment.is_three_way is True

    two_way = align_sheets_three_way(None, left, right, get_merge_rule("full_keep_left"))
    assert two_way.base_sheet is None


def test_three_way_all_equal_yields_same_status():
    base = _sheet("Skill", ["id", "name"], [["1", "A"]])
    left = _sheet("Skill", ["id", "name"], [["1", "A"]])
    right = _sheet("Skill", ["id", "name"], [["1", "A"]])

    alignment = align_sheets_three_way(base, left, right, get_merge_rule("full_keep_left"))
    row = _row_by_key(alignment, "1")

    assert row.status == "same"
    assert row.conflict_columns == set()
    assert row.conflict_kind == CONFLICT_KIND_TWO_WAY
    assert row.merged_row.value_at(2) == "A"


def test_left_modified_only_auto_adopts_left_value():
    base = _sheet("Skill", ["id", "name"], [["1", "A"]])
    left = _sheet("Skill", ["id", "name"], [["1", "A2"]])
    right = _sheet("Skill", ["id", "name"], [["1", "A"]])

    alignment = align_sheets_three_way(base, left, right, get_merge_rule("full_keep_left"))
    row = _row_by_key(alignment, "1")

    assert row.status == "same"
    assert row.merged_row.value_at(2) == "A2"
    assert row.conflict_kind == CONFLICT_KIND_THREE_WAY_LEFT_MODIFIED
    assert 2 in row.column_conflict_kinds
    assert row.column_conflict_kinds[2] == CONFLICT_KIND_THREE_WAY_LEFT_MODIFIED


def test_right_modified_only_auto_adopts_right_value():
    base = _sheet("Skill", ["id", "name"], [["1", "A"]])
    left = _sheet("Skill", ["id", "name"], [["1", "A"]])
    right = _sheet("Skill", ["id", "name"], [["1", "A3"]])

    alignment = align_sheets_three_way(base, left, right, get_merge_rule("full_keep_left"))
    row = _row_by_key(alignment, "1")

    assert row.status == "same"
    assert row.merged_row.value_at(2) == "A3"
    assert row.conflict_kind == CONFLICT_KIND_THREE_WAY_RIGHT_MODIFIED


def test_both_sides_modified_to_same_value_is_not_conflict():
    base = _sheet("Skill", ["id", "name"], [["1", "A"]])
    left = _sheet("Skill", ["id", "name"], [["1", "B"]])
    right = _sheet("Skill", ["id", "name"], [["1", "B"]])

    alignment = align_sheets_three_way(base, left, right, get_merge_rule("full_keep_left"))
    row = _row_by_key(alignment, "1")

    assert row.status == "same"
    assert row.merged_row.value_at(2) == "B"
    assert row.conflict_kind == CONFLICT_KIND_THREE_WAY_SAME_EDIT


def test_three_way_divergence_is_conflict():
    base = _sheet("Skill", ["id", "name"], [["1", "A"]])
    left = _sheet("Skill", ["id", "name"], [["1", "B"]])
    right = _sheet("Skill", ["id", "name"], [["1", "C"]])

    alignment = align_sheets_three_way(base, left, right, get_merge_rule("full_keep_left"))
    row = _row_by_key(alignment, "1")

    assert row.status == "conflict"
    assert 2 in row.conflict_columns
    assert row.conflict_kind == CONFLICT_KIND_THREE_WAY_DIVERGE
    # full_keep_left defaults to left on conflict
    assert row.merged_row.value_at(2) == "B"


def test_left_deleted_right_unchanged_is_clean_delete():
    base = _sheet("Skill", ["id", "name"], [["1", "A"]])
    left = _sheet("Skill", ["id", "name"], [])
    right = _sheet("Skill", ["id", "name"], [["1", "A"]])

    alignment = align_sheets_three_way(base, left, right, get_merge_rule("full_keep_left"))
    row = _row_by_key(alignment, "1")

    assert row.status == "deleted"
    assert row.base_row is not None
    assert row.left_row is None


def test_left_deleted_right_modified_escalates_to_conflict():
    base = _sheet("Skill", ["id", "name"], [["1", "A"]])
    left = _sheet("Skill", ["id", "name"], [])
    right = _sheet("Skill", ["id", "name"], [["1", "A-modified"]])

    alignment = align_sheets_three_way(base, left, right, get_merge_rule("full_keep_left"))
    row = _row_by_key(alignment, "1")

    assert row.status == "conflict"
    assert row.conflict_kind == CONFLICT_KIND_THREE_WAY_DIVERGE
    assert 2 in row.conflict_columns


def test_both_sides_add_same_row_is_merged_cleanly():
    base = _sheet("Skill", ["id", "name"], [])
    left = _sheet("Skill", ["id", "name"], [["9", "新行"]])
    right = _sheet("Skill", ["id", "name"], [["9", "新行"]])

    alignment = align_sheets_three_way(base, left, right, get_merge_rule("full_keep_left"))
    row = _row_by_key(alignment, "9")

    assert row.status == "same"
    assert row.merged_row.value_at(2) == "新行"


def test_both_sides_add_divergent_rows_is_conflict():
    base = _sheet("Skill", ["id", "name"], [])
    left = _sheet("Skill", ["id", "name"], [["9", "L新行"]])
    right = _sheet("Skill", ["id", "name"], [["9", "R新行"]])

    alignment = align_sheets_three_way(base, left, right, get_merge_rule("full_keep_left"))
    row = _row_by_key(alignment, "9")

    assert row.status == "conflict"
    assert row.conflict_kind == CONFLICT_KIND_THREE_WAY_DIVERGE


def test_three_way_mixed_columns_classifies_per_column():
    base = _sheet("Skill", ["id", "name", "desc"], [["1", "原名", "原描述"]])
    # left changes name, right changes desc -> each column is single-side modified
    left = _sheet("Skill", ["id", "name", "desc"], [["1", "新名", "原描述"]])
    right = _sheet("Skill", ["id", "name", "desc"], [["1", "原名", "新描述"]])

    alignment = align_sheets_three_way(base, left, right, get_merge_rule("full_keep_left"))
    row = _row_by_key(alignment, "1")

    assert row.status == "same"
    assert row.merged_row.value_at(2) == "新名"
    assert row.merged_row.value_at(3) == "新描述"
    assert row.column_conflict_kinds[2] == CONFLICT_KIND_THREE_WAY_LEFT_MODIFIED
    assert row.column_conflict_kinds[3] == CONFLICT_KIND_THREE_WAY_RIGHT_MODIFIED


def test_three_way_only_in_base_is_treated_as_both_sides_deleted():
    base = _sheet("Skill", ["id", "name"], [["1", "A"]])
    left = _sheet("Skill", ["id", "name"], [])
    right = _sheet("Skill", ["id", "name"], [])

    alignment = align_sheets_three_way(base, left, right, get_merge_rule("full_keep_left"))
    row = _row_by_key(alignment, "1")

    assert row.status == "deleted"
