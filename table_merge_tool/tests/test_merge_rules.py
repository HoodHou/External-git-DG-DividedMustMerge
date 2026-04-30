from __future__ import annotations

from table_merge_tool.alignment import ComparisonOptions, align_sheets
from table_merge_tool.exporter import build_diff_report_rows, format_diff_report_text
from table_merge_tool.exporter import _included_columns, _should_export_sheet
from table_merge_tool.merge_rules import get_merge_rule
from table_merge_tool.models import CellData, RowData, SheetData


def _row(row_index: int, values: list[str], kind: str) -> RowData:
    return RowData(
        row_index=row_index,
        cells=[CellData(column_index=index, value=value) for index, value in enumerate(values, start=1)],
        kind=kind,
    )


def _sheet(name: str, headers: list[str], data_rows: list[list[str]]) -> SheetData:
    rows = [
        _row(1, headers, "header"),
        _row(2, headers, "header"),
    ]
    for row_offset, values in enumerate(data_rows, start=3):
        rows.append(_row(row_offset, values, "data"))
    return SheetData(
        name=name,
        rows=rows,
        max_columns=len(headers),
        display_header_row=1,
        field_header_row=2,
        logical_headers=headers,
        display_headers=headers,
    )


def _sample_sheets():
    left = _sheet(
        "Skill",
        ["id", "name", "left_only"],
        [
            ["1", "左值", "L-flag"],
            ["2", "只在左边", "L-only"],
        ],
    )
    right = _sheet(
        "Skill",
        ["id", "name", "right_only"],
        [
            ["1", "右值", "R-flag"],
            ["3", "只在右边", "R-only"],
        ],
    )
    return left, right


def test_full_keep_left_preserves_both_sides_and_prefers_left_on_conflict():
    left, right = _sample_sheets()

    alignment = align_sheets(left, right, get_merge_rule("full_keep_left"))

    merged_data_rows = [row for row in alignment.rows if row.left_row or row.right_row]
    conflict_row = next(row for row in merged_data_rows if row.left_row and row.right_row and row.left_row.value_at(1) == "1")

    assert conflict_row.status == "conflict"
    assert conflict_row.merged_row.value_at(2) == "左值"
    assert conflict_row.merged_row.value_at(4) == "R-flag"
    assert any(row.status == "left_only" for row in alignment.rows)
    assert any(row.status == "right_only" for row in alignment.rows)


def test_preferred_key_aligns_same_id_even_when_blank_rows_shift_position():
    left = _sheet(
        "Skill",
        ["id", "name"],
        [
            ["39", "敌方"],
            ["40", "己方"],
            ["", ""],
            ["41", "敌方"],
        ],
    )
    right = _sheet(
        "Skill",
        ["id", "name"],
        [
            ["39", "敌方"],
            ["", ""],
            ["40", "己方"],
            ["41", "敌方"],
        ],
    )

    alignment = align_sheets(left, right, get_merge_rule("full_keep_left"), preferred_key_fields=["id"])

    id_40_rows = [
        row
        for row in alignment.rows
        if (row.left_row is not None and row.left_row.value_at(1) == "40")
        or (row.right_row is not None and row.right_row.value_at(1) == "40")
    ]
    assert len(id_40_rows) == 1
    assert id_40_rows[0].left_row is not None
    assert id_40_rows[0].right_row is not None
    assert id_40_rows[0].status == "same"


def test_right_priority_drops_left_only_additions_and_keeps_right_conflict_value():
    left, right = _sample_sheets()

    alignment = align_sheets(left, right, get_merge_rule("right_priority"))

    conflict_row = next(row for row in alignment.rows if row.left_row and row.right_row and row.left_row.value_at(1) == "1")
    left_only_row = next(row for row in alignment.rows if row.left_row and row.right_row is None and row.left_row.value_at(1) == "2")
    right_only_row = next(row for row in alignment.rows if row.right_row and row.left_row is None and row.right_row.value_at(1) == "3")

    assert conflict_row.merged_row.value_at(2) == "右值"
    assert conflict_row.merged_row.value_at(3) == ""
    assert left_only_row.status == "deleted"
    assert right_only_row.status == "right_only"

    included_titles = [binding.title for _, binding in _included_columns(alignment, get_merge_rule("right_priority"))]
    assert included_titles == ["id", "name", "right_only"]


def test_left_only_sheet_can_be_removed_by_right_priority_rule():
    left = _sheet("OnlyLeft", ["id", "name"], [["1", "left"]])
    alignment = align_sheets(left, None, get_merge_rule("right_priority"))

    assert _should_export_sheet(alignment, get_merge_rule("right_priority")) is False


def test_complete_left_keeps_left_baseline_without_auto_filling_right_values():
    left = _sheet("Skill", ["id", "name"], [["1", "",]])
    right = _sheet("Skill", ["id", "name"], [["1", "右侧补充值"]])

    alignment = align_sheets(left, right, get_merge_rule("complete_left"))

    conflict_row = next(row for row in alignment.rows if row.left_row and row.right_row and row.left_row.value_at(1) == "1")

    assert conflict_row.status == "conflict"
    assert conflict_row.merged_row.value_at(2) == ""


def test_complete_right_keeps_right_baseline_without_auto_filling_left_values():
    left = _sheet("Skill", ["id", "name"], [["1", "左侧补充值"]])
    right = _sheet("Skill", ["id", "name"], [["1", ""]])

    alignment = align_sheets(left, right, get_merge_rule("complete_right"))

    conflict_row = next(row for row in alignment.rows if row.left_row and row.right_row and row.left_row.value_at(1) == "1")

    assert conflict_row.status == "conflict"
    assert conflict_row.merged_row.value_at(2) == ""


def test_fill_empty_left_only_fills_left_blanks_and_drops_right_additions():
    left = _sheet("Skill", ["id", "name"], [["1", ""], ["2", "左侧行"]])
    right = _sheet("Skill", ["id", "name", "right_only"], [["1", "右侧补充值", "R"], ["3", "右侧新增", "R-only"]])

    alignment = align_sheets(left, right, get_merge_rule("fill_empty_left"))

    filled_row = next(row for row in alignment.rows if row.left_row and row.right_row and row.left_row.value_at(1) == "1")
    right_only_row = next(row for row in alignment.rows if row.right_row and row.left_row is None and row.right_row.value_at(1) == "3")

    assert filled_row.status == "same"
    assert filled_row.merged_row.value_at(2) == "右侧补充值"
    assert filled_row.merged_row.value_at(3) == ""
    assert right_only_row.status == "deleted"


def test_fill_empty_right_only_fills_right_blanks_and_drops_left_additions():
    left = _sheet("Skill", ["id", "name", "left_only"], [["1", "左侧补充值", "L"], ["2", "左侧新增", "L-only"]])
    right = _sheet("Skill", ["id", "name"], [["1", ""]])

    alignment = align_sheets(left, right, get_merge_rule("fill_empty_right"))

    filled_row = next(row for row in alignment.rows if row.left_row and row.right_row and row.left_row.value_at(1) == "1")
    left_only_row = next(row for row in alignment.rows if row.left_row and row.right_row is None and row.left_row.value_at(1) == "2")

    assert filled_row.status == "same"
    assert filled_row.merged_row.value_at(2) == "左侧补充值"
    assert filled_row.merged_row.value_at(3) == ""
    assert left_only_row.status == "deleted"


def test_full_conflict_blank_keeps_additions_but_blanks_conflict_values():
    left, right = _sample_sheets()

    alignment = align_sheets(left, right, get_merge_rule("full_conflict_blank"))

    conflict_row = next(row for row in alignment.rows if row.left_row and row.right_row and row.left_row.value_at(1) == "1")

    assert conflict_row.status == "conflict"
    assert conflict_row.merged_row.value_at(2) == ""
    assert conflict_row.merged_row.value_at(3) == "L-flag"
    assert conflict_row.merged_row.value_at(4) == "R-flag"
    assert any(row.status == "left_only" for row in alignment.rows)
    assert any(row.status == "right_only" for row in alignment.rows)


def test_append_rows_keep_existing_preserves_shared_rows_and_appends_right_rows_only():
    left = _sheet("Skill", ["id", "name", "left_only"], [["1", "左原值", "L"], ["2", "左侧已有", "L2"]])
    right = _sheet("Skill", ["id", "name", "right_only"], [["1", "右修改", "R"], ["3", "右侧新增", "R3"]])

    alignment = align_sheets(left, right, get_merge_rule("append_rows_keep_existing"))

    shared_row = next(row for row in alignment.rows if row.left_row and row.right_row and row.left_row.value_at(1) == "1")
    right_only_row = next(row for row in alignment.rows if row.right_row and row.left_row is None and row.right_row.value_at(1) == "3")

    assert shared_row.status == "same"
    assert shared_row.merged_row.value_at(2) == "左原值"
    assert shared_row.merged_row.value_at(3) == "L"
    assert shared_row.merged_row.value_at(4) == ""
    assert right_only_row.status == "right_only"


def test_ignore_trim_whitespace_option_treats_trim_only_changes_as_same():
    left = _sheet("Skill", ["id", "effect"], [["1", "  abc  "]])
    right = _sheet("Skill", ["id", "effect"], [["1", "abc"]])

    strict_alignment = align_sheets(left, right, get_merge_rule("full_keep_left"))
    relaxed_alignment = align_sheets(
        left,
        right,
        get_merge_rule("full_keep_left"),
        ComparisonOptions(ignore_trim_whitespace=True),
    )

    strict_row = next(row for row in strict_alignment.rows if row.left_row and row.right_row and row.left_row.value_at(1) == "1")
    relaxed_row = next(row for row in relaxed_alignment.rows if row.left_row and row.right_row and row.left_row.value_at(1) == "1")

    assert strict_row.status == "conflict"
    assert relaxed_row.status == "same"
    assert relaxed_row.conflict_columns == set()


def test_ignore_trim_whitespace_option_collapses_inner_newlines():
    left = _sheet("Skill", ["id", "remark"], [["1", "备注\n(有公式）"]])
    right = _sheet("Skill", ["id", "remark"], [["1", "备注 (有公式）"]])

    strict_alignment = align_sheets(left, right, get_merge_rule("full_keep_left"))
    relaxed_alignment = align_sheets(
        left,
        right,
        get_merge_rule("full_keep_left"),
        ComparisonOptions(ignore_trim_whitespace=True),
    )

    strict_row = next(row for row in strict_alignment.rows if row.left_row and row.right_row and row.left_row.value_at(1) == "1")
    relaxed_row = next(row for row in relaxed_alignment.rows if row.left_row and row.right_row and row.left_row.value_at(1) == "1")

    assert strict_row.status == "conflict"
    assert relaxed_row.status == "same"
    assert relaxed_row.conflict_columns == set()


def test_manual_key_fields_can_override_auto_id_matching():
    left = _sheet("Skill", ["id", "name", "value"], [["1", "火球", "左侧效果"]])
    right = _sheet("Skill", ["id", "name", "value"], [["99", "火球", "右侧效果"]])

    auto_alignment = align_sheets(left, right, get_merge_rule("full_keep_left"))
    manual_alignment = align_sheets(
        left,
        right,
        get_merge_rule("full_keep_left"),
        preferred_key_fields=["name"],
    )

    assert any(row.status == "left_only" for row in auto_alignment.rows)
    assert any(row.status == "right_only" for row in auto_alignment.rows)

    matched_row = next(
        row for row in manual_alignment.rows if row.left_row and row.right_row and row.left_row.kind == "data"
    )
    assert manual_alignment.key_fields == ["name"]
    assert matched_row.left_row.value_at(2) == "火球"
    assert matched_row.right_row.value_at(2) == "火球"


def test_duplicate_id_uses_unique_composite_key_instead_of_id_only():
    left = _sheet(
        "Skill",
        ["id", "level", "name", "value"],
        [
            ["1", "1", "火球1", "左1"],
            ["1", "2", "火球2", "左2"],
        ],
    )
    right = _sheet(
        "Skill",
        ["id", "level", "name", "value"],
        [
            ["1", "1", "火球1", "右1"],
            ["1", "2", "火球2", "右2"],
        ],
    )

    alignment = align_sheets(left, right, get_merge_rule("full_keep_left"))

    assert alignment.key_fields == ["id", "level"]
    level_2_row = next(row for row in alignment.rows if row.left_row and row.left_row.value_at(2) == "2")
    assert level_2_row.right_row is not None
    assert level_2_row.right_row.value_at(2) == "2"
    assert level_2_row.right_row.value_at(3) == "火球2"


def test_manual_duplicate_id_key_is_respected_and_pairs_by_occurrence():
    left = _sheet("Skill", ["id", "name", "value"], [["1", "火球1", "左1"], ["1", "火球2", "左2"]])
    right = _sheet("Skill", ["id", "name", "value"], [["1", "火球2", "右2"], ["1", "火球1", "右1"]])

    alignment = align_sheets(
        left,
        right,
        get_merge_rule("full_keep_left"),
        preferred_key_fields=["id"],
    )

    assert alignment.key_fields == ["id"]
    matched_rows = [row for row in alignment.rows if row.left_row and row.right_row and row.left_row.kind == "data"]
    assert len(matched_rows) == 2
    assert matched_rows[0].left_row.value_at(2) == "火球1"
    assert matched_rows[0].right_row.value_at(2) == "火球2"


def test_strict_single_id_pairs_rows_by_unique_id_regardless_of_order():
    left = _sheet("Skill", ["id", "name"], [["1", "火球"], ["2", "冰箭"]])
    right = _sheet("Skill", ["id", "name"], [["2", "冰箭右"], ["1", "火球右"]])

    alignment = align_sheets(
        left,
        right,
        get_merge_rule("full_keep_left"),
        preferred_key_fields=["id"],
        strict_single_key=True,
    )

    matched_rows = [row for row in alignment.rows if row.left_row and row.right_row and row.left_row.kind == "data"]
    assert alignment.key_fields == ["id"]
    assert [row.left_row.value_at(1) for row in matched_rows] == ["1", "2"]
    assert [row.right_row.value_at(1) for row in matched_rows] == ["1", "2"]


def test_strict_single_id_rejects_duplicate_id():
    left = _sheet("Skill", ["id", "name"], [["1", "火球"], ["1", "冰箭"]])
    right = _sheet("Skill", ["id", "name"], [["1", "火球"]])

    try:
        align_sheets(
            left,
            right,
            get_merge_rule("full_keep_left"),
            preferred_key_fields=["id"],
            strict_single_key=True,
        )
    except ValueError as exc:
        assert "不唯一" in str(exc)
    else:
        raise AssertionError("strict single id should reject duplicates")


def test_strict_single_id_allows_empty_id_rows_and_falls_back_to_regular_alignment():
    left = _sheet(
        "Skill",
        ["id", "name"],
        [["1", "火球"], ["", "备注行"], ["2", "冰箭"]],
    )
    right = _sheet(
        "Skill",
        ["id", "name"],
        [["2", "冰箭右"], ["1", "火球右"], ["", "备注行"]],
    )

    alignment = align_sheets(
        left,
        right,
        get_merge_rule("full_keep_left"),
        preferred_key_fields=["id"],
        strict_single_key=True,
    )

    matched_id_rows = [
        row
        for row in alignment.rows
        if row.left_row and row.right_row and row.left_row.kind == "data" and row.left_row.value_at(1)
    ]
    fallback_rows = [
        row
        for row in alignment.rows
        if row.left_row and row.right_row and row.left_row.kind == "data" and not row.left_row.value_at(1)
    ]
    assert [row.left_row.value_at(1) for row in matched_id_rows] == ["1", "2"]
    assert [row.right_row.value_at(1) for row in matched_id_rows] == ["1", "2"]
    assert fallback_rows
    assert fallback_rows[0].left_row.value_at(2) == "备注行"
    assert fallback_rows[0].right_row.value_at(2) == "备注行"


def test_strict_single_id_requires_one_manual_field():
    left = _sheet("Skill", ["id", "level", "name"], [["1", "1", "火球"]])
    right = _sheet("Skill", ["id", "level", "name"], [["1", "1", "火球"]])

    try:
        align_sheets(
            left,
            right,
            get_merge_rule("full_keep_left"),
            preferred_key_fields=["id", "level"],
            strict_single_key=True,
        )
    except ValueError as exc:
        assert "只能指定一个" in str(exc)
    else:
        raise AssertionError("strict single id should require exactly one field")


def test_diff_report_rows_export_only_changed_fields():
    left = _sheet("Skill", ["id", "name", "effect"], [["1", "火球", "造成10点伤害"]])
    right = _sheet("Skill", ["id", "name", "effect"], [["1", "火球", "造成12点伤害"]])

    alignment = align_sheets(left, right, get_merge_rule("full_keep_left"))
    rows = build_diff_report_rows({"Skill": alignment})
    text = format_diff_report_text(rows)

    assert len(rows) == 1
    assert rows[0]["sheet_name"] == "Skill"
    assert rows[0]["field_key"] == "effect"
    assert rows[0]["left_value"] == "造成10点伤害"
    assert rows[0]["right_value"] == "造成12点伤害"
    assert rows[0]["first_diff_char"] == "4"
    assert "[工作表] Skill" in text
    assert "字段: effect (effect)" in text
