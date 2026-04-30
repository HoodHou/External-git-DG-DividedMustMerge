from __future__ import annotations

import re
from dataclasses import dataclass
from difflib import SequenceMatcher
from functools import lru_cache
from itertools import combinations

from .excel_xml import clone_row_with_values
from .merge_rules import MergeRule, get_merge_rule, is_binding_included
from .models import (
    CONFLICT_KIND_THREE_WAY_DIVERGE,
    CONFLICT_KIND_THREE_WAY_LEFT_MODIFIED,
    CONFLICT_KIND_THREE_WAY_RIGHT_MODIFIED,
    CONFLICT_KIND_THREE_WAY_SAME_EDIT,
    CONFLICT_KIND_TWO_WAY,
    AlignedRow,
    ColumnBinding,
    RowData,
    SheetAlignment,
    SheetData,
)


ID_FIELD_RE = re.compile(r"(?:^|_)(id|level|key|code|name|type)(?:$|_)", re.IGNORECASE)
STRUCTURAL_ROW_KINDS = {"blank", "header", "group", "note"}
MATCH_THRESHOLD = 0.9


@dataclass(frozen=True, slots=True)
class ComparisonOptions:
    ignore_trim_whitespace: bool = False
    ignore_all_whitespace: bool = False
    ignore_case: bool = False
    normalize_fullwidth: bool = False
    numeric_tolerance: float = 0.0


def align_sheets(
    left_sheet: SheetData | None,
    right_sheet: SheetData | None,
    merge_rule: MergeRule | None = None,
    comparison_options: ComparisonOptions | None = None,
    preferred_key_fields: list[str] | None = None,
    strict_single_key: bool = False,
) -> SheetAlignment:
    merge_rule = get_merge_rule(merge_rule.rule_id if merge_rule is not None else None)
    comparison_options = comparison_options or ComparisonOptions()
    sheet_name = left_sheet.name if left_sheet is not None else right_sheet.name
    columns = build_column_bindings(left_sheet, right_sheet)
    shared_keys = resolve_key_fields(left_sheet, right_sheet, preferred_key_fields)

    if strict_single_key:
        shared_keys = resolve_strict_single_key_fields(left_sheet, right_sheet, preferred_key_fields)
        return align_sheets_by_strict_key(
            left_sheet,
            right_sheet,
            columns,
            shared_keys,
            merge_rule,
            comparison_options,
            sheet_name or "",
        )

    left_rows = left_sheet.rows if left_sheet is not None else []
    right_rows = right_sheet.rows if right_sheet is not None else []
    left_resolver = _make_row_map_resolver(left_sheet)
    right_resolver = _make_row_map_resolver(right_sheet)
    left_signatures = [
        row_signature(left_sheet, row, shared_keys, comparison_options, logical_map_resolver=left_resolver)
        for row in left_rows
    ]
    right_signatures = [
        row_signature(right_sheet, row, shared_keys, comparison_options, logical_map_resolver=right_resolver)
        for row in right_rows
    ]
    matcher = SequenceMatcher(a=left_signatures, b=right_signatures, autojunk=False)

    aligned_rows: list[AlignedRow] = []
    for opcode, left_a, left_b, right_a, right_b in matcher.get_opcodes():
        if opcode == "equal":
            for offset in range(left_b - left_a):
                aligned_rows.append(
                    make_aligned_row(
                        left_rows[left_a + offset],
                        right_rows[right_a + offset],
                        columns,
                        merge_rule,
                        comparison_options,
                    )
                )
        elif opcode == "replace":
            aligned_rows.extend(
                align_replace_block(
                    left_sheet,
                    right_sheet,
                    left_rows[left_a:left_b],
                    right_rows[right_a:right_b],
                    columns,
                    shared_keys,
                    merge_rule,
                    comparison_options,
                    left_logical_map_resolver=left_resolver,
                    right_logical_map_resolver=right_resolver,
                )
            )
        elif opcode == "delete":
            for left_row in left_rows[left_a:left_b]:
                aligned_rows.append(make_aligned_row(left_row, None, columns, merge_rule, comparison_options))
        elif opcode == "insert":
            for right_row in right_rows[right_a:right_b]:
                aligned_rows.append(make_aligned_row(None, right_row, columns, merge_rule, comparison_options))

    return SheetAlignment(
        sheet_name=sheet_name,
        columns=columns,
        rows=aligned_rows,
        left_sheet=left_sheet,
        right_sheet=right_sheet,
        key_fields=shared_keys,
    )


def _resolve_three_way_keys(
    base_sheet: SheetData | None,
    left_sheet: SheetData | None,
    right_sheet: SheetData | None,
    preferred_key_fields: list[str] | None = None,
) -> list[str]:
    """Resolve key fields for three-way alignment. Unlike the two-way variant,
    this accepts a key when *any* side can infer it and at least two sides
    (one of them base) contain the corresponding header.
    """
    base_headers = normalized_sheet_headers(base_sheet)
    left_headers = normalized_sheet_headers(left_sheet)
    right_headers = normalized_sheet_headers(right_sheet)

    if preferred_key_fields:
        preferred: list[str] = []
        seen: set[str] = set()
        for raw in preferred_key_fields:
            normalized = normalize_header(raw)
            if normalized and normalized not in seen and normalized in base_headers:
                preferred.append(normalized)
                seen.add(normalized)
        if preferred:
            return preferred

    candidates: list[list[str]] = []
    for keys in (
        resolve_key_fields(left_sheet, right_sheet, preferred_key_fields),
        infer_key_fields(left_sheet),
        infer_key_fields(right_sheet),
        infer_key_fields(base_sheet),
    ):
        if keys and keys not in candidates:
            candidates.append(keys)

    for keys in candidates:
        if all(key in base_headers for key in keys):
            # Require at least one non-base side to also carry the columns,
            # otherwise left/right rows can never match by key.
            left_has = all(key in left_headers for key in keys) if left_headers else False
            right_has = all(key in right_headers for key in keys) if right_headers else False
            if left_has or right_has:
                return keys
    return []


def _resolve_strict_three_way_key_fields(
    base_sheet: SheetData | None,
    left_sheet: SheetData | None,
    right_sheet: SheetData | None,
    preferred_key_fields: list[str] | None = None,
) -> list[str]:
    keys = resolve_strict_single_key_fields(left_sheet, right_sheet, preferred_key_fields)
    key_field = keys[0]
    if base_sheet is not None and key_field not in normalized_sheet_headers(base_sheet):
        raise ValueError(f"强单一ID字段“{key_field}”不存在于：Base。")
    return keys


def resolve_key_fields(
    left_sheet: SheetData | None,
    right_sheet: SheetData | None,
    preferred_key_fields: list[str] | None = None,
) -> list[str]:
    preferred = []
    seen: set[str] = set()
    for field in preferred_key_fields or []:
        normalized = normalize_header(field)
        if normalized and normalized not in seen:
            preferred.append(normalized)
            seen.add(normalized)
    if preferred:
        shared_headers = shared_sheet_headers(left_sheet, right_sheet)
        manual_keys = [field for field in preferred if field in shared_headers]
        if manual_keys:
            return manual_keys

    left_keys = infer_key_fields(left_sheet)
    right_keys = infer_key_fields(right_sheet)
    return [key for key in left_keys if key in right_keys]


def resolve_strict_single_key_fields(
    left_sheet: SheetData | None,
    right_sheet: SheetData | None,
    preferred_key_fields: list[str] | None = None,
) -> list[str]:
    preferred = []
    seen: set[str] = set()
    for field in preferred_key_fields or []:
        normalized = normalize_header(field)
        if normalized and normalized not in seen:
            preferred.append(normalized)
            seen.add(normalized)
    if len(preferred) != 1:
        raise ValueError("强单一ID模式需要且只能指定一个关键字段。")

    key_field = preferred[0]
    missing_sides: list[str] = []
    if left_sheet is not None and key_field not in normalized_sheet_headers(left_sheet):
        missing_sides.append("左侧")
    if right_sheet is not None and key_field not in normalized_sheet_headers(right_sheet):
        missing_sides.append("右侧")
    if missing_sides:
        raise ValueError(f"强单一ID字段“{key_field}”不存在于：{'、'.join(missing_sides)}。")
    return [key_field]


def shared_sheet_headers(left_sheet: SheetData | None, right_sheet: SheetData | None) -> set[str]:
    if left_sheet is None or right_sheet is None:
        return set()
    return normalized_sheet_headers(left_sheet) & normalized_sheet_headers(right_sheet)


def normalized_sheet_headers(sheet: SheetData | None) -> set[str]:
    if sheet is None:
        return set()
    return {normalize_header(header) for header in sheet.logical_headers if normalize_header(header)}


def build_column_bindings(
    left_sheet: SheetData | None,
    right_sheet: SheetData | None,
    base_sheet: SheetData | None = None,
) -> list[ColumnBinding]:
    bindings: list[ColumnBinding] = []
    seen_keys: set[str] = set()

    def append_sheet(sheet: SheetData | None, side: str) -> None:
        if sheet is None:
            return
        for index in range(1, sheet.max_columns + 1):
            logical = normalize_header(sheet.logical_headers[index - 1] if index - 1 < len(sheet.logical_headers) else f"col_{index}")
            display = sheet.display_headers[index - 1] if index - 1 < len(sheet.display_headers) else logical
            if logical in seen_keys:
                binding = next(item for item in bindings if item.key == logical)
            else:
                binding = ColumnBinding(key=logical, title=display)
                bindings.append(binding)
                seen_keys.add(logical)

            if side == "left" and binding.left_index is None:
                binding.left_index = index
            if side == "right" and binding.right_index is None:
                binding.right_index = index
            if side == "base" and binding.base_index is None:
                binding.base_index = index

    append_sheet(base_sheet, "base")
    append_sheet(left_sheet, "left")
    append_sheet(right_sheet, "right")
    return bindings


def infer_key_fields(sheet: SheetData | None) -> list[str]:
    if sheet is None:
        return []

    data_rows = [row for row in sheet.rows if row.kind == "data"]
    if not data_rows:
        return []

    exact_id_indexes = [
        index
        for index, header in enumerate(sheet.logical_headers, start=1)
        if normalize_header(header) == "id"
    ]
    for index in exact_id_indexes:
        keys = [row.value_at(index).strip() for row in data_rows if row.value_at(index).strip()]
        if keys and len(set(keys)) == len(keys):
            return ["id"]

    candidate_indexes = [
        index
        for index, header in enumerate(sheet.logical_headers, start=1)
        if ID_FIELD_RE.search(header) is not None
    ]
    if not candidate_indexes:
        return []

    candidate_indexes = candidate_indexes[:6]
    best_fields: list[str] = []
    best_score = -1.0

    for width in range(1, min(3, len(candidate_indexes)) + 1):
        for indexes in combinations(candidate_indexes, width):
            keys = [tuple(row.value_at(index).strip() for index in indexes) for row in data_rows]
            keys = [key for key in keys if all(key)]
            if not keys or len(set(keys)) != len(keys):
                continue

            coverage = len(keys) / max(len(data_rows), 1)
            id_bonus = sum(0.15 for index in indexes if "id" in sheet.logical_headers[index - 1].lower())
            score = coverage + id_bonus
            if score > best_score:
                best_score = score
                best_fields = [normalize_header(sheet.logical_headers[index - 1]) for index in indexes]

    return best_fields if best_score >= 0.8 else []


def _key_fields_are_unique(sheet: SheetData | None, key_fields: list[str]) -> bool:
    if sheet is None or not key_fields:
        return False
    rows = [row for row in sheet.rows if row.kind == "data"]
    if not rows:
        return False
    resolver = _make_row_map_resolver(sheet)
    keys: list[tuple[str, ...]] = []
    for row in rows:
        logical_map = resolver(row)
        key = tuple(logical_map.get(field, "").strip() for field in key_fields)
        if not all(key):
            continue
        keys.append(key)
    return bool(keys) and len(set(keys)) == len(keys) and len(keys) / len(rows) >= 0.8


def row_signature(
    sheet: SheetData | None,
    row: RowData,
    key_fields: list[str],
    comparison_options: ComparisonOptions | None = None,
    *,
    logical_map_resolver=None,
) -> tuple:
    if sheet is None:
        return ("missing",)
    if row.kind == "blank":
        return ("blank",)
    if row.kind == "header":
        return ("header", row.row_index)
    if row.kind == "group":
        return ("group", compact_text(row, comparison_options))

    logical_map = logical_map_resolver(row) if logical_map_resolver is not None else row_logical_map(sheet, row)
    if key_fields:
        key_values = tuple(normalize_compare_text(logical_map.get(field, ""), comparison_options) for field in key_fields)
        if all(key_values):
            return ("key",) + key_values

    return (row.kind, content_fingerprint(sheet, row, comparison_options))


def align_sheets_by_strict_key(
    left_sheet: SheetData | None,
    right_sheet: SheetData | None,
    columns: list[ColumnBinding],
    key_fields: list[str],
    merge_rule: MergeRule,
    comparison_options: ComparisonOptions,
    sheet_name: str,
) -> SheetAlignment:
    _validate_strict_unique_key(left_sheet, key_fields, "左侧", comparison_options)
    _validate_strict_unique_key(right_sheet, key_fields, "右侧", comparison_options)

    left_rows = left_sheet.rows if left_sheet is not None else []
    right_rows = right_sheet.rows if right_sheet is not None else []
    left_leading_count = _leading_structural_count(left_rows)
    right_leading_count = _leading_structural_count(right_rows)
    left_data_rows = [row for row in left_rows if row.kind == "data"]
    right_data_rows = [row for row in right_rows if row.kind == "data"]
    left_index = _strict_key_index(left_sheet, left_data_rows, key_fields, comparison_options)
    right_index = _strict_key_index(right_sheet, right_data_rows, key_fields, comparison_options)

    aligned_rows: list[AlignedRow] = []
    for index in range(max(left_leading_count, right_leading_count)):
        left_row = left_rows[index] if index < left_leading_count else None
        right_row = right_rows[index] if index < right_leading_count else None
        aligned_rows.append(make_aligned_row(left_row, right_row, columns, merge_rule, comparison_options))

    seen_keys: set[tuple[str, ...]] = set()
    left_unkeyed_rows: list[RowData] = []
    right_unkeyed_rows: list[RowData] = []
    for left_row in left_data_rows:
        key = _strict_key_for_row(left_sheet, left_row, key_fields, comparison_options)
        if not all(key):
            left_unkeyed_rows.append(left_row)
            continue
        seen_keys.add(key)
        aligned_rows.append(
            make_aligned_row(left_row, right_index.get(key), columns, merge_rule, comparison_options)
        )
    for right_row in right_data_rows:
        key = _strict_key_for_row(right_sheet, right_row, key_fields, comparison_options)
        if not all(key):
            right_unkeyed_rows.append(right_row)
            continue
        if key in seen_keys:
            continue
        aligned_rows.append(make_aligned_row(None, right_row, columns, merge_rule, comparison_options))

    if left_unkeyed_rows or right_unkeyed_rows:
        aligned_rows.extend(
            align_replace_block(
                left_sheet,
                right_sheet,
                left_unkeyed_rows,
                right_unkeyed_rows,
                columns,
                [],
                merge_rule,
                comparison_options,
            )
        )

    for row in left_rows[left_leading_count:]:
        if row.kind != "data":
            aligned_rows.append(make_aligned_row(row, None, columns, merge_rule, comparison_options))
    for row in right_rows[right_leading_count:]:
        if row.kind != "data":
            aligned_rows.append(make_aligned_row(None, row, columns, merge_rule, comparison_options))

    return SheetAlignment(
        sheet_name=sheet_name,
        columns=columns,
        rows=aligned_rows,
        left_sheet=left_sheet,
        right_sheet=right_sheet,
        key_fields=key_fields,
    )


def _leading_structural_count(rows: list[RowData]) -> int:
    count = 0
    for row in rows:
        if row.kind == "data":
            break
        count += 1
    return count


def _strict_key_index(
    sheet: SheetData | None,
    rows: list[RowData],
    key_fields: list[str],
    comparison_options: ComparisonOptions,
) -> dict[tuple[str, ...], RowData]:
    return {
        _strict_key_for_row(sheet, row, key_fields, comparison_options): row
        for row in rows
        if all(_strict_key_for_row(sheet, row, key_fields, comparison_options))
    }


def _strict_key_for_row(
    sheet: SheetData | None,
    row: RowData,
    key_fields: list[str],
    comparison_options: ComparisonOptions,
) -> tuple[str, ...]:
    logical_map = row_logical_map(sheet, row) if sheet is not None else {}
    return tuple(normalize_compare_text(logical_map.get(field, ""), comparison_options) for field in key_fields)


def _validate_strict_unique_key(
    sheet: SheetData | None,
    key_fields: list[str],
    side_label: str,
    comparison_options: ComparisonOptions,
) -> None:
    if sheet is None:
        return
    rows = [row for row in sheet.rows if row.kind == "data"]
    if not rows:
        return
    seen: dict[tuple[str, ...], int] = {}
    duplicate_samples: list[str] = []
    for row in rows:
        key = _strict_key_for_row(sheet, row, key_fields, comparison_options)
        if not all(key):
            continue
        previous = seen.get(key)
        if previous is not None:
            duplicate_samples.append(f"{'/'.join(key)}(行{previous},行{row.row_index})")
        else:
            seen[key] = row.row_index
    if duplicate_samples:
        sample = "、".join(duplicate_samples[:6])
        suffix = "..." if len(duplicate_samples) > 6 else ""
        raise ValueError(f"强单一ID字段“{key_fields[0]}”在{side_label}不唯一：{sample}{suffix}。")


def make_aligned_row(
    left_row: RowData | None,
    right_row: RowData | None,
    columns: list[ColumnBinding],
    merge_rule: MergeRule,
    comparison_options: ComparisonOptions | None = None,
) -> AlignedRow:
    merged_values: dict[int, str] = {}
    conflict_columns: set[int] = set()

    if left_row is None and right_row is not None:
        return _build_single_side_row(
            None,
            right_row,
            columns,
            merge_rule.keep_right_only_rows,
            "仅右侧存在，已按规则保留",
            "按规则忽略仅右侧新增",
        )

    if right_row is None and left_row is not None:
        return _build_single_side_row(
            left_row,
            None,
            columns,
            merge_rule.keep_left_only_rows,
            "仅左侧存在，已按规则保留",
            "按规则忽略仅左侧新增",
        )

    left_row = left_row or RowData(row_index=0)
    right_row = right_row or RowData(row_index=0)
    status = "same"
    reason = "自动对齐"

    for index, binding in enumerate(columns, start=1):
        left_value = left_row.value_at(binding.left_index or -1) if binding.left_index else ""
        right_value = right_row.value_at(binding.right_index or -1) if binding.right_index else ""
        if not is_binding_included(binding, merge_rule):
            merged_value = ""
        elif merge_rule.preserve_existing_rows:
            merged_value = left_value
        elif compare_text_values(left_value, right_value, comparison_options):
            merged_value = left_value
        elif not left_value and merge_rule.fill_missing_from_other_side:
            merged_value = right_value
        elif not right_value and merge_rule.fill_missing_from_other_side:
            merged_value = left_value
        else:
            status = "conflict"
            if merge_rule.conflict_resolution == "blank":
                preferred_side = "空值"
                merged_value = ""
            elif merge_rule.conflict_resolution == "left":
                preferred_side = "左侧"
                merged_value = left_value
            else:
                preferred_side = "右侧"
                merged_value = right_value
            reason = f"存在冲突，默认取{preferred_side}，请人工确认"
            conflict_columns.add(index)
        merged_values[index] = merged_value

    merged_row = clone_row_with_values(left_row, merged_values)
    merged_row.kind = left_row.kind if left_row.kind != "blank" else right_row.kind
    return AlignedRow(left_row, right_row, merged_row, status, reason, conflict_columns=conflict_columns)


def make_aligned_row_three_way(
    base_row: RowData | None,
    left_row: RowData | None,
    right_row: RowData | None,
    columns: list[ColumnBinding],
    merge_rule: MergeRule,
    comparison_options: ComparisonOptions | None = None,
) -> AlignedRow:
    """Align a single row across base/left/right, producing per-column
    ``conflict_kind`` classification and auto-resolving single-side edits.
    """
    # Delegate pure left/right-only additions (no base) back to the 2-way path
    # so the existing behaviour stays untouched.
    if base_row is None and (left_row is None or right_row is None):
        return make_aligned_row(left_row, right_row, columns, merge_rule, comparison_options)

    # Both-side deletion -> propagate as deleted row.
    if left_row is None and right_row is None:
        anchor = base_row or RowData(row_index=0)
        merged_row = clone_row_with_values(anchor, {index: "" for index in range(1, len(columns) + 1)})
        merged_row.kind = anchor.kind
        row = AlignedRow(
            left_row=None,
            right_row=None,
            merged_row=merged_row,
            status="deleted",
            reason="三方均视为删除",
            base_row=base_row,
            conflict_kind=CONFLICT_KIND_TWO_WAY,
        )
        return row

    # Base exists, one side deleted -> conflict if the other side modified it,
    # otherwise treat as clean delete.
    if base_row is not None and (left_row is None or right_row is None):
        return _build_three_way_deletion_row(
            base_row, left_row, right_row, columns, merge_rule, comparison_options
        )

    anchor_left = left_row if left_row is not None else RowData(row_index=0)
    anchor_right = right_row if right_row is not None else RowData(row_index=0)
    merged_values: dict[int, str] = {}
    conflict_columns: set[int] = set()
    column_conflict_kinds: dict[int, str] = {}
    worst_kind = CONFLICT_KIND_TWO_WAY
    has_divergence = False
    auto_resolved_count = 0

    for index, binding in enumerate(columns, start=1):
        left_value = anchor_left.value_at(binding.left_index or -1) if binding.left_index else ""
        right_value = anchor_right.value_at(binding.right_index or -1) if binding.right_index else ""
        base_value = base_row.value_at(binding.base_index or -1) if base_row is not None and binding.base_index else ""

        if not is_binding_included(binding, merge_rule):
            merged_values[index] = ""
            continue

        left_eq_right = compare_text_values(left_value, right_value, comparison_options)
        base_eq_left = compare_text_values(base_value, left_value, comparison_options)
        base_eq_right = compare_text_values(base_value, right_value, comparison_options)

        if left_eq_right:
            merged_values[index] = left_value
            if not base_eq_left:
                column_conflict_kinds[index] = CONFLICT_KIND_THREE_WAY_SAME_EDIT
                worst_kind = _worse_kind(worst_kind, CONFLICT_KIND_THREE_WAY_SAME_EDIT)
                auto_resolved_count += 1
            continue

        if base_eq_left and not base_eq_right:
            merged_values[index] = right_value
            column_conflict_kinds[index] = CONFLICT_KIND_THREE_WAY_RIGHT_MODIFIED
            worst_kind = _worse_kind(worst_kind, CONFLICT_KIND_THREE_WAY_RIGHT_MODIFIED)
            auto_resolved_count += 1
            continue

        if base_eq_right and not base_eq_left:
            merged_values[index] = left_value
            column_conflict_kinds[index] = CONFLICT_KIND_THREE_WAY_LEFT_MODIFIED
            worst_kind = _worse_kind(worst_kind, CONFLICT_KIND_THREE_WAY_LEFT_MODIFIED)
            auto_resolved_count += 1
            continue

        # Diverged: base/left/right all differ (or base missing and left != right).
        has_divergence = True
        column_conflict_kinds[index] = CONFLICT_KIND_THREE_WAY_DIVERGE
        worst_kind = CONFLICT_KIND_THREE_WAY_DIVERGE
        if merge_rule.conflict_resolution == "blank":
            merged_values[index] = ""
        elif merge_rule.conflict_resolution == "left":
            merged_values[index] = left_value or (right_value if merge_rule.fill_missing_from_other_side else "")
        else:
            merged_values[index] = right_value or (left_value if merge_rule.fill_missing_from_other_side else "")
        conflict_columns.add(index)

    if has_divergence:
        status = "conflict"
        if merge_rule.conflict_resolution == "blank":
            reason = "三方分歧，默认留空，请人工确认"
        else:
            reason = "三方分歧，默认取{0}，请人工确认".format(
                "左侧" if merge_rule.conflict_resolution == "left" else "右侧"
            )
    elif auto_resolved_count > 0:
        status = "same"
        reason = "三方合并，单侧修改已自动采纳"
    else:
        status = "same"
        reason = "三方一致"

    merged_row = clone_row_with_values(anchor_left, merged_values)
    merged_row.kind = (
        anchor_left.kind
        if anchor_left.kind != "blank"
        else (anchor_right.kind if anchor_right.kind != "blank" else (base_row.kind if base_row else "data"))
    )
    return AlignedRow(
        left_row=left_row,
        right_row=right_row,
        merged_row=merged_row,
        status=status,
        reason=reason,
        conflict_columns=conflict_columns,
        base_row=base_row,
        conflict_kind=worst_kind,
        column_conflict_kinds=column_conflict_kinds,
    )


_CONFLICT_KIND_PRIORITY = {
    CONFLICT_KIND_TWO_WAY: 0,
    CONFLICT_KIND_THREE_WAY_SAME_EDIT: 1,
    CONFLICT_KIND_THREE_WAY_LEFT_MODIFIED: 2,
    CONFLICT_KIND_THREE_WAY_RIGHT_MODIFIED: 2,
    CONFLICT_KIND_THREE_WAY_DIVERGE: 3,
}


def _worse_kind(current: str, candidate: str) -> str:
    if _CONFLICT_KIND_PRIORITY.get(candidate, 0) > _CONFLICT_KIND_PRIORITY.get(current, 0):
        return candidate
    return current


def _build_three_way_deletion_row(
    base_row: RowData,
    left_row: RowData | None,
    right_row: RowData | None,
    columns: list[ColumnBinding],
    merge_rule: MergeRule,
    comparison_options: ComparisonOptions | None,
) -> AlignedRow:
    # Exactly one of left/right is None here; the surviving side may have
    # modified values compared to base.
    surviving: RowData | None = left_row if left_row is not None else right_row
    assert surviving is not None
    deleted_side = "左侧" if left_row is None else "右侧"
    surviving_side = "右侧" if left_row is None else "左侧"
    surviving_is_left = right_row is None
    conflict_columns: set[int] = set()
    column_conflict_kinds: dict[int, str] = {}
    merged_values: dict[int, str] = {}
    has_survivor_edit = False

    for index, binding in enumerate(columns, start=1):
        base_value = base_row.value_at(binding.base_index or -1) if binding.base_index else ""
        if surviving_is_left:
            surv_value = surviving.value_at(binding.left_index or -1) if binding.left_index else ""
        else:
            surv_value = surviving.value_at(binding.right_index or -1) if binding.right_index else ""
        if not is_binding_included(binding, merge_rule):
            merged_values[index] = ""
            continue
        if not compare_text_values(base_value, surv_value, comparison_options):
            has_survivor_edit = True
            conflict_columns.add(index)
            column_conflict_kinds[index] = CONFLICT_KIND_THREE_WAY_DIVERGE
        merged_values[index] = surv_value  # keep surviving side visible for conflict rendering

    if has_survivor_edit:
        status = "conflict"
        reason = f"{deleted_side}删除、{surviving_side}修改，三方冲突需人工确认"
        worst_kind = CONFLICT_KIND_THREE_WAY_DIVERGE
    else:
        status = "deleted"
        reason = f"{deleted_side}已删除，{surviving_side}未修改，按规则视为删除"
        worst_kind = CONFLICT_KIND_TWO_WAY

    anchor = surviving or base_row
    merged_row = clone_row_with_values(anchor, merged_values)
    merged_row.kind = anchor.kind
    return AlignedRow(
        left_row=left_row,
        right_row=right_row,
        merged_row=merged_row,
        status=status,
        reason=reason,
        conflict_columns=conflict_columns,
        base_row=base_row,
        conflict_kind=worst_kind,
        column_conflict_kinds=column_conflict_kinds,
    )


def align_sheets_three_way(
    base_sheet: SheetData | None,
    left_sheet: SheetData | None,
    right_sheet: SheetData | None,
    merge_rule: MergeRule | None = None,
    comparison_options: ComparisonOptions | None = None,
    preferred_key_fields: list[str] | None = None,
    strict_single_key: bool = False,
) -> SheetAlignment:
    """Align three sheets using shared key fields.

    Falls back to :func:`align_sheets` (two-way) when ``base_sheet`` is absent
    or when no key fields can be resolved across the three sides — three-way
    merging without reliable keys is not safe.
    """
    merge_rule = get_merge_rule(merge_rule.rule_id if merge_rule is not None else None)
    comparison_options = comparison_options or ComparisonOptions()

    if base_sheet is None:
        return align_sheets(
            left_sheet,
            right_sheet,
            merge_rule=merge_rule,
            comparison_options=comparison_options,
            preferred_key_fields=preferred_key_fields,
            strict_single_key=strict_single_key,
        )

    # Resolve key fields that exist in at least one side so we can anchor
    # three-way alignment. Fall back to two-way only when no side can supply
    # a key (e.g. all sides empty or no recognisable id columns).
    if strict_single_key:
        shared_keys = _resolve_strict_three_way_key_fields(
            base_sheet, left_sheet, right_sheet, preferred_key_fields
        )
        _validate_strict_unique_key(base_sheet, shared_keys, "Base", comparison_options)
        _validate_strict_unique_key(left_sheet, shared_keys, "左侧", comparison_options)
        _validate_strict_unique_key(right_sheet, shared_keys, "右侧", comparison_options)
    else:
        shared_keys = _resolve_three_way_keys(
            base_sheet, left_sheet, right_sheet, preferred_key_fields
        )
    if not shared_keys:
        return align_sheets(
            left_sheet,
            right_sheet,
            merge_rule=merge_rule,
            comparison_options=comparison_options,
            preferred_key_fields=preferred_key_fields,
            strict_single_key=strict_single_key,
        )

    columns = build_column_bindings(left_sheet, right_sheet, base_sheet)
    sheet_name = (
        (left_sheet.name if left_sheet is not None else None)
        or (right_sheet.name if right_sheet is not None else None)
        or base_sheet.name
    )

    base_rows = base_sheet.rows
    left_rows = left_sheet.rows if left_sheet is not None else []
    right_rows = right_sheet.rows if right_sheet is not None else []

    base_resolver = _make_row_map_resolver(base_sheet)
    left_resolver = _make_row_map_resolver(left_sheet)
    right_resolver = _make_row_map_resolver(right_sheet)

    def key_of(sheet: SheetData | None, row: RowData, resolver) -> tuple[str, ...] | None:
        if sheet is None or row.kind != "data":
            return None
        logical_map = resolver(row)
        values = tuple(
            normalize_compare_text(logical_map.get(field, ""), comparison_options) for field in shared_keys
        )
        if not all(values):
            return None
        return values

    base_index: dict[tuple[str, ...], RowData] = {}
    for row in base_rows:
        key = key_of(base_sheet, row, base_resolver)
        if key is not None and key not in base_index:
            base_index[key] = row
    left_index: dict[tuple[str, ...], RowData] = {}
    for row in left_rows:
        key = key_of(left_sheet, row, left_resolver)
        if key is not None and key not in left_index:
            left_index[key] = row
    right_index: dict[tuple[str, ...], RowData] = {}
    for row in right_rows:
        key = key_of(right_sheet, row, right_resolver)
        if key is not None and key not in right_index:
            right_index[key] = row

    aligned_rows: list[AlignedRow] = []
    seen_keys: set[tuple[str, ...]] = set()

    # Walk left rows first to preserve user-visible left ordering; then any
    # base-only/right-only keys are appended in their native order.
    for row in left_rows:
        key = key_of(left_sheet, row, left_resolver)
        if key is None:
            # Structural row (header/blank/group) with no key: carry left-side
            # intact, fall back to the existing 2-way builder so status stays
            # consistent.
            aligned_rows.append(make_aligned_row(row, None, columns, merge_rule, comparison_options))
            continue
        if key in seen_keys:
            continue
        seen_keys.add(key)
        aligned_rows.append(
            make_aligned_row_three_way(
                base_index.get(key),
                row,
                right_index.get(key),
                columns,
                merge_rule,
                comparison_options,
            )
        )

    for row in right_rows:
        key = key_of(right_sheet, row, right_resolver)
        if key is None or key in seen_keys:
            continue
        seen_keys.add(key)
        aligned_rows.append(
            make_aligned_row_three_way(
                base_index.get(key),
                left_index.get(key),
                row,
                columns,
                merge_rule,
                comparison_options,
            )
        )

    for row in base_rows:
        key = key_of(base_sheet, row, base_resolver)
        if key is None or key in seen_keys:
            continue
        seen_keys.add(key)
        aligned_rows.append(
            make_aligned_row_three_way(
                row,
                left_index.get(key),
                right_index.get(key),
                columns,
                merge_rule,
                comparison_options,
            )
        )

    return SheetAlignment(
        sheet_name=sheet_name,
        columns=columns,
        rows=aligned_rows,
        left_sheet=left_sheet,
        right_sheet=right_sheet,
        key_fields=shared_keys,
        base_sheet=base_sheet,
    )


def align_replace_block(
    left_sheet: SheetData | None,
    right_sheet: SheetData | None,
    left_rows: list[RowData],
    right_rows: list[RowData],
    columns: list[ColumnBinding],
    key_fields: list[str],
    merge_rule: MergeRule,
    comparison_options: ComparisonOptions | None = None,
    *,
    left_logical_map_resolver=None,
    right_logical_map_resolver=None,
) -> list[AlignedRow]:
    left_count = len(left_rows)
    right_count = len(right_rows)
    if left_logical_map_resolver is None:
        left_logical_map_resolver = _make_row_map_resolver(left_sheet)
    if right_logical_map_resolver is None:
        right_logical_map_resolver = _make_row_map_resolver(right_sheet)

    keyed_path = _try_align_by_keys(
        left_rows,
        right_rows,
        columns,
        key_fields,
        merge_rule,
        comparison_options,
        left_logical_map_resolver,
        right_logical_map_resolver,
    )
    if keyed_path is not None:
        return keyed_path

    scores = [
        [
            row_match_score(
                left_sheet,
                left_row,
                right_sheet,
                right_row,
                key_fields,
                comparison_options,
                left_logical_map=left_logical_map_resolver(left_row),
                right_logical_map=right_logical_map_resolver(right_row),
            )
            for right_row in right_rows
        ]
        for left_row in left_rows
    ]

    dp = [[0.0 for _ in range(right_count + 1)] for _ in range(left_count + 1)]
    choice = [["" for _ in range(right_count + 1)] for _ in range(left_count + 1)]

    for left_index in range(left_count - 1, -1, -1):
        for right_index in range(right_count - 1, -1, -1):
            best_score = dp[left_index + 1][right_index]
            best_choice = "skip_left"

            if dp[left_index][right_index + 1] > best_score:
                best_score = dp[left_index][right_index + 1]
                best_choice = "skip_right"

            pair_score = scores[left_index][right_index]
            if pair_score >= MATCH_THRESHOLD:
                candidate_score = pair_score + dp[left_index + 1][right_index + 1]
                if candidate_score > best_score:
                    best_score = candidate_score
                    best_choice = "pair"

            dp[left_index][right_index] = best_score
            choice[left_index][right_index] = best_choice

    aligned_rows: list[AlignedRow] = []
    left_index = 0
    right_index = 0
    while left_index < left_count and right_index < right_count:
        current_choice = choice[left_index][right_index]
        if current_choice == "pair":
            aligned_rows.append(
                make_aligned_row(left_rows[left_index], right_rows[right_index], columns, merge_rule, comparison_options)
            )
            left_index += 1
            right_index += 1
        elif current_choice == "skip_right":
            aligned_rows.append(make_aligned_row(None, right_rows[right_index], columns, merge_rule, comparison_options))
            right_index += 1
        else:
            aligned_rows.append(make_aligned_row(left_rows[left_index], None, columns, merge_rule, comparison_options))
            left_index += 1

    while left_index < left_count:
        aligned_rows.append(make_aligned_row(left_rows[left_index], None, columns, merge_rule, comparison_options))
        left_index += 1

    while right_index < right_count:
        aligned_rows.append(make_aligned_row(None, right_rows[right_index], columns, merge_rule, comparison_options))
        right_index += 1

    return aligned_rows


def _try_align_by_keys(
    left_rows: list[RowData],
    right_rows: list[RowData],
    columns: list[ColumnBinding],
    key_fields: list[str],
    merge_rule: MergeRule,
    comparison_options: ComparisonOptions | None,
    left_logical_map_resolver,
    right_logical_map_resolver,
) -> list[AlignedRow] | None:
    if not key_fields:
        return None

    def collect_keys(rows: list[RowData], resolver) -> list[tuple[str, ...]] | None:
        keys: list[tuple[str, ...]] = []
        for row in rows:
            if row.kind in STRUCTURAL_ROW_KINDS:
                return None
            logical_map = resolver(row)
            tuple_value = tuple(
                normalize_compare_text(logical_map.get(field, ""), comparison_options)
                for field in key_fields
            )
            if not all(tuple_value):
                return None
            keys.append(tuple_value)
        return keys

    left_keys = collect_keys(left_rows, left_logical_map_resolver)
    if left_keys is None:
        return None
    right_keys = collect_keys(right_rows, right_logical_map_resolver)
    if right_keys is None:
        return None

    matcher = SequenceMatcher(a=left_keys, b=right_keys, autojunk=False)
    aligned_rows: list[AlignedRow] = []
    for opcode, left_a, left_b, right_a, right_b in matcher.get_opcodes():
        if opcode == "equal":
            for offset in range(left_b - left_a):
                aligned_rows.append(
                    make_aligned_row(
                        left_rows[left_a + offset],
                        right_rows[right_a + offset],
                        columns,
                        merge_rule,
                        comparison_options,
                    )
                )
        elif opcode == "delete":
            for left_row in left_rows[left_a:left_b]:
                aligned_rows.append(make_aligned_row(left_row, None, columns, merge_rule, comparison_options))
        elif opcode == "insert":
            for right_row in right_rows[right_a:right_b]:
                aligned_rows.append(make_aligned_row(None, right_row, columns, merge_rule, comparison_options))
        else:  # replace — keys differ across the block, so emit each side independently
            for left_row in left_rows[left_a:left_b]:
                aligned_rows.append(make_aligned_row(left_row, None, columns, merge_rule, comparison_options))
            for right_row in right_rows[right_a:right_b]:
                aligned_rows.append(make_aligned_row(None, right_row, columns, merge_rule, comparison_options))
    return aligned_rows


def _build_single_side_row(
    left_row: RowData | None,
    right_row: RowData | None,
    columns: list[ColumnBinding],
    keep_row: bool,
    keep_reason: str,
    drop_reason: str,
) -> AlignedRow:
    source_row = left_row or right_row or RowData(row_index=0)
    status = "left_only" if right_row is None else "right_only"
    if not keep_row:
        merged_row = clone_row_with_values(source_row, {})
        merged_row.cells = []
        merged_row.kind = "blank"
        return AlignedRow(left_row, right_row, merged_row, "deleted", drop_reason)

    values: dict[int, str] = {}
    for index, binding in enumerate(columns, start=1):
        source_index = binding.left_index if left_row is not None else binding.right_index
        values[index] = source_row.value_at(source_index or -1) if source_index else ""
    merged_row = clone_row_with_values(source_row, values)
    merged_row.kind = source_row.kind
    return AlignedRow(left_row, right_row, merged_row, status, keep_reason)


def row_match_score(
    left_sheet: SheetData | None,
    left_row: RowData,
    right_sheet: SheetData | None,
    right_row: RowData,
    key_fields: list[str],
    comparison_options: ComparisonOptions | None = None,
    *,
    left_logical_map: dict[str, str] | None = None,
    right_logical_map: dict[str, str] | None = None,
) -> float:
    if left_sheet is None or right_sheet is None:
        return 0.0
    if left_row.kind in STRUCTURAL_ROW_KINDS or right_row.kind in STRUCTURAL_ROW_KINDS:
        if left_row.kind != right_row.kind:
            return 0.0
        return 1.0 if compare_text_values(compact_text(left_row, comparison_options), compact_text(right_row, comparison_options)) else 0.0

    left_map = left_logical_map if left_logical_map is not None else row_logical_map(left_sheet, left_row)
    right_map = right_logical_map if right_logical_map is not None else row_logical_map(right_sheet, right_row)
    left_keys = tuple(normalize_compare_text(left_map.get(field, ""), comparison_options) for field in key_fields)
    right_keys = tuple(normalize_compare_text(right_map.get(field, ""), comparison_options) for field in key_fields)
    left_has_key = bool(key_fields) and all(left_keys)
    right_has_key = bool(key_fields) and all(right_keys)

    if left_has_key and right_has_key:
        return 1.0 if left_keys == right_keys else 0.0

    score = 0.0
    left_name = pick_identity_value(left_map, "name", "remark", "title", "skill_name")
    right_name = pick_identity_value(right_map, "name", "remark", "title", "skill_name")
    if left_name and right_name:
        if compare_text_values(left_name, right_name, comparison_options):
            score += 0.7
        else:
            score += SequenceMatcher(
                None,
                normalize_compare_text(left_name, comparison_options),
                normalize_compare_text(right_name, comparison_options),
            ).ratio() * 0.35

    left_level = pick_identity_value(left_map, "level")
    right_level = pick_identity_value(right_map, "level")
    if left_level and right_level and compare_text_values(left_level, right_level, comparison_options):
        score += 0.1

    left_type = pick_identity_value(left_map, "type", "show_type")
    right_type = pick_identity_value(right_map, "type", "show_type")
    if left_type and right_type and compare_text_values(left_type, right_type, comparison_options):
        score += 0.1

    common_fields = sorted(set(left_map) & set(right_map))
    comparable_fields = [
        field
        for field in common_fields
        if field
        and "id" not in field
        and normalize_compare_text(left_map.get(field, ""), comparison_options)
        and normalize_compare_text(right_map.get(field, ""), comparison_options)
    ]
    if comparable_fields:
        exact_matches = sum(
            1 for field in comparable_fields if compare_text_values(left_map[field], right_map[field], comparison_options)
        )
        score += min(0.2, exact_matches / len(comparable_fields) * 0.2)

    compact_score = SequenceMatcher(
        None,
        compact_text(left_row, comparison_options),
        compact_text(right_row, comparison_options),
    ).ratio()
    score += compact_score * 0.15

    if left_has_key != right_has_key:
        score = min(score, 0.95)

    return min(score, 1.0)


def row_logical_map(sheet: SheetData, row: RowData) -> dict[str, str]:
    headers = _sheet_logical_columns(sheet)
    return {logical: row.value_at(index) for index, logical in headers}


def _sheet_logical_columns(sheet: SheetData) -> list[tuple[int, str]]:
    return [
        (
            index,
            normalize_header(
                sheet.logical_headers[index - 1] if index - 1 < len(sheet.logical_headers) else f"col_{index}"
            ),
        )
        for index in range(1, sheet.max_columns + 1)
    ]


def _make_row_map_resolver(sheet: SheetData | None):
    if sheet is None:
        empty: dict[str, str] = {}
        return lambda row: empty
    headers = _sheet_logical_columns(sheet)
    cache: dict[int, dict[str, str]] = {}

    def resolve(row: RowData) -> dict[str, str]:
        key = id(row)
        cached = cache.get(key)
        if cached is not None:
            return cached
        result = {logical: row.value_at(index) for index, logical in headers}
        cache[key] = result
        return result

    return resolve


def compact_text(row: RowData, comparison_options: ComparisonOptions | None = None) -> str:
    values = [
        normalize_compare_text(cell.value, comparison_options)
        for cell in row.cells
        if normalize_compare_text(cell.value, comparison_options)
    ]
    return " | ".join(values)


def content_fingerprint(sheet: SheetData, row: RowData, comparison_options: ComparisonOptions | None = None) -> str:
    values: list[str] = []
    for index in range(1, sheet.max_columns + 1):
        value = normalize_compare_text(row.value_at(index), comparison_options)
        if value:
            header = normalize_header(sheet.logical_headers[index - 1] if index - 1 < len(sheet.logical_headers) else f"col_{index}")
            values.append(f"{header}={value}")
    return " | ".join(values[:8])


_FULLWIDTH_TRANS = {0x3000: 0x20}
for _fw in range(0xFF01, 0xFF5F):
    _FULLWIDTH_TRANS[_fw] = _fw - 0xFEE0


_DEFAULT_OPTIONS = ComparisonOptions()
_WS_RE = re.compile(r"\s+")


@lru_cache(maxsize=65536)
def _normalize_compare_cached(text: str, options: ComparisonOptions) -> str:
    if options.normalize_fullwidth:
        text = text.translate(_FULLWIDTH_TRANS)
    if options.ignore_all_whitespace:
        text = _WS_RE.sub("", text)
    elif options.ignore_trim_whitespace:
        text = _WS_RE.sub(" ", text.strip())
    if options.ignore_case:
        text = text.casefold()
    return text


def normalize_compare_text(value: str, comparison_options: ComparisonOptions | None = None) -> str:
    text = value if isinstance(value, str) else str(value or "")
    if comparison_options is None or comparison_options == _DEFAULT_OPTIONS:
        return text
    return _normalize_compare_cached(text, comparison_options)


_NUMBER_RE = re.compile(r"^-?\d+(?:\.\d+)?$")


def compare_text_values(left_value: str, right_value: str, comparison_options: ComparisonOptions | None = None) -> bool:
    left_norm = normalize_compare_text(left_value, comparison_options)
    right_norm = normalize_compare_text(right_value, comparison_options)
    if left_norm == right_norm:
        return True
    if comparison_options is not None and comparison_options.numeric_tolerance > 0:
        if _NUMBER_RE.match(left_norm) and _NUMBER_RE.match(right_norm):
            try:
                return abs(float(left_norm) - float(right_norm)) <= comparison_options.numeric_tolerance
            except (TypeError, ValueError):
                return False
    return False


_HEADER_NON_WORD_RE = re.compile(r"[^\w\u4e00-\u9fff]+")
_HEADER_UNDERSCORE_RE = re.compile(r"_+")


@lru_cache(maxsize=8192)
def normalize_header(value: str) -> str:
    text = value.strip().lower()
    text = _HEADER_NON_WORD_RE.sub("_", text)
    text = _HEADER_UNDERSCORE_RE.sub("_", text).strip("_")
    return text or "unknown"


def pick_identity_value(logical_map: dict[str, str], *keys: str) -> str:
    for key in keys:
        value = logical_map.get(key, "").strip()
        if value:
            return value
    for map_key, value in logical_map.items():
        if any(token in map_key for token in keys) and value.strip():
            return value.strip()
    return ""
