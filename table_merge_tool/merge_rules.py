from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True, slots=True)
class MergeRule:
    rule_id: str
    title: str
    summary: str
    conflict_resolution: str
    fill_missing_from_other_side: bool
    keep_left_only_rows: bool
    keep_right_only_rows: bool
    keep_left_only_columns: bool
    keep_right_only_columns: bool
    keep_left_only_sheets: bool
    keep_right_only_sheets: bool
    preserve_existing_rows: bool = False


DEFAULT_RULE_ID = "full_keep_left"

MERGE_RULES: dict[str, MergeRule] = {
    "full_keep_left": MergeRule(
        rule_id="full_keep_left",
        title="全合-冲突保左",
        summary="两侧新增都保留，字段冲突默认取左侧值。",
        conflict_resolution="left",
        fill_missing_from_other_side=True,
        keep_left_only_rows=True,
        keep_right_only_rows=True,
        keep_left_only_columns=True,
        keep_right_only_columns=True,
        keep_left_only_sheets=True,
        keep_right_only_sheets=True,
    ),
    "full_keep_right": MergeRule(
        rule_id="full_keep_right",
        title="全合-冲突保右",
        summary="两侧新增都保留，字段冲突默认取右侧值。",
        conflict_resolution="right",
        fill_missing_from_other_side=True,
        keep_left_only_rows=True,
        keep_right_only_rows=True,
        keep_left_only_columns=True,
        keep_right_only_columns=True,
        keep_left_only_sheets=True,
        keep_right_only_sheets=True,
    ),
    "left_priority": MergeRule(
        rule_id="left_priority",
        title="左侧优先",
        summary="以左侧为主，忽略右侧新增的行、列、子表；共享字段仍可用右侧补空。",
        conflict_resolution="left",
        fill_missing_from_other_side=True,
        keep_left_only_rows=True,
        keep_right_only_rows=False,
        keep_left_only_columns=True,
        keep_right_only_columns=False,
        keep_left_only_sheets=True,
        keep_right_only_sheets=False,
    ),
    "right_priority": MergeRule(
        rule_id="right_priority",
        title="右侧优先",
        summary="以右侧为主，忽略左侧新增的行、列、子表；共享字段仍可用左侧补空。",
        conflict_resolution="right",
        fill_missing_from_other_side=True,
        keep_left_only_rows=False,
        keep_right_only_rows=True,
        keep_left_only_columns=False,
        keep_right_only_columns=True,
        keep_left_only_sheets=False,
        keep_right_only_sheets=True,
    ),
    "complete_left": MergeRule(
        rule_id="complete_left",
        title="完全左",
        summary="中间结果完全以左侧为基底，不自动带入右侧新增行、列、子表，也不自动补右侧非空字段。",
        conflict_resolution="left",
        fill_missing_from_other_side=False,
        keep_left_only_rows=True,
        keep_right_only_rows=False,
        keep_left_only_columns=True,
        keep_right_only_columns=False,
        keep_left_only_sheets=True,
        keep_right_only_sheets=False,
    ),
    "complete_right": MergeRule(
        rule_id="complete_right",
        title="完全右",
        summary="中间结果完全以右侧为基底，不自动带入左侧新增行、列、子表，也不自动补左侧非空字段。",
        conflict_resolution="right",
        fill_missing_from_other_side=False,
        keep_left_only_rows=False,
        keep_right_only_rows=True,
        keep_left_only_columns=False,
        keep_right_only_columns=True,
        keep_left_only_sheets=False,
        keep_right_only_sheets=True,
    ),
    "fill_empty_left": MergeRule(
        rule_id="fill_empty_left",
        title="只补空-以左为主",
        summary="以左侧为主；左侧为空且右侧有值时才补右侧，忽略右侧新增行、列、子表。",
        conflict_resolution="left",
        fill_missing_from_other_side=True,
        keep_left_only_rows=True,
        keep_right_only_rows=False,
        keep_left_only_columns=True,
        keep_right_only_columns=False,
        keep_left_only_sheets=True,
        keep_right_only_sheets=False,
    ),
    "fill_empty_right": MergeRule(
        rule_id="fill_empty_right",
        title="只补空-以右为主",
        summary="以右侧为主；右侧为空且左侧有值时才补左侧，忽略左侧新增行、列、子表。",
        conflict_resolution="right",
        fill_missing_from_other_side=True,
        keep_left_only_rows=False,
        keep_right_only_rows=True,
        keep_left_only_columns=False,
        keep_right_only_columns=True,
        keep_left_only_sheets=False,
        keep_right_only_sheets=True,
    ),
    "full_conflict_blank": MergeRule(
        rule_id="full_conflict_blank",
        title="全合-冲突留空",
        summary="两侧新增都保留，空值仍可互补；真正冲突的字段留空并要求人工确认。",
        conflict_resolution="blank",
        fill_missing_from_other_side=True,
        keep_left_only_rows=True,
        keep_right_only_rows=True,
        keep_left_only_columns=True,
        keep_right_only_columns=True,
        keep_left_only_sheets=True,
        keep_right_only_sheets=True,
    ),
    "append_rows_keep_existing": MergeRule(
        rule_id="append_rows_keep_existing",
        title="只合新增行-不改已有行",
        summary="已有行完全保持左侧结果，不自动补空或处理冲突；只追加右侧新增行，不追加右侧新增列、子表。",
        conflict_resolution="left",
        fill_missing_from_other_side=False,
        keep_left_only_rows=True,
        keep_right_only_rows=True,
        keep_left_only_columns=True,
        keep_right_only_columns=False,
        keep_left_only_sheets=True,
        keep_right_only_sheets=False,
        preserve_existing_rows=True,
    ),
}


def list_merge_rules() -> list[MergeRule]:
    return list(MERGE_RULES.values())


def get_merge_rule(rule_id: str | None) -> MergeRule:
    if rule_id and rule_id in MERGE_RULES:
        return MERGE_RULES[rule_id]
    return MERGE_RULES[DEFAULT_RULE_ID]


def binding_side(binding) -> str:
    if binding.left_index is not None and binding.right_index is not None:
        return "shared"
    if binding.left_index is not None:
        return "left_only"
    return "right_only"


def is_binding_included(binding, merge_rule: MergeRule) -> bool:
    side = binding_side(binding)
    if side == "shared":
        return True
    if side == "left_only":
        return merge_rule.keep_left_only_columns
    return merge_rule.keep_right_only_columns
