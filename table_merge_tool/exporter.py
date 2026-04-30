from __future__ import annotations

import copy
import csv
import html
from pathlib import Path

from lxml import etree as LET

from .alignment import ComparisonOptions, compare_text_values
from .merge_rules import MergeRule, is_binding_included
from .models import CellData, RowData, SS, SheetAlignment, WorkbookData, XML_NS
from .text_diff import (
    compute_char_diff,
    compute_inline_diff,
    render_diff_html,
    render_diff_plain,
)


NS = {"ss": XML_NS}
LXML_NS = {
    "ss": "urn:schemas-microsoft-com:office:spreadsheet",
    "o": "urn:schemas-microsoft-com:office:office",
    "x": "urn:schemas-microsoft-com:office:excel",
}


def export_workbook(
    base_workbook: WorkbookData,
    right_workbook: WorkbookData | None,
    alignments: dict[str, SheetAlignment],
    output_path: str | Path,
    merge_rule: MergeRule,
) -> None:
    root = LET.fromstring(base_workbook.xml_bytes)
    tree = LET.ElementTree(root)
    base_sheet_map = base_workbook.sheet_map
    right_sheet_map = right_workbook.sheet_map if right_workbook is not None else {}
    right_tree = LET.ElementTree(LET.fromstring(right_workbook.xml_bytes)) if right_workbook is not None else None
    styles_node = _get_or_create_styles_node(root)
    left_style_map = _style_map(styles_node)
    right_styles_node = _get_or_create_styles_node(right_tree.getroot()) if right_tree is not None else None
    right_style_map = _style_map(right_styles_node) if right_styles_node is not None else {}
    right_worksheets = {
        node.attrib.get(f"{{{XML_NS}}}Name", f"Sheet{index}"): node
        for index, node in enumerate(right_tree.xpath("//ss:Worksheet", namespaces=LXML_NS), start=1)
    } if right_tree is not None else {}

    worksheet_by_name = {
        node.attrib.get(f"{{{XML_NS}}}Name", f"Sheet{index}"): node
        for index, node in enumerate(root.xpath("//ss:Worksheet", namespaces=LXML_NS), start=1)
    }

    for sheet_name, alignment in alignments.items():
        worksheet = worksheet_by_name.get(sheet_name)
        if worksheet is not None and not _should_export_sheet(alignment, merge_rule):
            root.remove(worksheet)
            worksheet_by_name.pop(sheet_name, None)
            continue
        if sheet_name not in worksheet_by_name and sheet_name in right_worksheets and _should_export_sheet(alignment, merge_rule):
            cloned_sheet = copy.deepcopy(right_worksheets[sheet_name])
            _remap_sheet_styles(cloned_sheet, styles_node, left_style_map, right_style_map)
            root.append(cloned_sheet)
            worksheet_by_name[sheet_name] = cloned_sheet

    for sheet_name, alignment in alignments.items():
        if not _should_export_sheet(alignment, merge_rule):
            continue
        worksheet = worksheet_by_name.get(sheet_name)
        if worksheet is None:
            continue

        table_matches = worksheet.xpath("./ss:Table", namespaces=LXML_NS)
        table = table_matches[0] if table_matches else None
        if table is None:
            continue

        existing_columns = list(table.xpath("./ss:Column", namespaces=LXML_NS))
        for node in list(table.xpath("./ss:Row", namespaces=LXML_NS)):
            table.remove(node)
        for node in list(table.xpath("./ss:Column", namespaces=LXML_NS)):
            table.remove(node)

        included_columns = _included_columns(alignment, merge_rule)
        export_rows = [
            _build_export_row(aligned_row, included_columns)
            for aligned_row in alignment.rows
            if aligned_row.status != "deleted"
        ]
        final_column_count = max(len(included_columns), 1)

        for index in range(max(final_column_count, 1)):
            if index < len(existing_columns):
                table.append(existing_columns[index])
            else:
                table.append(_clone_column(existing_columns[-1]) if existing_columns else LET.Element(LET.QName(XML_NS, "Column")))

        table.set(LET.QName(XML_NS, "ExpandedRowCount"), str(len(export_rows)))
        table.set(LET.QName(XML_NS, "ExpandedColumnCount"), str(final_column_count))

        for row_index, row in enumerate(export_rows, start=1):
            row.row_index = row_index
            table.append(row_to_lxml(row))

    _strip_formulas_from_tree(root)

    path = Path(output_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    tree.write(str(path), encoding="UTF-8", xml_declaration=True, standalone=True)


def export_diff_report(
    alignments: dict[str, SheetAlignment],
    output_path: str | Path,
    comparison_options: ComparisonOptions | None = None,
) -> None:
    path = Path(output_path)
    rows = build_diff_report_rows(alignments, comparison_options)
    path.parent.mkdir(parents=True, exist_ok=True)
    suffix = path.suffix.lower()
    if suffix == ".txt":
        path.write_text(format_diff_report_text(rows), encoding="utf-8")
        return
    if suffix == ".html":
        path.write_text(format_diff_report_html(rows, comparison_options), encoding="utf-8")
        return
    if suffix == ".md":
        path.write_text(format_diff_report_markdown(rows, comparison_options), encoding="utf-8")
        return
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(
            handle,
            fieldnames=[
                "sheet_name",
                "alignment_row",
                "row_status",
                "row_conflict_kind",
                "row_reason",
                "row_note",
                "left_row",
                "right_row",
                "base_row",
                "field_key",
                "field_title",
                "first_diff_char",
                "left_value",
                "right_value",
                "base_value",
                "left_vs_right_inline_plain",
            ],
            extrasaction="ignore",
        )
        writer.writeheader()
        writer.writerows(rows)


def build_diff_report_rows(
    alignments: dict[str, SheetAlignment],
    comparison_options: ComparisonOptions | None = None,
) -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []
    for sheet_name in sorted(alignments):
        alignment = alignments[sheet_name]
        for alignment_row, row in enumerate(alignment.rows, start=1):
            rows.extend(_aligned_row_diff_entries(alignment, row, alignment_row, comparison_options))
    return rows


def format_diff_report_text(rows: list[dict[str, str]]) -> str:
    if not rows:
        return "差异报告\n\n没有可导出的差异。"

    lines = ["差异报告", "", f"差异项数量: {len(rows)}", ""]
    current_sheet = ""
    current_row = ""
    for item in rows:
        sheet_name = item["sheet_name"]
        row_key = f"{item['sheet_name']}#{item['alignment_row']}"
        has_base = bool(item.get("has_base"))
        if sheet_name != current_sheet:
            if lines and lines[-1] != "":
                lines.append("")
            lines.append(f"[工作表] {sheet_name}")
            current_sheet = sheet_name
            current_row = ""
        if row_key != current_row:
            row_parts = [
                f"- 行 {item['alignment_row']}",
                f"状态: {_row_status_label(item['row_status'])}",
            ]
            if has_base:
                row_parts.append(f"base:{item.get('base_row') or '-'}")
            row_parts.append(f"左:{item['left_row'] or '-'}")
            row_parts.append(f"右:{item['right_row'] or '-'}")
            lines.append(" | ".join(row_parts))
            if item["row_reason"]:
                lines.append(f"  说明: {item['row_reason']}")
            if item["row_note"]:
                lines.append(f"  备注: {item['row_note']}")
            kind = item.get("row_conflict_kind", "")
            if kind and kind != "two_way":
                lines.append(f"  冲突类型: {CONFLICT_KIND_LABELS.get(kind, kind)}")
            current_row = row_key
        lines.append(f"  字段: {item['field_title']} ({item['field_key']})")
        inline_plain = item.get("left_vs_right_inline_plain") or ""
        if inline_plain:
            lines.append(f"    对比: {inline_plain}")
        lines.append(f"    左: {item['left_value']}")
        lines.append(f"    右: {item['right_value']}")
        if has_base:
            lines.append(f"    base: {item.get('base_value', '')}")
    return "\n".join(lines)


ROW_STATUS_LABELS = {
    "same": "相同",
    "conflict": "冲突",
    "left_only": "左独有",
    "right_only": "右独有",
    "deleted": "删除",
}

CONFLICT_KIND_LABELS = {
    "two_way": "两方",
    "three_way_diverge": "三方分歧",
    "three_way_left_modified": "左单方改",
    "three_way_right_modified": "右单方改",
    "three_way_same_edit": "双方同改",
}

DIFF_REPORT_CSS = """

:root {
  --bg-color: #f8fafc;
  --card-bg: #ffffff;
  --border-color: #e2e8f0;
  --text-main: #334155;
  --text-muted: #64748b;
  --accent: #0ea5e9;
  --conflict-bg: #fee2e2;
  --conflict-text: #991b1b;
  --left-bg: #e0f2fe;
  --left-text: #075985;
  --right-bg: #dcfce7;
  --right-text: #166534;
  --del-bg: #f1f5f9;
  --del-text: #475569;
}
body { background-color: var(--bg-color); color: var(--text-main); font-family: system-ui, -apple-system, sans-serif; margin: 0; }
.page { max-width: 1480px; margin: 0 auto; padding: 24px 32px 48px; }
h1 { font-size: 26px; margin: 0 0 20px; font-weight: 700; color: #0f172a; display: flex; justify-content: space-between; align-items: center; }
.toolbar { display: flex; gap: 8px; font-size: 14px; }
.toolbar button { background: #fff; border: 1px solid var(--border-color); border-radius: 6px; padding: 6px 12px; cursor: pointer; color: var(--text-main); font-weight: 500; font-family: inherit; box-shadow: 0 1px 2px rgba(0,0,0,0.05); }
.toolbar button:hover { background: #f1f5f9; }
h2 { font-size: 18px; margin: 32px 0 16px; border-left: 4px solid var(--accent); padding-left: 10px; font-weight: 600; color: #1e293b; }
h3 { font-size: 15px; margin: 0; }
section.sheet, details.sheet, section.file, details.file, .overview-panel { background: var(--card-bg); border: 1px solid var(--border-color); border-radius: 8px; padding: 16px 20px; margin-bottom: 20px; box-shadow: 0 1px 3px rgba(0,0,0,0.04); }
details.file > summary, details.sheet > summary, .overview-panel > summary { cursor: pointer; list-style: none; font-weight: 600; outline: none; }
details.file > summary::-webkit-details-marker, details.sheet > summary::-webkit-details-marker, .overview-panel > summary::-webkit-details-marker { display: none; }
table { border-collapse: collapse; width: 100%; margin-top: 12px; background: #fff; font-size: 14px; text-align: left; }
th, td { border: 1px solid var(--border-color); padding: 10px 12px; }
th { background: #f8fafc; font-weight: 600; color: var(--text-muted); }
a { color: var(--accent); text-decoration: none; }
a:hover { text-decoration: underline; }
.badge { display: inline-flex; align-items: center; padding: 2px 10px; margin: 0 6px 4px 0; border-radius: 12px; font-size: 12px; font-weight: 500; }
.badge-conflict { background: var(--conflict-bg); color: var(--conflict-text); }
.badge-left_only { background: var(--left-bg); color: var(--left-text); }
.badge-right_only { background: var(--right-bg); color: var(--right-text); }
.badge-deleted, .badge-same { background: var(--del-bg); color: var(--del-text); }
.sheet-header { display: flex; align-items: center; justify-content: space-between; gap: 12px; margin-bottom: 12px; }
.sheet-summary, .file-summary { display: flex; align-items: center; justify-content: space-between; gap: 10px; flex-wrap: wrap; }
.file-title, .sheet-title { font-size: 16px; font-weight: 700; color: #0f172a; }
.summary-meta { color: var(--text-muted); font-size: 13px; font-weight: 400; }
.row-block { border: 1px solid var(--border-color); border-radius: 8px; margin-top: 16px; overflow: hidden; background: #fff; box-shadow: 0 1px 2px rgba(0,0,0,0.02); }
.row-block > summary { cursor: pointer; list-style: none; padding: 12px 16px; background: #f8fafc; display: flex; align-items: center; gap: 12px; flex-wrap: wrap; border-bottom: 1px solid transparent; transition: background 0.15s; }
.row-block > summary:hover { background: #f1f5f9; }
.row-block[open] > summary { border-bottom-color: var(--border-color); }
.row-title { font-size: 15px; font-weight: 600; color: #1e293b; }
.row-meta { color: var(--text-muted); font-size: 13px; }
.row-note { color: var(--text-muted); font-size: 13px; padding: 0 16px 12px; background: #f8fafc; border-bottom: 1px solid var(--border-color); }
.metric-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(160px, 1fr)); gap: 16px; margin: 16px 0; }
.metric-card { background: var(--card-bg); border: 1px solid var(--border-color); border-radius: 8px; padding: 16px; box-shadow: 0 1px 3px rgba(0,0,0,0.04); display: flex; flex-direction: column; align-items: flex-start; justify-content: center; }
.metric-value { font-size: 28px; font-weight: 700; line-height: 1.2; color: #0f172a; }
.metric-label { color: var(--text-muted); font-size: 13px; margin-top: 4px; font-weight: 500; }

.status-bar { display: flex; height: 10px; overflow: hidden; border-radius: 999px; background: #e2e8f0; margin: 16px 0 12px; }
.status-segment { min-width: 4px; border-right: 1px solid #fff; }
.status-segment:last-child { border-right: none; }
.status-diff { background: #ef4444; }
.status-samefile { background: #22c55e; }
.status-skipped { background: #f59e0b; }
.status-failed { background: #64748b; }
.status-legend { display: flex; flex-wrap: wrap; gap: 16px; color: var(--text-muted); font-size: 13px; margin-bottom: 16px; }
.legend-dot { display: inline-block; width: 10px; height: 10px; border-radius: 50%; margin-right: 6px; }

/* Side-by-side diff block */
.field-diff-block { padding: 0; margin: 0; border-bottom: 1px solid var(--border-color); display: flex; flex-direction: column; }
.field-diff-block:last-child { border-bottom: none; }
.field-diff-header { padding: 10px 16px; display: flex; align-items: center; gap: 10px; background: #fdfdfd; border-bottom: 1px solid #f1f5f9; }
.field-diff-title { font-weight: 600; font-size: 14px; color: #0f172a; }
.field-diff-key { color: var(--text-muted); font-size: 12px; font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace; }

/* Grid for left and right columns */
.field-diff-grid { display: grid; grid-template-columns: 1fr 1fr; }
.field-diff-grid > div:not(:last-child) { border-right: 1px solid var(--border-color); }
.field-diff-grid.three-way { grid-template-columns: 1fr 1fr 1fr; }
.field-col { padding: 12px 16px; display: flex; flex-direction: column; gap: 6px; }
.field-col-header { font-size: 12px; font-weight: 600; color: var(--text-muted); text-transform: uppercase; letter-spacing: 0.5px; }
.field-col-value { font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace; font-size: 13px; line-height: 1.5; color: #334155; word-break: break-word; white-space: pre-wrap; }

/* Highlight styles */
.diff-del { background: #fee2e2; color: #991b1b; text-decoration: line-through; text-decoration-color: #ef4444; padding: 0 4px; border-radius: 4px; }
.diff-ins { background: #dcfce7; color: #166534; padding: 0 4px; border-radius: 4px; }

.empty-value { color: #cbd5e1; font-style: italic; }
.sheet-minibar { display: flex; height: 6px; border-radius: 999px; overflow: hidden; background: #e2e8f0; margin: 8px 0 12px; width: 200px; }
.sheet-minibar-seg.conflict { background: #ef4444; }
.sheet-minibar-seg.left_only { background: #0ea5e9; }
.sheet-minibar-seg.right_only { background: #22c55e; }
.sheet-minibar-seg.deleted { background: #94a3b8; }
.sheet-minibar-seg.same { background: transparent; }

""".strip()


def _summarize_rows_by_sheet(rows: list[dict[str, str]]) -> list[dict]:
    order: list[str] = []
    by_sheet: dict[str, dict] = {}
    for item in rows:
        sheet_name = item["sheet_name"]
        if sheet_name not in by_sheet:
            order.append(sheet_name)
            by_sheet[sheet_name] = {
                "sheet_name": sheet_name,
                "diff_row_keys": set(),
                "status_keys": {status: set() for status in ROW_STATUS_LABELS},
                "field_count": 0,
            }
        bucket = by_sheet[sheet_name]
        row_key = item["alignment_row"]
        bucket["diff_row_keys"].add(row_key)
        status = item["row_status"]
        bucket["status_keys"].setdefault(status, set()).add(row_key)
        bucket["field_count"] += 1
    summaries: list[dict] = []
    for sheet_name in order:
        bucket = by_sheet[sheet_name]
        summaries.append(
            {
                "sheet_name": sheet_name,
                "diff_rows": len(bucket["diff_row_keys"]),
                "field_count": bucket["field_count"],
                "conflict": len(bucket["status_keys"].get("conflict", set())),
                "left_only": len(bucket["status_keys"].get("left_only", set())),
                "right_only": len(bucket["status_keys"].get("right_only", set())),
                "deleted": len(bucket["status_keys"].get("deleted", set())),
                "same": len(bucket["status_keys"].get("same", set())),
            }
        )
    return summaries


def _escape_html(value: str | None) -> str:
    return html.escape(value or "", quote=True)


def _html_anchor_id(prefix: str, value: str) -> str:
    safe = "".join(char if char.isalnum() else "-" for char in value).strip("-").lower()
    return f"{prefix}-{safe or 'item'}"


def _row_status_label(status: str) -> str:
    return ROW_STATUS_LABELS.get(status, status or "-")


def _render_html_sheet_section(sheet_rows: list[dict[str, str]], summary: dict, *, default_open: bool = True) -> str:
    sheet_name = summary["sheet_name"]
    anchor = _html_anchor_id("sheet", sheet_name)
    heading = (
        f'<div class="sheet-header">'
        f'<span class="sheet-title">工作表: {_escape_html(sheet_name)}</span>'
        f'<span class="summary-meta">差异行 {summary["diff_rows"]} ｜ 冲突 {summary["conflict"]} ｜ 差异字段 {summary["field_count"]}</span>'
        '<a class="sheet-jump" href="#top">返回总览</a>'
        "</div>"
    )
    minibar = _render_sheet_minibar(summary)
    badges = (
        '<p class="badges">'
        f'<span class="badge badge-conflict">冲突 {summary["conflict"]}</span>'
        f'<span class="badge badge-left_only">左独有 {summary["left_only"]}</span>'
        f'<span class="badge badge-right_only">右独有 {summary["right_only"]}</span>'
        f'<span class="badge badge-deleted">删除 {summary["deleted"]}</span>'
        f'<span class="badge badge-same">差异字段 {summary["field_count"]}</span>'
        "</p>"
    )
    has_issue = summary["conflict"] > 0 or summary["left_only"] > 0 or summary["right_only"] > 0
    open_attr = " open" if default_open and has_issue else ""
    if not sheet_rows:
        return (
            f'<details class="sheet" id="{_escape_html(anchor)}"{open_attr}>'
            f'<summary class="sheet-summary">{heading}</summary>{minibar}{badges}<p class="summary-empty">无差异。</p></details>'
        )

    rows_by_alignment: dict[str, list[dict[str, str]]] = {}
    row_order: list[str] = []
    for item in sheet_rows:
        key = item["alignment_row"]
        if key not in rows_by_alignment:
            rows_by_alignment[key] = []
            row_order.append(key)
        rows_by_alignment[key].append(item)

    has_base = any(bool(item.get("has_base")) for item in sheet_rows)
    table_head = (
        '<table class="diff-table">'
        '<thead><tr>'
        '<th style="width:80px">ID/行号</th>'
        '<th style="width:70px">状态</th>'
        '<th style="width:16%">字段 (Key)</th>'
        + ('<th style="width:20%">Base</th>' if has_base else '') +
        '<th>左侧 (Left)</th>'
        '<th>右侧 (Right)</th>'
        '</tr></thead><tbody>'
    )
    
    row_blocks: list[str] = [table_head]

    for alignment_row in row_order:
        diff_items = rows_by_alignment[alignment_row]
        first_item = diff_items[0]
        status = first_item["row_status"]
        rowspan = len(diff_items)
        
        status_label = _escape_html(_row_status_label(status))
        status_badge = f'<span class="badge badge-{_escape_html(status)}" style="margin:0">{status_label}</span>'
        
        # Meta info
        meta_parts = []
        if first_item.get("has_base") and first_item.get("base_row"):
            meta_parts.append(f"b:{first_item.get('base_row')}")
        if first_item.get("left_row"): meta_parts.append(f"l:{first_item.get('left_row')}")
        if first_item.get("right_row"): meta_parts.append(f"r:{first_item.get('right_row')}")
        
        row_id_html = f"<b>{_escape_html(alignment_row)}</b><br><small style='color:#64748b;'>{_escape_html(', '.join(meta_parts))}</small>"
        
        for i, item in enumerate(diff_items):
            field_title = item["field_title"] or item["field_key"] or ""
            field_key = item["field_key"] or ""
            field_html = f"<b>{_escape_html(field_title)}</b>"
            if field_key and field_key != field_title:
                field_html += f"<br><small style='color:#64748b; font-family:monospace'>{_escape_html(field_key)}</small>"
            
            conflict_kind = item.get("column_conflict_kind") or item.get("row_conflict_kind") or ""
            if conflict_kind and conflict_kind != "two_way":
                field_html += f"<br><span class='badge badge-conflict' style='font-size:10px; padding:2px 6px; margin-top:4px;'>{_escape_html(CONFLICT_KIND_LABELS.get(conflict_kind, conflict_kind))}</span>"

            tr_class = " class='conflict-row'" if status == "conflict" else ""
            tr_html = [f"<tr{tr_class}>"]
            
            if i == 0:
                tr_html.append(f"<td rowspan='{rowspan}'>{row_id_html}</td>")
                tr_html.append(f"<td rowspan='{rowspan}'>{status_badge}</td>")
            
            tr_html.append(f"<td>{field_html}</td>")
            
            if has_base:
                base_html = _render_html_value(item.get("base_value", ""))
                tr_html.append(f"<td class='value-cell'>{base_html}</td>")
                
            left_html = item.get("left_highlight_html") or _render_html_value(item["left_value"])
            right_html = item.get("right_highlight_html") or _render_html_value(item["right_value"])
            
            tr_html.append(f"<td class='value-cell'>{left_html}</td>")
            tr_html.append(f"<td class='value-cell'>{right_html}</td>")
            tr_html.append("</tr>")
            
            row_blocks.append("".join(tr_html))
            
    row_blocks.append("</tbody></table>")
    
    return (
        f'<details class="sheet" id="{_escape_html(anchor)}"{open_attr}>'
        f'<summary class="sheet-summary">{heading}</summary>{minibar}{badges}{"".join(row_blocks)}</details>'
    )


    rows_by_alignment: dict[str, list[dict[str, str]]] = {}
    row_order: list[str] = []
    for item in sheet_rows:
        key = item["alignment_row"]
        if key not in rows_by_alignment:
            rows_by_alignment[key] = []
            row_order.append(key)
        rows_by_alignment[key].append(item)

    row_blocks: list[str] = []
    for alignment_row in row_order:
        diff_items = rows_by_alignment[alignment_row]
        first_item = diff_items[0]
        status = first_item["row_status"]
        row_meta_parts = [
            f"左:{first_item['left_row'] or '-'}",
            f"右:{first_item['right_row'] or '-'}",
        ]
        if first_item.get("has_base"):
            row_meta_parts.insert(0, f"base:{first_item.get('base_row') or '-'}")
        row_meta = " ".join(row_meta_parts)
        detail_open = " open" if status in {"conflict", "left_only", "right_only"} else ""
        summary_line = (
            f"<summary>"
            f'<span class="row-title">行 {_escape_html(alignment_row)}</span>'
            f'<span class="badge badge-{_escape_html(status)}">{_escape_html(_row_status_label(status))}</span>'
            f'<span class="row-meta">{_escape_html(row_meta)} ｜ 差异字段 {len(diff_items)}</span>'
            f"</summary>"
        )
        notes: list[str] = []
        if first_item["row_reason"]:
            notes.append(f"说明: {first_item['row_reason']}")
        if first_item["row_note"]:
            notes.append(f"备注: {first_item['row_note']}")
        conflict_kind = first_item.get("row_conflict_kind", "")
        if conflict_kind and conflict_kind != "two_way":
            notes.append(f"冲突类型: {CONFLICT_KIND_LABELS.get(conflict_kind, conflict_kind)}")
        note_line = f'<div class="row-note">{_escape_html(" ｜ ".join(notes))}</div>' if notes else ""
        blocks: list[str] = [_render_html_field_block(item) for item in diff_items]
        row_blocks.append(
            f'<details class="row-block"{detail_open}>{summary_line}{note_line}'
            f'<div class="row-block-body">{"".join(blocks)}</div></details>'
        )
    return (
        f'<details class="sheet" id="{_escape_html(anchor)}"{open_attr}>'
        f'<summary class="sheet-summary">{heading}</summary>{minibar}{badges}{"".join(row_blocks)}</details>'
    )


def _render_sheet_minibar(summary: dict) -> str:
    total = (
        summary.get("conflict", 0)
        + summary.get("left_only", 0)
        + summary.get("right_only", 0)
        + summary.get("deleted", 0)
        + summary.get("same", 0)
    )
    if total <= 0:
        return ""
    segments: list[str] = []
    for key in ("conflict", "left_only", "right_only", "deleted", "same"):
        count = summary.get(key, 0)
        if count <= 0:
            continue
        percent = count / total * 100
        segments.append(f'<span class="sheet-minibar-seg {key}" style="width:{percent:.2f}%" title="{_row_status_label(key)} {count}"></span>')
    return f'<div class="sheet-minibar">{"".join(segments)}</div>'


def _render_html_field_block(item: dict[str, str]) -> str:
    field_title = item["field_title"] or item["field_key"] or ""
    field_key = item["field_key"] or ""
    has_base = bool(item.get("has_base"))
    is_conflict = item["row_status"] == "conflict"
    conflict_kind = item.get("column_conflict_kind") or item.get("row_conflict_kind") or ""
    kind_badge = ""
    if conflict_kind and conflict_kind != "two_way":
        kind_badge = (
            f'<span class="badge badge-conflict">'
            f'{_escape_html(CONFLICT_KIND_LABELS.get(conflict_kind, conflict_kind))}'
            "</span>"
        )
    header_parts = [
        f'<span class="field-diff-title">{_escape_html(field_title)}</span>',
    ]
    if field_key and field_key != field_title:
        header_parts.append(f'<span class="field-diff-key">{_escape_html(field_key)}</span>')
    if kind_badge:
        header_parts.append(kind_badge)
    header = f'<div class="field-diff-header">{"".join(header_parts)}</div>'

    grid_class = "field-diff-grid three-way" if has_base else "field-diff-grid"
    block_classes = ["field-diff-block"]
    if is_conflict:
        block_classes.append("conflict")
    if has_base:
        block_classes.append("three-way")
    cols = []
    if has_base:
        base_val = _render_html_value(item.get("base_value", ""))
        cols.append(f'<div class="field-col"><div class="field-col-header">Base</div><div class="field-col-value">{base_val}</div></div>')
    
    left_val = item.get("left_highlight_html") or _render_html_value(item["left_value"])
    right_val = item.get("right_highlight_html") or _render_html_value(item["right_value"])
    
    cols.append(f'<div class="field-col"><div class="field-col-header">Left (左)</div><div class="field-col-value">{left_val}</div></div>')
    cols.append(f'<div class="field-col"><div class="field-col-header">Right (右)</div><div class="field-col-value">{right_val}</div></div>')

    grid_html = f'<div class="{grid_class}">{"".join(cols)}</div>'
    return f'<div class="{" ".join(block_classes)}">{header}{grid_html}</div>'


def _render_html_side_row(label: str, value_html: str, css_class: str) -> str:
    return ""  # No longer used, but kept for compatibility just in case


def _render_html_value(value: str | None) -> str:
    if value:
        return _escape_html(value)
    return '<span class="empty-value">(空)</span>'


def _render_html_sheet_summary_table(summaries: list[dict]) -> str:
    if not summaries:
        return '<p class="summary-empty">无差异工作表。</p>'
    header = (
        "<tr>"
        "<th>工作表</th><th>差异行</th><th>冲突</th>"
        "<th>左独有</th><th>右独有</th><th>删除</th><th>差异字段</th>"
        "</tr>"
    )
    body_rows = [
        "<tr>"
        f'<td><a href="#{_escape_html(_html_anchor_id("sheet", summary["sheet_name"]))}">{_escape_html(summary["sheet_name"])}</a></td>'
        f'<td class="num">{summary["diff_rows"]}</td>'
        f'<td class="num">{summary["conflict"]}</td>'
        f'<td class="num">{summary["left_only"]}</td>'
        f'<td class="num">{summary["right_only"]}</td>'
        f'<td class="num">{summary["deleted"]}</td>'
        f'<td class="num">{summary["field_count"]}</td>'
        "</tr>"
        for summary in summaries
    ]
    return f"<table><thead>{header}</thead><tbody>{''.join(body_rows)}</tbody></table>"


def format_diff_report_html(
    rows: list[dict[str, str]],
    comparison_options: ComparisonOptions | None = None,
    *,
    title: str = "差异报告",
    file_summaries: list[dict] | None = None,
) -> str:
    del comparison_options  # reserved for future use; row shape already incorporates options
    parts: list[str] = [
        "<!doctype html>",
        '<html lang="zh"><head><meta charset="utf-8">',
        f"<title>{_escape_html(title)}</title>",
        f"<style>{DIFF_REPORT_CSS}</style>",
        "</head><body>",
        '<div class="page" id="top">',
        f"<h1><span>{_escape_html(title)}</span> <div class='toolbar'><button onclick='toggleAllDetails(true)'>全部展开</button><button onclick='toggleAllDetails(false)'>全部折叠</button></div></h1>", 
    ]

    if file_summaries is None:
        sheet_summaries = _summarize_rows_by_sheet(rows)
        parts.append("<h2>总览</h2>")
        parts.append(_render_html_sheet_summary_table(sheet_summaries))
        if not sheet_summaries:
            parts.append('<p class="summary-empty">没有可导出的差异。</p>')
        else:
            rows_by_sheet: dict[str, list[dict[str, str]]] = {}
            for item in rows:
                rows_by_sheet.setdefault(item["sheet_name"], []).append(item)
            parts.append("<h2>详情</h2>")
            for summary in sheet_summaries:
                parts.append(_render_html_sheet_section(rows_by_sheet.get(summary["sheet_name"], []), summary))
    else:
        parts.append("<h2>文件汇总</h2>")
        parts.append(_render_html_batch_visualization(file_summaries))
        parts.append(
            '<details class="overview-panel" open><summary>文件汇总表</summary>'
            f"{_render_html_file_summary_table(file_summaries)}</details>"
        )
        for entry in file_summaries:
            parts.append(_render_html_file_section(entry))

    parts.append("<script>function toggleAllDetails(state) { document.querySelectorAll('details').forEach(d => { d.open = state; }); }</script></div></body></html>")
    return "".join(parts)


def _batch_status_bucket(entry: dict) -> str:
    status = str(entry.get("status", ""))
    rows = entry.get("rows") or []
    if "失败" in status:
        return "failed"
    if "跳过" in status or "缺少" in status or "独有" in status:
        return "skipped"
    if rows or "差异" in status and "无差异" not in status:
        return "diff"
    return "same"


def _render_html_batch_visualization(entries: list[dict]) -> str:
    total = len(entries)
    buckets = {"diff": 0, "same": 0, "skipped": 0, "failed": 0}
    total_sheets = 0
    total_diff_rows = 0
    total_fields = 0
    for entry in entries:
        buckets[_batch_status_bucket(entry)] += 1
        totals = _aggregate_file_entry_totals(entry)
        total_sheets += totals["sheets"]
        total_diff_rows += totals["diff_rows"]
        total_fields += sum(1 for _ in entry.get("rows") or [])
    cards = (
        '<div class="metric-grid">'
        f'<div class="metric-card"><span class="metric-value">{total}</span><span class="metric-label">文件总数</span></div>'
        f'<div class="metric-card"><span class="metric-value">{buckets["diff"]}</span><span class="metric-label">有差异</span></div>'
        f'<div class="metric-card"><span class="metric-value">{buckets["same"]}</span><span class="metric-label">无差异</span></div>'
        f'<div class="metric-card"><span class="metric-value">{buckets["failed"] + buckets["skipped"]}</span><span class="metric-label">失败/跳过</span></div>'
        f'<div class="metric-card"><span class="metric-value">{total_sheets}</span><span class="metric-label">涉及工作表</span></div>'
        f'<div class="metric-card"><span class="metric-value">{total_diff_rows}</span><span class="metric-label">差异行</span></div>'
        f'<div class="metric-card"><span class="metric-value">{total_fields}</span><span class="metric-label">差异字段</span></div>'
        "</div>"
    )
    if total <= 0:
        return cards
    segments: list[str] = []
    for key, css_class in (
        ("diff", "status-diff"),
        ("same", "status-samefile"),
        ("skipped", "status-skipped"),
        ("failed", "status-failed"),
    ):
        count = buckets[key]
        if count <= 0:
            continue
        percent = count / total * 100
        segments.append(f'<span class="status-segment {css_class}" style="width:{percent:.2f}%"></span>')
    legend = (
        '<div class="status-legend">'
        f'<span><i class="legend-dot status-diff"></i>有差异 {buckets["diff"]}</span>'
        f'<span><i class="legend-dot status-samefile"></i>无差异 {buckets["same"]}</span>'
        f'<span><i class="legend-dot status-skipped"></i>跳过 {buckets["skipped"]}</span>'
        f'<span><i class="legend-dot status-failed"></i>失败 {buckets["failed"]}</span>'
        "</div>"
    )
    return cards + f'<div class="status-bar">{"".join(segments)}</div>' + legend


def _render_html_file_summary_table(entries: list[dict]) -> str:
    if not entries:
        return '<p class="summary-empty">没有可导出的文件。</p>'
    header = (
        "<tr>"
        "<th>文件</th><th>状态</th><th>工作表</th><th>差异行</th>"
        "<th>冲突</th><th>左独有</th><th>右独有</th><th>删除</th><th>备注</th>"
        "</tr>"
    )
    body_rows: list[str] = []
    for entry in entries:
        totals = _aggregate_file_entry_totals(entry)
        anchor = _html_anchor_id("file", str(entry.get("file", "")))
        body_rows.append(
            "<tr>"
            f'<td><a href="#{_escape_html(anchor)}">{_escape_html(entry.get("file", ""))}</a></td>'
            f'<td>{_escape_html(entry.get("status", ""))}</td>'
            f'<td class="num">{totals["sheets"]}</td>'
            f'<td class="num">{totals["diff_rows"]}</td>'
            f'<td class="num">{totals["conflict"]}</td>'
            f'<td class="num">{totals["left_only"]}</td>'
            f'<td class="num">{totals["right_only"]}</td>'
            f'<td class="num">{totals["deleted"]}</td>'
            f'<td>{_escape_html(entry.get("note", ""))}</td>'
            "</tr>"
        )
    return f"<table><thead>{header}</thead><tbody>{''.join(body_rows)}</tbody></table>"


def _render_html_file_section(entry: dict) -> str:
    file_name = entry.get("file", "")
    status = entry.get("status", "")
    note = entry.get("note", "")
    rows = entry.get("rows") or []
    anchor = _html_anchor_id("file", str(file_name))
    totals = _aggregate_file_entry_totals(entry)
    status_line_parts: list[str] = []
    if status:
        status_line_parts.append(f"状态: {_escape_html(status)}")
    if note:
        status_line_parts.append(_escape_html(note))
    summary_meta = (
        f'工作表 {totals["sheets"]} ｜ 差异行 {totals["diff_rows"]} ｜ '
        f'冲突 {totals["conflict"]} ｜ 差异字段 {len(rows)}'
    )
    summary_html = (
        '<summary class="file-summary">'
        f'<span class="file-title">{_escape_html(file_name)}</span>'
        f'<span class="badge badge-{_escape_html(_file_status_badge_class(entry))}">{_escape_html(status or "未分类")}</span>'
        f'<span class="summary-meta">{_escape_html(summary_meta)}</span>'
        "</summary>"
    )
    note_line = f'<p class="note">{" ｜ ".join(status_line_parts)}</p>' if status_line_parts else ""
    sheet_summaries = _summarize_rows_by_sheet(rows)
    body: list[str] = []
    if sheet_summaries:
        body.append(
            '<details class="overview-panel"><summary>工作表汇总</summary>'
            f"{_render_html_sheet_summary_table(sheet_summaries)}</details>"
        )
        rows_by_sheet: dict[str, list[dict[str, str]]] = {}
        for item in rows:
            rows_by_sheet.setdefault(item["sheet_name"], []).append(item)
        for sheet_summary in sheet_summaries:
            body.append(_render_html_sheet_section(rows_by_sheet.get(sheet_summary["sheet_name"], []), sheet_summary, default_open=False))
    else:
        body.append('<p class="summary-empty">无差异详情。</p>')
    open_attr = " open" if _batch_status_bucket(entry) in {"diff", "failed"} else ""
    return f'<details class="file" id="{_escape_html(anchor)}"{open_attr}>{summary_html}{note_line}{"".join(body)}</details>'


def _file_status_badge_class(entry: dict) -> str:
    bucket = _batch_status_bucket(entry)
    if bucket == "diff":
        return "conflict"
    if bucket == "skipped":
        return "right_only"
    if bucket == "failed":
        return "deleted"
    return "same"


def _aggregate_file_entry_totals(entry: dict) -> dict[str, int]:
    rows = entry.get("rows") or []
    sheet_summaries = _summarize_rows_by_sheet(rows)
    totals = {
        "sheets": len(sheet_summaries),
        "diff_rows": sum(summary["diff_rows"] for summary in sheet_summaries),
        "conflict": sum(summary["conflict"] for summary in sheet_summaries),
        "left_only": sum(summary["left_only"] for summary in sheet_summaries),
        "right_only": sum(summary["right_only"] for summary in sheet_summaries),
        "deleted": sum(summary["deleted"] for summary in sheet_summaries),
    }
    return totals


def _escape_markdown_cell(value: str | None) -> str:
    text = value or ""
    return text.replace("|", "\\|").replace("\r\n", " ").replace("\n", "<br>")


def format_diff_report_markdown(
    rows: list[dict[str, str]],
    comparison_options: ComparisonOptions | None = None,
    *,
    title: str = "差异报告",
    file_summaries: list[dict] | None = None,
) -> str:
    del comparison_options
    lines: list[str] = [f"# {title}", ""]

    if file_summaries is None:
        sheet_summaries = _summarize_rows_by_sheet(rows)
        lines.append("## 总览")
        lines.append("")
        lines.extend(_render_markdown_sheet_summary_table(sheet_summaries))
        lines.append("")
        if not sheet_summaries:
            lines.append("> 没有可导出的差异。")
            return "\n".join(lines)
        rows_by_sheet: dict[str, list[dict[str, str]]] = {}
        for item in rows:
            rows_by_sheet.setdefault(item["sheet_name"], []).append(item)
        lines.append("## 详情")
        lines.append("")
        for summary in sheet_summaries:
            lines.extend(_render_markdown_sheet_section(rows_by_sheet.get(summary["sheet_name"], []), summary))
    else:
        lines.append("## 文件汇总")
        lines.append("")
        lines.extend(_render_markdown_file_summary_table(file_summaries))
        lines.append("")
        for entry in file_summaries:
            lines.extend(_render_markdown_file_section(entry))

    return "\n".join(lines).rstrip() + "\n"


def _render_markdown_sheet_summary_table(summaries: list[dict]) -> list[str]:
    if not summaries:
        return ["> 无差异工作表。"]
    lines = [
        "| 工作表 | 差异行 | 冲突 | 左独有 | 右独有 | 删除 | 差异字段 |",
        "| --- | ---: | ---: | ---: | ---: | ---: | ---: |",
    ]
    for summary in summaries:
        lines.append(
            "| "
            f"{_escape_markdown_cell(summary['sheet_name'])} | "
            f"{summary['diff_rows']} | {summary['conflict']} | "
            f"{summary['left_only']} | {summary['right_only']} | "
            f"{summary['deleted']} | {summary['field_count']} |"
        )
    return lines


def _render_markdown_sheet_section(sheet_rows: list[dict[str, str]], summary: dict) -> list[str]:
    lines = [
        f"### 工作表: {summary['sheet_name']}",
        "",
        (
            f"- 差异行 {summary['diff_rows']} ｜ 冲突 {summary['conflict']} ｜ "
            f"左独有 {summary['left_only']} ｜ 右独有 {summary['right_only']} ｜ "
            f"删除 {summary['deleted']} ｜ 差异字段 {summary['field_count']}"
        ),
        "",
    ]
    if not sheet_rows:
        lines.append("> 无差异。")
        lines.append("")
        return lines

    rows_by_alignment: dict[str, list[dict[str, str]]] = {}
    row_order: list[str] = []
    for item in sheet_rows:
        key = item["alignment_row"]
        if key not in rows_by_alignment:
            rows_by_alignment[key] = []
            row_order.append(key)
        rows_by_alignment[key].append(item)

    for alignment_row in row_order:
        diff_items = rows_by_alignment[alignment_row]
        first_item = diff_items[0]
        has_base = bool(first_item.get("has_base"))
        header_parts = [f"行 {alignment_row}", _row_status_label(first_item["row_status"])]
        if has_base:
            header_parts.append(f"base:{first_item.get('base_row') or '-'}")
        header_parts.append(f"左:{first_item['left_row'] or '-'}")
        header_parts.append(f"右:{first_item['right_row'] or '-'}")
        lines.append(f"#### {' ｜ '.join(header_parts)}")
        kind = first_item.get("row_conflict_kind", "")
        meta_parts: list[str] = []
        if first_item["row_reason"]:
            meta_parts.append(f"说明: {first_item['row_reason']}")
        if first_item["row_note"]:
            meta_parts.append(f"备注: {first_item['row_note']}")
        if kind and kind != "two_way":
            meta_parts.append(f"冲突类型: {CONFLICT_KIND_LABELS.get(kind, kind)}")
        if meta_parts:
            lines.append("")
            lines.append(f"> {' ｜ '.join(meta_parts)}")
        lines.append("")
        for item in diff_items:
            field_title = item["field_title"] or item["field_key"] or ""
            field_key = item["field_key"] or ""
            field_label = field_title if field_key == field_title or not field_key else f"{field_title} ({field_key})"
            lines.append(f"- **{field_label}**")
            inline_plain = item.get("left_vs_right_inline_plain") or ""
            if inline_plain:
                lines.append(f"    - 对比: `{inline_plain}`")
            lines.append(f"    - 左: {item['left_value'] or '(空)'}")
            lines.append(f"    - 右: {item['right_value'] or '(空)'}")
            if has_base:
                lines.append(f"    - base: {item.get('base_value', '') or '(空)'}")
        lines.append("")
    return lines


def _render_markdown_file_summary_table(entries: list[dict]) -> list[str]:
    if not entries:
        return ["> 没有可导出的文件。"]
    lines = [
        "| 文件 | 状态 | 工作表 | 差异行 | 冲突 | 左独有 | 右独有 | 删除 | 备注 |",
        "| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | --- |",
    ]
    for entry in entries:
        totals = _aggregate_file_entry_totals(entry)
        lines.append(
            "| "
            f"{_escape_markdown_cell(entry.get('file', ''))} | "
            f"{_escape_markdown_cell(entry.get('status', ''))} | "
            f"{totals['sheets']} | {totals['diff_rows']} | "
            f"{totals['conflict']} | {totals['left_only']} | "
            f"{totals['right_only']} | {totals['deleted']} | "
            f"{_escape_markdown_cell(entry.get('note', ''))} |"
        )
    return lines


def _render_markdown_file_section(entry: dict) -> list[str]:
    file_name = entry.get("file", "")
    status = entry.get("status", "")
    note = entry.get("note", "")
    rows = entry.get("rows") or []
    lines: list[str] = [f"## {file_name}", ""]
    tail: list[str] = []
    if status:
        tail.append(f"状态: {status}")
    if note:
        tail.append(note)
    if tail:
        lines.append("- " + " ｜ ".join(tail))
        lines.append("")
    sheet_summaries = _summarize_rows_by_sheet(rows)
    if not sheet_summaries:
        lines.append("> 无差异详情。")
        lines.append("")
        return lines
    lines.extend(_render_markdown_sheet_summary_table(sheet_summaries))
    lines.append("")
    rows_by_sheet: dict[str, list[dict[str, str]]] = {}
    for item in rows:
        rows_by_sheet.setdefault(item["sheet_name"], []).append(item)
    for summary in sheet_summaries:
        lines.extend(_render_markdown_sheet_section(rows_by_sheet.get(summary["sheet_name"], []), summary))
    return lines


def _aligned_row_diff_entries(
    alignment: SheetAlignment,
    row,
    alignment_row: int,
    comparison_options: ComparisonOptions | None = None,
) -> list[dict[str, str]]:
    report_all_fields = row.status in {"left_only", "right_only", "deleted"}
    has_base = getattr(alignment, "base_sheet", None) is not None
    entries: list[dict[str, str]] = []
    for logical_index, binding in enumerate(alignment.columns, start=1):
        left_value = row.left_row.value_at(binding.left_index or -1) if row.left_row is not None and binding.left_index else ""
        merged_value = row.merged_row.value_at(logical_index)
        right_value = row.right_row.value_at(binding.right_index or -1) if row.right_row is not None and binding.right_index else ""
        base_index = getattr(binding, "base_index", None)
        base_row_obj = getattr(row, "base_row", None)
        base_value = (
            base_row_obj.value_at(base_index or -1)
            if base_row_obj is not None and base_index
            else ""
        )
        has_diff = (
            not compare_text_values(left_value, right_value, comparison_options)
            or logical_index in row.conflict_columns
        )
        if report_all_fields:
            if not (left_value or merged_value or right_value or base_value):
                continue
        elif not has_diff:
            continue

        column_conflict_kind = ""
        column_kinds = getattr(row, "column_conflict_kinds", None)
        if column_kinds:
            column_conflict_kind = column_kinds.get(logical_index, "")

        inline_spans = compute_inline_diff(left_value, right_value)
        left_vs_right_html = render_diff_html(inline_spans)
        left_vs_right_plain = render_diff_plain(inline_spans)
        _, right_highlight_spans = compute_char_diff(left_value, right_value)
        left_highlight_spans, _ = compute_char_diff(left_value, right_value)
        left_highlight_html = render_diff_html(left_highlight_spans)
        right_highlight_html = render_diff_html(right_highlight_spans)
        if has_base:
            base_vs_left_spans = compute_inline_diff(base_value, left_value)
            base_vs_right_spans = compute_inline_diff(base_value, right_value)
            base_vs_left_html = render_diff_html(base_vs_left_spans)
            base_vs_right_html = render_diff_html(base_vs_right_spans)
            base_vs_left_plain = render_diff_plain(base_vs_left_spans)
            base_vs_right_plain = render_diff_plain(base_vs_right_spans)
        else:
            base_vs_left_html = ""
            base_vs_right_html = ""
            base_vs_left_plain = ""
            base_vs_right_plain = ""

        entries.append(
            {
                "sheet_name": alignment.sheet_name,
                "alignment_row": str(alignment_row),
                "row_status": row.status,
                "row_reason": row.reason,
                "row_note": row.note,
                "row_conflict_kind": getattr(row, "conflict_kind", "two_way"),
                "column_conflict_kind": column_conflict_kind,
                "left_row": str(row.left_row.row_index) if row.left_row is not None else "",
                "right_row": str(row.right_row.row_index) if row.right_row is not None else "",
                "merged_row": str(row.merged_row.row_index) if row.merged_row is not None else "",
                "base_row": str(base_row_obj.row_index) if base_row_obj is not None else "",
                "field_key": binding.key,
                "field_title": binding.title or binding.key or f"col_{logical_index}",
                "first_diff_char": str(_first_difference_index(left_value, right_value) or ""),
                "left_value": left_value,
                "merged_value": merged_value,
                "right_value": right_value,
                "base_value": base_value,
                "left_vs_right_inline_html": left_vs_right_html,
                "left_vs_right_inline_plain": left_vs_right_plain,
                "left_highlight_html": left_highlight_html,
                "right_highlight_html": right_highlight_html,
                "base_vs_left_inline_html": base_vs_left_html,
                "base_vs_right_inline_html": base_vs_right_html,
                "base_vs_left_inline_plain": base_vs_left_plain,
                "base_vs_right_inline_plain": base_vs_right_plain,
                "has_base": "1" if has_base else "",
            }
        )
    return entries


def _first_difference_index(left_value: str, right_value: str) -> int:
    left_text = left_value or ""
    right_text = right_value or ""
    limit = min(len(left_text), len(right_text))
    for index in range(limit):
        if left_text[index] != right_text[index]:
            return index + 1
    if len(left_text) != len(right_text):
        return limit + 1
    return 0


def _build_export_row(aligned_row, included_columns) -> RowData:
    source_row = aligned_row.left_row or aligned_row.right_row or aligned_row.merged_row
    export_row = RowData(
        row_index=source_row.row_index,
        attrs=dict(source_row.attrs),
        kind=source_row.kind,
    )

    for actual_index, (logical_index, binding) in enumerate(included_columns, start=1):
        template_cell = None
        if aligned_row.left_row is not None and binding.left_index is not None:
            template_cell = aligned_row.left_row.cell_at(binding.left_index)
        if template_cell is None and aligned_row.right_row is not None and binding.right_index is not None:
            template_cell = aligned_row.right_row.cell_at(binding.right_index)
        if template_cell is None:
            template_cell = aligned_row.merged_row.cell_at(logical_index)

        cell = template_cell.clone() if template_cell is not None else CellData(column_index=actual_index)
        cell.column_index = actual_index
        cell.value = aligned_row.merged_row.value_at(logical_index)
        export_row.cells.append(cell)

    export_row.cells.sort(key=lambda item: item.column_index)
    return export_row


def _should_export_sheet(alignment: SheetAlignment, merge_rule: MergeRule) -> bool:
    if alignment.left_sheet is not None and alignment.right_sheet is None:
        return merge_rule.keep_left_only_sheets
    if alignment.right_sheet is not None and alignment.left_sheet is None:
        return merge_rule.keep_right_only_sheets
    return True


def _included_columns(alignment: SheetAlignment, merge_rule: MergeRule):
    return [
        (logical_index, binding)
        for logical_index, binding in enumerate(alignment.columns, start=1)
        if is_binding_included(binding, merge_rule)
    ]


def _clone_column(column_node):
    return copy.deepcopy(column_node)


def row_to_lxml(row: RowData):
    row_element = LET.Element(LET.QName(XML_NS, "Row"))
    for key, value in row.attrs.items():
        row_element.set(key, value)

    previous_column = 0
    for cell in sorted(row.cells, key=lambda item: item.column_index):
        cell_element = LET.SubElement(row_element, LET.QName(XML_NS, "Cell"))
        if cell.column_index != previous_column + 1:
            cell_element.set(LET.QName(XML_NS, "Index"), str(cell.column_index))
        for key, value in cell.attrs.items():
            cell_element.set(key, value)

        data_element = LET.SubElement(cell_element, LET.QName(XML_NS, "Data"))
        data_element.set(LET.QName(XML_NS, "Type"), cell.data_type or _guess_type(cell.value))
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


def _strip_formulas_from_tree(root) -> None:
    for cell in root.xpath("//ss:Cell", namespaces=LXML_NS):
        for attr_name in list(cell.attrib):
            local_name = LET.QName(attr_name).localname
            if local_name in {"Formula", "ArrayRange"}:
                del cell.attrib[attr_name]


def _get_or_create_styles_node(root):
    matches = root.xpath("./ss:Styles", namespaces=LXML_NS)
    if matches:
        return matches[0]
    styles = LET.Element(LET.QName(XML_NS, "Styles"))
    root.insert(0, styles)
    return styles


def _style_map(styles_node) -> dict[str, LET._Element]:
    return {
        style.attrib.get(f"{{{XML_NS}}}ID"): style
        for style in styles_node.xpath("./ss:Style", namespaces=LXML_NS)
        if style.attrib.get(f"{{{XML_NS}}}ID")
    }


def _remap_sheet_styles(sheet_node, left_styles_node, left_style_map, right_style_map) -> None:
    style_attrs = [
        attr_name
        for node in sheet_node.xpath(".//*[@ss:StyleID]", namespaces=LXML_NS)
        for attr_name in node.attrib
        if LET.QName(attr_name).localname == "StyleID"
    ]
    if not style_attrs:
        return

    remap: dict[str, str] = {}
    used_style_ids = {
        node.attrib.get(f"{{{XML_NS}}}StyleID")
        for node in sheet_node.xpath(".//*[@ss:StyleID]", namespaces=LXML_NS)
        if node.attrib.get(f"{{{XML_NS}}}StyleID")
    }

    for style_id in sorted(used_style_ids):
        left_style = left_style_map.get(style_id)
        right_style = right_style_map.get(style_id)
        if right_style is None:
            continue
        if left_style is None:
            left_styles_node.append(copy.deepcopy(right_style))
            left_style_map[style_id] = left_styles_node[-1]
            continue
        if LET.tostring(left_style) == LET.tostring(right_style):
            continue

        new_style_id = _allocate_style_id(left_style_map, style_id)
        new_style = copy.deepcopy(right_style)
        new_style.attrib[f"{{{XML_NS}}}ID"] = new_style_id
        left_styles_node.append(new_style)
        left_style_map[new_style_id] = new_style
        remap[style_id] = new_style_id

    if not remap:
        return

    for node in sheet_node.xpath(".//*[@ss:StyleID]", namespaces=LXML_NS):
        style_id = node.attrib.get(f"{{{XML_NS}}}StyleID")
        if style_id in remap:
            node.attrib[f"{{{XML_NS}}}StyleID"] = remap[style_id]


def _allocate_style_id(existing_styles: dict[str, LET._Element], base_style_id: str) -> str:
    seed = f"{base_style_id}_r"
    candidate = seed
    counter = 1
    while candidate in existing_styles:
        counter += 1
        candidate = f"{seed}{counter}"
    return candidate
