from __future__ import annotations

from table_merge_tool import exporter


def _diff_row(**overrides: str) -> dict[str, str]:
    row = {
        "sheet_name": "Skill<Main>",
        "alignment_row": "3",
        "row_status": "conflict",
        "row_conflict_kind": "two_way",
        "column_conflict_kind": "",
        "row_reason": "值不同",
        "row_note": "人工确认",
        "left_row": "3",
        "right_row": "3",
        "merged_row": "3",
        "base_row": "",
        "field_key": "desc|text",
        "field_title": "描述",
        "first_diff_char": "2",
        "left_value": "A < B",
        "merged_value": "A < B",
        "right_value": "A | B\nC",
        "base_value": "",
        "left_vs_right_inline_html": 'A <span class="diff-ins">| B<br>C</span>',
        "left_vs_right_inline_plain": "A {+| B\nC+}",
        "left_highlight_html": "A ",
        "right_highlight_html": 'A <span class="diff-ins">| B<br>C</span>',
        "base_vs_left_inline_html": "",
        "base_vs_right_inline_html": "",
        "base_vs_left_inline_plain": "",
        "base_vs_right_inline_plain": "",
        "has_base": "",
    }
    row.update(overrides)
    return row


def _three_way_row(**overrides: str) -> dict[str, str]:
    row = _diff_row(
        row_conflict_kind="three_way_diverge",
        column_conflict_kind="three_way_diverge",
        base_row="3",
        base_value="A = B",
        base_vs_left_inline_html='A <span class="diff-del">=</span><span class="diff-ins">&lt;</span> B',
        base_vs_right_inline_html='A <span class="diff-del">=</span><span class="diff-ins">|</span> B<br>C',
        base_vs_left_inline_plain="A [-=-]{+<+} B",
        base_vs_right_inline_plain="A [-=-]{+|+} B\nC",
        has_base="1",
    )
    row.update(overrides)
    return row


def test_format_diff_report_html_escapes_values_and_renders_summary():
    html = exporter.format_diff_report_html([_diff_row()])

    assert html.startswith("<!doctype html>")
    assert "<span>差异报告</span>" in html
    assert "Skill&lt;Main&gt;" in html
    assert '<details class="sheet" id="sheet-skill-main" open>' in html
    assert '<a href="#sheet-skill-main">' in html
    assert '<span class="badge badge-conflict">冲突 1</span>' in html
    assert "desc|text" in html
    assert "toggleAllDetails" in html


def test_format_diff_report_html_renders_compact_field_blocks_with_diff_highlight():
    html = exporter.format_diff_report_html([_diff_row()])
    assert '<table class="diff-table">' in html
    assert "conflict-row" in html
    assert '<span class="diff-ins">| B<br>C</span>' in html
    assert 'field-diff-side merged' not in html
    assert "合并:" not in html
    assert '<div class="sheet-minibar">' in html


def test_format_diff_report_html_renders_three_way_base_row():
    html = exporter.format_diff_report_html([_three_way_row()])
    assert "<th style=\"width:20%\">Base</th>" in html
    assert "A = B" in html
    assert "三方分歧" in html


def test_format_diff_report_html_handles_empty_rows():
    html = exporter.format_diff_report_html([])

    assert "无差异工作表" in html
    assert "没有可导出的差异" in html


def test_format_diff_report_markdown_uses_bullet_layout_and_inline_diff():
    markdown = exporter.format_diff_report_markdown([_diff_row()])

    assert markdown.startswith("# 差异报告\n")
    assert "| Skill<Main> | 1 | 1 | 0 | 0 | 0 | 1 |" in markdown
    assert "### 工作表: Skill<Main>" in markdown
    assert "#### 行 3" in markdown
    assert "- **描述 (desc|text)**" in markdown
    assert "- 对比: `A {+| B\nC+}`" in markdown
    assert "- 左: A < B" in markdown
    assert "- 中:" not in markdown


def test_format_diff_report_markdown_renders_three_way_base():
    markdown = exporter.format_diff_report_markdown([_three_way_row()])
    assert "base:3" in markdown
    assert "- base: A = B" in markdown
    assert "冲突类型: 三方分歧" in markdown
    assert "- 中:" not in markdown


def test_format_diff_report_markdown_renders_file_summaries():
    rows = [_diff_row(sheet_name="SheetA")]
    markdown = exporter.format_diff_report_markdown(
        [],
        title="批量差异报告",
        file_summaries=[
            {
                "file": "a|b.xml",
                "status": "已对比",
                "note": "1处差异",
                "rows": rows,
            }
        ],
    )

    assert markdown.startswith("# 批量差异报告")
    assert "| a\\|b.xml | 已对比 | 1 | 1 | 1 | 0 | 0 | 0 | 1处差异 |" in markdown
    assert "## a|b.xml" in markdown
    assert "### 工作表: SheetA" in markdown


def test_format_diff_report_text_uses_inline_plain_and_base_line():
    text = exporter.format_diff_report_text([_three_way_row()])
    assert "[工作表] Skill<Main>" in text
    assert "base:3" in text
    assert "冲突类型: 三方分歧" in text
    assert "对比: A {+| B" in text
    assert "base: A = B" in text
    assert "    中:" not in text


def test_format_diff_report_html_renders_batch_collapsible_visual_summary():
    html = exporter.format_diff_report_html(
        [],
        title="批量差异报告",
        file_summaries=[
            {"file": "changed.xml", "status": "有差异", "note": "1处差异", "rows": [_diff_row(sheet_name="SheetA")]},
            {"file": "same.xml", "status": "无差异", "note": "", "rows": []},
            {"file": "failed.xml", "status": "失败", "note": "读取失败", "rows": []},
        ],
    )

    assert '<div class="metric-grid">' in html
    assert '<div class="status-bar">' in html
    assert '<details class="overview-panel" open><summary>文件汇总表</summary>' in html
    assert '<details class="file" id="file-changed-xml" open>' in html
    assert '<span class="file-title">changed.xml</span>' in html
    assert '<summary class="file-summary">' in html
    assert "{'sheet_name'" not in html
    assert '<details class="file" id="file-same-xml">' in html
    assert '<details class="sheet" id="sheet-sheeta"' in html
    assert "工作表汇总" in html


def test_export_diff_report_routes_by_extension(tmp_path, monkeypatch):
    monkeypatch.setattr(exporter, "build_diff_report_rows", lambda alignments, comparison_options=None: [_diff_row()])

    html_path = tmp_path / "report.html"
    md_path = tmp_path / "report.md"
    txt_path = tmp_path / "report.txt"
    csv_path = tmp_path / "report.csv"

    for path in (html_path, md_path, txt_path, csv_path):
        exporter.export_diff_report({}, path)

    assert html_path.read_text(encoding="utf-8").startswith("<!doctype html>")
    assert md_path.read_text(encoding="utf-8").startswith("# 差异报告")
    assert txt_path.read_text(encoding="utf-8").startswith("差异报告")
    csv_header = csv_path.read_text(encoding="utf-8-sig").splitlines()[0]
    assert csv_header.startswith("sheet_name,alignment_row")
    assert "base_value" in csv_header
    assert "merged_value" not in csv_header
    assert "merged_row" not in csv_header
    assert "left_vs_right_inline_plain" in csv_header
