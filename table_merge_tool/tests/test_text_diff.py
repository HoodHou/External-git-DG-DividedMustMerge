from __future__ import annotations

from table_merge_tool.text_diff import (
    TextDiffSpan,
    compute_char_diff,
    compute_inline_diff,
    render_diff_html,
    render_diff_inline_html,
    render_diff_plain,
)


def _ops(spans):
    return [(span.op, span.text) for span in spans]


def test_identical_strings_yield_only_equal_spans():
    left_spans, right_spans = compute_char_diff("hello", "hello")
    assert _ops(left_spans) == [("equal", "hello")]
    assert _ops(right_spans) == [("equal", "hello")]


def test_full_replacement():
    left_spans, right_spans = compute_char_diff("abc", "xyz")
    assert _ops(left_spans) == [("delete", "abc")]
    assert _ops(right_spans) == [("insert", "xyz")]


def test_prefix_change_only():
    left_spans, right_spans = compute_char_diff("今天天不好", "今天天挺好")
    assert ("equal", "今天天") in _ops(left_spans)
    assert ("equal", "好") in _ops(left_spans)
    assert ("delete", "不") in _ops(left_spans)
    assert ("insert", "挺") in _ops(right_spans)


def test_empty_inputs():
    left_spans, right_spans = compute_char_diff("", "abc")
    assert _ops(left_spans) == []
    assert _ops(right_spans) == [("insert", "abc")]

    left_spans, right_spans = compute_char_diff("abc", "")
    assert _ops(left_spans) == [("delete", "abc")]
    assert _ops(right_spans) == []

    left_spans, right_spans = compute_char_diff("", "")
    assert _ops(left_spans) == []
    assert _ops(right_spans) == []


def test_single_char_replace_in_cjk():
    left_spans, right_spans = compute_char_diff("今天天气不错", "今天天气真棒")
    assert ("equal", "今天天气") in _ops(left_spans)
    assert ("delete", "不错") in _ops(left_spans)
    assert ("insert", "真棒") in _ops(right_spans)


def test_word_granularity():
    left_spans, right_spans = compute_char_diff(
        "hello world foo", "hello there foo", granularity="word"
    )
    left_ops = _ops(left_spans)
    right_ops = _ops(right_spans)
    assert ("delete", "world") in left_ops
    assert ("insert", "there") in right_ops
    assert any(op == "equal" and "hello" in text for op, text in left_ops)


def test_line_granularity():
    left_text = "line1\nline2\nline3\n"
    right_text = "line1\nline2-modified\nline3\n"
    left_spans, right_spans = compute_char_diff(left_text, right_text, granularity="line")
    left_ops = _ops(left_spans)
    right_ops = _ops(right_spans)
    assert ("delete", "line2\n") in left_ops
    assert ("insert", "line2-modified\n") in right_ops


def test_html_rendering_escapes_special_chars():
    spans = [
        TextDiffSpan("equal", "a "),
        TextDiffSpan("delete", "<b>&"),
        TextDiffSpan("insert", "\"c\""),
    ]
    result = render_diff_html(spans)
    assert "&lt;b&gt;&amp;" in result
    assert "&quot;c&quot;" in result
    assert '<span class="diff-del">' in result
    assert '<span class="diff-ins">' in result


def test_html_rendering_converts_newlines_to_br():
    spans = [TextDiffSpan("equal", "a\nb")]
    assert render_diff_html(spans) == "a<br>b"


def test_plain_rendering_marks_changes():
    spans = [
        TextDiffSpan("equal", "hello "),
        TextDiffSpan("delete", "old"),
        TextDiffSpan("insert", "new"),
        TextDiffSpan("equal", " world"),
    ]
    assert render_diff_plain(spans) == "hello [-old-]{+new+} world"


def test_inline_diff_single_stream():
    spans = compute_inline_diff("今天天不好", "今天天挺好")
    ops = _ops(spans)
    assert ops == [
        ("equal", "今天天"),
        ("delete", "不"),
        ("insert", "挺"),
        ("equal", "好"),
    ]


def test_inline_diff_full_replace():
    spans = compute_inline_diff("abc", "xyz")
    ops = _ops(spans)
    assert ops == [("delete", "abc"), ("insert", "xyz")]


def test_empty_spans_are_skipped_in_rendering():
    spans = [TextDiffSpan("equal", ""), TextDiffSpan("delete", "x")]
    assert render_diff_html(spans) == '<span class="diff-del">x</span>'
    assert render_diff_plain(spans) == "[-x-]"


def test_render_inline_html_uses_style_attributes_for_qtooltip():
    spans = [
        TextDiffSpan("equal", "hello "),
        TextDiffSpan("delete", "old"),
        TextDiffSpan("insert", "new"),
    ]
    html = render_diff_inline_html(spans)
    assert "class=" not in html
    assert 'style="background-color:#ffebe9' in html
    assert 'style="background-color:#dafbe1' in html
    assert "hello " in html
