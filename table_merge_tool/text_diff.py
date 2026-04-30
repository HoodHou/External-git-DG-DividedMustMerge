from __future__ import annotations

import html
from dataclasses import dataclass
from difflib import SequenceMatcher


@dataclass(frozen=True, slots=True)
class TextDiffSpan:
    op: str  # "equal" | "insert" | "delete"
    text: str


Granularity = str  # "char" | "word" | "line"


def _tokenize(value: str, granularity: str) -> list[str]:
    if not value:
        return []
    if granularity == "line":
        parts = value.splitlines(keepends=True)
        return parts if parts else [value]
    if granularity == "word":
        tokens: list[str] = []
        buffer: list[str] = []
        mode: str | None = None
        for ch in value:
            kind = "word" if ch.isalnum() or ch == "_" else ("space" if ch.isspace() else "punct")
            if mode is None:
                mode = kind
                buffer.append(ch)
                continue
            if kind == mode and kind == "word":
                buffer.append(ch)
            else:
                tokens.append("".join(buffer))
                buffer = [ch]
                mode = kind
        if buffer:
            tokens.append("".join(buffer))
        return tokens
    return list(value)


def compute_char_diff(
    left: str,
    right: str,
    *,
    granularity: Granularity = "char",
) -> tuple[list[TextDiffSpan], list[TextDiffSpan]]:
    """Return ``(left_spans, right_spans)``.

    ``left_spans`` is what to render under the left value: ``equal`` segments
    are shared content, ``delete`` segments are text that only exists on the
    left (shown with strikethrough). ``right_spans`` mirrors: ``equal`` shared,
    ``insert`` only on the right (highlighted green).

    ``replace`` operations from :class:`difflib.SequenceMatcher` are expanded
    into paired ``delete`` (left side) + ``insert`` (right side), which is what
    human readers expect when comparing two near-identical strings.
    """
    left_tokens = _tokenize(left, granularity)
    right_tokens = _tokenize(right, granularity)
    matcher = SequenceMatcher(a=left_tokens, b=right_tokens, autojunk=False)
    left_spans: list[TextDiffSpan] = []
    right_spans: list[TextDiffSpan] = []
    for op, i1, i2, j1, j2 in matcher.get_opcodes():
        left_segment = "".join(left_tokens[i1:i2])
        right_segment = "".join(right_tokens[j1:j2])
        if op == "equal":
            if left_segment:
                left_spans.append(TextDiffSpan("equal", left_segment))
            if right_segment:
                right_spans.append(TextDiffSpan("equal", right_segment))
        elif op == "delete":
            if left_segment:
                left_spans.append(TextDiffSpan("delete", left_segment))
        elif op == "insert":
            if right_segment:
                right_spans.append(TextDiffSpan("insert", right_segment))
        elif op == "replace":
            if left_segment:
                left_spans.append(TextDiffSpan("delete", left_segment))
            if right_segment:
                right_spans.append(TextDiffSpan("insert", right_segment))
    return left_spans, right_spans


def compute_inline_diff(
    left: str,
    right: str,
    *,
    granularity: Granularity = "char",
) -> list[TextDiffSpan]:
    """Single-stream diff: a unified sequence of ``equal``/``delete``/``insert``
    spans suitable for rendering an inline ``[-del-]{+ins+}`` view (e.g. on the
    merged value row of a report).
    """
    left_tokens = _tokenize(left, granularity)
    right_tokens = _tokenize(right, granularity)
    matcher = SequenceMatcher(a=left_tokens, b=right_tokens, autojunk=False)
    spans: list[TextDiffSpan] = []
    for op, i1, i2, j1, j2 in matcher.get_opcodes():
        if op == "equal":
            segment = "".join(left_tokens[i1:i2])
            if segment:
                spans.append(TextDiffSpan("equal", segment))
        elif op == "delete":
            segment = "".join(left_tokens[i1:i2])
            if segment:
                spans.append(TextDiffSpan("delete", segment))
        elif op == "insert":
            segment = "".join(right_tokens[j1:j2])
            if segment:
                spans.append(TextDiffSpan("insert", segment))
        elif op == "replace":
            left_segment = "".join(left_tokens[i1:i2])
            right_segment = "".join(right_tokens[j1:j2])
            if left_segment:
                spans.append(TextDiffSpan("delete", left_segment))
            if right_segment:
                spans.append(TextDiffSpan("insert", right_segment))
    return spans


_CSS_CLASS = {
    "equal": "diff-eq",
    "delete": "diff-del",
    "insert": "diff-ins",
}


def render_diff_html(spans: list[TextDiffSpan]) -> str:
    parts: list[str] = []
    for span in spans:
        if not span.text:
            continue
        escaped = html.escape(span.text, quote=True).replace("\n", "<br>")
        if span.op == "equal":
            parts.append(escaped)
        else:
            parts.append(f'<span class="{_CSS_CLASS[span.op]}">{escaped}</span>')
    return "".join(parts)


_INLINE_STYLE = {
    "equal": "",
    "delete": "background-color:#ffebe9;color:#82071e;text-decoration:line-through;",
    "insert": "background-color:#dafbe1;color:#0b5d1e;",
}


def render_diff_inline_html(spans: list[TextDiffSpan]) -> str:
    """Like :func:`render_diff_html` but using inline ``style`` attributes so
    it renders correctly inside ``QToolTip`` (which does not honor ``class``).
    """
    parts: list[str] = []
    for span in spans:
        if not span.text:
            continue
        escaped = html.escape(span.text, quote=True).replace("\n", "<br>")
        style = _INLINE_STYLE.get(span.op, "")
        if style:
            parts.append(f'<span style="{style}">{escaped}</span>')
        else:
            parts.append(escaped)
    return "".join(parts)


def render_diff_plain(spans: list[TextDiffSpan]) -> str:
    parts: list[str] = []
    for span in spans:
        if not span.text:
            continue
        if span.op == "equal":
            parts.append(span.text)
        elif span.op == "delete":
            parts.append(f"[-{span.text}-]")
        elif span.op == "insert":
            parts.append(f"{{+{span.text}+}}")
    return "".join(parts)
