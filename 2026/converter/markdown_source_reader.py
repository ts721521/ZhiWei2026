# -*- coding: utf-8 -*-
"""Source-to-markdown text fallback reader extracted from office_converter.py."""


def convert_source_to_markdown_text(
    source_path,
    *,
    has_markitdown,
    markitdown_cls=None,
    open_fn=open,
):
    if has_markitdown and markitdown_cls is not None:
        md = markitdown_cls()
        result = md.convert(source_path)
        for attr in ("text_content", "markdown", "text"):
            value = getattr(result, attr, None)
            if isinstance(value, str) and value.strip():
                return value
        if isinstance(result, str) and result.strip():
            return result
    try:
        with open_fn(source_path, "r", encoding="utf-8", errors="ignore") as fh:
            return fh.read()
    except (OSError, RuntimeError, TypeError, ValueError, UnicodeError):
        return ""
