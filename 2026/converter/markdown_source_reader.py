# -*- coding: utf-8 -*-
"""Source-to-markdown text fallback reader extracted from office_converter.py."""

import os

_MARKITDOWN_CONVERT_ERRORS = (
    OSError,
    RuntimeError,
    TypeError,
    ValueError,
    AttributeError,
)
try:
    from markitdown._exceptions import (
        FileConversionException,
        MarkItDownException,
        MissingDependencyException,
        UnsupportedFormatException,
    )

    _MARKITDOWN_CONVERT_ERRORS = _MARKITDOWN_CONVERT_ERRORS + (
        MarkItDownException,
        FileConversionException,
        MissingDependencyException,
        UnsupportedFormatException,
    )
except (ImportError, ModuleNotFoundError, AttributeError):
    pass

try:
    # Some markitdown builds expose FileConversionException under _markitdown.
    from markitdown._markitdown import FileConversionException as _MDFileConversionException

    _MARKITDOWN_CONVERT_ERRORS = _MARKITDOWN_CONVERT_ERRORS + (_MDFileConversionException,)
except (ImportError, ModuleNotFoundError, AttributeError):
    pass


_TEXT_FALLBACK_EXTS = {
    ".txt",
    ".md",
    ".markdown",
    ".csv",
    ".tsv",
    ".json",
    ".yaml",
    ".yml",
    ".xml",
    ".html",
    ".htm",
    ".log",
    ".ini",
    ".cfg",
    ".conf",
    ".rst",
}


def _can_use_plain_text_fallback(source_path):
    ext = os.path.splitext(str(source_path or ""))[1].lower()
    return ext in _TEXT_FALLBACK_EXTS


def convert_source_to_markdown_text(
    source_path,
    *,
    has_markitdown,
    markitdown_cls=None,
    open_fn=open,
):
    if has_markitdown and markitdown_cls is not None:
        md = markitdown_cls()
        try:
            result = md.convert(source_path)
        except _MARKITDOWN_CONVERT_ERRORS as exc:
            raise ValueError(f"markitdown_convert_failed: {exc}")
        except BaseException as exc:
            # markitdown 0.0.x defines some conversion exceptions under BaseException.
            raise ValueError(f"markitdown_convert_failed: {exc}")
        for attr in ("text_content", "markdown", "text"):
            value = getattr(result, attr, None)
            if isinstance(value, str) and value.strip():
                return value
        if isinstance(result, str) and result.strip():
            return result
        raise ValueError("markitdown_empty_output")
    if not _can_use_plain_text_fallback(source_path):
        raise ValueError("markitdown_unavailable_for_binary_source")
    try:
        with open_fn(source_path, "r", encoding="utf-8", errors="ignore") as fh:
            return fh.read()
    except (OSError, RuntimeError, TypeError, ValueError, UnicodeError):
        return ""
