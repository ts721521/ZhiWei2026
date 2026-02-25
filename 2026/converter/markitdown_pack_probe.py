# -*- coding: utf-8 -*-
"""MarkItDown packaging smoke probe helpers for V6.0 module-1."""

from __future__ import annotations

import os
from typing import Callable, Optional


def _extract_markdown_text(result) -> str:
    for attr in ("text_content", "markdown", "text"):
        value = getattr(result, attr, None)
        if isinstance(value, str) and value.strip():
            return value
    if isinstance(result, str) and result.strip():
        return result
    return ""


def run_markitdown_probe(
    input_path: str,
    output_path: str,
    *,
    markitdown_cls=None,
    makedirs_fn: Callable[..., None] = os.makedirs,
    open_fn=open,
) -> dict:
    src = str(input_path or "").strip()
    dst = str(output_path or "").strip()
    if not src or not dst:
        return {"status": "invalid", "message": "input/output required"}
    if not os.path.exists(src):
        return {"status": "missing_input", "message": f"input not found: {src}"}

    cls = markitdown_cls
    if cls is None:
        try:
            from markitdown import MarkItDown as _MarkItDown

            cls = _MarkItDown
        except (ImportError, ModuleNotFoundError, OSError) as exc:
            return {"status": "markitdown_missing", "message": str(exc)}

    try:
        converter = cls()
        result = converter.convert(src)
        text = _extract_markdown_text(result)
        if not text.strip():
            text = "(empty)"
        makedirs_fn(os.path.dirname(os.path.abspath(dst)), exist_ok=True)
        with open_fn(dst, "w", encoding="utf-8") as f:
            f.write(text.rstrip() + "\n")
        return {"status": "ok", "message": "", "output_path": os.path.abspath(dst)}
    except (OSError, RuntimeError, TypeError, ValueError, AttributeError) as exc:
        return {"status": "probe_failed", "message": str(exc)}
