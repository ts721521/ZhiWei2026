# -*- coding: utf-8 -*-
"""Excel chart helper functions extracted from office_converter.py."""

from converter.excel_json_utils import col_index_to_label


def extract_chart_title_text(chart):
    title_obj = getattr(chart, "title", None)
    if title_obj is None:
        return ""
    if isinstance(title_obj, str):
        return title_obj.strip()

    # openpyxl chart title is often rich text; fall back to str(title_obj).
    try:
        tx = getattr(title_obj, "tx", None)
        rich = getattr(tx, "rich", None) if tx is not None else None
        if rich is not None:
            parts = []
            for para in getattr(rich, "p", []) or []:
                for run in getattr(para, "r", []) or []:
                    txt = getattr(run, "t", None)
                    if txt:
                        parts.append(str(txt))
            if parts:
                return "".join(parts).strip()
    except Exception:
        pass

    try:
        return str(title_obj).strip()
    except Exception:
        return ""


def stringify_chart_anchor(anchor):
    if anchor is None:
        return ""
    try:
        marker = getattr(anchor, "_from", None)
        if marker is not None:
            col = int(getattr(marker, "col", 0)) + 1
            row = int(getattr(marker, "row", 0)) + 1
            return f"{col_index_to_label(col)}{row}"
    except Exception:
        pass
    try:
        return str(anchor)
    except Exception:
        return ""
