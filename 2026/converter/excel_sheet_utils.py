# -*- coding: utf-8 -*-
"""Excel worksheet styling helpers extracted from office_converter.py."""


def style_header_row(ws, has_openpyxl=False, font_cls=None):
    if not has_openpyxl or font_cls is None:
        return
    for cell in ws[1]:
        cell.font = font_cls(bold=True)


def auto_fit_sheet(ws, max_width=90):
    for col in ws.columns:
        max_length = 0
        col_letter = col[0].column_letter
        for cell in col:
            try:
                v = str(cell.value) if cell.value is not None else ""
                max_length = max(max_length, len(v))
            except (TypeError, ValueError, AttributeError):
                pass
        ws.column_dimensions[col_letter].width = min(max_length + 2, max_width)
