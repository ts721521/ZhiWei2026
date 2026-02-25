# -*- coding: utf-8 -*-
"""Excel content scan helpers extracted from office_converter.py."""


def scan_excel_content_in_thread(
    workbook,
    *,
    price_keywords,
    log_info_fn,
    log_warning_fn,
):
    try:
        for sheet in workbook.Worksheets:
            try:
                data = sheet.UsedRange.Value
                if not data:
                    continue
                if not isinstance(data, tuple):
                    data = ((data,),)
                for row in data:
                    if not row:
                        continue
                    for cell in row:
                        if cell and isinstance(cell, str):
                            for kw in price_keywords:
                                if kw in cell:
                                    log_info_fn(
                                        f"Excel matched keyword [{kw}] in sheet: {sheet.Name}"
                                    )
                                    return True
            except (AttributeError, RuntimeError, TypeError, ValueError):
                continue
    except (AttributeError, RuntimeError, TypeError, ValueError) as exc:
        log_warning_fn(f"scan Excel content failed: {exc}")
    return False
