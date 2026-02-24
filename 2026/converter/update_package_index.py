# -*- coding: utf-8 -*-
"""Update-package index XLSX helper extracted from office_converter.py."""


def write_update_package_index_xlsx(
    xlsx_path,
    records,
    *,
    has_openpyxl,
    workbook_cls=None,
    font_cls=None,
    make_file_hyperlink_fn=None,
    log_error=None,
):
    if not has_openpyxl or workbook_cls is None or font_cls is None:
        return None
    try:
        wb = workbook_cls()
        ws = wb.active
        ws.title = "IncrementalIndex"
        headers = [
            "seq",
            "change_state",
            "process_status",
            "source_file",
            "source_path",
            "source_md5",
            "source_sha256",
            "renamed_from",
            "rename_match_type",
            "packaged_pdf",
            "packaged_pdf_path",
            "packaged_pdf_md5",
            "note",
        ]
        ws.append(headers)
        for cell in ws[1]:
            cell.font = font_cls(bold=True)

        for rec in records:
            ws.append(
                [
                    rec.get("seq", 0),
                    rec.get("change_state", ""),
                    rec.get("process_status", ""),
                    rec.get("source_file", ""),
                    rec.get("source_path", ""),
                    rec.get("source_md5", ""),
                    rec.get("source_sha256", ""),
                    rec.get("renamed_from", ""),
                    rec.get("rename_match_type", ""),
                    rec.get("packaged_pdf", ""),
                    rec.get("packaged_pdf_path", ""),
                    rec.get("packaged_pdf_md5", ""),
                    rec.get("note", ""),
                ]
            )
            row_idx = ws.max_row
            src_cell = ws.cell(row=row_idx, column=5)
            if rec.get("source_path") and callable(make_file_hyperlink_fn):
                src_cell.hyperlink = make_file_hyperlink_fn(rec["source_path"])
                src_cell.style = "Hyperlink"
            renamed_cell = ws.cell(row=row_idx, column=8)
            if rec.get("renamed_from") and callable(make_file_hyperlink_fn):
                renamed_cell.hyperlink = make_file_hyperlink_fn(rec["renamed_from"])
                renamed_cell.style = "Hyperlink"
            pdf_cell = ws.cell(row=row_idx, column=11)
            if rec.get("packaged_pdf_path") and callable(make_file_hyperlink_fn):
                pdf_cell.hyperlink = make_file_hyperlink_fn(rec["packaged_pdf_path"])
                pdf_cell.style = "Hyperlink"

        for col in ws.columns:
            col_letter = col[0].column_letter
            max_len = 0
            for cell in col:
                value = "" if cell.value is None else str(cell.value)
                if len(value) > max_len:
                    max_len = len(value)
            ws.column_dimensions[col_letter].width = min(max_len + 2, 80)

        wb.save(xlsx_path)
        return xlsx_path
    except Exception as e:
        if callable(log_error):
            log_error(f"[update_package] failed to write XLSX index: {e}")
        return None
