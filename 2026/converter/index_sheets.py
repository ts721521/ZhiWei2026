# -*- coding: utf-8 -*-
"""Excel index sheet writers extracted from office_converter.py."""


def write_conversion_index_sheet(
    ws,
    records,
    *,
    style_header_row_fn,
    auto_fit_sheet_fn,
    make_file_hyperlink_fn,
):
    headers = [
        "No.",
        "Source File",
        "Source Path",
        "Source MD5",
        "Output PDF",
        "Output PDF Path",
        "Output PDF MD5",
        "Status",
    ]
    ws.append(headers)
    style_header_row_fn(ws)

    for idx, rec in enumerate(records, 1):
        ws.append(
            [
                idx,
                rec.get("source_filename", ""),
                rec.get("source_abspath", ""),
                rec.get("source_md5", ""),
                rec.get("pdf_filename", ""),
                rec.get("pdf_abspath", ""),
                rec.get("pdf_md5", ""),
                rec.get("status", ""),
            ]
        )
        src_cell = ws.cell(row=idx + 1, column=3)
        src_path = rec.get("source_abspath", "")
        if src_path:
            src_cell.hyperlink = make_file_hyperlink_fn(src_path)
            src_cell.style = "Hyperlink"

        pdf_cell = ws.cell(row=idx + 1, column=6)
        pdf_path = rec.get("pdf_abspath", "")
        if pdf_path:
            pdf_cell.hyperlink = make_file_hyperlink_fn(pdf_path)
            pdf_cell.style = "Hyperlink"

    auto_fit_sheet_fn(ws)


def write_merge_index_sheet(
    ws,
    records,
    *,
    style_header_row_fn,
    auto_fit_sheet_fn,
    make_file_hyperlink_fn,
):
    headers = [
        "No.",
        "Merged PDF",
        "Merged PDF Path",
        "Merged PDF MD5",
        "Source Order",
        "Source PDF",
        "Source PDF Path",
        "Source PDF MD5",
        "Short ID",
        "Page Range",
        "Start Page",
        "End Page",
        "Page Count",
    ]
    ws.append(headers)
    style_header_row_fn(ws)

    for idx, rec in enumerate(records, 1):
        start_page = rec.get("start_page_1based", "")
        end_page = rec.get("end_page_1based", "")
        position = f"{start_page}-{end_page}" if start_page and end_page else ""

        ws.append(
            [
                idx,
                rec.get("merged_pdf_name", ""),
                rec.get("merged_pdf_path", ""),
                rec.get("merged_pdf_md5", ""),
                rec.get("source_index", ""),
                rec.get("source_filename", ""),
                rec.get("source_abspath", ""),
                rec.get("source_md5", ""),
                rec.get("source_short_id", ""),
                position,
                start_page,
                end_page,
                rec.get("page_count", ""),
            ]
        )

        merged_cell = ws.cell(row=idx + 1, column=3)
        merged_path = rec.get("merged_pdf_path", "")
        if merged_path:
            merged_cell.hyperlink = make_file_hyperlink_fn(merged_path)
            merged_cell.style = "Hyperlink"

        source_cell = ws.cell(row=idx + 1, column=7)
        source_path = rec.get("source_abspath", "")
        if source_path:
            source_cell.hyperlink = make_file_hyperlink_fn(source_path)
            source_cell.style = "Hyperlink"

    auto_fit_sheet_fn(ws)
