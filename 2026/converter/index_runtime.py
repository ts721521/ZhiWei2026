# -*- coding: utf-8 -*-
"""Index and merge-map runtime helpers extracted from office_converter.py."""


def write_merge_map(output_path, records, *, csv_module, json_module, open_fn=open):
    base_no_ext = output_path[:-4] if output_path.lower().endswith(".pdf") else output_path
    csv_path = f"{base_no_ext}.map.csv"
    json_path = f"{base_no_ext}.map.json"
    fields = [
        "merge_batch_id",
        "merged_pdf_name",
        "merged_pdf_path",
        "merged_pdf_md5",
        "source_index",
        "source_filename",
        "source_abspath",
        "source_relpath",
        "source_md5",
        "source_short_id",
        "start_page_1based",
        "end_page_1based",
        "page_count",
        "bookmark_title",
    ]

    with open_fn(csv_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv_module.DictWriter(f, fieldnames=fields)
        writer.writeheader()
        writer.writerows(records)

    with open_fn(json_path, "w", encoding="utf-8") as f:
        json_module.dump(
            {
                "version": 1,
                "record_count": len(records),
                "records": records,
            },
            f,
            ensure_ascii=False,
            indent=2,
        )
    return csv_path, json_path


def append_conversion_index_record(
    source_path,
    pdf_path,
    *,
    status="",
    exists_fn,
    abspath_fn,
    basename_fn,
    relpath_fn,
    compute_md5_fn,
    get_source_root_for_path_fn,
    target_folder,
    conversion_index_records,
):
    if not source_path or not pdf_path:
        return
    if not exists_fn(pdf_path):
        return

    src_abs = abspath_fn(source_path)
    pdf_abs = abspath_fn(pdf_path)

    try:
        src_md5 = compute_md5_fn(src_abs)
    except (OSError, TypeError, ValueError, RuntimeError):
        src_md5 = ""
    try:
        pdf_md5 = compute_md5_fn(pdf_abs)
    except (OSError, TypeError, ValueError, RuntimeError):
        pdf_md5 = ""

    try:
        src_rel = relpath_fn(src_abs, get_source_root_for_path_fn(src_abs))
    except (OSError, TypeError, ValueError, RuntimeError):
        src_rel = src_abs
    try:
        pdf_rel = relpath_fn(pdf_abs, target_folder)
    except (OSError, TypeError, ValueError, RuntimeError):
        pdf_rel = pdf_abs

    conversion_index_records.append(
        {
            "source_filename": basename_fn(src_abs),
            "source_abspath": src_abs,
            "source_relpath": src_rel,
            "source_md5": src_md5,
            "pdf_filename": basename_fn(pdf_abs),
            "pdf_abspath": pdf_abs,
            "pdf_relpath": pdf_rel,
            "pdf_md5": pdf_md5,
            "status": status or "",
        }
    )


def write_conversion_index_workbook(
    conversion_index_records,
    *,
    config,
    has_openpyxl,
    workbook_cls,
    write_conversion_index_sheet_fn,
    now_fn,
    join_path_fn,
    print_fn=print,
    log_warning_fn=None,
    log_info_fn=None,
):
    if not conversion_index_records:
        return None
    if not has_openpyxl:
        print_fn("\n[INFO] openpyxl not found. Conversion index Excel will not be generated.")
        if log_warning_fn:
            log_warning_fn("openpyxl missing; conversion index Excel skipped.")
        return None

    timestamp = now_fn().strftime("%Y%m%d_%H%M%S")
    index_path = join_path_fn(config["target_folder"], f"Convert_Index_{timestamp}.xlsx")
    wb = workbook_cls()
    ws = wb.active
    ws.title = "ConvertedPDFs"
    write_conversion_index_sheet_fn(ws, conversion_index_records)
    wb.save(index_path)
    print_fn(f"\nConversion index saved: {index_path}")
    if log_info_fn:
        log_info_fn(f"Conversion index saved: {index_path}")
    return index_path


def write_conversion_index_workbook_for_converter(
    converter,
    *,
    has_openpyxl,
    workbook_cls,
    now_fn,
    join_path_fn,
    print_fn=print,
    log_warning_fn=None,
    log_info_fn=None,
):
    index_path = write_conversion_index_workbook(
        converter.conversion_index_records,
        config=converter.config,
        has_openpyxl=has_openpyxl,
        workbook_cls=workbook_cls,
        write_conversion_index_sheet_fn=converter._write_conversion_index_sheet,
        now_fn=now_fn,
        join_path_fn=join_path_fn,
        print_fn=print_fn,
        log_warning_fn=log_warning_fn,
        log_info_fn=log_info_fn,
    )
    if not index_path:
        return None
    converter.convert_index_path = index_path
    return index_path


def write_merge_map_for_converter(converter, output_path, records, *, csv_module, json_module, open_fn=open):
    if not records:
        return None, None
    return write_merge_map(
        output_path,
        records,
        csv_module=csv_module,
        json_module=json_module,
        open_fn=open_fn,
    )
