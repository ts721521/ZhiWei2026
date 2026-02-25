# -*- coding: utf-8 -*-
"""Traceability helpers for short-id normalization and trace map export."""

import os
from datetime import datetime


def apply_short_id_prefix(short_id, prefix="ZW-"):
    sid = str(short_id or "").strip().upper()
    pfx = str(prefix or "").strip().upper()
    if not sid:
        return ""
    if not pfx:
        return sid
    return sid if sid.startswith(pfx) else f"{pfx}{sid}"


def strip_short_id_prefix(short_id, prefix="ZW-"):
    sid = str(short_id or "").strip().upper()
    pfx = str(prefix or "").strip().upper()
    if pfx and sid.startswith(pfx):
        return sid[len(pfx) :]
    return sid


def normalize_short_id_for_match(short_id, prefix="ZW-"):
    return strip_short_id_prefix(short_id, prefix=prefix)


def write_trace_map_xlsx(
    trace_records,
    *,
    output_path,
    has_openpyxl,
    workbook_cls=None,
    load_workbook_fn=None,
    font_cls=None,
    style_header_row_fn=None,
    auto_fit_sheet_fn=None,
    log_warning_fn=None,
):
    if not trace_records:
        return None
    if not has_openpyxl or workbook_cls is None or load_workbook_fn is None:
        if log_warning_fn:
            log_warning_fn("openpyxl missing; trace_map.xlsx skipped.")
        return None

    fields = [
        "source_short_id",
        "source_filename",
        "source_abspath",
        "source_relpath",
        "source_md5",
        "source_type",
        "updated_at",
    ]

    existing_by_sid = {}
    if os.path.exists(output_path):
        try:
            wb_old = load_workbook_fn(output_path)
            ws_old = wb_old.active
            header = [str(c.value or "").strip() for c in ws_old[1]]
            idx = {name: i for i, name in enumerate(header)}
            for row in ws_old.iter_rows(min_row=2, values_only=True):
                sid = str(row[idx.get("source_short_id", -1)] or "").strip().upper()
                if not sid:
                    continue
                existing_by_sid[sid] = {
                    "source_short_id": sid,
                    "source_filename": str(
                        row[idx.get("source_filename", -1)] or ""
                    ).strip(),
                    "source_abspath": str(
                        row[idx.get("source_abspath", -1)] or ""
                    ).strip(),
                    "source_relpath": str(
                        row[idx.get("source_relpath", -1)] or ""
                    ).strip(),
                    "source_md5": str(row[idx.get("source_md5", -1)] or "").strip(),
                    "source_type": str(
                        row[idx.get("source_type", -1)] or ""
                    ).strip(),
                    "updated_at": str(row[idx.get("updated_at", -1)] or "").strip(),
                }
        except (OSError, RuntimeError, TypeError, ValueError, KeyError, IndexError, AttributeError):
            existing_by_sid = {}

    for rec in trace_records:
        sid = str(rec.get("source_short_id", "") or "").strip().upper()
        if not sid:
            continue
        existing_by_sid[sid] = {
            "source_short_id": sid,
            "source_filename": str(rec.get("source_filename", "") or "").strip(),
            "source_abspath": str(rec.get("source_abspath", "") or "").strip(),
            "source_relpath": str(rec.get("source_relpath", "") or "").strip(),
            "source_md5": str(rec.get("source_md5", "") or "").strip(),
            "source_type": str(rec.get("source_type", "") or "").strip(),
            "updated_at": str(rec.get("updated_at", "") or "").strip(),
        }

    wb = workbook_cls()
    ws = wb.active
    ws.title = "TraceMap"
    ws.append(fields)
    for sid in sorted(existing_by_sid.keys()):
        row = existing_by_sid[sid]
        ws.append([row.get(k, "") for k in fields])

    if style_header_row_fn:
        style_header_row_fn(ws)
    elif font_cls:
        for c in ws[1]:
            c.font = font_cls(bold=True)
    if auto_fit_sheet_fn:
        auto_fit_sheet_fn(ws)

    wb.save(output_path)
    return output_path


def write_trace_map_for_converter(
    converter,
    *,
    output_name="trace_map.xlsx",
    now_fn=datetime.now,
    join_fn=os.path.join,
    has_openpyxl=False,
    workbook_cls=None,
    load_workbook_fn=None,
    font_cls=None,
    style_header_row_fn=None,
    auto_fit_sheet_fn=None,
    log_warning_fn=None,
    log_info_fn=None,
):
    if not converter.config.get("enable_traceability_anchor_and_map", True):
        return None
    target_root = converter.config.get("target_folder", "")
    if not target_root:
        return None

    now_iso = now_fn().isoformat(timespec="seconds")
    prefix = converter.config.get("short_id_prefix", "ZW-")
    trace_records = []

    for rec in converter.merge_index_records:
        sid = apply_short_id_prefix(rec.get("source_short_id", ""), prefix)
        if not sid:
            continue
        trace_records.append(
            {
                "source_short_id": sid,
                "source_filename": rec.get("source_filename", ""),
                "source_abspath": rec.get("source_abspath", ""),
                "source_relpath": rec.get("source_relpath", ""),
                "source_md5": rec.get("source_md5", ""),
                "source_type": "merge_map",
                "updated_at": now_iso,
            }
        )

    for rec in converter.markdown_quality_records:
        sid = apply_short_id_prefix(rec.get("source_short_id", ""), prefix)
        if not sid:
            continue
        trace_records.append(
            {
                "source_short_id": sid,
                "source_filename": rec.get("source_filename", ""),
                "source_abspath": rec.get("source_abspath", rec.get("source_pdf", "")),
                "source_relpath": "",
                "source_md5": rec.get("source_md5", ""),
                "source_type": "markdown",
                "updated_at": now_iso,
            }
        )

    out_path = write_trace_map_xlsx(
        trace_records,
        output_path=join_fn(target_root, output_name),
        has_openpyxl=has_openpyxl,
        workbook_cls=workbook_cls,
        load_workbook_fn=load_workbook_fn,
        font_cls=font_cls,
        style_header_row_fn=style_header_row_fn,
        auto_fit_sheet_fn=auto_fit_sheet_fn,
        log_warning_fn=log_warning_fn,
    )
    if out_path:
        converter.trace_map_path = out_path
        converter.generated_trace_map_outputs = [out_path]
        if log_info_fn:
            log_info_fn(f"trace_map generated: {out_path}")
    return out_path
