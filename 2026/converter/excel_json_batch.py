# -*- coding: utf-8 -*-
"""Batch Excel JSON export orchestration extracted from office_converter.py."""

import os


def write_excel_structured_json_exports(
    converter,
    *,
    abspath_fn=os.path.abspath,
    exists_fn=os.path.exists,
    splitext_fn=os.path.splitext,
    log_info_fn=None,
    log_error_fn=None,
):
    if not converter.config.get("enable_excel_json", False):
        return []

    excel_exts = set(
        ext.lower()
        for ext in converter.config.get("allowed_extensions", {}).get("excel", [])
    )
    source_paths = []
    seen = set()
    for rec in converter.conversion_index_records:
        src = abspath_fn(rec.get("source_abspath", "") or "")
        if not src or not exists_fn(src):
            continue
        if splitext_fn(src)[1].lower() not in excel_exts:
            continue
        if src in seen:
            continue
        seen.add(src)
        source_paths.append(src)

    outputs = []
    for src in source_paths:
        try:
            out_path = converter._export_single_excel_json(src)
            if out_path and exists_fn(out_path):
                outputs.append(out_path)
                if log_info_fn:
                    log_info_fn(f"Excel JSON generated: {out_path}")
        except (OSError, RuntimeError, TypeError, ValueError, AttributeError) as exc:
            if log_error_fn:
                log_error_fn(f"Excel JSON export failed {src}: {exc}")

    converter.generated_excel_json_outputs = outputs
    return outputs
