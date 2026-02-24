# -*- coding: utf-8 -*-
"""Excel JSON export logic extracted from office_converter.py."""

import json
import os
from datetime import datetime

from converter.excel_json_utils import (
    build_column_profiles,
    col_index_to_label,
    extract_formula_sheet_refs,
    is_effectively_empty_row,
    is_empty_json_cell,
    json_safe_value,
    looks_like_header_row,
    normalize_header_row,
)


def _read_positive_int(config, key, default_value):
    try:
        return max(1, int(config.get(key, default_value) or default_value))
    except Exception:
        return max(1, int(default_value))


def _safe_source_relpath(source_excel_path, source_root_resolver):
    try:
        return os.path.relpath(
            os.path.abspath(source_excel_path),
            source_root_resolver(source_excel_path),
        )
    except Exception:
        return os.path.basename(source_excel_path)


def _read_formula_map(
    ws_formula,
    row_index_1based,
    max_cols,
    row_values_json,
    *,
    include_formulas,
    extract_sheet_links,
    formula_sample_limit,
    sheet_name,
):
    formula_map = {}
    formula_cells_count = 0
    formula_samples = []
    sheet_link_counts = {}

    if ws_formula is None or (not include_formulas and not extract_sheet_links):
        return formula_map, formula_cells_count, formula_samples, sheet_link_counts

    for col_index_1based in range(1, max_cols + 1):
        cell_value = ws_formula.cell(row=row_index_1based, column=col_index_1based).value
        if not (isinstance(cell_value, str) and cell_value.startswith("=")):
            continue

        idx0 = col_index_1based - 1
        formula_map[idx0] = cell_value
        formula_cells_count += 1

        if len(formula_samples) < formula_sample_limit:
            address = f"{col_index_to_label(col_index_1based)}{row_index_1based}"
            formula_samples.append(
                {
                    "cell": address,
                    "formula": cell_value,
                    "value": row_values_json[idx0] if idx0 < len(row_values_json) else None,
                }
            )

        if extract_sheet_links:
            refs = extract_formula_sheet_refs(cell_value, sheet_name)
            for ref_sheet in refs:
                sheet_link_counts[ref_sheet] = sheet_link_counts.get(ref_sheet, 0) + 1

    return formula_map, formula_cells_count, formula_samples, sheet_link_counts


def _trim_row_values(row_values_json, row_values_raw, formula_map):
    last_keep_idx = -1
    for idx in range(len(row_values_json)):
        if (not is_empty_json_cell(row_values_json[idx])) or (idx in formula_map):
            last_keep_idx = idx

    if last_keep_idx >= 0:
        return (
            row_values_json[: last_keep_idx + 1],
            row_values_raw[: last_keep_idx + 1],
            formula_map,
        )

    return [], [], {}


def export_single_excel_json(
    source_excel_path,
    *,
    config,
    build_ai_output_path_from_source_fn,
    source_root_resolver,
    has_openpyxl,
    load_workbook_fn,
    extract_workbook_defined_names_fn,
    extract_sheet_charts_fn,
    extract_sheet_pivot_tables_fn,
):
    out_path = build_ai_output_path_from_source_fn(source_excel_path, "ExcelJSON", ".json")
    if not out_path:
        return None

    max_rows = _read_positive_int(config, "excel_json_max_rows", 2000)
    max_cols = _read_positive_int(config, "excel_json_max_cols", 80)
    records_preview_limit = _read_positive_int(config, "excel_json_records_preview", 200)
    profile_rows_limit = _read_positive_int(config, "excel_json_profile_rows", 500)
    include_formulas = bool(config.get("excel_json_include_formulas", True))
    extract_sheet_links = bool(config.get("excel_json_extract_sheet_links", True))
    include_merged_ranges = bool(config.get("excel_json_include_merged_ranges", True))
    formula_sample_limit = _read_positive_int(config, "excel_json_formula_sample_limit", 200)
    merged_range_limit = _read_positive_int(config, "excel_json_merged_range_limit", 500)

    payload = {
        "version": 1,
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "source_excel": os.path.abspath(source_excel_path),
        "source_relpath": _safe_source_relpath(source_excel_path, source_root_resolver),
        "parse_status": "ok",
        "error": "",
        "limits": {
            "max_rows": max_rows,
            "max_cols": max_cols,
            "records_preview_limit": records_preview_limit,
            "profile_rows_limit": profile_rows_limit,
            "include_formulas": include_formulas,
            "extract_sheet_links": extract_sheet_links,
            "include_merged_ranges": include_merged_ranges,
            "formula_sample_limit": formula_sample_limit,
            "merged_range_limit": merged_range_limit,
        },
        "sheets": [],
        "workbook_links": [],
        "workbook_defined_name_count": 0,
        "workbook_defined_names": [],
        "chart_count_total": 0,
        "pivot_table_count_total": 0,
    }

    ext = os.path.splitext(source_excel_path)[1].lower()
    if ext == ".xls":
        payload["parse_status"] = "unsupported_format_xls"
        payload["error"] = "xls not supported by openpyxl"
    elif not has_openpyxl:
        payload["parse_status"] = "openpyxl_missing"
        payload["error"] = "openpyxl not installed"
    else:
        wb_values = None
        wb_formula = None
        try:
            wb_values = load_workbook_fn(source_excel_path, data_only=True, read_only=True)
            if include_formulas or extract_sheet_links or include_merged_ranges:
                wb_formula = load_workbook_fn(
                    source_excel_path,
                    data_only=False,
                    read_only=False,
                )

            workbook_link_counts = {}
            if wb_formula is not None and callable(extract_workbook_defined_names_fn):
                payload["workbook_defined_names"] = extract_workbook_defined_names_fn(wb_formula)
                payload["workbook_defined_name_count"] = len(payload["workbook_defined_names"])

            for ws in wb_values.worksheets:
                ws_formula = None
                if wb_formula is not None and ws.title in wb_formula.sheetnames:
                    ws_formula = wb_formula[ws.title]

                rows_json = []
                rows_raw = []
                rows_meta = []
                row_count_scanned = 0
                row_count_empty_skipped = 0
                truncated = False
                formula_cells_count = 0
                formula_samples = []
                sheet_link_counts = {}

                for row in ws.iter_rows(values_only=True):
                    row_count_scanned += 1
                    if row_count_scanned > max_rows:
                        truncated = True
                        break

                    row_values_raw = list(row)[:max_cols]
                    row_values_json = [json_safe_value(v) for v in row_values_raw]
                    (
                        formula_map,
                        formula_count_row,
                        formula_samples_row,
                        sheet_link_counts_row,
                    ) = _read_formula_map(
                        ws_formula,
                        row_count_scanned,
                        max_cols,
                        row_values_json,
                        include_formulas=include_formulas,
                        extract_sheet_links=extract_sheet_links,
                        formula_sample_limit=formula_sample_limit,
                        sheet_name=ws.title,
                    )
                    formula_cells_count += formula_count_row
                    if formula_samples_row:
                        remaining = formula_sample_limit - len(formula_samples)
                        if remaining > 0:
                            formula_samples.extend(formula_samples_row[:remaining])
                    for ref_sheet, ref_count in sheet_link_counts_row.items():
                        sheet_link_counts[ref_sheet] = (
                            sheet_link_counts.get(ref_sheet, 0) + ref_count
                        )

                    if len(row) > max_cols:
                        truncated = True

                    row_values_json, row_values_raw, formula_map = _trim_row_values(
                        row_values_json,
                        row_values_raw,
                        formula_map,
                    )

                    if is_effectively_empty_row(row_values_json) and not formula_map:
                        row_count_empty_skipped += 1
                        continue

                    rows_json.append(row_values_json)
                    rows_raw.append(row_values_raw)
                    rows_meta.append(
                        {
                            "source_row_index_1based": row_count_scanned,
                            "formulas_by_col_index0": formula_map,
                        }
                    )

                width = 0
                for row_values in rows_json:
                    width = max(width, len(row_values))

                header_detected = False
                header_row_index = None
                header_raw = []
                data_rows_json = rows_json
                data_rows_raw = rows_raw
                data_rows_meta = rows_meta
                if rows_json and looks_like_header_row(rows_json[0]):
                    header_detected = True
                    header_row_index = 0
                    header_raw = rows_json[0]
                    data_rows_json = rows_json[1:]
                    data_rows_raw = rows_raw[1:]
                    data_rows_meta = rows_meta[1:]
                    width = max(width, len(header_raw))

                header = normalize_header_row(header_raw, width)

                records_preview = []
                for row_vals, row_meta in zip(
                    data_rows_json[:records_preview_limit],
                    data_rows_meta[:records_preview_limit],
                ):
                    record = {}
                    formula_named = {}
                    for idx, col_name in enumerate(header):
                        value = row_vals[idx] if idx < len(row_vals) else None
                        record[col_name] = value
                        formula_text = row_meta.get("formulas_by_col_index0", {}).get(idx)
                        if formula_text:
                            formula_named[col_name] = formula_text
                    record["_source_row_index_1based"] = row_meta.get("source_row_index_1based")
                    if formula_named:
                        record["__formulas"] = formula_named
                    records_preview.append(record)

                column_profiles = build_column_profiles(
                    header,
                    data_rows_raw,
                    profile_rows_limit,
                )

                merged_ranges = []
                merged_ranges_truncated = False
                if include_merged_ranges and ws_formula is not None:
                    all_ranges = list(ws_formula.merged_cells.ranges)
                    if len(all_ranges) > merged_range_limit:
                        merged_ranges_truncated = True
                    for merged_range in all_ranges[:merged_range_limit]:
                        merged_ranges.append(
                            {
                                "range": str(merged_range),
                                "top_left_row_1based": int(merged_range.min_row),
                                "top_left_col_1based": int(merged_range.min_col),
                                "top_left_value": json_safe_value(
                                    ws_formula.cell(
                                        row=merged_range.min_row,
                                        column=merged_range.min_col,
                                    ).value
                                ),
                            }
                        )

                linked_sheets = sorted(sheet_link_counts.keys())
                for to_sheet, ref_count in sheet_link_counts.items():
                    edge_key = (ws.title, to_sheet)
                    workbook_link_counts[edge_key] = (
                        workbook_link_counts.get(edge_key, 0) + ref_count
                    )

                charts = (
                    extract_sheet_charts_fn(ws_formula)
                    if callable(extract_sheet_charts_fn)
                    else []
                )
                pivots = (
                    extract_sheet_pivot_tables_fn(ws_formula)
                    if callable(extract_sheet_pivot_tables_fn)
                    else []
                )
                payload["chart_count_total"] += len(charts)
                payload["pivot_table_count_total"] += len(pivots)

                payload["sheets"].append(
                    {
                        "name": ws.title,
                        "row_count_scanned": row_count_scanned,
                        "row_count_exported": len(rows_json),
                        "row_count_empty_skipped": row_count_empty_skipped,
                        "max_cols_exported": max_cols,
                        "truncated": truncated,
                        "header_detected": header_detected,
                        "header_row_index_exported": header_row_index,
                        "header": header,
                        "data_row_count": len(data_rows_json),
                        "rows": rows_json,
                        "records_preview": records_preview,
                        "column_profiles": column_profiles,
                        "formula_stats": {
                            "formula_cells_count": formula_cells_count,
                            "formula_sample_count": len(formula_samples),
                            "formula_samples": formula_samples,
                            "linked_sheets": linked_sheets,
                            "linked_sheet_ref_counts": sheet_link_counts,
                        },
                        "merged_ranges_count": len(merged_ranges),
                        "merged_ranges_truncated": merged_ranges_truncated,
                        "merged_ranges": merged_ranges,
                        "charts_count": len(charts),
                        "charts": charts,
                        "pivot_tables_count": len(pivots),
                        "pivot_tables": pivots,
                    }
                )

            payload["workbook_links"] = [
                {
                    "from_sheet": edge_key[0],
                    "to_sheet": edge_key[1],
                    "ref_count": ref_count,
                }
                for edge_key, ref_count in sorted(
                    workbook_link_counts.items(),
                    key=lambda item: (item[0][0], item[0][1]),
                )
            ]
        except Exception as exc:
            payload["parse_status"] = "parse_failed"
            payload["error"] = str(exc)
        finally:
            if wb_values is not None:
                try:
                    wb_values.close()
                except Exception:
                    pass
            if wb_formula is not None:
                try:
                    wb_formula.close()
                except Exception:
                    pass

    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    return out_path
