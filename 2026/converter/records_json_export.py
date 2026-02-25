# -*- coding: utf-8 -*-
"""Records JSON export helper extracted from office_converter.py."""

import json
import os


def write_records_json_exports(
    *,
    config,
    conversion_index_records,
    merge_index_records,
    now_fn,
    log_info=None,
):
    if not config.get("enable_excel_json", False):
        return []

    target_root = config.get("target_folder", "")
    if not target_root:
        return []

    ts = now_fn().strftime("%Y%m%d_%H%M%S")
    out_dir = os.path.join(target_root, "_AI", "Records")
    os.makedirs(out_dir, exist_ok=True)

    outputs = []

    if conversion_index_records:
        convert_path = os.path.join(out_dir, f"Convert_Records_{ts}.json")
        with open(convert_path, "w", encoding="utf-8") as f:
            json.dump(
                {
                    "version": 1,
                    "record_type": "convert_index",
                    "record_count": len(conversion_index_records),
                    "records": conversion_index_records,
                },
                f,
                ensure_ascii=False,
                indent=2,
            )
        outputs.append(convert_path)

    if merge_index_records:
        merge_path = os.path.join(out_dir, f"Merge_Records_{ts}.json")
        with open(merge_path, "w", encoding="utf-8") as f:
            json.dump(
                {
                    "version": 1,
                    "record_type": "merge_index",
                    "record_count": len(merge_index_records),
                    "records": merge_index_records,
                },
                f,
                ensure_ascii=False,
                indent=2,
            )
        outputs.append(merge_path)

    if log_info:
        for p in outputs:
            log_info(f"Records JSON generated: {p}")
    return outputs


def write_records_json_exports_for_converter(converter, *, now_fn, log_info=None):
    outputs = write_records_json_exports(
        config=converter.config,
        conversion_index_records=converter.conversion_index_records,
        merge_index_records=converter.merge_index_records,
        now_fn=now_fn,
        log_info=log_info,
    )
    converter.generated_records_json_outputs = outputs
    return outputs
