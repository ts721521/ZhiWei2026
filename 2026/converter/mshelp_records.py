# -*- coding: utf-8 -*-
"""MSHelp record/index helpers extracted from office_converter.py."""

import csv
import json
import os
from datetime import datetime
from pathlib import Path


def build_mshelp_record(
    source_cab_path,
    markdown_path,
    topic_count,
    *,
    folder_name="MSHelpViewer",
    get_source_root_for_path_fn=None,
):
    src_abs = os.path.abspath(source_cab_path)
    md_abs = os.path.abspath(markdown_path)
    folder_name = str(folder_name or "MSHelpViewer").strip() or "MSHelpViewer"
    folder_name_lower = folder_name.lower()

    mshelp_dir = ""
    try:
        p = Path(src_abs)
        for parent in [p.parent, *p.parents]:
            if parent.name.lower() == folder_name_lower:
                mshelp_dir = str(parent)
                break
    except Exception:
        mshelp_dir = ""

    try:
        root = (
            get_source_root_for_path_fn(src_abs)
            if callable(get_source_root_for_path_fn)
            else ""
        )
        src_rel = os.path.relpath(src_abs, root) if root else src_abs
    except Exception:
        src_rel = src_abs

    return {
        "source_cab": src_abs,
        "source_cab_relpath": src_rel,
        "mshelpviewer_dir": mshelp_dir,
        "markdown_path": md_abs,
        "topic_count": int(topic_count or 0),
        "status": "success",
    }


def write_mshelp_index_files(
    mshelp_records,
    target_root,
    *,
    generated_mshelp_outputs=None,
    log_info=None,
):
    if not mshelp_records:
        return []
    if not target_root:
        return []

    out_dir = os.path.join(target_root, "_AI", "MSHelp")
    os.makedirs(out_dir, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    json_path = os.path.join(out_dir, f"MSHelp_Index_{ts}.json")
    csv_path = os.path.join(out_dir, f"MSHelp_Index_{ts}.csv")

    records = sorted(
        mshelp_records,
        key=lambda x: (x.get("mshelpviewer_dir", ""), x.get("source_cab", "")),
    )
    payload = {
        "version": 1,
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "record_count": len(records),
        "records": records,
    }
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)

    fields = [
        "source_cab",
        "source_cab_relpath",
        "mshelpviewer_dir",
        "markdown_path",
        "topic_count",
        "status",
    ]
    with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writeheader()
        writer.writerows(records)

    outputs = [json_path, csv_path]
    if isinstance(generated_mshelp_outputs, list):
        generated_mshelp_outputs.extend(outputs)
    if callable(log_info):
        log_info(f"MSHelp index generated: {json_path}")
        log_info(f"MSHelp index generated: {csv_path}")
    return outputs
