# -*- coding: utf-8 -*-
"""Failed-file report export helpers extracted from office_converter.py."""

import json
import os
from datetime import datetime


def export_failed_files_report(
    detailed_error_records,
    output_dir,
    *,
    run_mode,
    now_fn=None,
    log_error=None,
):
    if not detailed_error_records:
        return {"json_path": None, "txt_path": None, "summary": "no_failed_records"}

    if now_fn is None:
        now_fn = datetime.now

    os.makedirs(output_dir, exist_ok=True)
    timestamp = now_fn().strftime("%Y%m%d_%H%M%S")

    json_path = os.path.join(output_dir, f"failed_files_report_{timestamp}.json")
    report_data = {
        "generated_at": now_fn().isoformat(),
        "run_mode": run_mode,
        "total_failed": len(detailed_error_records),
        "statistics": {
            "by_error_type": {},
            "by_category": {"retryable": 0, "needs_manual": 0, "unrecoverable": 0},
            "retryable_count": 0,
            "manual_action_count": 0,
        },
        "records": detailed_error_records,
    }

    for record in detailed_error_records:
        et = str(record.get("error_type", "unknown"))
        report_data["statistics"]["by_error_type"][et] = (
            report_data["statistics"]["by_error_type"].get(et, 0) + 1
        )
        cat = str(record.get("error_category", ""))
        if cat in report_data["statistics"]["by_category"]:
            report_data["statistics"]["by_category"][cat] += 1
        if bool(record.get("is_retryable", False)):
            report_data["statistics"]["retryable_count"] += 1
        if bool(record.get("requires_manual_action", False)):
            report_data["statistics"]["manual_action_count"] += 1

    try:
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(report_data, f, indent=2, ensure_ascii=False)
    except Exception as e:
        if callable(log_error):
            log_error(f"Failed to write JSON report: {e}")
        json_path = None

    txt_path = os.path.join(output_dir, f"failed_files_report_{timestamp}.txt")
    try:
        lines = _build_failed_report_text_lines(
            detailed_error_records,
            report_data,
            run_mode=run_mode,
            json_path=json_path,
            now_fn=now_fn,
        )
        with open(txt_path, "w", encoding="utf-8") as f:
            f.write("\n".join(lines))
    except Exception as e:
        if callable(log_error):
            log_error(f"Failed to write TXT report: {e}")
        txt_path = None

    return {
        "json_path": json_path,
        "txt_path": txt_path,
        "summary": (
            f"total_failed={len(detailed_error_records)}, "
            f"retryable={report_data['statistics']['retryable_count']}"
        ),
    }


def _build_failed_report_text_lines(
    detailed_error_records,
    report_data,
    *,
    run_mode,
    json_path,
    now_fn,
):
    lines = []
    lines.append("=" * 70)
    lines.append("ZhiWei - Failed Conversion Report")
    lines.append(f"Generated At: {now_fn().strftime('%Y-%m-%d %H:%M:%S')}")
    lines.append(f"Run Mode: {run_mode}")
    lines.append("=" * 70)
    lines.append("")

    lines.append("## Summary")
    lines.append("-" * 70)
    lines.append(f"Total Failed: {len(detailed_error_records)}")
    lines.append(f"Retryable: {report_data['statistics']['retryable_count']}")
    lines.append(f"Needs Manual Action: {report_data['statistics']['manual_action_count']}")
    lines.append("")

    lines.append("### By Error Type")
    for et, count in sorted(
        report_data["statistics"]["by_error_type"].items(), key=lambda x: -x[1]
    ):
        lines.append(f"  - {et}: {count}")
    lines.append("")

    lines.append("## Failed Files")
    lines.append("-" * 70)
    for i, record in enumerate(detailed_error_records, 1):
        lines.append(f"\n[{i}] {record.get('file_name', '')}")
        lines.append(f"    source: {record.get('source_path', '')}")
        lines.append(f"    error_type: {record.get('error_type', '')}")
        lines.append(f"    message: {record.get('message', '')}")
        lines.append(f"    retryable: {'yes' if record.get('is_retryable') else 'no'}")
        if record.get("requires_manual_action"):
            lines.append("    suggestion:")
            for line in str(record.get("suggestion", "")).split("\n"):
                lines.append(f"        {line}")

    lines.append("")
    lines.append("=" * 70)
    lines.append("Tips:")
    lines.append("- Retryable files can be retried from the UI.")
    lines.append("- Manual-action files should be fixed before rerun.")
    lines.append("- JSON detail report: " + (json_path or "generation_failed"))
    lines.append("=" * 70)
    return lines
