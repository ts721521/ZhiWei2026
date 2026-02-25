# -*- coding: utf-8 -*-
"""Markdown quality report helper extracted from office_converter.py."""

import json
import os


def write_markdown_quality_report(
    *,
    config,
    markdown_quality_records,
    now_fn,
):
    if not config.get("enable_markdown_quality_report", True):
        return None
    if not markdown_quality_records:
        return None

    target_root = config.get("target_folder", "")
    if not target_root:
        return None

    sample_limit = max(1, int(config.get("markdown_quality_sample_limit", 20) or 20))
    now = now_fn()
    ts = now.strftime("%Y%m%d_%H%M%S")
    out_dir = os.path.join(target_root, "_AI", "MarkdownQuality")
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, f"Markdown_Quality_Report_{ts}.json")

    total_pages = 0
    non_empty_pages = 0
    header_removed = 0
    footer_removed = 0
    page_no_removed = 0
    heading_count = 0
    samples = []

    for rec in markdown_quality_records:
        total_pages += int(rec.get("page_count", 0) or 0)
        non_empty_pages += int(rec.get("non_empty_page_count", 0) or 0)
        header_removed += int(rec.get("removed_header_lines", 0) or 0)
        footer_removed += int(rec.get("removed_footer_lines", 0) or 0)
        page_no_removed += int(rec.get("removed_page_number_lines", 0) or 0)
        heading_count += int(rec.get("heading_count", 0) or 0)
        if len(samples) < sample_limit:
            samples.append(
                {
                    "source_pdf": rec.get("source_pdf", ""),
                    "markdown_path": rec.get("markdown_path", ""),
                    "page_count": rec.get("page_count", 0),
                    "non_empty_page_count": rec.get("non_empty_page_count", 0),
                    "removed_header_lines": rec.get("removed_header_lines", 0),
                    "removed_footer_lines": rec.get("removed_footer_lines", 0),
                    "removed_page_number_lines": rec.get("removed_page_number_lines", 0),
                    "heading_count": rec.get("heading_count", 0),
                }
            )

    payload = {
        "version": 1,
        "generated_at": now_fn().isoformat(timespec="seconds"),
        "record_count": len(markdown_quality_records),
        "summary": {
            "total_pages": total_pages,
            "non_empty_pages": non_empty_pages,
            "removed_header_lines": header_removed,
            "removed_footer_lines": footer_removed,
            "removed_page_number_lines": page_no_removed,
            "heading_count": heading_count,
        },
        "sample_limit": sample_limit,
        "samples": samples,
        "records": markdown_quality_records,
    }
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    return out_path


def write_markdown_quality_report_for_converter(converter, *, now_fn, log_info_fn=None):
    out_path = write_markdown_quality_report(
        config=converter.config,
        markdown_quality_records=converter.markdown_quality_records,
        now_fn=now_fn,
    )
    if not out_path:
        return None
    converter.markdown_quality_report_path = out_path
    converter.generated_markdown_quality_outputs.append(out_path)
    if callable(log_info_fn):
        log_info_fn(f"Markdown quality report generated: {out_path}")
    return out_path
