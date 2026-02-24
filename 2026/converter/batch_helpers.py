# -*- coding: utf-8 -*-
"""Batch processing helper functions extracted from office_converter.py."""

import os


def get_progress_prefix(current, total):
    width = len(str(total)) if total > 0 else 1
    percent = current / total if total else 0
    bar_len = 20
    filled = int(bar_len * percent)
    bar = "#" * filled + "-" * (bar_len - filled)
    return f"[{int(percent * 100):>3}%]{bar} [{str(current).rjust(width)}/{total}]"


def collect_retry_candidates(
    failed_dir,
    allowed_extensions,
    error_records,
    detailed_error_records=None,
):
    retry_files = []
    retry_alias_map = {}

    if not os.path.exists(failed_dir):
        return retry_files, retry_alias_map

    valid_exts = {
        str(e).lower()
        for sub in (allowed_extensions or {}).values()
        for e in (sub or [])
        if isinstance(e, str) and e
    }
    if not valid_exts:
        return retry_files, retry_alias_map

    retryable_sources = None
    if detailed_error_records:
        retryable_sources = {
            os.path.abspath(str(rec.get("source_path", "")))
            for rec in detailed_error_records
            if rec.get("is_retryable")
        }

    name_map = {}
    for src in error_records or []:
        abs_src = os.path.abspath(src)
        if retryable_sources is not None and abs_src not in retryable_sources:
            continue
        name_map.setdefault(os.path.basename(src), []).append(src)

    for f in os.listdir(failed_dir):
        if f.startswith("~$"):
            continue
        ext = os.path.splitext(f)[1].lower()
        if ext not in valid_exts:
            continue

        retry_path = os.path.join(failed_dir, f)
        name = os.path.basename(retry_path)
        mapped_list = name_map.get(name) or []
        if not mapped_list:
            continue

        retry_files.append(retry_path)
        retry_alias_map[retry_path] = mapped_list.pop(0)

    return retry_files, retry_alias_map
