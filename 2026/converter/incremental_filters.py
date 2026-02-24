# -*- coding: utf-8 -*-
"""Incremental/scan filter helpers extracted from office_converter.py."""

import os


def apply_source_priority_filter(files, config, is_win_fn=None, log_info=None):
    if is_win_fn is None:
        is_win_fn = lambda: False

    if not config.get("source_priority_skip_same_name_pdf", False):
        return files, []

    office_exts = set()
    office_exts.update(e.lower() for e in config.get("allowed_extensions", {}).get("word", []))
    office_exts.update(e.lower() for e in config.get("allowed_extensions", {}).get("excel", []))
    office_exts.update(
        e.lower() for e in config.get("allowed_extensions", {}).get("powerpoint", [])
    )

    office_keys = {}
    for p in files:
        ext = os.path.splitext(p)[1].lower()
        if ext not in office_exts:
            continue
        parent = os.path.abspath(os.path.dirname(p))
        if is_win_fn():
            parent = parent.lower()
        stem = os.path.splitext(os.path.basename(p))[0].lower()
        office_keys[(parent, stem)] = p

    kept = []
    skipped = []
    for p in files:
        ext = os.path.splitext(p)[1].lower()
        if ext != ".pdf":
            kept.append(p)
            continue

        parent = os.path.abspath(os.path.dirname(p))
        if is_win_fn():
            parent = parent.lower()
        stem = os.path.splitext(os.path.basename(p))[0].lower()
        office_path = office_keys.get((parent, stem))
        if office_path:
            skipped.append(
                {
                    "source_path": os.path.abspath(p),
                    "status": "source_priority_skipped",
                    "detail": "same_dir_same_stem_office_exists",
                    "final_path": "",
                    "preferred_source": os.path.abspath(office_path),
                }
            )
            continue

        kept.append(p)

    if skipped and callable(log_info):
        log_info(f"[source_priority] 跳过同目录同名 PDF: {len(skipped)}，保留 Office 版本")
    return kept, skipped


def apply_global_md5_dedup(
    files,
    enabled,
    ext_bucket_fn,
    compute_md5_fn,
    log_warning=None,
    log_info=None,
):
    if not enabled:
        return files, []

    seen = {}
    kept = []
    skipped = []

    for path in files:
        bucket = ext_bucket_fn(path)
        try:
            md5_value = compute_md5_fn(path)
        except Exception as e:
            if callable(log_warning):
                log_warning(f"[global_md5] failed to compute MD5, keep file: {path} | {e}")
            kept.append(path)
            continue

        key = (bucket, md5_value)
        if key in seen:
            skipped.append(
                {
                    "source_path": os.path.abspath(path),
                    "status": "dedup_skipped",
                    "detail": "same_type_same_md5",
                    "final_path": "",
                    "md5": md5_value,
                    "duplicate_of": os.path.abspath(seen[key]),
                }
            )
            continue

        seen[key] = path
        kept.append(path)

    if skipped and callable(log_info):
        log_info(f"[global_md5] skipped duplicate source files: {len(skipped)}")
    return kept, skipped
