# -*- coding: utf-8 -*-
"""Convert-scan candidate helper extracted from office_converter.py."""

import os
from datetime import datetime


def scan_convert_candidates(
    config,
    configured_roots,
    *,
    probe_source_root_access_fn,
    record_scan_access_skip_fn,
    filter_date=None,
    filter_mode="after",
    isdir_fn=None,
    walk_fn=None,
    getctime_fn=None,
    print_warn_fn=None,
    log_error_fn=None,
):
    if isdir_fn is None:
        isdir_fn = os.path.isdir
    if walk_fn is None:
        walk_fn = os.walk
    if getctime_fn is None:
        getctime_fn = os.path.getctime

    files = []
    scan_skip_seen = set()
    source_roots = []
    for source_root in configured_roots:
        if probe_source_root_access_fn(
            source_root, context={"scan_scope": "convert"}, seen_keys=scan_skip_seen
        ):
            source_roots.append(os.path.abspath(source_root))

    if not source_roots:
        single = config.get("source_folder", "")
        msg = f"\n[WARN] Source directory does not exist or empty: {single}"
        if callable(print_warn_fn):
            print_warn_fn(msg)
        if callable(log_error_fn):
            log_error_fn("source directory does not exist or source_folders empty")
        return files

    excl_config = config.get("excluded_folders", [])
    excl_names = {
        x.lower() for x in excl_config if not os.path.isabs(x) and os.sep not in x and "/" not in x
    }
    excl_paths = {
        os.path.abspath(x).lower()
        for x in excl_config
        if os.path.isabs(x) or os.sep in x or "/" in x
    }

    valid_exts = set()
    for sub in config.get("allowed_extensions", {}).values():
        if isinstance(sub, list):
            for ext in sub:
                if isinstance(ext, str) and ext:
                    valid_exts.add(ext.lower())

    for source_folder in source_roots:
        if not isdir_fn(source_folder):
            continue
        for root, dirs, filenames in walk_fn(
            source_folder,
            onerror=lambda e, sf=source_folder: record_scan_access_skip_fn(
                getattr(e, "filename", sf),
                e,
                context={"scan_scope": "convert", "source_root": sf},
                seen_keys=scan_skip_seen,
            ),
        ):
            dirs[:] = [
                d
                for d in dirs
                if d.lower() not in excl_names
                and os.path.abspath(os.path.join(root, d)).lower() not in excl_paths
            ]
            for fname in filenames:
                if fname.startswith("~$"):
                    continue
                ext = os.path.splitext(fname)[1].lower()
                if ext not in valid_exts:
                    continue

                full_path = os.path.join(root, fname)
                if filter_date:
                    try:
                        ctime = getctime_fn(full_path)
                        file_date = datetime.fromtimestamp(ctime).date()
                        filter_d = filter_date.date()
                        if filter_mode == "after" and file_date < filter_d:
                            continue
                        if filter_mode == "before" and file_date > filter_d:
                            continue
                    except Exception:
                        pass
                files.append(full_path)
    return files
