# -*- coding: utf-8 -*-
"""Merge candidate scan/build helpers extracted from office_converter.py."""

import os
from datetime import datetime

from converter.constants import MERGE_MODE_ALL_IN_ONE


def scan_candidates_by_ext(ext, scan_roots, exclude_abs_paths=None):
    ext = str(ext or "").lower()
    if not ext.startswith("."):
        ext = "." + ext

    excludes = set(map(os.path.abspath, exclude_abs_paths or []))
    files = []
    for scan_folder in scan_roots or []:
        if not scan_folder or not os.path.isdir(scan_folder):
            continue
        for root, dirs, names in os.walk(scan_folder):
            dirs[:] = [d for d in dirs if os.path.abspath(os.path.join(root, d)) not in excludes]
            if os.path.abspath(root) in excludes:
                continue
            for name in names:
                if name.lower().endswith(ext):
                    files.append(os.path.join(root, name))
    files.sort()
    return files


def build_markdown_merge_tasks(md_files, merge_mode, now=None):
    if not md_files:
        return []
    tasks = []
    if merge_mode == MERGE_MODE_ALL_IN_ONE:
        ts = (now or datetime.now()).strftime("%Y%m%d_%H%M%S")
        tasks.append((f"Merged_All_{ts}.md", md_files))
        return tasks

    categories = {
        "Price": "Price_",
        "Word": "Word_",
        "Excel": "Excel_",
        "PPT": "PPT_",
        "PDF": "PDF_",
    }
    matched = set()
    ts = (now or datetime.now()).strftime("%Y%m%d_%H%M%S")
    for cat_label, prefix in categories.items():
        group = [p for p in md_files if os.path.basename(p).startswith(prefix)]
        if not group:
            continue
        matched.update(group)
        tasks.append((f"Merged_{cat_label}_{ts}_001.md", group))

    others = [p for p in md_files if p not in matched]
    if others:
        tasks.append((f"Merged_Markdown_{ts}_001.md", others))
    return tasks
