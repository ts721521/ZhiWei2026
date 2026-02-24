# -*- coding: utf-8 -*-
"""MSHelp scan helpers extracted from office_converter.py."""

import os

from converter.platform_utils import is_win


def find_mshelpviewer_dirs(
    root_dir,
    folder_name="MSHelpViewer",
    is_dir_fn=None,
    walk_fn=None,
    is_win_fn=None,
):
    if is_dir_fn is None:
        is_dir_fn = os.path.isdir
    if walk_fn is None:
        walk_fn = os.walk
    if is_win_fn is None:
        is_win_fn = is_win

    result = []
    if not root_dir or not is_dir_fn(root_dir):
        return result

    folder_name = str(folder_name or "MSHelpViewer").strip() or "MSHelpViewer"
    folder_name_lower = folder_name.lower()

    for current_root, dirs, _ in walk_fn(root_dir):
        if os.path.basename(current_root).lower() == folder_name_lower:
            result.append(current_root)
            dirs[:] = []
            continue
        matches = [d for d in dirs if d.lower() == folder_name_lower]
        for d in matches:
            result.append(os.path.join(current_root, d))

    seen = set()
    unique = []
    for d in result:
        ad = os.path.abspath(d)
        key = ad.lower() if is_win_fn() else ad
        if key in seen:
            continue
        seen.add(key)
        unique.append(ad)
    return unique


def scan_mshelp_cab_candidates(
    config,
    source_roots,
    *,
    find_mshelpviewer_dirs_fn,
    find_files_recursive_fn,
    is_win_fn=None,
):
    if is_win_fn is None:
        is_win_fn = is_win

    dirs = []
    for source_root in source_roots:
        dirs.extend(find_mshelpviewer_dirs_fn(source_root))

    cab_exts = tuple(
        e.lower() for e in config.get("allowed_extensions", {}).get("cab", [".cab"])
    )
    if not cab_exts:
        cab_exts = (".cab",)

    files = []
    seen = set()
    for d in dirs:
        for cab_path in find_files_recursive_fn(d, cab_exts):
            abs_cab = os.path.abspath(cab_path)
            key = abs_cab.lower() if is_win_fn() else abs_cab
            if key in seen:
                continue
            seen.add(key)
            files.append(abs_cab)
    files.sort()
    return dirs, files
