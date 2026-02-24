# -*- coding: utf-8 -*-
"""Source-root resolution helpers extracted from office_converter.py."""

import os


def get_configured_source_roots(config):
    roots = config.get("source_folders")
    if isinstance(roots, list) and roots:
        resolved = [os.path.abspath(str(p).strip()) for p in roots if str(p).strip()]
        if resolved:
            return resolved
    single = str(config.get("source_folder", "")).strip()
    if single:
        return [os.path.abspath(single)]
    return []


def get_source_roots(config, is_dir_func=None):
    """Return list of accessible source folder paths to scan."""
    if is_dir_func is None:
        is_dir_func = os.path.isdir
    return [r for r in get_configured_source_roots(config) if is_dir_func(r)]


def probe_source_root_access(
    source_root,
    record_skip_fn,
    context=None,
    seen_keys=None,
    is_dir_func=None,
    listdir_fn=None,
):
    if is_dir_func is None:
        is_dir_func = os.path.isdir
    if listdir_fn is None:
        listdir_fn = os.listdir

    abs_root = os.path.abspath(str(source_root))
    if not is_dir_func(abs_root):
        record_skip_fn(
            abs_root,
            FileNotFoundError(f"source folder not found or inaccessible: {abs_root}"),
            context=context,
            seen_keys=seen_keys,
        )
        return False
    try:
        listdir_fn(abs_root)
        return True
    except Exception as e:
        record_skip_fn(abs_root, e, context=context, seen_keys=seen_keys)
        return False


def get_source_root_for_path(abs_path, roots, fallback=""):
    """Return the source root containing abs_path (prefer longest prefix)."""
    abs_path = os.path.abspath(abs_path)
    if not roots:
        return fallback or ""
    match = ""
    for r in roots:
        r = os.path.abspath(r)
        if abs_path == r or abs_path.startswith(r + os.sep) or abs_path.startswith(r + "/"):
            if len(r) > len(match):
                match = r
    return match or (roots[0] if roots else (fallback or ""))
