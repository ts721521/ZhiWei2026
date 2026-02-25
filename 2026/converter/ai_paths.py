# -*- coding: utf-8 -*-
"""AI output path helpers extracted from office_converter.py."""

import os


def build_ai_output_path(source_path, sub_dir, ext, target_root):
    target_root = str(target_root or "")
    if not target_root:
        return None

    source_abs = os.path.abspath(source_path)
    try:
        rel = os.path.relpath(source_abs, target_root)
    except (TypeError, ValueError, OSError):
        rel = os.path.basename(source_abs)
    if rel.startswith(".."):
        rel = os.path.basename(source_abs)
    rel_no_ext = os.path.splitext(rel)[0]

    output_path = os.path.join(target_root, "_AI", sub_dir, rel_no_ext + ext)
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    return output_path


def build_ai_output_path_from_source(
    source_path, sub_dir, ext, target_root, source_root_resolver=None
):
    target_root = str(target_root or "")
    if not target_root:
        return None

    source_abs = os.path.abspath(source_path)
    rel = os.path.basename(source_abs)
    if callable(source_root_resolver):
        try:
            src_root = source_root_resolver(source_abs)
            rel_try = os.path.relpath(source_abs, src_root)
            if not rel_try.startswith(".."):
                rel = rel_try
        except (TypeError, ValueError, OSError):
            pass
    rel_no_ext = os.path.splitext(rel)[0]

    output_path = os.path.join(target_root, "_AI", sub_dir, rel_no_ext + ext)
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    return output_path
