# -*- coding: utf-8 -*-
"""Artifact metadata helpers extracted from office_converter.py."""

import os
from datetime import datetime

from converter.hash_utils import compute_file_hash, compute_md5


def safe_file_meta(path, target_folder):
    if not path:
        return None
    abs_path = os.path.abspath(path)
    if not os.path.exists(abs_path):
        return None
    try:
        stat = os.stat(abs_path)
    except OSError:
        return None

    try:
        rel_path = os.path.relpath(abs_path, target_folder)
    except (TypeError, ValueError, OSError):
        rel_path = abs_path

    md5_value = ""
    sha256_value = ""
    try:
        md5_value = compute_md5(abs_path)
    except (TypeError, ValueError, OSError):
        pass
    try:
        sha256_value = compute_file_hash(abs_path)
    except (TypeError, ValueError, OSError):
        pass

    return {
        "path_abs": abs_path,
        "path_rel_to_target": rel_path,
        "size_bytes": int(stat.st_size),
        "mtime": datetime.fromtimestamp(stat.st_mtime).isoformat(timespec="seconds"),
        "md5": md5_value,
        "sha256": sha256_value,
    }


def add_artifact(artifacts, kind, path, target_folder):
    meta = safe_file_meta(path, target_folder)
    if not meta:
        return
    item = {"kind": kind}
    item.update(meta)
    artifacts.append(item)
