# -*- coding: utf-8 -*-
"""CAB extraction helper extracted from office_converter.py."""

import os
import shutil


def extract_cab_with_fallback(
    cab_path,
    extract_dir,
    *,
    is_win_fn,
    run_cmd,
    find_files_recursive_fn,
    cab_7z_path,
    get_app_path_fn,
    which_fn,
):
    cab_abs = os.path.abspath(cab_path)
    extract_abs = os.path.abspath(extract_dir)
    os.makedirs(extract_abs, exist_ok=True)

    expand_ok = False
    if is_win_fn():
        try:
            cmd_expand = ["expand", cab_abs, "-F:*", extract_abs]
            run_cmd(
                cmd_expand,
                capture_output=True,
                text=True,
                encoding="gbk",
                errors="ignore",
                check=True,
            )
            expand_ok = True
        except (OSError, RuntimeError, TypeError, ValueError):
            expand_ok = False

    if expand_ok and find_files_recursive_fn(extract_abs, (".mshc", ".htm", ".html")):
        return

    seven_zip = str(cab_7z_path or "").strip()
    if seven_zip:
        if not os.path.isabs(seven_zip):
            seven_zip = os.path.abspath(os.path.join(get_app_path_fn(), seven_zip))
        if not os.path.isfile(seven_zip):
            seven_zip = ""
    if not seven_zip:
        seven_zip = which_fn("7z") or which_fn("7za") or ""
    if not seven_zip:
        raise RuntimeError(
            "CAB extraction fallback requires 7z. Please install 7-Zip or set cab_7z_path."
        )

    cmd_7z = [seven_zip, "x", cab_abs, f"-o{extract_abs}", "-y"]
    run_cmd(
        cmd_7z,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="ignore",
        check=True,
    )

    if not find_files_recursive_fn(extract_abs, (".mshc", ".htm", ".html")):
        raise RuntimeError(f"CAB extraction produced no MSHC/HTML payload: {cab_path}")
