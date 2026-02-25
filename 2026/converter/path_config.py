# -*- coding: utf-8 -*-
"""Path/runtime config helpers extracted from office_converter.py."""

import json
import os


def get_path_from_config(cfg, key_base, prefer_win=False, prefer_mac=False):
    cfg = cfg or {}
    val = None
    if prefer_win:
        val = cfg.get(f"{key_base}_win")
    elif prefer_mac:
        val = cfg.get(f"{key_base}_mac")
    if not val:
        val = cfg.get(key_base)
    if val:
        return os.path.abspath(val)
    return ""


def init_paths_from_config(
    cfg,
    get_app_path_fn,
    gettempdir_fn,
    *,
    isabs_fn=os.path.isabs,
    abspath_fn=os.path.abspath,
    join_fn=os.path.join,
    makedirs_fn=os.makedirs,
):
    """Build runtime temp/failed/merge paths from config and ensure directories."""
    cfg = cfg or {}

    temp_root = str(cfg.get("temp_sandbox_root", "")).strip()
    if temp_root:
        if not isabs_fn(temp_root):
            temp_root = abspath_fn(join_fn(get_app_path_fn(), temp_root))
    else:
        temp_root = gettempdir_fn()

    target_folder = cfg.get("target_folder", "")
    temp_sandbox = join_fn(temp_root, "OfficeToPDF_Sandbox")
    failed_dir = join_fn(target_folder, "_FAILED_FILES")
    merge_output_dir = join_fn(target_folder, "_MERGED")

    makedirs_fn(temp_sandbox, exist_ok=True)
    makedirs_fn(failed_dir, exist_ok=True)
    makedirs_fn(merge_output_dir, exist_ok=True)

    return {
        "temp_sandbox_root": temp_root,
        "temp_sandbox": temp_sandbox,
        "failed_dir": failed_dir,
        "merge_output_dir": merge_output_dir,
    }


def init_paths_from_config_for_converter(
    converter,
    *,
    get_app_path_fn,
    gettempdir_fn,
    isabs_fn=os.path.isabs,
    abspath_fn=os.path.abspath,
    join_fn=os.path.join,
    makedirs_fn=os.makedirs,
):
    paths = init_paths_from_config(
        converter.config,
        get_app_path_fn,
        gettempdir_fn,
        isabs_fn=isabs_fn,
        abspath_fn=abspath_fn,
        join_fn=join_fn,
        makedirs_fn=makedirs_fn,
    )
    converter.temp_sandbox_root = paths["temp_sandbox_root"]
    converter.temp_sandbox = paths["temp_sandbox"]
    converter.failed_dir = paths["failed_dir"]
    converter.merge_output_dir = paths["merge_output_dir"]
    return paths


def save_config(config_path, config, *, open_fn=open, dump_fn=json.dump, log_error_fn=None):
    """Persist config json with stable formatting."""
    try:
        with open_fn(config_path, "w", encoding="utf-8") as fh:
            dump_fn(config, fh, indent=4, ensure_ascii=False)
        return True
    except (OSError, TypeError, ValueError) as exc:
        if log_error_fn is not None:
            log_error_fn(f"failed to save config: {exc}")
        return False
