# -*- coding: utf-8 -*-
"""Config load helpers extracted from office_converter.py."""

from converter.config_validation import validate_runtime_config_or_raise


def load_config(
    converter,
    path,
    *,
    open_fn,
    json_loads_fn,
    abspath_fn,
    print_fn,
    exit_fn,
):
    try:
        with open_fn(path, "r", encoding="utf-8") as f:
            content = f.read().replace("\\", "/")
    except (OSError, UnicodeDecodeError) as exc:
        print_fn(f"[ERROR] Failed to load config file: {exc}")
        exit_fn(1)
        return
    try:
        converter.config = json_loads_fn(content)
    except ValueError as exc:
        print_fn(f"[ERROR] Invalid JSON in config file: {exc}")
        exit_fn(1)
        return

    converter.config["source_folder"] = converter._get_path_from_config("source_folder")
    converter.config["target_folder"] = converter._get_path_from_config("target_folder")
    converter.config["obsidian_root"] = converter._get_path_from_config("obsidian_root")

    src_list = converter.config.get("source_folders")
    if isinstance(src_list, list) and src_list:
        converter.config["source_folders"] = [
            abspath_fn(str(path_item).strip())
            for path_item in src_list
            if str(path_item).strip()
        ]
        if converter.config["source_folders"]:
            converter.config["source_folder"] = converter.config["source_folders"][0]
    else:
        converter.config["source_folders"] = (
            [converter.config["source_folder"]]
            if converter.config.get("source_folder")
            else []
        )

    converter._apply_config_defaults()
    effective_cfg = dict(converter.config)
    if hasattr(converter, "run_mode"):
        effective_cfg["run_mode"] = converter.run_mode
    if hasattr(converter, "collect_mode"):
        effective_cfg["collect_mode"] = converter.collect_mode
    if hasattr(converter, "content_strategy"):
        effective_cfg["content_strategy"] = converter.content_strategy
    if hasattr(converter, "merge_mode"):
        effective_cfg["merge_mode"] = converter.merge_mode
    try:
        validate_runtime_config_or_raise(effective_cfg)
    except (TypeError, ValueError, AttributeError) as exc:
        print_fn(f"[ERROR] Invalid config schema: {exc}")
        exit_fn(1)
        return
