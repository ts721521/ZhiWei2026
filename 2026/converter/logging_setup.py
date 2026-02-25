# -*- coding: utf-8 -*-
"""Logging setup helper extracted from office_converter.py."""

import os


def setup_logging(
    *,
    config,
    engine_type,
    run_mode,
    content_strategy,
    merge_mode,
    temp_sandbox,
    merge_output_dir,
    app_version,
    get_readable_run_mode_fn,
    get_readable_content_strategy_fn,
    get_readable_merge_mode_fn,
    should_reuse_office_app_fn,
    get_office_restart_every_fn,
    mode_convert_only,
    mode_convert_then_merge,
    mode_merge_only,
    now_fn,
    get_app_path_fn,
    logging_module,
):
    log_dir = config.get("log_folder", "./logs")
    if not os.path.isabs(log_dir):
        log_dir = os.path.join(get_app_path_fn(), log_dir)
    os.makedirs(log_dir, exist_ok=True)

    log_path = os.path.join(log_dir, f"conversion_log_{now_fn().strftime('%Y%m%d_%H%M%S')}.txt")

    logging_module.basicConfig(
        filename=log_path,
        level=logging_module.INFO,
        format="%(message)s",
        encoding="utf-8",
        force=True,
    )
    console = logging_module.StreamHandler()
    console.setLevel(logging_module.INFO)
    console.setFormatter(logging_module.Formatter("%(message)s"))
    logging_module.getLogger("").addHandler(console)

    engine_label = engine_type.upper() if engine_type else "N/A"
    with open(log_path, "a", encoding="utf-8") as f:
        f.write(f"{now_fn()} === Task Start (v{app_version}) ===\n")
        f.write(f"run_mode: {run_mode} ({get_readable_run_mode_fn()})\n")
        if run_mode in (mode_convert_only, mode_convert_then_merge):
            f.write(
                f"content_strategy: {content_strategy} ({get_readable_content_strategy_fn()})\n"
            )
        if run_mode in (mode_convert_then_merge, mode_merge_only):
            f.write(f"merge_mode: {merge_mode} ({get_readable_merge_mode_fn()})\n")
        f.write(f"engine: {engine_label}\n")
        f.write(f"source_folder: {config['source_folder']}\n")
        f.write(f"target_folder: {config['target_folder']}\n")
        f.write(f"office_reuse_app: {should_reuse_office_app_fn()}\n")
        f.write(f"office_restart_every_n_files: {get_office_restart_every_fn()}\n")
        if config.get("enable_sandbox", True):
            f.write(f"sandbox: enabled | temp: {temp_sandbox}\n")
        else:
            f.write(f"sandbox: disabled | temp: {temp_sandbox} (PDF temp only)\n")
        if run_mode in (mode_convert_then_merge, mode_merge_only):
            f.write(f"merge_output_dir: {merge_output_dir}\n")
        f.write("=" * 60 + "\n")

    return log_path
