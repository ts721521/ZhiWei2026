# -*- coding: utf-8 -*-
"""CLI wizard flow orchestration extracted from office_converter.py."""


def run_cli_wizard(
    *,
    interactive,
    print_welcome_fn,
    confirm_config_in_terminal_fn,
    ask_for_subfolder_fn,
    select_run_mode_fn,
    get_run_mode_fn,
    select_collect_mode_fn,
    select_content_strategy_fn,
    select_merge_mode_fn,
    select_engine_mode_fn,
    check_and_handle_running_processes_fn,
    init_paths_from_config_fn,
    config,
    mode_collect_only,
    mode_convert_only,
    mode_convert_then_merge,
    mode_merge_only,
    mode_mshelp_only,
):
    if not interactive:
        return
    print_welcome_fn()
    confirm_config_in_terminal_fn()
    ask_for_subfolder_fn()
    select_run_mode_fn()

    run_mode = get_run_mode_fn()
    if run_mode == mode_collect_only:
        select_collect_mode_fn()
    elif run_mode in (mode_convert_only, mode_convert_then_merge):
        select_content_strategy_fn()

    run_mode = get_run_mode_fn()
    if run_mode in (mode_convert_then_merge, mode_merge_only) and (config or {}).get(
        "enable_merge", True
    ):
        select_merge_mode_fn()

    run_mode = get_run_mode_fn()
    if run_mode not in (mode_merge_only, mode_collect_only, mode_mshelp_only):
        select_engine_mode_fn()
        check_and_handle_running_processes_fn()

    init_paths_from_config_fn()
