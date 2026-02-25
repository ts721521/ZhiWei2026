# -*- coding: utf-8 -*-
"""Runtime lifecycle helpers extracted from office_converter.py."""


def cleanup_all_processes(
    engine_type,
    *,
    process_names_for_engine_fn,
    kill_process_by_name_fn,
):
    for app in process_names_for_engine_fn(engine_type):
        kill_process_by_name_fn(app)


def close_office_apps(
    *,
    reuse_process,
    run_mode,
    mode_merge_only,
    mode_collect_only,
    cleanup_all_processes_fn,
):
    if not reuse_process and run_mode not in (mode_merge_only, mode_collect_only):
        cleanup_all_processes_fn()


def kill_current_app(
    app_type,
    *,
    reuse_process,
    force,
    engine_type,
    engine_wps,
    engine_ms,
    kill_process_by_name_fn,
):
    if reuse_process and not force:
        return
    name_map = {
        engine_wps: {"word": "wps", "excel": "et", "ppt": "wpp"},
        engine_ms: {"word": "winword", "excel": "excel", "ppt": "powerpnt"},
    }
    if engine_type not in name_map:
        return
    app_name = name_map[engine_type].get(app_type, "")
    kill_process_by_name_fn(app_name)


def on_office_file_processed(
    ext,
    *,
    should_reuse_office_app_fn,
    reuse_process,
    get_office_restart_every_fn,
    get_app_type_for_ext_fn,
    office_file_counter,
    kill_current_app_fn,
    log_info_fn,
):
    if not should_reuse_office_app_fn():
        return office_file_counter
    if reuse_process:
        return office_file_counter
    restart_every = get_office_restart_every_fn()
    if restart_every <= 0:
        return office_file_counter
    app_type = get_app_type_for_ext_fn(ext)
    if not app_type:
        return office_file_counter

    office_file_counter += 1
    if office_file_counter % restart_every != 0:
        return office_file_counter

    log_info_fn(
        f"[perf] periodic office restart ({app_type}) at file #{office_file_counter}"
    )
    kill_current_app_fn(app_type, force=True)
    return office_file_counter


def check_and_handle_running_processes(
    *,
    run_mode,
    config,
    interactive,
    resolve_process_handling_fn,
    cleanup_all_processes_fn,
):
    decision = resolve_process_handling_fn(
        run_mode=run_mode,
        kill_process_mode=config.get("kill_process_mode", "ask"),
        interactive=interactive,
    )
    if decision["skip"]:
        return None
    if decision["cleanup_all"]:
        cleanup_all_processes_fn()
    return decision["reuse_process"]


def check_and_handle_running_processes_for_converter(
    converter,
    *,
    resolve_process_handling_fn,
):
    reuse_process = check_and_handle_running_processes(
        run_mode=converter.run_mode,
        config=converter.config,
        interactive=converter.interactive,
        resolve_process_handling_fn=resolve_process_handling_fn,
        cleanup_all_processes_fn=converter.cleanup_all_processes,
    )
    if reuse_process is not None:
        converter.reuse_process = reuse_process
    return reuse_process
