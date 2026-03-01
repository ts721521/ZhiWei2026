# -*- coding: utf-8 -*-
"""Bootstrap state helpers extracted from office_converter.py."""


def build_default_perf_metrics():
    return {
        "scan_seconds": 0.0,
        "batch_seconds": 0.0,
        "merge_seconds": 0.0,
        "postprocess_seconds": 0.0,
        "total_seconds": 0.0,
        "convert_core_seconds": 0.0,
        "pdf_wait_seconds": 0.0,
        "markdown_seconds": 0.0,
        "mshelp_merge_seconds": 0.0,
    }


def initialize_runtime_state(
    converter,
    *,
    config_path,
    interactive,
    mode_convert_then_merge,
    collect_mode_copy_and_index,
    merge_mode_category,
    strategy_standard,
):
    converter.config_path = config_path
    converter.interactive = interactive

    converter.temp_sandbox = None
    converter.temp_sandbox_root = None
    converter.failed_dir = None
    converter.merge_output_dir = None

    converter.engine_type = None
    converter.is_running = True
    converter.reuse_process = False

    converter.run_mode = mode_convert_then_merge
    converter.collect_mode = collect_mode_copy_and_index
    converter.merge_mode = merge_mode_category
    converter.content_strategy = strategy_standard

    converter.price_keywords = []
    converter.excluded_folders = []

    converter.filter_date = None
    converter.filter_mode = "after"

    converter.enable_merge_index = False
    converter.enable_merge_excel = False

    converter.progress_callback = None
    converter.file_plan_callback = None
    converter.file_done_callback = None


def initialize_output_tracking_state(converter):
    converter.generated_pdfs = []
    converter.generated_merge_outputs = []
    converter.generated_merge_markdown_outputs = []
    converter.generated_map_outputs = []
    converter.generated_markdown_outputs = []
    converter.generated_markdown_manifest_outputs = []
    converter.generated_markdown_quality_outputs = []
    converter.generated_excel_json_outputs = []
    converter.generated_records_json_outputs = []
    converter.generated_chromadb_outputs = []
    converter.generated_update_package_outputs = []
    converter.generated_mshelp_outputs = []
    converter.generated_trace_map_outputs = []
    converter.generated_fast_md_outputs = []
    converter.generated_prompt_outputs = []
    converter.markdown_quality_records = []
    converter.trace_short_id_taken = set()
    converter.conversion_index_records = []
    converter.merge_index_records = []
    converter.mshelp_records = []
    converter.collect_index_path = None
    converter.convert_index_path = None
    converter.merge_excel_path = None
    converter.corpus_manifest_path = None
    converter.update_package_manifest_path = None
    converter.markdown_quality_report_path = None
    converter.chromadb_export_manifest_path = None
    converter.trace_map_path = None
    converter.prompt_ready_path = None
    converter.incremental_registry_path = ""
    converter._incremental_context = None
    converter._office_file_counter = 0
    converter.perf_metrics = build_default_perf_metrics()


def initialize_error_tracking_state(converter):
    converter.stats = {
        "total": 0,
        "success": 0,
        "failed": 0,
        "timeout": 0,
        "skipped": 0,
        "permission_denied": 0,
        "file_locked": 0,
        "file_corrupted": 0,
        "com_error": 0,
    }
    converter.error_records = []
    converter.detailed_error_records = []
    converter.failed_report_path = None


def register_signal_handlers(
    *,
    signal_module,
    current_thread_fn,
    main_thread_fn,
    signal_handler_fn,
):
    if current_thread_fn() is not main_thread_fn():
        return
    try:
        signal_module.signal(signal_module.SIGINT, signal_handler_fn)
        signal_module.signal(signal_module.SIGTERM, signal_handler_fn)
    except (AttributeError, OSError, RuntimeError, ValueError):
        pass


def handle_stop_signal(signum, *, set_running_fn, log_warning_fn):
    log_warning_fn(f"received signal {signum}, stopping gracefully...")
    set_running_fn(False)


def initialize_converter_for_runtime(
    converter,
    *,
    config_path,
    interactive,
    mode_convert_then_merge,
    collect_mode_copy_and_index,
    merge_mode_category,
    strategy_standard,
    signal_module,
    current_thread_fn,
    main_thread_fn,
):
    initialize_runtime_state(
        converter,
        config_path=config_path,
        interactive=interactive,
        mode_convert_then_merge=mode_convert_then_merge,
        collect_mode_copy_and_index=collect_mode_copy_and_index,
        merge_mode_category=merge_mode_category,
        strategy_standard=strategy_standard,
    )
    initialize_output_tracking_state(converter)
    register_signal_handlers(
        signal_module=signal_module,
        current_thread_fn=current_thread_fn,
        main_thread_fn=main_thread_fn,
        signal_handler_fn=converter.signal_handler,
    )
    converter.load_config(config_path)
    converter._init_paths_from_config()
    initialize_error_tracking_state(converter)
