# -*- coding: utf-8 -*-
"""Config default normalization extracted from office_converter.py."""

from converter.constants import (
    ENGINE_ASK,
    KILL_MODE_ASK,
    MERGE_CONVERT_SUBMODE_MERGE_ONLY,
    MERGE_MODE_CATEGORY,
    MODE_CONVERT_THEN_MERGE,
)


def apply_config_defaults(
    cfg,
    run_mode_default,
    collect_mode_default,
    content_strategy_default,
    enable_merge_index_default,
    enable_merge_excel_default,
):
    cfg.setdefault("obsidian_sync_enabled", True)
    cfg.setdefault("obsidian_root", "")
    cfg.setdefault("obsidian_program_name", "ZhiWei")
    cfg.setdefault("enable_corpus_manifest", True)
    if "output_enable_md" not in cfg and "enable_markdown" in cfg:
        cfg["output_enable_md"] = bool(cfg.get("enable_markdown", True))
    cfg.setdefault("markdown_strip_header_footer", True)
    cfg.setdefault("markdown_structured_headings", True)
    cfg.setdefault("enable_markdown_quality_report", True)
    cfg.setdefault("markdown_quality_sample_limit", 20)
    cfg.setdefault("enable_excel_json", False)
    cfg.setdefault("excel_json_max_rows", 2000)
    cfg.setdefault("excel_json_max_cols", 80)
    cfg.setdefault("excel_json_records_preview", 200)
    cfg.setdefault("excel_json_profile_rows", 500)
    cfg.setdefault("excel_json_include_formulas", True)
    cfg.setdefault("excel_json_extract_sheet_links", True)
    cfg.setdefault("excel_json_include_merged_ranges", True)
    cfg.setdefault("excel_json_formula_sample_limit", 200)
    cfg.setdefault("excel_json_merged_range_limit", 500)
    cfg.setdefault("enable_chromadb_export", False)
    cfg.setdefault("chromadb_persist_dir", "")
    cfg.setdefault("chromadb_collection_name", "office_corpus")
    cfg.setdefault("chromadb_max_chars_per_chunk", 1800)
    cfg.setdefault("chromadb_chunk_overlap", 200)
    cfg.setdefault("chromadb_write_jsonl_fallback", True)
    cfg.setdefault("timeout_seconds", 60)
    cfg.setdefault("enable_sandbox", True)
    cfg.setdefault("default_engine", ENGINE_ASK)
    cfg.setdefault("kill_process_mode", KILL_MODE_ASK)
    cfg.setdefault("auto_retry_failed", False)
    cfg.setdefault("enable_failed_file_trace_log", True)
    cfg.setdefault("office_reuse_app", True)
    cfg.setdefault("office_restart_every_n_files", 25)
    cfg.setdefault("pdf_wait_seconds", 15)
    cfg.setdefault("ppt_timeout_seconds", cfg.get("timeout_seconds", 60))
    cfg.setdefault("ppt_pdf_wait_seconds", cfg.get("pdf_wait_seconds", 15))
    cfg.setdefault("enable_merge", True)
    cfg.setdefault("max_merge_size_mb", 80)
    cfg.setdefault("output_enable_pdf", True)
    cfg.setdefault("output_enable_md", True)
    cfg.setdefault("output_enable_merged", True)
    cfg.setdefault("output_enable_independent", False)
    # Keep legacy key aligned to avoid split-brain behavior across versions.
    cfg["enable_markdown"] = bool(cfg.get("output_enable_md", True))
    cfg.setdefault("merge_convert_submode", MERGE_CONVERT_SUBMODE_MERGE_ONLY)
    cfg.setdefault("temp_sandbox_root", "")
    cfg.setdefault("sandbox_min_free_gb", 10)
    cfg.setdefault("sandbox_low_space_policy", "block")
    cfg.setdefault("enable_llm_delivery_hub", True)
    cfg.setdefault("llm_delivery_root", "")
    cfg.setdefault("llm_delivery_flatten", True)
    cfg.setdefault("llm_delivery_include_pdf", False)
    cfg.setdefault("enable_gdrive_upload", False)
    cfg.setdefault("gdrive_client_secrets_path", "")
    cfg.setdefault("gdrive_folder_id", "")
    cfg.setdefault("gdrive_token_path", "")
    cfg.setdefault("overwrite_same_size", True)
    cfg.setdefault("merge_mode", MERGE_MODE_CATEGORY)
    cfg.setdefault("merge_source", "source")
    cfg.setdefault("enable_merge_index", False)
    cfg.setdefault("enable_merge_excel", False)
    cfg.setdefault("enable_merge_map", True)
    cfg.setdefault("bookmark_with_short_id", True)
    cfg.setdefault("enable_incremental_mode", False)
    cfg.setdefault("incremental_verify_hash", False)
    cfg.setdefault("incremental_reprocess_renamed", False)
    cfg.setdefault("incremental_registry_path", "")
    cfg.setdefault("source_priority_skip_same_name_pdf", False)
    cfg.setdefault("global_md5_dedup", False)
    cfg.setdefault("enable_update_package", True)
    cfg.setdefault("update_package_root", "")
    cfg.setdefault("cab_7z_path", "")
    cfg.setdefault("mshelpviewer_folder_name", "MSHelpViewer")
    cfg.setdefault("enable_mshelp_merge_output", True)
    cfg.setdefault("enable_mshelp_output_docx", False)
    cfg.setdefault("enable_mshelp_output_pdf", False)

    cfg.setdefault("enable_parallel_conversion", False)
    cfg.setdefault("parallel_workers", 4)
    cfg.setdefault("parallel_checkpoint_interval", 10)

    cfg.setdefault("enable_checkpoint", True)
    cfg.setdefault("checkpoint_auto_resume", True)

    if "privacy" not in cfg or not isinstance(cfg["privacy"], dict):
        cfg["privacy"] = {}
    cfg["privacy"].setdefault("mask_md5_in_logs", True)

    if "everything" not in cfg or not isinstance(cfg["everything"], dict):
        cfg["everything"] = {}
    cfg["everything"].setdefault("enabled", True)
    cfg["everything"].setdefault("es_path", "")
    cfg["everything"].setdefault("prefer_path_exact", True)
    cfg["everything"].setdefault("timeout_ms", 1500)

    if "listary" not in cfg or not isinstance(cfg["listary"], dict):
        cfg["listary"] = {}
    cfg["listary"].setdefault("enabled", True)
    cfg["listary"].setdefault("copy_query_on_locate", True)

    if "price_keywords" not in cfg or not isinstance(cfg["price_keywords"], list):
        cfg["price_keywords"] = ["报价", "价格表", "Price", "Quotation"]
    price_keywords = cfg["price_keywords"]

    if "excluded_folders" not in cfg or not isinstance(cfg["excluded_folders"], list):
        cfg["excluded_folders"] = ["temp", "backup", "archive"]
    excluded_folders = cfg["excluded_folders"]

    exts = cfg.setdefault("allowed_extensions", {})
    exts.setdefault("word", [".doc", ".docx"])
    exts.setdefault("excel", [".xls", ".xlsx"])
    exts.setdefault("powerpoint", [".ppt", ".pptx"])
    exts.setdefault("pdf", [".pdf"])
    exts.setdefault("cab", [".cab"])

    merge_mode = cfg.get("merge_mode", MERGE_MODE_CATEGORY)
    run_mode = cfg.get("run_mode", run_mode_default)
    collect_mode = cfg.get("collect_mode", collect_mode_default)
    content_strategy = cfg.get("content_strategy", content_strategy_default)
    enable_merge_index = bool(cfg.get("enable_merge_index", enable_merge_index_default))
    enable_merge_excel = bool(cfg.get("enable_merge_excel", enable_merge_excel_default))

    if run_mode == MODE_CONVERT_THEN_MERGE:
        cfg["merge_source"] = "target"

    return {
        "price_keywords": price_keywords,
        "excluded_folders": excluded_folders,
        "merge_mode": merge_mode,
        "run_mode": run_mode,
        "collect_mode": collect_mode,
        "content_strategy": content_strategy,
        "enable_merge_index": enable_merge_index,
        "enable_merge_excel": enable_merge_excel,
    }
