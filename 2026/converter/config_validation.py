# -*- coding: utf-8 -*-
"""Runtime config schema validation helpers."""

from converter.constants import (
    COLLECT_MODE_COPY_AND_INDEX,
    COLLECT_MODE_INDEX_ONLY,
    ENGINE_ASK,
    ENGINE_MS,
    ENGINE_WPS,
    KILL_MODE_ASK,
    KILL_MODE_AUTO,
    KILL_MODE_KEEP,
    MERGE_CONVERT_SUBMODE_MERGE_ONLY,
    MERGE_CONVERT_SUBMODE_PDF_TO_MD,
    MERGE_MODE_ALL_IN_ONE,
    MERGE_MODE_CATEGORY,
    MODE_COLLECT_ONLY,
    MODE_CONVERT_ONLY,
    MODE_CONVERT_THEN_MERGE,
    MODE_MERGE_ONLY,
    MODE_MSHELP_ONLY,
    STRATEGY_PRICE_ONLY,
    STRATEGY_SMART_TAG,
    STRATEGY_STANDARD,
)


def _is_int_like(value):
    return isinstance(value, int) and not isinstance(value, bool)


def _validate_enum(cfg, key, allowed, errors):
    value = cfg.get(key)
    if value not in allowed:
        errors.append(f"{key}={value!r} not in {sorted(allowed)!r}")


def _validate_int_range(cfg, key, min_value, max_value, errors):
    value = cfg.get(key)
    if not _is_int_like(value):
        errors.append(f"{key} must be int, got {type(value).__name__}")
        return
    if value < min_value or value > max_value:
        errors.append(f"{key}={value} out of range [{min_value}, {max_value}]")


def _validate_bool(cfg, key, errors):
    value = cfg.get(key)
    if not isinstance(value, bool):
        errors.append(f"{key} must be bool, got {type(value).__name__}")


def validate_runtime_config_or_raise(cfg):
    """Validate runtime config values and raise ValueError on schema mismatch."""
    errors = []
    _validate_enum(
        cfg,
        "run_mode",
        {
            MODE_CONVERT_ONLY,
            MODE_MERGE_ONLY,
            MODE_CONVERT_THEN_MERGE,
            MODE_COLLECT_ONLY,
            MODE_MSHELP_ONLY,
        },
        errors,
    )
    _validate_enum(
        cfg,
        "collect_mode",
        {COLLECT_MODE_COPY_AND_INDEX, COLLECT_MODE_INDEX_ONLY},
        errors,
    )
    _validate_enum(
        cfg,
        "content_strategy",
        {STRATEGY_STANDARD, STRATEGY_SMART_TAG, STRATEGY_PRICE_ONLY},
        errors,
    )
    _validate_enum(cfg, "default_engine", {ENGINE_WPS, ENGINE_MS, ENGINE_ASK}, errors)
    _validate_enum(
        cfg,
        "kill_process_mode",
        {KILL_MODE_ASK, KILL_MODE_AUTO, KILL_MODE_KEEP},
        errors,
    )
    _validate_enum(
        cfg,
        "merge_mode",
        {MERGE_MODE_CATEGORY, MERGE_MODE_ALL_IN_ONE},
        errors,
    )
    _validate_enum(
        cfg,
        "merge_convert_submode",
        {MERGE_CONVERT_SUBMODE_MERGE_ONLY, MERGE_CONVERT_SUBMODE_PDF_TO_MD},
        errors,
    )

    _validate_int_range(cfg, "timeout_seconds", 1, 36000, errors)
    _validate_int_range(cfg, "pdf_wait_seconds", 1, 36000, errors)
    _validate_int_range(cfg, "ppt_timeout_seconds", 1, 36000, errors)
    _validate_int_range(cfg, "ppt_pdf_wait_seconds", 1, 36000, errors)
    _validate_int_range(cfg, "max_merge_size_mb", 1, 4096, errors)
    _validate_int_range(cfg, "markdown_max_size_mb", 1, 4096, errors)
    _validate_int_range(cfg, "parallel_workers", 1, 128, errors)
    _validate_int_range(cfg, "parallel_checkpoint_interval", 1, 1000000, errors)
    if "parallel_max_pending" in cfg:
        _validate_int_range(cfg, "parallel_max_pending", 1, 1000000, errors)

    for key in (
        "enable_parallel_conversion",
        "enable_merge",
        "output_enable_pdf",
        "output_enable_md",
        "output_enable_merged",
        "output_enable_independent",
        "enable_fast_md_engine",
    ):
        _validate_bool(cfg, key, errors)

    if not isinstance(cfg.get("source_folders", []), list):
        errors.append("source_folders must be list")
    if not isinstance(cfg.get("excluded_folders", []), list):
        errors.append("excluded_folders must be list")
    if not isinstance(cfg.get("price_keywords", []), list):
        errors.append("price_keywords must be list")
    if not isinstance(cfg.get("allowed_extensions", {}), dict):
        errors.append("allowed_extensions must be dict")

    if errors:
        raise ValueError("; ".join(errors))
