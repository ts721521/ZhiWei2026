# -*- coding: utf-8 -*-
"""Failed-file trace helpers extracted from office_converter.py."""

import json
import os
from datetime import datetime


def build_failed_file_trace_payload(
    *,
    source_path,
    error_detail,
    status,
    elapsed,
    is_retry,
    failed_copy_path=None,
    extra_context=None,
    get_failure_output_expectation_fn=None,
    get_readable_run_mode_fn=None,
    get_readable_engine_type_fn=None,
    infer_failure_stage_fn=None,
):
    expected_outputs = (
        get_failure_output_expectation_fn() if callable(get_failure_output_expectation_fn) else {}
    )
    source_abs = os.path.abspath(str(source_path or ""))
    raw_error = str(error_detail.get("raw_error", "") or "")
    ctx = dict(error_detail.get("context") or {})
    if extra_context:
        ctx.update(dict(extra_context))

    failure_stage = ""
    if callable(infer_failure_stage_fn):
        failure_stage = infer_failure_stage_fn(source_abs, raw_error=raw_error, context=ctx)

    run_mode = get_readable_run_mode_fn() if callable(get_readable_run_mode_fn) else ""
    engine = get_readable_engine_type_fn() if callable(get_readable_engine_type_fn) else ""

    return {
        "version": 1,
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "status": status,
        "is_retry": bool(is_retry),
        "source_path": source_abs,
        "source_file_name": os.path.basename(source_abs),
        "source_ext": os.path.splitext(source_abs)[1].lower(),
        "failed_copy_path": os.path.abspath(failed_copy_path) if failed_copy_path else "",
        "run_mode": run_mode,
        "engine": engine,
        "elapsed_seconds": float(elapsed or 0.0),
        "failure_stage": failure_stage,
        "expected_outputs": expected_outputs,
        "error_type": error_detail.get("error_type", ""),
        "error_category": error_detail.get("error_category", ""),
        "is_retryable": bool(error_detail.get("is_retryable", False)),
        "requires_manual_action": bool(error_detail.get("requires_manual_action", False)),
        "message": error_detail.get("message", ""),
        "suggestion": error_detail.get("suggestion", ""),
        "raw_error": raw_error,
        "context": ctx,
    }


def write_failed_file_trace_log(
    payload,
    failed_copy_path=None,
    *,
    enable_trace_log=True,
    failed_dir="",
    target_folder="",
    sanitize_stem_fn=None,
    log_error=None,
):
    if not enable_trace_log:
        return None

    base_dir = ""
    if failed_copy_path:
        try:
            base_dir = os.path.dirname(os.path.abspath(failed_copy_path))
        except (OSError, RuntimeError, TypeError, ValueError):
            base_dir = ""
    if not base_dir:
        base_dir = failed_dir or target_folder
    if not base_dir:
        return None

    try:
        os.makedirs(base_dir, exist_ok=True)
    except OSError:
        return None

    source_name = os.path.basename(str(payload.get("source_path", "") or ""))
    if failed_copy_path:
        source_name = os.path.basename(str(failed_copy_path))
    stem = os.path.splitext(source_name)[0]
    if callable(sanitize_stem_fn):
        stem = sanitize_stem_fn(stem)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    log_path = os.path.join(base_dir, f"{stem}.{ts}.failure.json")

    try:
        with open(log_path, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
        return log_path
    except (OSError, TypeError, ValueError) as e:
        if callable(log_error):
            log_error(f"failed to write failed-file trace log: {e}")
        return None
