# -*- coding: utf-8 -*-
"""Error record and scan-skip runtime helpers extracted from office_converter.py."""


def record_detailed_error(
    source_path,
    exception,
    *,
    context,
    classify_conversion_error_fn,
    abspath_fn,
    basename_fn,
    now_fn,
    infer_failure_stage_fn,
    get_failure_output_expectation_fn,
    detailed_error_records,
    stats,
):
    error_info = classify_conversion_error_fn(exception, str(source_path))
    record = {
        "source_path": abspath_fn(source_path),
        "file_name": basename_fn(source_path),
        "error_type": error_info["error_type"],
        "error_category": error_info["error_category"],
        "message": error_info["message"],
        "suggestion": error_info["suggestion"],
        "is_retryable": error_info["is_retryable"],
        "requires_manual_action": error_info["requires_manual_action"],
        "raw_error": str(exception)[:500] if exception else "",
        "timestamp": now_fn().isoformat(),
        "context": context or {},
    }
    record["failure_stage"] = infer_failure_stage_fn(
        record["source_path"],
        raw_error=record["raw_error"],
        context=record["context"],
    )
    record["expected_outputs"] = get_failure_output_expectation_fn()
    detailed_error_records.append(record)

    error_type = error_info["error_type"]
    if error_type in stats:
        stats[error_type] += 1
    return record


def record_scan_access_skip(
    path,
    exception,
    *,
    context,
    seen_keys,
    silent,
    is_win_fn,
    abspath_fn,
    record_detailed_error_fn,
    stats,
    build_failed_file_trace_payload_fn,
    write_failed_file_trace_log_fn,
    print_fn=print,
    log_warning_fn=None,
):
    abs_path = abspath_fn(path) if path else ""
    key = abs_path.lower() if is_win_fn() else abs_path
    if seen_keys is not None:
        if key in seen_keys:
            return None
        seen_keys.add(key)

    detail = record_detailed_error_fn(
        abs_path or "<unknown>",
        exception,
        context=dict(context or {}, phase="scan", skip_only=True),
    )
    stats["skipped"] += 1

    msg = (
        f"[scan_skip] inaccessible path skipped: {abs_path or '<unknown>'} | "
        f"type={detail.get('error_type')} | err={exception}"
    )
    trace_payload = build_failed_file_trace_payload_fn(
        source_path=abs_path or "<unknown>",
        error_detail=detail,
        status="skipped_scan",
        elapsed=(detail.get("context") or {}).get("elapsed", 0.0),
        is_retry=False,
        failed_copy_path=None,
        extra_context={"scan_only": True},
    )
    write_failed_file_trace_log_fn(trace_payload, failed_copy_path=None)
    if not silent:
        print_fn(f"[WARN] {msg}")
    if log_warning_fn:
        log_warning_fn(msg)
    return detail
