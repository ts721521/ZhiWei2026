# -*- coding: utf-8 -*-
"""Sequential batch processing extracted from office_converter.py."""

import logging
import os
import time

from converter.errors import ConversionErrorType


def run_batch(converter, file_list, is_retry=False, source_alias_map=None):
    """Run sequential conversion with checkpoint support."""
    total = len(file_list)
    results = []
    source_alias_map = source_alias_map or {}

    checkpoint = None
    checkpoint_interval = converter.config.get("parallel_checkpoint_interval", 10)

    if converter.config.get("enable_checkpoint", True) and not is_retry:
        checkpoint, pending_files = converter._init_checkpoint(file_list)
        if checkpoint and len(pending_files) < total:
            file_list = pending_files
            total = len(file_list)

    completed_count = 0

    for idx, fpath in enumerate(file_list, 1):
        if not converter.is_running:
            if checkpoint:
                converter._save_checkpoint(checkpoint)
            break

        logical_source = source_alias_map.get(fpath, fpath)
        fname = os.path.basename(fpath)
        ext = os.path.splitext(fpath)[1].lower()
        target_path_initial = converter.get_target_path(logical_source, ext)

        progress_prefix = converter.get_progress_prefix(idx, total)
        if converter.progress_callback:
            converter.progress_callback(idx, total)

        label = "[retry]" if is_retry else "processing"
        print(f"\r{progress_prefix} {label}: {fname}" + " " * 20, end="", flush=True)

        started_at = time.time()
        try:
            status, final_path = converter.process_single_file(
                fpath,
                target_path_initial,
                ext,
                progress_prefix,
                is_retry,
            )
            elapsed = time.time() - started_at
            completed_count += 1

            if status.startswith("skip"):
                print(f"\r{progress_prefix} {status}: {fname} ({elapsed:.2f}s)    ")
                logging.info(f"{status}: {logical_source}")
                record = {
                    "source_path": os.path.abspath(logical_source),
                    "status": "skipped",
                    "detail": status,
                    "final_path": final_path,
                    "elapsed": elapsed,
                }
                results.append(record)
                converter._emit_file_done(record)
            else:
                converter.stats["success"] += 1
                print(f"\r{progress_prefix} {status}: {fname} ({elapsed:.2f}s)    ")
                logging.info(f"{status}: {logical_source} -> {final_path}")
                is_pdf_output = str(final_path).lower().endswith(".pdf")
                result_status = "success" if is_pdf_output else "success_non_pdf"
                if is_pdf_output:
                    converter._append_conversion_index_record(
                        logical_source,
                        final_path,
                        status,
                    )
                record = {
                    "source_path": os.path.abspath(logical_source),
                    "status": result_status,
                    "detail": status,
                    "final_path": final_path,
                    "elapsed": elapsed,
                }
                results.append(record)
                converter._emit_file_done(record)

            if checkpoint and completed_count % checkpoint_interval == 0:
                checkpoint = converter._mark_file_done_in_checkpoint(checkpoint, fpath)
                converter._save_checkpoint(checkpoint)
            elif checkpoint:
                checkpoint = converter._mark_file_done_in_checkpoint(checkpoint, fpath)

        except (OSError, RuntimeError, TypeError, ValueError, AttributeError) as exc:
            elapsed = time.time() - started_at
            err_msg = str(exc)

            error_detail = converter.record_detailed_error(
                logical_source,
                exc,
                context={
                    "run_mode": converter.get_readable_run_mode(),
                    "engine": converter.get_readable_engine_type(),
                    "elapsed": elapsed,
                },
            )

            is_timeout = error_detail["error_type"] == ConversionErrorType.TIMEOUT
            is_timeout = is_timeout or ("timeout" in err_msg.lower()) or ("瓒呮椂" in err_msg)
            if is_timeout:
                converter.stats["timeout"] += 1
                print(f"\r{progress_prefix} timeout: {fname} ({elapsed:.2f}s)    ")
                record = {
                    "source_path": os.path.abspath(logical_source),
                    "status": "timeout",
                    "detail": err_msg,
                    "final_path": "",
                    "elapsed": elapsed,
                    "error": err_msg,
                    "error_type": error_detail["error_type"],
                    "error_category": error_detail["error_category"],
                    "suggestion": error_detail["suggestion"],
                    "failure_stage": error_detail.get("failure_stage", ""),
                }
                results.append(record)
                converter._emit_file_done(record)
            else:
                converter.stats["failed"] += 1
                print(
                    f"\r{progress_prefix} failed({error_detail['error_type']}): {fname}    "
                )
                record = {
                    "source_path": os.path.abspath(logical_source),
                    "status": "failed",
                    "detail": err_msg,
                    "final_path": "",
                    "elapsed": elapsed,
                    "error": err_msg,
                    "error_type": error_detail["error_type"],
                    "error_category": error_detail["error_category"],
                    "suggestion": error_detail["suggestion"],
                    "is_retryable": error_detail["is_retryable"],
                    "requires_manual_action": error_detail["requires_manual_action"],
                    "failure_stage": error_detail.get("failure_stage", ""),
                }
                results.append(record)
                converter._emit_file_done(record)

            logging.error(
                f"failed: {logical_source} | reason: {exc} | type: {error_detail['error_type']}"
            )

            failed_copy_path = None
            if not is_retry:
                failed_copy_path = converter.quarantine_failed_file(fpath)
                converter.error_records.append(logical_source)

            trace_payload = converter._build_failed_file_trace_payload(
                source_path=logical_source,
                error_detail=error_detail,
                status=record["status"],
                elapsed=elapsed,
                is_retry=is_retry,
                failed_copy_path=failed_copy_path,
                extra_context={"parallel": False},
            )
            trace_path = converter._write_failed_file_trace_log(
                trace_payload,
                failed_copy_path=failed_copy_path,
            )
            if trace_path:
                record["failure_trace_path"] = trace_path

            completed_count += 1
            if checkpoint:
                checkpoint = converter._mark_file_done_in_checkpoint(checkpoint, fpath)
                if completed_count % checkpoint_interval == 0:
                    converter._save_checkpoint(checkpoint)

    if checkpoint and checkpoint.get("status") == "completed":
        converter._clear_checkpoint()

    return results
