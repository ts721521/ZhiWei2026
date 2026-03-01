# -*- coding: utf-8 -*-
"""Parallel batch processing extracted from office_converter.py."""

import logging
import os
import threading
import time


def run_batch_parallel(converter, file_list, is_retry=False, source_alias_map=None):
    """Run conversion in parallel by delegating work through converter hooks."""
    from concurrent.futures import FIRST_COMPLETED, ThreadPoolExecutor, wait

    total = len(file_list)
    results = []
    source_alias_map = source_alias_map or {}

    max_workers = converter.config.get("parallel_workers", 4)
    max_pending = converter.config.get("parallel_max_pending", max_workers * 2)
    if not isinstance(max_pending, int) or max_pending <= 0:
        max_pending = max_workers * 2
    max_pending = max(max_pending, max_workers)
    checkpoint_interval = converter.config.get("parallel_checkpoint_interval", 10)

    checkpoint, pending_files = converter._init_checkpoint(file_list)

    if checkpoint and len(pending_files) < total:
        file_list = pending_files
        total = len(file_list)
        print(f"[parallel] resume pending files: {total}")

    print(f"[parallel] start {max_workers} workers to process {total} files")

    completed_count = 0
    lock = threading.Lock()

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        future_to_file = {}
        pending = set()

        def _consume_done(done_futures):
            nonlocal checkpoint, completed_count
            for future in done_futures:
                pending.discard(future)
                meta = future_to_file.pop(future, None)
                if not meta:
                    continue
                fpath, logical_source, fname, started_at = meta

                try:
                    status, final_path = future.result()
                    elapsed = time.time() - started_at

                    with lock:
                        completed_count += 1

                    if status.startswith("skip"):
                        print(
                            f"\r[{completed_count}/{total}] {status}: {fname} ({elapsed:.2f}s)    "
                        )
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
                        with lock:
                            converter.stats["success"] += 1
                        print(
                            f"\r[{completed_count}/{total}] {status}: {fname} ({elapsed:.2f}s)    "
                        )
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

                    if checkpoint:
                        checkpoint = converter._mark_file_done_in_checkpoint(checkpoint, fpath)
                        if completed_count % checkpoint_interval == 0:
                            converter._save_checkpoint(checkpoint)

                except (OSError, RuntimeError, TypeError, ValueError, AttributeError) as exc:
                    elapsed = time.time() - started_at
                    err_msg = str(exc)

                    with lock:
                        completed_count += 1
                        converter.stats["failed"] += 1

                    error_detail = converter.record_detailed_error(
                        logical_source,
                        exc,
                        context={
                            "run_mode": converter.get_readable_run_mode(),
                            "engine": converter.get_readable_engine_type(),
                            "elapsed": elapsed,
                            "parallel": True,
                        },
                    )

                    print(
                        f"\r[{completed_count}/{total}] failed({error_detail['error_type']}): {fname}    "
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
                        "root_error_type": error_detail.get("root_error_type", ""),
                        "root_error_message": error_detail.get("root_error_message", ""),
                        "root_error_stage": error_detail.get("root_error_stage", ""),
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
                        extra_context={"parallel": True},
                    )
                    trace_path = converter._write_failed_file_trace_log(
                        trace_payload,
                        failed_copy_path=failed_copy_path,
                    )
                    if trace_path:
                        record["failure_trace_path"] = trace_path

        for idx, fpath in enumerate(file_list, 1):
            if not converter.is_running:
                break
            while len(pending) >= max_pending and converter.is_running:
                done, _ = wait(pending, return_when=FIRST_COMPLETED)
                _consume_done(done)

            logical_source = source_alias_map.get(fpath, fpath)
            ext = os.path.splitext(fpath)[1].lower()
            target_path_initial = converter.get_target_path(logical_source, ext)
            progress_prefix = converter.get_progress_prefix(idx, total)

            future = executor.submit(
                converter._convert_single_file_threadsafe,
                fpath,
                target_path_initial,
                ext,
                progress_prefix,
                is_retry,
            )
            pending.add(future)
            future_to_file[future] = (
                fpath,
                logical_source,
                os.path.basename(fpath),
                time.time(),
            )

        while pending and converter.is_running:
            done, _ = wait(pending, return_when=FIRST_COMPLETED)
            _consume_done(done)

        if not converter.is_running and pending:
            executor.shutdown(wait=False, cancel_futures=True)

    if checkpoint and checkpoint.get("status") == "completed":
        converter._clear_checkpoint()

    return results


def convert_single_file_threadsafe_for_converter(
    converter,
    fpath,
    target_path_initial,
    ext,
    progress_prefix,
    is_retry=False,
    *,
    has_win32,
    co_initialize_fn,
    co_uninitialize_fn,
):
    if has_win32:
        co_initialize_fn()
    try:
        return converter.process_single_file(
            fpath,
            target_path_initial,
            ext,
            progress_prefix,
            is_retry,
        )
    finally:
        if has_win32:
            co_uninitialize_fn()
