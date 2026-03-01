# -*- coding: utf-8 -*-
"""Single-file conversion pipeline extracted from office_converter.py."""

import logging
import os
import shutil
import threading
import time
import traceback
import uuid
import zipfile

from converter.constants import STRATEGY_PRICE_ONLY, STRATEGY_STANDARD
from converter.display_helpers import safe_console_print


def process_single_file(
    converter,
    file_path,
    target_path_initial,
    ext,
    progress_str,
    is_retry=False,
):
    if os.path.getsize(file_path) == 0:
        converter.stats["skipped"] += 1
        logging.warning(f"skip empty file: {file_path}")
        return "skip_empty", target_path_initial

    is_word = ext in converter.config["allowed_extensions"].get("word", [])
    is_excel = ext in converter.config["allowed_extensions"].get("excel", [])
    is_ppt = ext in converter.config["allowed_extensions"].get("powerpoint", [])
    is_pdf = ext == ".pdf"
    is_cab = ext in converter.config["allowed_extensions"].get("cab", [])
    is_office = is_word or is_excel or is_ppt

    filename = os.path.basename(file_path)

    is_filename_match = False
    if converter.content_strategy != STRATEGY_STANDARD:
        for kw in converter.price_keywords:
            if kw in filename:
                is_filename_match = True
                break

    if converter.content_strategy == STRATEGY_PRICE_ONLY and not is_filename_match:
        if is_word or is_ppt:
            converter.stats["skipped"] += 1
            return "skip_strategy", target_path_initial

    sandbox_pdf = os.path.join(converter.temp_sandbox, f"{uuid.uuid4()}.pdf")

    use_sandbox = converter.config.get("enable_sandbox", True)
    working_src = file_path
    sandbox_src_path = None

    final_target_path = target_path_initial

    result_context = {
        "is_price": is_filename_match,
        "scan_aborted": False,
        "skip_scan": is_filename_match,
    }
    output_plan = converter.compute_convert_output_plan(converter.run_mode, converter.config)

    if is_filename_match:
        final_target_path = converter.get_target_path(
            file_path,
            ext,
            prefix_override="Price_",
        )

    base_timeout = converter.config.get("timeout_seconds", 60)
    word_timeout = converter.config.get("word_timeout_seconds", base_timeout)
    excel_timeout = converter.config.get("excel_timeout_seconds", base_timeout)
    ppt_timeout = converter.config.get("ppt_timeout_seconds", base_timeout)
    if is_ppt:
        current_timeout = ppt_timeout
    elif is_excel:
        current_timeout = excel_timeout
    elif is_word:
        current_timeout = word_timeout
    else:
        current_timeout = base_timeout

    base_wait = converter.config.get("pdf_wait_seconds", 15)
    word_wait = converter.config.get("word_pdf_wait_seconds", base_wait)
    excel_wait = converter.config.get("excel_pdf_wait_seconds", base_wait)
    ppt_wait = converter.config.get("ppt_pdf_wait_seconds", base_wait)
    if is_ppt:
        current_pdf_wait = ppt_wait
    elif is_excel:
        current_pdf_wait = excel_wait
    elif is_word:
        current_pdf_wait = word_wait
    else:
        current_pdf_wait = base_wait
    office_retry_limit = int(converter.config.get("office_com_retry_times", 1) or 0)
    if office_retry_limit < 0:
        office_retry_limit = 0
    if is_ppt:
        ppt_retry_limit = int(
            converter.config.get("ppt_com_retry_times", max(office_retry_limit, 3)) or 0
        )
        office_retry_limit = max(office_retry_limit, ppt_retry_limit)
    office_retry_delay = float(converter.config.get("office_retry_delay_seconds", 0.8) or 0.8)
    retry_timeout_scale = float(converter.config.get("office_retry_timeout_scale", 1.5) or 1.0)
    retry_pdf_wait_scale = float(converter.config.get("office_retry_pdf_wait_scale", 1.5) or 1.0)
    retry_timeout_cap = int(converter.config.get("office_retry_timeout_cap_seconds", 0) or 0)
    retry_pdf_wait_cap = int(converter.config.get("office_retry_pdf_wait_cap_seconds", 0) or 0)
    if retry_timeout_scale < 1.0:
        retry_timeout_scale = 1.0
    if retry_pdf_wait_scale < 1.0:
        retry_pdf_wait_scale = 1.0

    def _kill_office_app_for_ext():
        if is_word:
            converter._kill_current_app("word", force=True)
        elif is_excel:
            converter._kill_current_app("excel", force=True)
        elif is_ppt:
            converter._kill_current_app("ppt", force=True)

    def _is_retryable_office_exception(exc):
        signal = " ".join(
            [
                str(exc or "").lower(),
                str(getattr(exc, "root_error_type", "") or "").lower(),
                str(getattr(exc, "root_error_message", "") or "").lower(),
            ]
        )
        return any(
            kw in signal
            for kw in (
                "com error",
                "com_error",
                "word.application",
                "excel.application",
                "powerpoint.application",
                "<unknown>.open",
                "conversion worker failed",
                "timeout",
                "pdfnotgenerated",
            )
        )

    try:
        if use_sandbox:
            sandbox_src_path = os.path.join(converter.temp_sandbox, f"{uuid.uuid4()}{ext}")
            shutil.copy2(file_path, sandbox_src_path)
            converter._unblock_file(sandbox_src_path)
            working_src = sandbox_src_path

        # Fast-fail obviously broken OOXML containers to avoid long COM hangs/retries.
        ooxml_required_entries = {
            ".docx": "word/document.xml",
            ".xlsx": "xl/workbook.xml",
            ".pptx": "ppt/presentation.xml",
        }
        if ext in ooxml_required_entries:
            try:
                with open(working_src, "rb") as raw_f:
                    signature = raw_f.read(4)
                if signature != b"PK\x03\x04":
                    raise ValueError(
                        f"file appears corrupted: invalid {ext} package signature ({file_path})"
                    )
                with zipfile.ZipFile(working_src, "r") as zf:
                    names = set(zf.namelist())
                    required_entry = ooxml_required_entries[ext]
                    if "[Content_Types].xml" not in names or required_entry not in names:
                        raise ValueError(
                            f"file appears corrupted: invalid {ext} package content ({file_path})"
                        )
            except (zipfile.BadZipFile, zipfile.LargeZipFile, OSError) as exc:
                raise ValueError(
                    f"file appears corrupted: invalid {ext} package ({file_path})"
                ) from exc

        attempt = 0
        while True:
            attempt_timeout = max(1, int(current_timeout * (retry_timeout_scale ** attempt)))
            attempt_pdf_wait = max(1, int(current_pdf_wait * (retry_pdf_wait_scale ** attempt)))
            if retry_timeout_cap > 0:
                attempt_timeout = min(attempt_timeout, retry_timeout_cap)
            if retry_pdf_wait_cap > 0:
                attempt_pdf_wait = min(attempt_pdf_wait, retry_pdf_wait_cap)
            convert_core_start = time.perf_counter()
            try:
                if is_cab:
                    md_path, rendered_count = converter._convert_cab_to_markdown(
                        working_src,
                        file_path,
                    )
                    converter._add_perf_seconds(
                        "convert_core_seconds",
                        time.perf_counter() - convert_core_start,
                    )
                    return f"success_cab_md[{rendered_count}]", md_path

                if is_pdf:
                    if not is_filename_match and converter.content_strategy != STRATEGY_STANDARD:
                        has_kw = converter.scan_pdf_content(working_src)
                        if has_kw:
                            result_context["is_price"] = True
                        elif converter.content_strategy == STRATEGY_PRICE_ONLY:
                            converter.stats["skipped"] += 1
                            return "skip_content", target_path_initial

                    converter.copy_pdf_direct(working_src, sandbox_pdf)
                    converter._add_perf_seconds(
                        "convert_core_seconds",
                        time.perf_counter() - convert_core_start,
                    )

                else:
                    thread_state = {"error": None, "traceback": ""}

                    def _convert_worker():
                        try:
                            converter.convert_logic_in_thread(
                                working_src,
                                sandbox_pdf,
                                ext,
                                result_context,
                            )
                        except BaseException as exc:  # pylint: disable=broad-except
                            thread_state["error"] = exc
                            thread_state["traceback"] = traceback.format_exc()
                            logging.error(
                                "conversion worker failed: %s",
                                file_path,
                                exc_info=True,
                            )

                    convert_thread = threading.Thread(
                        target=_convert_worker,
                        daemon=True,
                    )
                    convert_thread.start()

                    wait_start = time.time()
                    while convert_thread.is_alive():
                        elapsed = time.time() - wait_start
                        if elapsed > attempt_timeout:
                            break
                        safe_console_print(
                            f"{progress_str} converting: {filename} ({elapsed:.1f}s)    ",
                            end="",
                            flush=True,
                        )
                        time.sleep(0.1)

                    convert_thread.join(timeout=0.1)

                    if convert_thread.is_alive():
                        converter._add_perf_seconds(
                            "convert_core_seconds",
                            time.perf_counter() - convert_core_start,
                        )
                        logging.error(f"timeout skip (>{attempt_timeout}s)")
                        _kill_office_app_for_ext()
                        timeout_error = RuntimeError("timeout")
                        timeout_error.root_error_type = "TimeoutError"
                        timeout_error.root_error_stage = "convert_thread"
                        timeout_error.root_error_message = f"conversion exceeded timeout ({attempt_timeout}s)"
                        raise timeout_error

                    root_exc = thread_state.get("error")
                    if root_exc is not None:
                        converter._add_perf_seconds(
                            "convert_core_seconds",
                            time.perf_counter() - convert_core_start,
                        )
                        root_type = type(root_exc).__name__
                        root_msg = str(root_exc)
                        wrapped = RuntimeError(
                            f"conversion worker failed: {root_type}: {root_msg}"
                        )
                        wrapped.root_error_type = root_type
                        wrapped.root_error_message = root_msg[:500]
                        wrapped.root_error_stage = "convert_thread"
                        wrapped.root_error_traceback = (thread_state.get("traceback") or "")[:4000]
                        raise wrapped

                converter._add_perf_seconds(
                    "convert_core_seconds",
                    time.perf_counter() - convert_core_start,
                )

                if result_context["scan_aborted"]:
                    converter.stats["skipped"] += 1
                    return "skip_content", target_path_initial

                if result_context["is_price"]:
                    final_target_path = converter.get_target_path(
                        file_path,
                        ext,
                        prefix_override="Price_",
                    )

                wait_pdf_start = time.perf_counter()
                while time.perf_counter() - wait_pdf_start < attempt_pdf_wait:
                    if os.path.exists(sandbox_pdf):
                        time.sleep(0.5)
                        converter._add_perf_seconds(
                            "pdf_wait_seconds",
                            time.perf_counter() - wait_pdf_start,
                        )
                        final_path_res = ""
                        md_path_res = ""
                        if output_plan.get("need_final_pdf"):
                            result_status, final_path_res = converter.handle_file_conflict(
                                sandbox_pdf,
                                final_target_path,
                            )
                            converter.generated_pdfs.append(final_path_res)
                        else:
                            result_status = (
                                "success_no_output"
                                if not output_plan.get("need_markdown")
                                else "success_md_only"
                            )

                        if output_plan.get("need_markdown"):
                            markdown_start = time.perf_counter()
                            if final_path_res and os.path.exists(final_path_res):
                                md_path_res = (
                                    converter._export_pdf_markdown(
                                        final_path_res,
                                        source_path_hint=file_path,
                                    )
                                    or ""
                                )
                            else:
                                md_path_res = (
                                    converter._export_pdf_markdown(
                                        sandbox_pdf,
                                        source_path_hint=file_path,
                                    )
                                    or ""
                                )
                            converter._add_perf_seconds(
                                "markdown_seconds",
                                time.perf_counter() - markdown_start,
                            )

                        tag_info = ""
                        if is_filename_match:
                            tag_info = " [name_hit]"
                        elif result_context["is_price"]:
                            tag_info = " [content_hit]"
                        final_output_path = final_path_res or md_path_res
                        return f"{result_status}{tag_info}", final_output_path
                    time.sleep(0.5)

                converter._add_perf_seconds(
                    "pdf_wait_seconds",
                    time.perf_counter() - wait_pdf_start,
                )
                wait_error = RuntimeError(
                    f"conversion command sent but PDF not generated ({attempt_pdf_wait}s)"
                )
                wait_error.root_error_type = "PdfNotGenerated"
                wait_error.root_error_stage = "pdf_wait"
                wait_error.root_error_message = (
                    f"pdf file was not observed within wait window ({attempt_pdf_wait}s)"
                )
                raise wait_error
            except RuntimeError as exc:
                should_retry = is_office and attempt < office_retry_limit and _is_retryable_office_exception(exc)
                if not should_retry:
                    raise
                attempt += 1
                logging.warning(
                    "office retry %s/%s for %s due to: %s",
                    attempt,
                    office_retry_limit,
                    file_path,
                    exc,
                )
                _kill_office_app_for_ext()
                try:
                    if os.path.exists(sandbox_pdf):
                        os.remove(sandbox_pdf)
                except OSError:
                    pass
                if office_retry_delay > 0:
                    time.sleep(office_retry_delay)
                continue

    finally:
        if is_office:
            converter._on_office_file_processed(ext)
        try:
            if sandbox_src_path and os.path.exists(sandbox_src_path):
                os.remove(sandbox_src_path)
            if os.path.exists(sandbox_pdf):
                os.remove(sandbox_pdf)
        except OSError:
            pass
