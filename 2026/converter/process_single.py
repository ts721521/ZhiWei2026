# -*- coding: utf-8 -*-
"""Single-file conversion pipeline extracted from office_converter.py."""

import logging
import os
import shutil
import threading
import time
import uuid

from converter.constants import STRATEGY_PRICE_ONLY, STRATEGY_STANDARD


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
    ppt_timeout = converter.config.get("ppt_timeout_seconds", base_timeout)
    current_timeout = ppt_timeout if is_ppt else base_timeout

    base_wait = converter.config.get("pdf_wait_seconds", 15)
    ppt_wait = converter.config.get("ppt_pdf_wait_seconds", base_wait)
    current_pdf_wait = ppt_wait if is_ppt else base_wait

    try:
        convert_core_start = time.perf_counter()
        if use_sandbox:
            sandbox_src_path = os.path.join(converter.temp_sandbox, f"{uuid.uuid4()}{ext}")
            shutil.copy2(file_path, sandbox_src_path)
            converter._unblock_file(sandbox_src_path)
            working_src = sandbox_src_path

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
            convert_thread = threading.Thread(
                target=converter.convert_logic_in_thread,
                args=(working_src, sandbox_pdf, ext, result_context),
                daemon=True,
            )
            convert_thread.start()

            wait_start = time.time()
            while convert_thread.is_alive():
                elapsed = time.time() - wait_start
                if elapsed > current_timeout:
                    break
                print(
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
                converter.stats["timeout"] += 1
                logging.error(f"timeout skip (>{current_timeout}s)")
                if is_word:
                    converter._kill_current_app("word", force=True)
                elif is_excel:
                    converter._kill_current_app("excel", force=True)
                elif is_ppt:
                    converter._kill_current_app("ppt", force=True)
                raise Exception("timeout")
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
        while time.perf_counter() - wait_pdf_start < current_pdf_wait:
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
                        md_path_res = converter._export_pdf_markdown(final_path_res) or ""
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
        raise Exception(
            f"conversion command sent but PDF not generated ({current_pdf_wait}s)"
        )

    finally:
        if is_office:
            converter._on_office_file_processed(ext)
        try:
            if sandbox_src_path and os.path.exists(sandbox_src_path):
                os.remove(sandbox_src_path)
            if os.path.exists(sandbox_pdf):
                os.remove(sandbox_pdf)
        except Exception:
            pass
