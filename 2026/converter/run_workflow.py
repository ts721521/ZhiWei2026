# -*- coding: utf-8 -*-
"""Top-level run workflow extracted from office_converter.py."""

import logging
import os
import shutil
import time

from converter.constants import (
    MODE_COLLECT_ONLY,
    MODE_CONVERT_ONLY,
    MODE_CONVERT_THEN_MERGE,
    MODE_MERGE_ONLY,
    MODE_MSHELP_ONLY,
)


def run(converter, resume_file_list=None, app_version=""):
    converter.setup_logging()
    converter.print_runtime_summary()
    converter._reset_perf_metrics()
    converter._office_file_counter = 0
    total_start = time.perf_counter()

    merge_outputs = []
    batch_results = []
    converter.generated_pdfs = []
    converter.generated_merge_outputs = []
    converter.generated_merge_markdown_outputs = []
    converter.generated_map_outputs = []
    converter.generated_markdown_outputs = []
    converter.generated_markdown_quality_outputs = []
    converter.generated_excel_json_outputs = []
    converter.generated_records_json_outputs = []
    converter.generated_chromadb_outputs = []
    converter.generated_update_package_outputs = []
    converter.generated_mshelp_outputs = []
    converter.markdown_quality_records = []
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
    converter.llm_hub_root = ""
    converter.incremental_registry_path = ""
    converter._incremental_context = None

    try:
        converter._check_sandbox_free_space_or_raise()
    except Exception as exc:
        logging.error(f"sandbox free space precheck failed: {exc}")
        raise

    if converter.run_mode == MODE_COLLECT_ONLY:
        converter.collect_office_files_and_build_excel()
    elif converter.run_mode == MODE_MSHELP_ONLY:
        (
            batch_results,
            mshelp_dirs,
            mshelp_index_outputs,
            mshelp_merged_outputs,
        ) = converter._run_mshelp_only()
        summary = (
            f"\n=== MSHelp Final Summary v{app_version}) ===\n"
            f"MSHelpViewer dirs: {len(mshelp_dirs)}\n"
            f"total CAB files: {converter.stats['total']}\n"
            f"success: {converter.stats['success']}\n"
            f"failed: {converter.stats['failed']}\n"
            f"timeout: {converter.stats['timeout']}\n"
            f"skipped: {converter.stats['skipped']}\n"
            f"index outputs: {len(mshelp_index_outputs)}\n"
            f"merged outputs: {len(mshelp_merged_outputs)}\n"
        )
        logging.info(summary)
        print(summary)
    else:
        incremental_ctx = {}
        if converter.run_mode in (MODE_CONVERT_ONLY, MODE_CONVERT_THEN_MERGE):
            if resume_file_list is not None:
                files = []
                for path in resume_file_list:
                    path = str(path or "").strip()
                    if not path:
                        continue
                    files.append(os.path.abspath(path))
                logging.info(
                    "resume mode active: using provided file list count=%s",
                    len(files),
                )
            else:
                scan_start = time.perf_counter()
                logging.info("scanning files...")
                files = converter._scan_convert_candidates()
                logging.info("scan candidate file count: %s", len(files))

                files, source_priority_skips = converter._apply_source_priority_filter(files)
                if source_priority_skips:
                    converter.stats["skipped"] += len(source_priority_skips)
                    batch_results.extend(source_priority_skips)

                files, incremental_ctx = converter._apply_incremental_filter(files)
                if incremental_ctx.get("enabled"):
                    converter.stats["skipped"] += incremental_ctx.get("unchanged_count", 0)
                    renamed_pairs = incremental_ctx.get("renamed_pairs", []) or []
                    if renamed_pairs and not incremental_ctx.get("reprocess_renamed", False):
                        converter.stats["skipped"] += len(renamed_pairs)
                        for item in renamed_pairs:
                            batch_results.append(
                                {
                                    "source_path": item.get("to_path", ""),
                                    "status": "renamed_detected",
                                    "detail": "rename detected; no reconvert",
                                    "final_path": "",
                                    "renamed_from": item.get("from_path", ""),
                                }
                            )

                files, dedup_skips = converter._apply_global_md5_dedup(files)
                if dedup_skips:
                    converter.stats["skipped"] += len(dedup_skips)
                    batch_results.extend(dedup_skips)
                converter._add_perf_seconds("scan_seconds", time.perf_counter() - scan_start)

            converter._emit_file_plan(files)

            converter.stats["total"] = len(files)
            if files:
                logging.info("start processing %s files", len(files))
                batch_start = time.perf_counter()

                if converter.config.get("enable_parallel_conversion", False):
                    print(
                        f"[parallel] using {converter.config.get('parallel_workers', 4)} workers"
                    )
                    batch_results.extend(converter.run_batch_parallel(files))
                else:
                    batch_results.extend(converter.run_batch(files))

                converter._add_perf_seconds("batch_seconds", time.perf_counter() - batch_start)
            else:
                if resume_file_list is not None:
                    print("\n[INFO] Resume mode: no pending files.")
                elif incremental_ctx.get("enabled"):
                    print("\n[INFO] Incremental mode: no added/modified files found.")
                else:
                    print("\n[INFO] No convertible Office files found in source directory.")

            converter.close_office_apps()

            failed_count = converter.stats["failed"] + converter.stats["timeout"]
            should_retry = False
            if failed_count > 0:
                if converter.config.get("auto_retry_failed", False):
                    should_retry = True
                    print(f"\n[CONFIG] Auto retry failed files ({failed_count})...")
                elif converter.interactive:
                    should_retry = converter.ask_retry_failed_files(failed_count, timeout=20)

            if should_retry:
                print("\n" + "=" * 60)
                print("  Start retrying failed files...")
                print("  Re-checking and cleaning related processes...")
                print("=" * 60)

                if not converter.reuse_process:
                    converter.cleanup_all_processes()

                if os.path.exists(converter.failed_dir):
                    if converter.config.get("enable_sandbox", True) and not os.path.exists(
                        converter.temp_sandbox
                    ):
                        os.makedirs(converter.temp_sandbox)

                retry_files, retry_alias_map = converter._collect_retry_candidates()

                if retry_files:
                    retry_start = time.perf_counter()
                    batch_results.extend(
                        converter.run_batch(
                            retry_files,
                            is_retry=True,
                            source_alias_map=retry_alias_map,
                        )
                    )
                    converter._add_perf_seconds(
                        "batch_seconds",
                        time.perf_counter() - retry_start,
                    )
                else:
                    print("No retryable files found in failed directory.")

                converter.close_office_apps()

            if converter.run_mode == MODE_CONVERT_ONLY and converter.enable_merge_excel:
                try:
                    converter._write_conversion_index_workbook()
                except Exception as exc:
                    logging.error(f"failed to write conversion index: {exc}")

            try:
                converter._flush_incremental_registry(batch_results)
            except Exception as exc:
                logging.error(f"failed to write incremental registry: {exc}")
            try:
                converter._generate_update_package(batch_results)
            except Exception as exc:
                logging.error(f"failed to generate update package: {exc}")

        elif converter.run_mode == MODE_MERGE_ONLY:
            print("Current mode is merge_and_convert. Conversion step skipped.")
            merge_start = time.perf_counter()
            merge_outputs = converter._run_merge_mode_pipeline(batch_results) or []
            converter._add_perf_seconds("merge_seconds", time.perf_counter() - merge_start)

        if (
            converter.run_mode == MODE_CONVERT_THEN_MERGE
            and converter.config.get("enable_merge", True)
            and bool(converter.config.get("output_enable_merged", True))
        ):
            merge_start = time.perf_counter()
            merge_outputs = []
            if bool(converter.config.get("output_enable_pdf", True)):
                merge_outputs.extend(converter.merge_pdfs() or [])
            if bool(converter.config.get("output_enable_md", True)):
                md_outputs = (
                    converter.merge_markdowns(candidates=converter.generated_markdown_outputs)
                    or []
                )
                merge_outputs.extend(md_outputs)
            converter._add_perf_seconds("merge_seconds", time.perf_counter() - merge_start)

        summary = (
            f"\n=== Final Summary v{app_version}) ===\n"
            f"total: {converter.stats['total']}\n"
            f"success: {converter.stats['success']}\n"
            f"failed: {converter.stats['failed']}\n"
            f"timeout: {converter.stats['timeout']}\n"
            f"skipped(empty/strategy): {converter.stats['skipped']}\n"
        )

        if converter._incremental_context and converter._incremental_context.get("enabled"):
            inc = converter._incremental_context
            summary += (
                f"incremental scan: {inc.get('scanned_count', 0)} | "
                f"added: {inc.get('added_count', 0)} | "
                f"modified: {inc.get('modified_count', 0)} | "
                f"renamed: {inc.get('renamed_count', 0)} | "
                f"unchanged: {inc.get('unchanged_count', 0)} | "
                f"deleted: {inc.get('deleted_count', 0)}\n"
            )
            if converter.incremental_registry_path:
                summary += f"incremental registry: {converter.incremental_registry_path}\n"
            if converter.update_package_manifest_path:
                summary += f"update package manifest: {converter.update_package_manifest_path}\n"

        logging.info(summary)
        print(summary)

    postprocess_start = time.perf_counter()
    try:
        converter._write_excel_structured_json_exports()
    except Exception as exc:
        logging.error(f"failed to write Excel JSON: {exc}")

    try:
        converter._write_records_json_exports()
    except Exception as exc:
        logging.error(f"failed to write Records JSON: {exc}")

    try:
        converter._write_chromadb_export()
    except Exception as exc:
        logging.error(f"failed to write ChromaDB export: {exc}")

    try:
        converter._write_markdown_quality_report()
    except Exception as exc:
        logging.error(f"failed to write Markdown quality report: {exc}")

    try:
        converter._write_corpus_manifest(merge_outputs=merge_outputs)
    except Exception as exc:
        logging.error(f"failed to write corpus manifest: {exc}")

    if converter.detailed_error_records:
        try:
            report_result = converter.export_failed_files_report()
            if report_result.get("txt_path"):
                failed_summary = (
                    f"\n=== Failed Files Report ===\n"
                    f"report path: {report_result['txt_path']}\n"
                    f"summary: {report_result['summary']}\n"
                )
                print(failed_summary)
                logging.info(failed_summary)
        except Exception as exc:
            logging.error(f"failed to export failed files report: {exc}")

    converter._add_perf_seconds("postprocess_seconds", time.perf_counter() - postprocess_start)
    converter.perf_metrics["total_seconds"] = time.perf_counter() - total_start
    perf_summary = converter._build_perf_summary()
    logging.info(perf_summary)
    print(perf_summary)

    try:
        open_dir = converter.config["target_folder"]
        if merge_outputs and converter.merge_output_dir and os.path.isdir(converter.merge_output_dir):
            open_dir = converter.merge_output_dir
        if hasattr(os, "startfile"):
            os.startfile(open_dir)
    except Exception:
        pass

    if converter.temp_sandbox and os.path.exists(converter.temp_sandbox):
        try:
            shutil.rmtree(converter.temp_sandbox, ignore_errors=True)
        except Exception:
            pass
