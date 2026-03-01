# -*- coding: utf-8 -*-
"""Fast Markdown engine workflow (Office -> Markdown, no COM/PDF)."""

import logging
import os
from datetime import datetime

from converter.traceability import apply_short_id_prefix


def _safe_stem(name):
    stem = str(name or "file").strip()
    out = []
    for ch in stem:
        if ch.isalnum() or ch in ("-", "_", "."):
            out.append(ch)
        else:
            out.append("_")
    return "".join(out).strip("._") or "file"


def _resolve_markdown_max_size_bytes(config, default_mb=80):
    cfg = config or {}
    raw = cfg.get("markdown_max_size_mb", cfg.get("max_merge_size_mb", default_mb))
    try:
        mb = int(raw)
    except (TypeError, ValueError):
        mb = int(default_mb)
    if mb <= 0:
        return 0
    return mb * 1024 * 1024


def run_fast_md_pipeline(
    files,
    *,
    config,
    target_folder,
    generated_markdown_outputs,
    markdown_quality_records,
    trace_short_id_taken,
    stats,
    source_root_resolver_fn,
    compute_md5_fn,
    build_short_id_fn,
    convert_source_to_markdown_text_fn,
    progress_callback=None,
    emit_file_done_fn=None,
    on_file_done_for_checkpoint=None,
    now_fn=datetime.now,
    log_info_fn=logging.info,
    log_warning_fn=logging.warning,
):
    if not files:
        return {"batch_results": [], "bundle_path": None, "markdown_files": []}

    out_dir = os.path.join(target_folder, "_MD_Corpus")
    os.makedirs(out_dir, exist_ok=True)

    batch_results = []
    markdown_files = []
    short_id_prefix = config.get("short_id_prefix", "ZW-")
    markdown_max_size_bytes = _resolve_markdown_max_size_bytes(config)

    total = len(files)
    for idx, source_path in enumerate(files, 1):
        if callable(progress_callback):
            try:
                progress_callback(idx, total)
            except (TypeError, ValueError, AttributeError, RuntimeError) as exc:
                log_warning_fn(f"fast-md progress callback failed: {exc}")
        src_abs = os.path.abspath(source_path)
        src_name = os.path.basename(src_abs)
        source_md5 = ""
        short_id = ""
        try:
            source_md5 = compute_md5_fn(src_abs)
        except (OSError, RuntimeError, TypeError, ValueError, AttributeError):
            source_md5 = ""
        if source_md5:
            try:
                short_id = build_short_id_fn(source_md5, trace_short_id_taken)
            except (RuntimeError, TypeError, ValueError, AttributeError):
                short_id = source_md5[:8].upper()
        short_id = apply_short_id_prefix(short_id, short_id_prefix) if short_id else ""

        try:
            body = convert_source_to_markdown_text_fn(src_abs) or ""
            if not str(body).strip():
                body = "(empty)"
        except (OSError, RuntimeError, TypeError, ValueError, AttributeError) as exc:
            stats["failed"] += 1
            log_warning_fn(f"fast-md convert failed {src_abs}: {exc}")
            record = {
                "source_path": src_abs,
                "status": "failed",
                "detail": f"fast_md_failed: {exc}",
                "final_path": "",
                "elapsed": 0.0,
            }
            batch_results.append(record)
            if callable(emit_file_done_fn):
                emit_file_done_fn(dict(record))
            if callable(on_file_done_for_checkpoint):
                on_file_done_for_checkpoint(src_abs)
            continue

        stem = _safe_stem(os.path.splitext(src_name)[0])
        md_name = f"{short_id}_{stem}.md" if short_id else f"{stem}.md"
        md_path = os.path.join(out_dir, md_name)
        rel_path = src_abs
        try:
            root = source_root_resolver_fn(src_abs)
            rel_path = os.path.relpath(src_abs, root)
        except (OSError, RuntimeError, TypeError, ValueError, AttributeError):
            rel_path = src_abs

        md_text = (
            "---\n"
            f"source_file: {src_abs}\n"
            f"source_short_id: {short_id}\n"
            f"source_md5: {source_md5}\n"
            "---\n\n"
            f"# {src_name}\n\n"
            f"{body.rstrip()}\n"
        )
        md_size = len(md_text.encode("utf-8"))
        if markdown_max_size_bytes > 0 and md_size > markdown_max_size_bytes:
            stats["failed"] += 1
            log_warning_fn(
                f"fast-md output oversize {src_abs}: {md_size} > {markdown_max_size_bytes} bytes"
            )
            record = {
                "source_path": src_abs,
                "status": "failed",
                "detail": (
                    "fast_md_output_oversize:"
                    f" {md_size} > {markdown_max_size_bytes} bytes"
                ),
                "final_path": "",
                "elapsed": 0.0,
            }
            batch_results.append(record)
            if callable(emit_file_done_fn):
                emit_file_done_fn(dict(record))
            if callable(on_file_done_for_checkpoint):
                on_file_done_for_checkpoint(src_abs)
            continue

        with open(md_path, "w", encoding="utf-8") as f:
            f.write(md_text)

        generated_markdown_outputs.append(md_path)
        markdown_files.append(md_path)
        markdown_quality_records.append(
            {
                "source_pdf": src_abs,
                "markdown_path": os.path.abspath(md_path),
                "source_short_id": short_id,
                "source_md5": source_md5,
                "source_filename": src_name,
                "source_abspath": src_abs,
                "source_relpath": rel_path,
                "page_count": 0,
                "non_empty_page_count": 1,
                "removed_header_lines": 0,
                "removed_footer_lines": 0,
                "removed_page_number_lines": 0,
                "heading_count": 0,
                "header_candidate_count": 0,
                "footer_candidate_count": 0,
                "strip_header_footer": False,
                "structured_headings": False,
            }
        )
        stats["success"] += 1
        record = {
            "source_path": src_abs,
            "status": "success_md_only",
            "detail": "fast_md",
            "final_path": md_path,
            "elapsed": 0.0,
        }
        batch_results.append(record)
        if callable(emit_file_done_fn):
            emit_file_done_fn(dict(record))
        if callable(on_file_done_for_checkpoint):
            on_file_done_for_checkpoint(src_abs)
        log_info_fn(f"fast-md export success: {md_path}")

    bundle_path = None
    if markdown_files:
        groups = []
        if markdown_max_size_bytes <= 0:
            groups = [list(markdown_files)]
        else:
            current = []
            current_size = 0
            for path in markdown_files:
                try:
                    f_size = os.path.getsize(path)
                except (OSError, RuntimeError, TypeError, ValueError):
                    f_size = 0
                estimated = f_size + 32768
                if current and (current_size + estimated > markdown_max_size_bytes):
                    groups.append(current)
                    current = [path]
                    current_size = estimated
                else:
                    current.append(path)
                    current_size += estimated
            if current:
                groups.append(current)

        for idx, group in enumerate(groups, 1):
            bundle_name = (
                "_Knowledge_Bundle.md"
                if len(groups) == 1
                else f"_Knowledge_Bundle_{idx:03d}.md"
            )
            current_bundle_path = os.path.join(out_dir, bundle_name)
            with open(current_bundle_path, "w", encoding="utf-8") as out:
                out.write(
                    f"# _Knowledge_Bundle\n\n- generated_at: {now_fn().isoformat(timespec='seconds')}\n\n"
                )
                for doc_idx, path in enumerate(group, 1):
                    if doc_idx > 1:
                        out.write("\n---\n\n")
                    out.write(f"# {os.path.basename(path)}\n\n")
                    with open(path, "r", encoding="utf-8", errors="ignore") as src:
                        while True:
                            chunk = src.read(64 * 1024)
                            if not chunk:
                                break
                            out.write(chunk)
                    out.write("\n")
            generated_markdown_outputs.append(current_bundle_path)
            if bundle_path is None:
                bundle_path = current_bundle_path
            log_info_fn(f"fast-md bundle generated: {current_bundle_path}")

    return {
        "batch_results": batch_results,
        "bundle_path": bundle_path,
        "markdown_files": markdown_files,
    }
