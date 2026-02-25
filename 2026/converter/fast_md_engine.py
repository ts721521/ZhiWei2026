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

    for source_path in files:
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
            batch_results.append(
                {
                    "source_path": src_abs,
                    "status": "failed",
                    "detail": f"fast_md_failed: {exc}",
                    "final_path": "",
                    "elapsed": 0.0,
                }
            )
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

        with open(md_path, "w", encoding="utf-8") as f:
            f.write("---\n")
            f.write(f"source_file: {src_abs}\n")
            f.write(f"source_short_id: {short_id}\n")
            f.write(f"source_md5: {source_md5}\n")
            f.write("---\n\n")
            f.write(f"# {src_name}\n\n")
            f.write(body.rstrip() + "\n")

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
        batch_results.append(
            {
                "source_path": src_abs,
                "status": "success_md_only",
                "detail": "fast_md",
                "final_path": md_path,
                "elapsed": 0.0,
            }
        )
        log_info_fn(f"fast-md export success: {md_path}")

    bundle_path = None
    if markdown_files:
        bundle_path = os.path.join(out_dir, "_Knowledge_Bundle.md")
        with open(bundle_path, "w", encoding="utf-8") as out:
            out.write(f"# _Knowledge_Bundle\n\n- generated_at: {now_fn().isoformat(timespec='seconds')}\n\n")
            for idx, path in enumerate(markdown_files, 1):
                if idx > 1:
                    out.write("\n---\n\n")
                out.write(f"# {os.path.basename(path)}\n\n")
                with open(path, "r", encoding="utf-8", errors="ignore") as src:
                    while True:
                        chunk = src.read(64 * 1024)
                        if not chunk:
                            break
                        out.write(chunk)
                out.write("\n")
        generated_markdown_outputs.append(bundle_path)
        log_info_fn(f"fast-md bundle generated: {bundle_path}")

    return {
        "batch_results": batch_results,
        "bundle_path": bundle_path,
        "markdown_files": markdown_files,
    }
