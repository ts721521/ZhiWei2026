# -*- coding: utf-8 -*-
"""Corpus manifest and LLM delivery hub helpers extracted from office_converter.py."""

import json
import logging
import os
import shutil
from datetime import datetime


def maybe_build_llm_delivery_hub(converter, target_folder, artifacts):
    if not converter.config.get("enable_llm_delivery_hub", True):
        return None

    if not artifacts:
        return None

    obsidian_sync_enabled = bool(converter.config.get("obsidian_sync_enabled", False))
    obsidian_root = str(converter.config.get("obsidian_root", "") or "").strip()
    obsidian_program_name = (
        str(converter.config.get("obsidian_program_name", "ZhiWei") or "ZhiWei").strip()
        or "ZhiWei"
    )

    if obsidian_sync_enabled and obsidian_root:
        llm_root = os.path.join(obsidian_root, obsidian_program_name)
    else:
        llm_root = converter.config.get("llm_delivery_root") or os.path.join(
            target_folder, "_LLM_UPLOAD"
        )
    try:
        os.makedirs(llm_root, exist_ok=True)
    except (OSError, RuntimeError, TypeError, ValueError) as exc:
        logging.error(f"failed to create LLM hub root {llm_root}: {exc}")
        return None

    include_pdf = converter.config.get("llm_delivery_include_pdf", False)
    flatten = converter.config.get("llm_delivery_flatten", False)
    if obsidian_sync_enabled:
        flatten = False

    hub_items = []
    counts = {
        "markdown": 0,
        "json": 0,
        "pdf": 0,
        "other": 0,
    }

    llm_content_kinds = {
        "markdown_export",
        "merged_markdown",
        "mshelp_merged_markdown",
        "excel_structured_json",
    }
    llm_pdf_kinds = {
        "merged_pdf",
        "converted_pdf",
        "mshelp_merged_pdf",
    }

    dedup_enabled = converter.config.get("upload_dedup_merged", True)
    has_merged_pdf = dedup_enabled and any(
        artifact.get("kind") == "merged_pdf" for artifact in artifacts
    )
    has_merged_md = dedup_enabled and any(
        artifact.get("kind") == "merged_markdown" for artifact in artifacts
    )
    has_mshelp_merged = dedup_enabled and any(
        artifact.get("kind") == "mshelp_merged_markdown" for artifact in artifacts
    )

    mshelp_source_paths = set()
    if has_mshelp_merged:
        for record in getattr(converter, "mshelp_records", []) or []:
            md_path = record.get("markdown_path", "")
            if md_path:
                mshelp_source_paths.add(os.path.normcase(os.path.abspath(md_path)))

    for artifact in artifacts:
        kind = artifact.get("kind", "")
        rel = artifact.get("path_rel_to_target") or ""
        abs_path = artifact.get("path_abs") or ""
        if not rel or not abs_path:
            continue

        is_content = kind in llm_content_kinds
        is_pdf_kind = kind in llm_pdf_kinds

        if not (is_content or (include_pdf and is_pdf_kind)):
            continue

        if has_merged_pdf and kind == "converted_pdf":
            continue

        if has_merged_md and kind == "markdown_export":
            continue

        if has_mshelp_merged and kind == "markdown_export":
            norm = os.path.normcase(os.path.abspath(abs_path))
            if norm in mshelp_source_paths:
                continue

        ext = os.path.splitext(rel.lower())[1]
        if ext == ".md":
            counts["markdown"] += 1
            category = "Markdown"
        elif ext in (".json", ".jsonl"):
            counts["json"] += 1
            category = "JSON"
        elif ext == ".pdf":
            counts["pdf"] += 1
            category = "PDF"
        else:
            counts["other"] += 1
            category = "Files"

        if flatten:
            base_name = os.path.basename(rel)
            hub_rel = base_name
            hub_rel_base, hub_rel_ext = os.path.splitext(hub_rel)
            candidate = os.path.join(llm_root, hub_rel)
            collision_idx = 1
            while os.path.exists(candidate):
                hub_rel = f"{hub_rel_base}_{collision_idx}{hub_rel_ext}"
                candidate = os.path.join(llm_root, hub_rel)
                collision_idx += 1
        else:
            hub_rel = os.path.join(category, rel)

        hub_abs = os.path.join(llm_root, hub_rel)
        hub_dir = os.path.dirname(hub_abs)
        try:
            os.makedirs(hub_dir, exist_ok=True)
        except (OSError, RuntimeError, TypeError, ValueError) as exc:
            logging.error(f"failed to create LLM hub subdir {hub_dir}: {exc}")
            continue

        try:
            shutil.copy2(abs_path, hub_abs)
        except (OSError, RuntimeError, TypeError, ValueError) as exc:
            logging.error(f"failed to copy to LLM hub {hub_abs}: {exc}")
            continue

        try:
            size_bytes = os.path.getsize(hub_abs)
        except OSError:
            size_bytes = 0

        hub_items.append(
            {
                "kind": kind,
                "source_abs_path": abs_path,
                "delivery_rel_path": hub_rel.replace("\\", "/"),
                "size_bytes": int(size_bytes),
                "md5": artifact.get("md5", ""),
                "sha256": artifact.get("sha256", ""),
            }
        )

    if not hub_items:
        return None

    manifest = {
        "version": 1,
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "run_mode": converter.run_mode,
        "source_folder": converter.config.get("source_folder", ""),
        "target_folder": target_folder,
        "hub_root": llm_root,
        "items": hub_items,
        "summary": counts,
    }

    manifest_path = None
    if converter.config.get("enable_upload_json_manifest", True):
        manifest_path = os.path.join(llm_root, "llm_upload_manifest.json")
        try:
            with open(manifest_path, "w", encoding="utf-8") as f:
                json.dump(manifest, f, ensure_ascii=False, indent=2)
        except (OSError, RuntimeError, TypeError, ValueError) as exc:
            logging.error(f"failed to write LLM hub manifest {manifest_path}: {exc}")

    if converter.config.get("enable_upload_readme", True):
        readme_path = os.path.join(llm_root, "README_UPLOAD_LIST.txt")
        try:
            total_size = sum(item["size_bytes"] for item in hub_items)
            readme_lines = [
                "=== LLM Upload File List / 上传文件清单 ===",
                f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
                f"Total files: {len(hub_items)}",
                f"Total size: {total_size / 1024 / 1024:.1f} MB",
                f"  Markdown: {counts['markdown']}  |  JSON: {counts['json']}  |  PDF: {counts['pdf']}",
                "",
                "--- File List ---",
            ]
            for idx, item in enumerate(hub_items, 1):
                size_mb = item["size_bytes"] / 1024 / 1024
                readme_lines.append(
                    f"  {idx:3d}. [{item['kind']}] {item['delivery_rel_path']}  ({size_mb:.2f} MB)"
                )
            readme_lines.append("")
            readme_lines.append("--- Notes ---")
            if has_merged_pdf:
                readme_lines.append(
                    "* Individual converted PDFs are excluded (already in merged volumes)."
                )
            if has_merged_md:
                readme_lines.append(
                    "* Individual markdown files are excluded (already in merged markdown packages)."
                )
            if has_mshelp_merged:
                readme_lines.append(
                    "* Individual MSHelp markdowns are excluded (already in merged packages)."
                )
            readme_lines.append("* Metadata files (manifests, quality reports, index) are excluded.")
            readme_lines.append("")
            with open(readme_path, "w", encoding="utf-8") as f:
                f.write("\n".join(readme_lines))
        except (OSError, RuntimeError, TypeError, ValueError) as exc:
            logging.warning(f"failed to write upload readme: {exc}")

    logging.info(
        "LLM hub built at %s | files: %s (md=%s, json=%s, pdf=%s)",
        llm_root,
        len(hub_items),
        counts["markdown"],
        counts["json"],
        counts["pdf"],
    )

    converter.llm_hub_root = llm_root

    return {
        "kind": "llm_delivery_hub",
        "path_abs": llm_root,
        "path_rel_to_target": os.path.relpath(llm_root, target_folder).replace("\\", "/"),
        "size_bytes": 0,
        "mtime": datetime.now().isoformat(timespec="seconds"),
        "md5": "",
        "sha256": "",
        "manifest_path": manifest_path,
    }


def write_corpus_manifest(converter, merge_outputs=None):
    if not converter.config.get("enable_corpus_manifest", True):
        return None

    target_folder = converter.config.get("target_folder", "")
    if not target_folder:
        return None
    os.makedirs(target_folder, exist_ok=True)

    artifacts = []
    seen = set()

    def _append(kind, path):
        if not path:
            return
        abs_path = os.path.abspath(path)
        key = (kind, abs_path)
        if key in seen:
            return
        converter._add_artifact(artifacts, kind, abs_path)
        seen.add(key)

    for path in converter.generated_pdfs:
        _append("converted_pdf", path)
    for path in merge_outputs or converter.generated_merge_outputs or []:
        path_low = str(path).lower()
        if path_low.endswith(".md"):
            _append("merged_markdown", path)
        else:
            _append("merged_pdf", path)
    for path in converter.generated_merge_markdown_outputs:
        _append("merged_markdown", path)
    for path in converter.generated_map_outputs:
        path_low = str(path).lower()
        if path_low.endswith(".map.csv"):
            _append("merge_map_csv", path)
        elif path_low.endswith(".map.json"):
            _append("merge_map_json", path)
        else:
            _append("merge_map_file", path)
    for path in converter.generated_markdown_outputs:
        _append("markdown_export", path)
    for path in getattr(converter, "generated_markdown_manifest_outputs", []) or []:
        _append("markdown_image_manifest", path)
    for path in getattr(converter, "generated_fast_md_outputs", []) or []:
        _append("knowledge_bundle_markdown", path)
    for path in converter.generated_markdown_quality_outputs:
        _append("markdown_quality_report", path)
    for path in converter.generated_excel_json_outputs:
        _append("excel_structured_json", path)
    for path in converter.generated_records_json_outputs:
        _append("records_json", path)
    for path in getattr(converter, "generated_trace_map_outputs", []) or []:
        _append("trace_map_xlsx", path)
    for path in getattr(converter, "generated_prompt_outputs", []) or []:
        _append("prompt_ready_txt", path)
    for path in converter.generated_chromadb_outputs:
        path_low = str(path).lower()
        if path_low.endswith(".jsonl"):
            _append("chromadb_docs_jsonl", path)
        elif path_low.endswith(".json"):
            _append("chromadb_export_manifest", path)
        else:
            _append("chromadb_export_file", path)
    for path in converter.generated_update_package_outputs:
        path_low = str(path).lower()
        if path_low.endswith("incremental_manifest.json"):
            _append("update_package_manifest", path)
        elif path_low.endswith("incremental_index.xlsx"):
            _append("update_package_index_xlsx", path)
        elif path_low.endswith("incremental_index.csv"):
            _append("update_package_index_csv", path)
        elif path_low.endswith("incremental_index.json"):
            _append("update_package_index_json", path)
        else:
            _append("update_package_file", path)
    for path in converter.generated_mshelp_outputs:
        path_low = str(path).lower()
        if path_low.endswith(".json") and "mshelp_index_" in path_low:
            _append("mshelp_index_json", path)
        elif path_low.endswith(".csv") and "mshelp_index_" in path_low:
            _append("mshelp_index_csv", path)
        elif path_low.endswith(".docx"):
            _append("mshelp_merged_docx", path)
        elif path_low.endswith(".pdf"):
            _append("mshelp_merged_pdf", path)
        elif path_low.endswith(".md"):
            _append("mshelp_merged_markdown", path)
        else:
            _append("mshelp_output_file", path)

    _append("convert_index_excel", converter.convert_index_path)
    _append("collect_index_excel", converter.collect_index_path)
    _append("merge_list_excel", converter.merge_excel_path)

    manifest = {
        "version": 1,
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "run_mode": converter.run_mode,
        "collect_mode": converter.collect_mode,
        "merge_mode": converter.merge_mode,
        "content_strategy": converter.content_strategy,
        "source_folder": converter.config.get("source_folder", ""),
        "target_folder": target_folder,
        "artifacts": artifacts,
        "conversion_records": converter.conversion_index_records,
        "merge_records": converter.merge_index_records,
        "summary": {
            "converted_pdf_count": len(converter.generated_pdfs),
            "merged_pdf_count": len(
                [
                    path
                    for path in (merge_outputs or converter.generated_merge_outputs or [])
                    if not str(path).lower().endswith(".md")
                ]
            ),
            "merged_markdown_count": len(converter.generated_merge_markdown_outputs),
            "merge_map_count": len(converter.generated_map_outputs),
            "markdown_count": len(converter.generated_markdown_outputs),
            "markdown_image_manifest_count": len(
                getattr(converter, "generated_markdown_manifest_outputs", []) or []
            ),
            "knowledge_bundle_count": len(getattr(converter, "generated_fast_md_outputs", []) or []),
            "markdown_quality_report_count": len(converter.generated_markdown_quality_outputs),
            "excel_structured_json_count": len(converter.generated_excel_json_outputs),
            "records_json_count": len(converter.generated_records_json_outputs),
            "trace_map_count": len(getattr(converter, "generated_trace_map_outputs", []) or []),
            "prompt_ready_count": len(getattr(converter, "generated_prompt_outputs", []) or []),
            "chromadb_export_file_count": len(converter.generated_chromadb_outputs),
            "update_package_file_count": len(converter.generated_update_package_outputs),
            "mshelp_output_file_count": len(converter.generated_mshelp_outputs),
            "conversion_record_count": len(converter.conversion_index_records),
            "merge_record_count": len(converter.merge_index_records),
            "artifact_count": len(artifacts),
        },
    }

    try:
        llm_hub_meta = converter._maybe_build_llm_delivery_hub(target_folder, artifacts)
        if llm_hub_meta:
            artifacts.append(llm_hub_meta)
            manifest["summary"]["artifact_count"] = len(artifacts)
    except (OSError, RuntimeError, TypeError, ValueError, AttributeError) as exc:
        logging.error(f"failed to build LLM delivery hub: {exc}")

    manifest_path = os.path.join(target_folder, "corpus.json")
    with open(manifest_path, "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)

    converter.corpus_manifest_path = manifest_path
    print(f"\nCorpus manifest generated: {manifest_path}")
    logging.info(f"Corpus manifest generated: {manifest_path}")
    return manifest_path
