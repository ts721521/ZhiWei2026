# -*- coding: utf-8 -*-
"""Update package generation extracted from office_converter.py."""

import csv
import json
import os
import shutil
from datetime import datetime

from converter.file_registry import FileRegistry


UPDATE_INDEX_FIELDS = [
    "seq",
    "change_state",
    "process_status",
    "source_file",
    "source_path",
    "source_md5",
    "source_sha256",
    "renamed_from",
    "rename_match_type",
    "packaged_pdf",
    "packaged_pdf_path",
    "packaged_pdf_md5",
    "note",
]


def _collect_changed_sources(scan_meta):
    changed_sources = []
    for src_path, meta in (scan_meta or {}).items():
        state = str(meta.get("change_state", ""))
        if state in ("added", "modified", "renamed"):
            changed_sources.append(src_path)
    return changed_sources


def _build_result_map(process_results, registry):
    result_map = {}
    for item in process_results or []:
        src = item.get("source_path")
        if not src:
            continue
        result_map[registry.normalize_path(src)] = item
    return result_map


def _safe_compute_md5(compute_md5_fn, path):
    try:
        return compute_md5_fn(path)
    except (OSError, RuntimeError, TypeError, ValueError, AttributeError):
        return ""


def generate_update_package(
    process_results,
    *,
    incremental_context,
    config,
    run_mode,
    resolve_update_package_root_fn,
    compute_md5_fn,
    write_update_package_index_xlsx_fn,
    now_fn=None,
    logger_info=None,
):
    context = incremental_context or {}
    if not context.get("enabled"):
        return None, []
    if not config.get("enable_update_package", True):
        return None, []

    registry = context.get("registry")
    if registry is None:
        registry = FileRegistry(
            context.get("registry_path", ""),
            base_root=config.get("source_folder", ""),
        )

    scan_meta = context.get("scan_meta") or {}
    changed_sources = _collect_changed_sources(scan_meta)
    if not changed_sources:
        return None, []

    result_map = _build_result_map(process_results, registry)

    now_fn = now_fn or datetime.now
    package_root = resolve_update_package_root_fn()
    timestamp = now_fn().strftime("%Y%m%d_%H%M%S")
    package_dir = os.path.join(package_root, f"Update_Package_{timestamp}")
    pdf_dir = os.path.join(package_dir, "PDF")
    os.makedirs(pdf_dir, exist_ok=True)

    records = []
    packaged_pdfs = []
    used_pdf_names = {}

    for idx, src_path in enumerate(sorted(changed_sources), 1):
        norm_key = registry.normalize_path(src_path)
        meta = scan_meta.get(src_path, {})
        result = result_map.get(norm_key, {})

        status = str(result.get("status", "pending"))
        detail = str(result.get("detail", "") or "")
        final_path = str(result.get("final_path", "") or "")
        if status == "pending" and str(meta.get("change_state", "")) == "renamed":
            status = "renamed_detected"
            if not detail:
                detail = "rename detected; no reconvert"

        source_md5 = _safe_compute_md5(compute_md5_fn, src_path)

        packaged_pdf = ""
        packaged_pdf_path = ""
        packaged_pdf_md5 = ""
        if status == "success" and final_path and os.path.exists(final_path):
            base_name = os.path.basename(final_path)
            stem, ext = os.path.splitext(base_name)
            count = used_pdf_names.get(base_name, 0)
            if count > 0:
                base_name = f"{stem}_{count}{ext}"
            used_pdf_names[os.path.basename(final_path)] = count + 1
            packaged_pdf_path = os.path.join(pdf_dir, base_name)
            try:
                shutil.copy2(final_path, packaged_pdf_path)
                packaged_pdf = base_name
                packaged_pdfs.append(packaged_pdf_path)
                packaged_pdf_md5 = _safe_compute_md5(compute_md5_fn, packaged_pdf_path)
            except (OSError, shutil.Error, RuntimeError, TypeError, ValueError) as exc:
                detail = f"{detail}; copy_failed={exc}" if detail else f"copy_failed={exc}"

        records.append(
            {
                "seq": idx,
                "change_state": meta.get("change_state", ""),
                "process_status": status,
                "source_file": os.path.basename(src_path),
                "source_path": os.path.abspath(src_path),
                "source_md5": source_md5,
                "source_sha256": meta.get("source_hash_sha256", ""),
                "renamed_from": meta.get("renamed_from", ""),
                "rename_match_type": meta.get("rename_match_type", ""),
                "packaged_pdf": packaged_pdf,
                "packaged_pdf_path": os.path.abspath(packaged_pdf_path)
                if packaged_pdf_path
                else "",
                "packaged_pdf_md5": packaged_pdf_md5,
                "note": detail,
            }
        )

    index_json = os.path.join(package_dir, "incremental_index.json")
    index_csv = os.path.join(package_dir, "incremental_index.csv")

    with open(index_json, "w", encoding="utf-8") as f:
        json.dump(
            {
                "version": 1,
                "record_count": len(records),
                "records": records,
            },
            f,
            ensure_ascii=False,
            indent=2,
        )

    with open(index_csv, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=UPDATE_INDEX_FIELDS)
        writer.writeheader()
        writer.writerows(records)

    index_xlsx = os.path.join(package_dir, "incremental_index.xlsx")
    xlsx_path = write_update_package_index_xlsx_fn(index_xlsx, records)

    status_counts = {}
    for rec in records:
        key = rec.get("process_status", "unknown")
        status_counts[key] = status_counts.get(key, 0) + 1

    manifest = {
        "version": 1,
        "generated_at": now_fn().isoformat(timespec="seconds"),
        "package_dir": os.path.abspath(package_dir),
        "run_mode": run_mode,
        "target_folder": config.get("target_folder", ""),
        "incremental_registry_path": context.get("registry_path", ""),
        "scan_summary": {
            "scanned_count": context.get("scanned_count", 0),
            "added_count": context.get("added_count", 0),
            "modified_count": context.get("modified_count", 0),
            "renamed_count": context.get("renamed_count", 0),
            "unchanged_count": context.get("unchanged_count", 0),
            "deleted_count": context.get("deleted_count", 0),
        },
        "deleted_paths": context.get("deleted_paths", []),
        "renamed_pairs": context.get("renamed_pairs", []),
        "record_count": len(records),
        "packaged_pdf_count": len(packaged_pdfs),
        "status_counts": status_counts,
        "records": records,
    }

    package_manifest = os.path.join(package_dir, "incremental_manifest.json")
    with open(package_manifest, "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)

    outputs = [package_manifest, index_json, index_csv]
    if xlsx_path:
        outputs.append(xlsx_path)
    outputs.extend(packaged_pdfs)

    if callable(logger_info):
        logger_info(f"[update_package] update package generated: {package_dir}")

    return package_manifest, outputs


def generate_update_package_for_converter(converter, process_results):
    package_manifest, outputs = generate_update_package(
        process_results,
        incremental_context=converter._incremental_context,
        config=converter.config,
        run_mode=converter.run_mode,
        resolve_update_package_root_fn=converter._resolve_update_package_root,
        compute_md5_fn=converter._compute_md5,
        write_update_package_index_xlsx_fn=converter._write_update_package_index_xlsx,
        logger_info=converter._update_package_log_info,
    )
    if package_manifest:
        converter.generated_update_package_outputs = outputs
        converter.update_package_manifest_path = package_manifest
    return package_manifest
