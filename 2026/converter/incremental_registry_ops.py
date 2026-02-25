# -*- coding: utf-8 -*-
"""Incremental registry helpers extracted from office_converter.py."""

import os
from datetime import datetime


def build_source_meta(path, include_hash=False, compute_file_hash_fn=None, log_warning=None):
    abs_path = os.path.abspath(path)
    try:
        stat = os.stat(abs_path)
    except OSError:
        return None

    source_hash = ""
    if include_hash and callable(compute_file_hash_fn):
        try:
            source_hash = compute_file_hash_fn(abs_path)
        except (OSError, RuntimeError, TypeError, ValueError, AttributeError) as e:
            if callable(log_warning):
                log_warning(f"[incremental] failed to compute source hash: {abs_path} | {e}")

    return {
        "source_path": abs_path,
        "ext": os.path.splitext(abs_path)[1].lower(),
        "source_size": int(stat.st_size),
        "source_mtime_ns": int(
            getattr(stat, "st_mtime_ns", int(stat.st_mtime * 1_000_000_000))
        ),
        "source_mtime": datetime.fromtimestamp(stat.st_mtime).isoformat(timespec="seconds"),
        "source_hash_sha256": source_hash,
    }


def flush_incremental_registry(
    context,
    process_results,
    run_mode,
    compute_md5_fn,
    log_info=None,
):
    context = context or {}
    if not context.get("enabled"):
        return

    registry = context.get("registry")
    if not registry:
        return

    result_map = {}
    for item in process_results or []:
        src = item.get("source_path")
        if not src:
            continue
        result_map[registry.normalize_path(src)] = item

    now_iso = datetime.now().isoformat(timespec="seconds")
    new_entries = {}

    for src_path, meta in (context.get("scan_meta") or {}).items():
        key = registry.normalize_path(src_path)
        prev = registry.entries.get(key, {}) if isinstance(registry.entries, dict) else {}
        rename_from_key = str(meta.get("renamed_from_key", "") or "")
        if (not prev) and rename_from_key and isinstance(registry.entries, dict):
            prev = registry.entries.get(rename_from_key, {})
        entry = dict(prev) if isinstance(prev, dict) else {}

        entry.update(
            {
                "source_path": os.path.abspath(src_path),
                "ext": meta.get("ext", ""),
                "source_size": meta.get("source_size", 0),
                "source_mtime": meta.get("source_mtime", ""),
                "source_mtime_ns": meta.get("source_mtime_ns", 0),
                "source_hash_sha256": meta.get("source_hash_sha256", ""),
                "change_state": meta.get("change_state", ""),
                "renamed_from": meta.get("renamed_from", ""),
                "rename_match_type": meta.get("rename_match_type", ""),
                "last_seen_at": now_iso,
                "last_run_mode": run_mode,
            }
        )

        result = result_map.get(key)
        if result:
            entry["last_status"] = result.get("status", "")
            entry["last_error"] = result.get("error", "")
            entry["last_processed_at"] = now_iso
            final_path = result.get("final_path", "")
            if final_path and os.path.exists(final_path):
                entry["last_output_pdf"] = os.path.abspath(final_path)
                try:
                    entry["last_output_pdf_md5"] = compute_md5_fn(final_path)
                except (OSError, RuntimeError, TypeError, ValueError, AttributeError):
                    pass
        elif meta.get("change_state") == "unchanged":
            entry["last_status"] = entry.get("last_status") or "unchanged"
        elif meta.get("change_state") == "renamed":
            entry["last_status"] = "renamed_detected"
            entry["last_processed_at"] = now_iso

        new_entries[key] = entry

    registry.entries = new_entries
    run_summary = {
        "scanned_count": context.get("scanned_count", 0),
        "added_count": context.get("added_count", 0),
        "modified_count": context.get("modified_count", 0),
        "renamed_count": context.get("renamed_count", 0),
        "unchanged_count": context.get("unchanged_count", 0),
        "deleted_count": context.get("deleted_count", 0),
        "processed_result_count": len(result_map),
    }
    registry.save(run_summary=run_summary)
    if callable(log_info):
        log_info(f"[incremental] registry updated: {context.get('registry_path', '')}")
