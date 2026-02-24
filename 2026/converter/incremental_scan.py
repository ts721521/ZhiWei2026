# -*- coding: utf-8 -*-
"""Incremental scan/filter logic extracted from office_converter.py."""

import os

from converter.file_registry import FileRegistry


def apply_incremental_filter(
    files,
    config,
    resolve_registry_path_fn,
    build_source_meta_fn,
    compute_file_hash_fn,
    file_registry_cls=FileRegistry,
    log_info=None,
):
    context = {
        "enabled": False,
        "registry": None,
        "registry_path": "",
        "scan_meta": {},
        "scanned_count": len(files),
        "added_count": 0,
        "modified_count": 0,
        "renamed_count": 0,
        "unchanged_count": 0,
        "deleted_count": 0,
        "deleted_paths": [],
        "renamed_pairs": [],
        "reprocess_renamed": False,
    }

    if not config.get("enable_incremental_mode", False):
        return files, context

    verify_hash = bool(config.get("incremental_verify_hash", False))
    reprocess_renamed = bool(config.get("incremental_reprocess_renamed", False))
    registry_path = resolve_registry_path_fn()
    registry = file_registry_cls(registry_path, base_root=config.get("source_folder", ""))
    registry.load()

    process_files = []
    scan_meta = {}
    added = 0
    modified = 0
    renamed = 0
    unchanged = 0

    for path in files:
        meta = build_source_meta_fn(path, include_hash=verify_hash)
        if not meta:
            continue

        prev = registry.get(path)
        state = "added"
        if isinstance(prev, dict):
            prev_size = int(prev.get("source_size", -1))
            prev_mtime_ns = int(prev.get("source_mtime_ns", -1))
            same_size = prev_size == meta["source_size"]
            same_mtime = prev_mtime_ns == meta["source_mtime_ns"]
            if verify_hash:
                prev_hash = str(prev.get("source_hash_sha256", "") or "")
                curr_hash = meta.get("source_hash_sha256", "")
                same_hash = bool(prev_hash and curr_hash and prev_hash == curr_hash)
            else:
                same_hash = True

            if same_size and same_mtime and same_hash:
                state = "unchanged"
            else:
                state = "modified"

        meta["change_state"] = state
        scan_meta[path] = meta

        if state == "unchanged":
            unchanged += 1
        elif state == "added":
            added += 1
            process_files.append(path)
        else:
            modified += 1
            process_files.append(path)

    current_keys = {registry.normalize_path(p) for p in scan_meta.keys()}
    old_keys = set(registry.keys())
    deleted_keys = sorted(old_keys - current_keys)
    deleted_set = set(deleted_keys)

    # Rename detection: match current "added" files to previous deleted entries.
    renamed_pairs = []
    deleted_entry_map = {}
    if isinstance(registry.entries, dict):
        for key in deleted_keys:
            entry = registry.entries.get(key)
            if isinstance(entry, dict):
                deleted_entry_map[key] = entry

    added_paths = [
        p for p, m in scan_meta.items() if str(m.get("change_state", "")) == "added"
    ]

    for src_path in sorted(added_paths):
        meta = scan_meta.get(src_path, {})
        ext = str(meta.get("ext", "") or "")
        size = int(meta.get("source_size", -1))
        mtime_ns = int(meta.get("source_mtime_ns", -1))

        candidates = []
        for old_key in list(deleted_set):
            old_entry = deleted_entry_map.get(old_key) or {}
            if str(old_entry.get("ext", "")) != ext:
                continue
            if int(old_entry.get("source_size", -2)) != size:
                continue
            candidates.append((old_key, old_entry))

        if not candidates:
            continue

        curr_hash = str(meta.get("source_hash_sha256", "") or "")
        if not curr_hash:
            try:
                curr_hash = compute_file_hash_fn(src_path)
                meta["source_hash_sha256"] = curr_hash
            except Exception:
                curr_hash = ""

        matched = None
        if curr_hash:
            for old_key, old_entry in candidates:
                old_hash = str(old_entry.get("source_hash_sha256", "") or "")
                if old_hash and old_hash == curr_hash:
                    matched = (old_key, old_entry, "hash")
                    break

        if matched is None:
            for old_key, old_entry in candidates:
                old_mtime_ns = int(old_entry.get("source_mtime_ns", -1))
                if old_mtime_ns == mtime_ns and old_mtime_ns >= 0:
                    matched = (old_key, old_entry, "mtime")
                    break

        if matched is None and len(candidates) == 1:
            old_key, old_entry = candidates[0]
            matched = (old_key, old_entry, "ext_size_unique")

        if matched is None:
            continue

        old_key, old_entry, match_type = matched
        old_path = str(old_entry.get("source_path", "") or old_key)

        meta["change_state"] = "renamed"
        meta["renamed_from"] = old_path
        meta["renamed_from_key"] = old_key
        meta["rename_match_type"] = match_type

        deleted_set.discard(old_key)
        deleted_entry_map.pop(old_key, None)

        renamed += 1
        added = max(0, added - 1)

        renamed_pairs.append(
            {
                "from_path": old_path,
                "to_path": os.path.abspath(src_path),
                "match_type": match_type,
            }
        )

        if not reprocess_renamed and src_path in process_files:
            process_files.remove(src_path)

    deleted_keys = sorted(deleted_set)

    context = {
        "enabled": True,
        "registry": registry,
        "registry_path": registry_path,
        "scan_meta": scan_meta,
        "scanned_count": len(scan_meta),
        "added_count": added,
        "modified_count": modified,
        "renamed_count": renamed,
        "unchanged_count": unchanged,
        "deleted_count": len(deleted_keys),
        "deleted_paths": deleted_keys,
        "renamed_pairs": renamed_pairs,
        "reprocess_renamed": reprocess_renamed,
    }

    if callable(log_info):
        log_info(
            "[incremental] 鎵弿瀹屾垚: scanned=%s added=%s modified=%s renamed=%s unchanged=%s deleted=%s",
            context["scanned_count"],
            added,
            modified,
            renamed,
            unchanged,
            context["deleted_count"],
        )
    return process_files, context
