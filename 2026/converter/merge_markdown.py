# -*- coding: utf-8 -*-
"""Markdown merge workflow extracted from office_converter.py."""

import logging
import os
import json
from datetime import datetime

from converter.markdown_image_map import remap_markdown_images_for_merge

def _safe_positive_int(value, fallback):
    try:
        parsed = int(value)
        return parsed if parsed > 0 else fallback
    except (TypeError, ValueError):
        return fallback


def _resolve_markdown_max_size_bytes(config, default_mb=80):
    cfg = config or {}
    mb = _safe_positive_int(
        cfg.get("markdown_max_size_mb", cfg.get("max_merge_size_mb", default_mb)),
        default_mb,
    )
    return mb * 1024 * 1024


def _split_group_by_size(paths, max_size_bytes):
    if not paths:
        return []
    groups = []
    current = []
    current_size = 0
    for path in paths:
        try:
            size = os.path.getsize(path)
        except (OSError, RuntimeError, TypeError, ValueError):
            size = 0
        estimated = size + 16384
        if current and (current_size + estimated > max_size_bytes):
            groups.append(current)
            current = [path]
            current_size = estimated
        else:
            current.append(path)
            current_size += estimated
    if current:
        groups.append(current)
    return groups


def merge_markdowns(
    converter,
    candidates=None,
    *,
    now_fn=datetime.now,
    open_fn=open,
    join_fn=os.path.join,
    abspath_fn=os.path.abspath,
    basename_fn=os.path.basename,
    log_warning_fn=logging.warning,
    log_info_fn=logging.info,
):
    md_files = list(candidates or [])
    if not md_files:
        md_files = converter._scan_merge_candidates_by_ext(".md")
    if not md_files:
        return []

    tasks = converter._build_markdown_merge_tasks(md_files)
    if not tasks:
        return []

    generated = []
    generated_image_manifests = []
    max_size_bytes = _resolve_markdown_max_size_bytes(getattr(converter, "config", {}))
    merge_index_records = list(getattr(converter, "merge_index_records", []) or [])
    enable_md_image_manifest = bool(
        getattr(converter, "config", {}).get("enable_markdown_image_manifest", True)
    )
    for output_name, group in tasks:
        groups = _split_group_by_size(group, max_size_bytes) if max_size_bytes > 0 else [group]
        for part_idx, sub_group in enumerate(groups, 1):
            part_name = output_name
            if len(groups) > 1:
                part_name = output_name[:-3] + f"_part_{part_idx:03d}.md"

            lines = [
                f"# {part_name}",
                "",
                f"- generated_at: {now_fn().isoformat(timespec='seconds')}",
                f"- source_count: {len(sub_group)}",
                "",
                "## Source Map",
                "",
            ]
            for idx, path in enumerate(sub_group, 1):
                lines.append(f"{idx}. {path}")
            lines.extend(["", "## Documents", ""])

            image_manifest_records = []
            for idx, path in enumerate(sub_group, 1):
                try:
                    with open_fn(path, "r", encoding="utf-8", errors="ignore") as f:
                        content = f.read().strip()
                except (OSError, RuntimeError, TypeError, ValueError, UnicodeError) as exc:
                    log_warning_fn(f"skip markdown in merge {path}: {exc}")
                    continue

                if enable_md_image_manifest:
                    try:
                        content, image_records = remap_markdown_images_for_merge(
                            markdown_text=content,
                            source_markdown_path=path,
                            merge_output_dir=converter.merge_output_dir,
                            merged_markdown_name=part_name,
                            doc_index=idx,
                            merge_index_records=merge_index_records,
                            open_fn=open_fn,
                            exists_fn=os.path.exists,
                            makedirs_fn=os.makedirs,
                        )
                        if image_records:
                            image_manifest_records.extend(image_records)
                    except (
                        OSError,
                        RuntimeError,
                        TypeError,
                        ValueError,
                        UnicodeError,
                        AttributeError,
                    ) as exc:
                        log_warning_fn(f"image remap skipped for {path}: {exc}")

                lines.extend(
                    [
                        f"### [{idx}] {basename_fn(path)}",
                        "",
                        f"- source_markdown: {abspath_fn(path)}",
                        "",
                        "---",
                        "",
                        content,
                        "",
                    ]
                )

            out_path = join_fn(converter.merge_output_dir, part_name)
            with open_fn(out_path, "w", encoding="utf-8") as f:
                f.write("\n".join(lines).rstrip() + "\n")
            generated.append(out_path)
            log_info_fn(f"merged markdown generated: {out_path}")

            if enable_md_image_manifest and image_manifest_records:
                manifest_path = (
                    out_path[:-3] + ".image_manifest.json"
                    if out_path.lower().endswith(".md")
                    else out_path + ".image_manifest.json"
                )
                with open_fn(manifest_path, "w", encoding="utf-8") as f:
                    json.dump(
                        {
                            "version": 1,
                            "generated_at": now_fn().isoformat(timespec="seconds"),
                            "merged_markdown_name": os.path.basename(out_path),
                            "merged_markdown_path": out_path,
                            "record_count": len(image_manifest_records),
                            "records": image_manifest_records,
                        },
                        f,
                        ensure_ascii=False,
                        indent=2,
                    )
                generated_image_manifests.append(manifest_path)
                log_info_fn(f"merged markdown image manifest generated: {manifest_path}")

    converter.generated_merge_markdown_outputs = list(generated)
    converter.generated_markdown_manifest_outputs = list(generated_image_manifests)
    return generated
