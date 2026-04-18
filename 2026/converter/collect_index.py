# -*- coding: utf-8 -*-
"""Collect-only scan/index workflow extracted from office_converter.py."""

import logging
import os
import shutil
from collections import Counter
from datetime import datetime

from converter.constants import (
    COLLECT_COPY_LAYOUT_FLAT,
    COLLECT_COPY_LAYOUT_PRESERVE_TREE,
    COLLECT_MODE_COPY_AND_INDEX,
)


def _assign_flat_collect_targets(unique_records, duplicate_records, target_root):
    """Rewrite TargetPath to a single folder; disambiguate same basenames."""
    names = [os.path.basename(rec["src"]) for rec in unique_records]
    counts = Counter(names)
    per_base_seq = {}
    for rec in unique_records:
        base = os.path.basename(rec["src"])
        if counts[base] <= 1:
            flat_name = base
        else:
            per_base_seq[base] = per_base_seq.get(base, 0) + 1
            seq = per_base_seq[base]
            stem, ext = os.path.splitext(base)
            flat_name = f"{stem}__{seq}{ext}"
        rec["dst"] = os.path.join(target_root, flat_name)
    src_to_dst = {rec["src"]: rec["dst"] for rec in unique_records}
    for rec in duplicate_records:
        rec["keep_dst"] = src_to_dst.get(rec["keep_src"], rec["keep_dst"])


def collect_office_files_and_build_excel(
    converter,
    *,
    has_openpyxl,
    workbook_cls,
    font_cls,
):
    if not has_openpyxl:
        print("\n[ERROR] openpyxl not found. Cannot generate Excel report.")
        print("Run pip install openpyxl and retry.")
        logging.error("openpyxl missing; collect_only mode cannot continue.")
        return

    configured_roots = converter._get_configured_source_roots()
    scan_skip_seen = set()
    source_roots = []
    for source_root in configured_roots:
        if converter._probe_source_root_access(
            source_root,
            context={"scan_scope": "collect_only"},
            seen_keys=scan_skip_seen,
        ):
            source_roots.append(os.path.abspath(source_root))

    if not source_roots:
        print("[WARN] No source folder(s) to scan. collect_only skipped.")
        return
    target_root = converter.config["target_folder"]
    os.makedirs(target_root, exist_ok=True)
    copy_layout = converter.config.get(
        "collect_copy_layout", COLLECT_COPY_LAYOUT_PRESERVE_TREE
    )
    if copy_layout not in (COLLECT_COPY_LAYOUT_PRESERVE_TREE, COLLECT_COPY_LAYOUT_FLAT):
        copy_layout = COLLECT_COPY_LAYOUT_PRESERVE_TREE

    exts_word = converter.config["allowed_extensions"].get("word", [])
    exts_excel = converter.config["allowed_extensions"].get("excel", [])
    exts_ppt = converter.config["allowed_extensions"].get("powerpoint", [])
    office_exts = set(exts_word + exts_excel + exts_ppt)

    excl_config = converter.config.get("excluded_folders", [])
    excl_names = {
        x.lower()
        for x in excl_config
        if not os.path.isabs(x) and os.sep not in x and "/" not in x
    }
    excl_paths = {
        os.path.abspath(x).lower()
        for x in excl_config
        if os.path.isabs(x) or os.sep in x or "/" in x
    }

    print("\n" + "=" * 60)
    print(" File collection & dedup mode")
    print("=" * 60)
    print(f" Source dir(s) : {len(source_roots)} folder(s)")
    print(f" Target dir : {target_root}")
    print(f" Sub mode   : {converter.get_readable_collect_mode()} ({converter.collect_mode})")
    if converter.collect_mode == COLLECT_MODE_COPY_AND_INDEX:
        print(
            f" Copy layout: {copy_layout} "
            f"(flat=all files under target folder; preserve_tree=keep subfolders)"
        )
    print(f" Filter ext : {office_exts}")
    print("=" * 60)

    all_files = []
    for source_root in source_roots:
        if not os.path.isdir(source_root):
            continue
        for root, dirs, files in os.walk(
            source_root,
            onerror=lambda e, sf=source_root: converter._record_scan_access_skip(
                getattr(e, "filename", sf),
                e,
                context={"scan_scope": "collect_only", "source_root": sf},
                seen_keys=scan_skip_seen,
            ),
        ):
            dirs[:] = [
                d
                for d in dirs
                if d.lower() not in excl_names
                and os.path.abspath(os.path.join(root, d)).lower() not in excl_paths
            ]
            for name in files:
                if name.startswith("~$"):
                    continue
                ext = os.path.splitext(name)[1].lower()
                if ext in office_exts:
                    full_path = os.path.join(root, name)
                    try:
                        size = os.path.getsize(full_path)
                    except OSError:
                        continue
                    all_files.append((full_path, size, ext))

    total = len(all_files)
    print(f"Scanned Office files: {total}")
    logging.info(f"[collect_only] scanned Office files: {total}")

    if total == 0:
        print("[INFO] No Office files found. collect_only finished.")
        return

    size_groups = {}
    for path, size, ext in all_files:
        size_groups.setdefault(size, []).append((path, ext))

    unique_records = []
    duplicate_records = []
    group_id_counter = 1

    for size, files in size_groups.items():
        if not converter.is_running:
            break
        if len(files) == 1:
            src_path, ext = files[0]
            rel = os.path.relpath(src_path, converter._get_source_root_for_path(src_path))
            dst_path = os.path.join(target_root, rel)
            unique_records.append(
                {
                    "group_id": None,
                    "src": src_path,
                    "dst": dst_path,
                    "size": size,
                    "ext": ext,
                }
            )
            continue

        hash_groups = {}
        for src_path, ext in files:
            file_hash = converter._compute_file_hash(src_path)
            hash_groups.setdefault(file_hash, []).append((src_path, ext))

        for _, same_hash_files in hash_groups.items():
            if len(same_hash_files) == 1:
                src_path, ext = same_hash_files[0]
                rel = os.path.relpath(src_path, converter._get_source_root_for_path(src_path))
                dst_path = os.path.join(target_root, rel)
                unique_records.append(
                    {
                        "group_id": None,
                        "src": src_path,
                        "dst": dst_path,
                        "size": size,
                        "ext": ext,
                    }
                )
            else:
                group_id = f"G{group_id_counter}"
                group_id_counter += 1

                keep_src, keep_ext = same_hash_files[0]
                keep_rel = os.path.relpath(keep_src, converter._get_source_root_for_path(keep_src))
                keep_dst = os.path.join(target_root, keep_rel)

                unique_records.append(
                    {
                        "group_id": group_id,
                        "src": keep_src,
                        "dst": keep_dst,
                        "size": size,
                        "ext": keep_ext,
                    }
                )

                for dup_src, dup_ext in same_hash_files[1:]:
                    duplicate_records.append(
                        {
                            "group_id": group_id,
                            "src": dup_src,
                            "size": size,
                            "ext": dup_ext,
                            "keep_src": keep_src,
                            "keep_dst": keep_dst,
                        }
                    )

    print("\nDedup completed:")
    print(f"  Unique files    : {len(unique_records)}")
    print(f"  Duplicate files : {len(duplicate_records)}")
    logging.info(
        f"[collect_only] unique={len(unique_records)}, duplicate={len(duplicate_records)}"
    )

    if (
        converter.collect_mode == COLLECT_MODE_COPY_AND_INDEX
        and copy_layout == COLLECT_COPY_LAYOUT_FLAT
    ):
        _assign_flat_collect_targets(unique_records, duplicate_records, target_root)

    copied_count = 0
    if converter.collect_mode == COLLECT_MODE_COPY_AND_INDEX:
        print("\nCopying unique files to target directory...")
        for idx, rec in enumerate(unique_records, 1):
            if not converter.is_running:
                break
            src = rec["src"]
            dst = rec["dst"]
            dst_dir = os.path.dirname(dst)
            os.makedirs(dst_dir, exist_ok=True)
            try:
                if not os.path.exists(dst):
                    shutil.copy2(src, dst)
                rec["copied"] = True
                copied_count += 1
            except (OSError, RuntimeError, TypeError, ValueError) as exc:
                logging.error(f"[collect_only] copy failed: {src} -> {dst} | {exc}")
                rec["copied"] = False

            if idx % 20 == 0 or idx == len(unique_records):
                print(
                    f"\rProcessed {idx}/{len(unique_records)} unique files...",
                    end="",
                    flush=True,
                )
        print(f"\rCopy finished, copied {copied_count} files.         ")
    else:
        print("\nCurrent mode is [index_only]; skip file copy.")
        for rec in unique_records:
            rec["copied"] = False

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    excel_path = os.path.join(target_root, f"office_index_{timestamp}.xlsx")

    wb = workbook_cls()
    ws_unique = wb.active
    ws_unique.title = "UniqueFiles"
    ws_dup = wb.create_sheet("Duplicates")

    if converter.collect_mode == COLLECT_MODE_COPY_AND_INDEX:
        headers_unique = [
            "No.",
            "GroupID",
            "FileName",
            "Ext",
            "Size(KB)",
            "SourcePath",
            "TargetPath",
        ]
    else:
        headers_unique = [
            "No.",
            "GroupID",
            "FileName",
            "Ext",
            "Size(KB)",
            "SourcePath",
        ]

    ws_unique.append(headers_unique)
    for cell in ws_unique[1]:
        cell.font = font_cls(bold=True)

    for idx, rec in enumerate(unique_records, 1):
        src = rec["src"]
        dst = rec["dst"]
        size_kb = round(rec["size"] / 1024, 2)
        group_id = rec["group_id"] or ""
        file_name = os.path.basename(src)
        ext = rec["ext"]

        if converter.collect_mode == COLLECT_MODE_COPY_AND_INDEX:
            row = [idx, group_id, file_name, ext, size_kb, src, dst]
            ws_unique.append(row)
            dst_cell = ws_unique.cell(row=idx + 1, column=7)
            dst_cell.hyperlink = converter._make_file_hyperlink(dst)
            dst_cell.style = "Hyperlink"
        else:
            row = [idx, group_id, file_name, ext, size_kb, src]
            ws_unique.append(row)
            src_cell = ws_unique.cell(row=idx + 1, column=6)
            src_cell.hyperlink = converter._make_file_hyperlink(src)
            src_cell.style = "Hyperlink"

    if converter.collect_mode == COLLECT_MODE_COPY_AND_INDEX:
        headers_dup = [
            "No.",
            "GroupID",
            "FileName",
            "Ext",
            "Size(KB)",
            "SourcePath",
            "KeptTargetPath",
        ]
    else:
        headers_dup = [
            "No.",
            "GroupID",
            "FileName",
            "Ext",
            "Size(KB)",
            "SourcePath",
            "KeptSourcePath",
        ]

    ws_dup.append(headers_dup)
    for cell in ws_dup[1]:
        cell.font = font_cls(bold=True)

    for idx, rec in enumerate(duplicate_records, 1):
        src = rec["src"]
        keep_src = rec["keep_src"]
        keep_dst = rec["keep_dst"]
        size_kb = round(rec["size"] / 1024, 2)
        group_id = rec["group_id"]
        file_name = os.path.basename(src)
        ext = rec["ext"]

        if converter.collect_mode == COLLECT_MODE_COPY_AND_INDEX:
            row = [idx, group_id, file_name, ext, size_kb, src, keep_dst]
            ws_dup.append(row)
            src_cell = ws_dup.cell(row=idx + 1, column=6)
            src_cell.hyperlink = converter._make_file_hyperlink(src)
            src_cell.style = "Hyperlink"
        else:
            row = [idx, group_id, file_name, ext, size_kb, src, keep_src]
            ws_dup.append(row)
            src_cell = ws_dup.cell(row=idx + 1, column=6)
            src_cell.hyperlink = converter._make_file_hyperlink(src)
            src_cell.style = "Hyperlink"
            keep_cell = ws_dup.cell(row=idx + 1, column=7)
            keep_cell.hyperlink = converter._make_file_hyperlink(keep_src)
            keep_cell.style = "Hyperlink"

    for ws in (ws_unique, ws_dup):
        for col in ws.columns:
            max_length = 0
            col_letter = col[0].column_letter
            for cell in col:
                try:
                    value = str(cell.value) if cell.value is not None else ""
                    max_length = max(max_length, len(value))
                except (TypeError, ValueError, AttributeError):
                    pass
            ws.column_dimensions[col_letter].width = min(max_length + 2, 80)

    wb.save(excel_path)
    converter.collect_index_path = excel_path
    print(f"\nExcel index generated: {excel_path}")
    logging.info(f"[collect_only] Excel index generated: {excel_path}")

    print("\n=== collect_only summary ===")
    print(f"scanned          : {total}")
    print(
        f"unique files     : {len(unique_records)} (copied {copied_count}; valid in copy mode)"
    )
    print(f"duplicate files  : {len(duplicate_records)}")
    print(f"index file       : {excel_path}")
    print("========================\n")
