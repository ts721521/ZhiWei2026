# -*- coding: utf-8 -*-
"""PDF merge task builder extracted from office_converter.py."""

import os


def get_merge_tasks(
    *,
    run_mode,
    merge_source,
    target_folder,
    get_source_roots_fn,
    failed_dir,
    merge_output_dir,
    merge_mode,
    merge_mode_all_in_one,
    mode_merge_only,
    merge_filename_pattern,
    max_merge_size_mb,
    now_fn,
    format_merge_filename_fn,
    print_fn=print,
    is_dir_fn=os.path.isdir,
    walk_fn=os.walk,
    abspath_fn=os.path.abspath,
    basename_fn=os.path.basename,
    getsize_fn=os.path.getsize,
):
    scan_source_type = "target"
    if run_mode == mode_merge_only:
        scan_source_type = merge_source or "source"

    if scan_source_type == "source":
        scan_roots = get_source_roots_fn()
        print_fn(f"  [merge_only/source] scanning {len(scan_roots)} source folder(s)")
    else:
        scan_roots = [target_folder]
        print_fn(f"  [merge scan] scanning target: {scan_roots[0]}")

    all_pdfs = []
    exclude_abs_paths = set(map(abspath_fn, [failed_dir, merge_output_dir]))
    if scan_source_type == "source":
        exclude_abs_paths.add(abspath_fn(target_folder))

    for scan_folder in scan_roots:
        if not scan_folder or not is_dir_fn(scan_folder):
            continue
        for root, dirs, files in walk_fn(scan_folder):
            dirs[:] = [
                d
                for d in dirs
                if abspath_fn(os.path.join(root, d)) not in exclude_abs_paths
            ]

            if abspath_fn(root) in exclude_abs_paths:
                continue
            for f in files:
                if f.lower().endswith(".pdf"):
                    all_pdfs.append(os.path.join(root, f))

    if not all_pdfs:
        print_fn("[INFO] no PDF files found for merge.")
        return []

    all_pdfs.sort()

    merge_tasks = []
    now = now_fn()
    pattern = (merge_filename_pattern or "Merged_{category}_{timestamp}_{idx}").strip()
    if not pattern:
        pattern = "Merged_{category}_{timestamp}_{idx}"

    if merge_mode == merge_mode_all_in_one:
        output_name = format_merge_filename_fn(pattern, category="All", idx=1, now=now)
        merge_tasks.append((output_name, all_pdfs))
        return merge_tasks

    categories = {
        "Price Documents": "Price_",
        "Word Documents": "Word_",
        "Excel Sheets": "Excel_",
        "PPT Slides": "PPT_",
        "Original PDF": "PDF_",
    }
    max_size_bytes = max_merge_size_mb * 1024 * 1024

    for _cat_name, prefix in categories.items():
        current_cat_files = [p for p in all_pdfs if basename_fn(p).startswith(prefix)]
        if not current_cat_files:
            continue
        current_cat_files.sort()

        groups = []
        current_group = []
        current_size = 0
        for pdf_path in current_cat_files:
            try:
                f_size = getsize_fn(pdf_path)
            except (OSError, RuntimeError, TypeError, ValueError):
                continue

            if f_size > max_size_bytes:
                if current_group:
                    groups.append(current_group)
                    current_group = []
                    current_size = 0
                groups.append([pdf_path])
                continue

            if current_size + f_size > max_size_bytes:
                groups.append(current_group)
                current_group = [pdf_path]
                current_size = f_size
            else:
                current_group.append(pdf_path)
                current_size += f_size
        if current_group:
            groups.append(current_group)

        cat_label = prefix.rstrip("_")
        for idx, group in enumerate(groups, 1):
            output_filename = format_merge_filename_fn(
                pattern, category=cat_label, idx=idx, now=now
            )
            merge_tasks.append((output_filename, group))

    return merge_tasks
