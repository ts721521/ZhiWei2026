# -*- coding: utf-8 -*-
"""Markdown merge workflow extracted from office_converter.py."""

import logging
import os
from datetime import datetime


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
    for output_name, group in tasks:
        lines = [
            f"# {output_name}",
            "",
            f"- generated_at: {now_fn().isoformat(timespec='seconds')}",
            f"- source_count: {len(group)}",
            "",
            "## Source Map",
            "",
        ]
        for idx, path in enumerate(group, 1):
            lines.append(f"{idx}. {path}")
        lines.extend(["", "## Documents", ""])

        for idx, path in enumerate(group, 1):
            try:
                with open_fn(path, "r", encoding="utf-8", errors="ignore") as f:
                    content = f.read().strip()
            except (OSError, RuntimeError, TypeError, ValueError, UnicodeError) as exc:
                log_warning_fn(f"skip markdown in merge {path}: {exc}")
                continue
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

        out_path = join_fn(converter.merge_output_dir, output_name)
        with open_fn(out_path, "w", encoding="utf-8") as f:
            f.write("\n".join(lines).rstrip() + "\n")
        generated.append(out_path)
        log_info_fn(f"merged markdown generated: {out_path}")

    converter.generated_merge_markdown_outputs = list(generated)
    return generated
