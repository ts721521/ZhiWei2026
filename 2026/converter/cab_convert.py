# -*- coding: utf-8 -*-
"""MSHelp CAB -> Markdown conversion helper extracted from office_converter.py."""

import os
import shutil


def convert_cab_to_markdown(
    cab_path,
    source_path_for_output,
    *,
    has_bs4,
    temp_sandbox,
    uuid4_hex_fn,
    extract_cab_with_fallback_fn,
    find_files_recursive_fn,
    extract_mshc_payload_fn,
    parse_mshelp_topics_fn,
    build_ai_output_path_from_source_fn,
    normalize_md_line_fn,
    render_html_to_markdown_fn,
    append_mshelp_record_fn,
    now_fn,
    generated_markdown_outputs,
    log_warning_fn,
):
    if not os.path.exists(cab_path):
        raise RuntimeError(f"CAB file not found: {cab_path}")

    if not has_bs4:
        log_warning_fn(
            "beautifulsoup4 is not installed; CAB markdown quality may be limited."
        )

    temp_root = os.path.join(temp_sandbox, f"cab_{uuid4_hex_fn()}")
    extract_dir = os.path.join(temp_root, "cab_extract")
    content_dir = os.path.join(temp_root, "mshc_content")

    try:
        extract_cab_with_fallback_fn(cab_path, extract_dir)

        mshc_files = find_files_recursive_fn(extract_dir, (".mshc",))
        html_root = extract_dir
        if mshc_files:
            extract_mshc_payload_fn(mshc_files[0], content_dir)
            html_root = content_dir

        topics = parse_mshelp_topics_fn(html_root)
        if not topics:
            raise RuntimeError(f"no parseable help topics in CAB: {cab_path}")

        md_path = build_ai_output_path_from_source_fn(
            source_path_for_output, "Markdown", ".md"
        )
        if not md_path:
            raise RuntimeError(f"failed to build markdown path for CAB: {cab_path}")

        lines = [
            f"# {os.path.basename(source_path_for_output)}",
            "",
            f"- source_cab: {os.path.abspath(source_path_for_output)}",
            f"- topic_count: {len(topics)}",
            f"- generated_at: {now_fn().isoformat(timespec='seconds')}",
            "",
            "## 鐩綍",
            "",
        ]
        for idx, topic in enumerate(topics, 1):
            title = normalize_md_line_fn(topic.get("title", "") or topic.get("id", ""))
            lines.append(f"{idx}. {title or 'Untitled'}")
        lines.append("")

        rendered_count = 0
        for idx, topic in enumerate(topics, 1):
            title = normalize_md_line_fn(topic.get("title", "") or topic.get("id", ""))
            html_file = topic.get("file", "")
            if not html_file or not os.path.exists(html_file):
                continue
            body_md = render_html_to_markdown_fn(html_file)
            if not body_md:
                continue
            lines.extend([f"## {idx}. {title or 'Untitled'}", "", body_md, ""])
            rendered_count += 1

        if rendered_count <= 0:
            raise RuntimeError(
                f"CAB topics parsed but no readable content rendered: {cab_path}"
            )

        with open(md_path, "w", encoding="utf-8") as f:
            f.write("\n".join(lines).rstrip() + "\n")

        generated_markdown_outputs.append(md_path)
        append_mshelp_record_fn(source_path_for_output, md_path, rendered_count)
        return md_path, rendered_count
    finally:
        try:
            if os.path.exists(temp_root):
                shutil.rmtree(temp_root, ignore_errors=True)
        except (OSError, TypeError, ValueError):
            pass
