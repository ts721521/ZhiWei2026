# -*- coding: utf-8 -*-
"""PDF -> Markdown export helper extracted from office_converter.py."""

import os

from converter.traceability import apply_short_id_prefix


def export_pdf_markdown(
    pdf_path,
    *,
    config,
    has_pypdf,
    pdf_reader_cls=None,
    source_path_hint=None,
    build_ai_output_path_fn=None,
    build_ai_output_path_from_source_fn=None,
    collect_margin_candidates_fn=None,
    clean_markdown_page_lines_fn=None,
    render_markdown_blocks_fn=None,
    compute_md5_fn=None,
    build_short_id_fn=None,
    short_id_taken_ids=None,
    short_id_prefix="ZW-",
    now_fn=None,
):
    if not config.get("output_enable_md", config.get("enable_markdown", True)):
        return None, None
    if not has_pypdf or pdf_reader_cls is None:
        return None, None
    if not pdf_path or not os.path.exists(pdf_path):
        return None, None

    if source_path_hint:
        if build_ai_output_path_from_source_fn is None:
            return None, None
        md_path = build_ai_output_path_from_source_fn(source_path_hint, "Markdown", ".md")
    else:
        if build_ai_output_path_fn is None:
            return None, None
        md_path = build_ai_output_path_fn(pdf_path, "Markdown", ".md")
    if not md_path:
        return None, None

    if now_fn is None:
        raise RuntimeError("now_fn is required")
    if collect_margin_candidates_fn is None:
        collect_margin_candidates_fn = lambda _pages: (set(), set())
    if clean_markdown_page_lines_fn is None:
        clean_markdown_page_lines_fn = lambda text, _h, _f: ([text], {})
    if render_markdown_blocks_fn is None:
        render_markdown_blocks_fn = lambda lines, structured_headings=True: (
            [str(x) for x in lines],
            0,
        )

    try:
        reader = pdf_reader_cls(pdf_path)
        page_count = len(reader.pages)
        raw_page_texts = []
        for page in reader.pages:
            text = ""
            try:
                text = page.extract_text() or ""
            except (AttributeError, RuntimeError, TypeError, ValueError):
                text = ""
            raw_page_texts.append(text)
    except (OSError, RuntimeError, TypeError, ValueError, AttributeError, Exception):
        return None, None

    strip_header_footer = bool(config.get("markdown_strip_header_footer", True))
    structured_headings = bool(config.get("markdown_structured_headings", True))
    enable_traceability = bool(config.get("enable_traceability_anchor_and_map", True))
    header_keys, footer_keys = (
        collect_margin_candidates_fn(raw_page_texts) if strip_header_footer else (set(), set())
    )

    source_file_for_meta = os.path.abspath(source_path_hint or pdf_path)
    source_short_id = ""
    source_md5 = ""
    if enable_traceability and compute_md5_fn is not None:
        try:
            source_md5 = compute_md5_fn(source_file_for_meta)
        except (OSError, RuntimeError, TypeError, ValueError, AttributeError):
            source_md5 = ""
    if enable_traceability and source_md5:
        try:
            if build_short_id_fn is not None:
                taken = short_id_taken_ids if short_id_taken_ids is not None else set()
                source_short_id = build_short_id_fn(source_md5, taken)
            else:
                source_short_id = source_md5[:8].upper()
        except (RuntimeError, TypeError, ValueError, AttributeError):
            source_short_id = source_md5[:8].upper()
        source_short_id = apply_short_id_prefix(source_short_id, short_id_prefix)

    lines = [
        "---",
        f"source_file: {source_file_for_meta}",
        f"source_short_id: {source_short_id}",
        f"source_md5: {source_md5}",
        "---",
        "",
        f"# {os.path.basename(pdf_path)}",
        "",
        f"- source_pdf: {os.path.abspath(pdf_path)}",
        f"- page_count: {page_count}",
        f"- generated_at: {now_fn().isoformat(timespec='seconds')}",
        "",
    ]

    removed_header_total = 0
    removed_footer_total = 0
    removed_page_no_total = 0
    heading_total = 0
    non_empty_pages = 0

    for idx, raw_text in enumerate(raw_page_texts, 1):
        page_lines, page_stats = clean_markdown_page_lines_fn(raw_text, header_keys, footer_keys)
        blocks, heading_count = render_markdown_blocks_fn(
            page_lines, structured_headings=structured_headings
        )
        page_body = "\n\n".join(blocks).strip()
        if page_body:
            non_empty_pages += 1
        else:
            page_body = "(empty)"

        removed_header_total += page_stats.get("removed_header_lines", 0)
        removed_footer_total += page_stats.get("removed_footer_lines", 0)
        removed_page_no_total += page_stats.get("removed_page_number_lines", 0)
        heading_total += heading_count

        lines.extend([f"## Page {idx}", "", page_body, ""])

    with open(md_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines).rstrip() + "\n")

    quality_record = {
        "source_pdf": source_file_for_meta,
        "markdown_path": os.path.abspath(md_path),
        "source_short_id": source_short_id,
        "source_md5": source_md5,
        "source_filename": os.path.basename(source_file_for_meta),
        "source_abspath": source_file_for_meta,
        "page_count": page_count,
        "non_empty_page_count": non_empty_pages,
        "removed_header_lines": removed_header_total,
        "removed_footer_lines": removed_footer_total,
        "removed_page_number_lines": removed_page_no_total,
        "heading_count": heading_total,
        "header_candidate_count": len(header_keys),
        "footer_candidate_count": len(footer_keys),
        "strip_header_footer": strip_header_footer,
        "structured_headings": structured_headings,
    }
    return md_path, quality_record
