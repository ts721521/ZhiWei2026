# -*- coding: utf-8 -*-
"""Runtime wrapper for PDF->Markdown export extracted from office_converter.py."""


def export_pdf_markdown_for_converter(
    converter,
    pdf_path,
    *,
    source_path_hint=None,
    export_pdf_markdown_fn,
    has_pypdf,
    pdf_reader_cls,
    now_fn,
    log_info_fn,
    log_error_fn,
):
    try:
        md_path, quality_record = export_pdf_markdown_fn(
            pdf_path,
            config=converter.config,
            has_pypdf=has_pypdf,
            pdf_reader_cls=pdf_reader_cls if has_pypdf else None,
            source_path_hint=source_path_hint,
            build_ai_output_path_fn=converter._build_ai_output_path,
            build_ai_output_path_from_source_fn=converter._build_ai_output_path_from_source,
            collect_margin_candidates_fn=converter._collect_margin_candidates,
            clean_markdown_page_lines_fn=converter._clean_markdown_page_lines,
            render_markdown_blocks_fn=converter._render_markdown_blocks,
            compute_md5_fn=converter._compute_md5,
            build_short_id_fn=converter._build_short_id,
            short_id_taken_ids=converter.trace_short_id_taken,
            short_id_prefix=converter.config.get("short_id_prefix", "ZW-"),
            now_fn=now_fn,
        )
        if not md_path:
            return None
        converter.generated_markdown_outputs.append(md_path)
        if quality_record:
            converter.markdown_quality_records.append(quality_record)
        log_info_fn(f"Markdown export success: {md_path}")
        return md_path
    except (OSError, RuntimeError, TypeError, ValueError, AttributeError) as exc:
        log_error_fn(f"Markdown export failed {pdf_path}: {exc}")
        return None
