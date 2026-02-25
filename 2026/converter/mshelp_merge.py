# -*- coding: utf-8 -*-
"""MSHelp merged-markdown export helper extracted from office_converter.py."""

import os


def merge_mshelp_markdowns(
    mshelp_records,
    config,
    *,
    generated_outputs=None,
    export_markdown_to_docx_fn=None,
    export_markdown_to_pdf_fn=None,
    now_fn=None,
    log_info=None,
    log_warning=None,
):
    if not mshelp_records:
        return []
    if not bool(config.get("enable_mshelp_merge_output", True)):
        return []
    if now_fn is None:
        from datetime import datetime as _dt  # noqa: PLC0415

        now_fn = _dt.now

    target_root = config.get("target_folder", "")
    if not target_root:
        return []
    out_dir = os.path.join(target_root, "_AI", "MSHelp", "Merged")
    os.makedirs(out_dir, exist_ok=True)

    try:
        max_size_mb = int(config.get("max_merge_size_mb", 80) or 80)
    except (TypeError, ValueError):
        max_size_mb = 80
    max_size_bytes = max(1, max_size_mb) * 1024 * 1024

    valid = []
    for rec in mshelp_records:
        mdp = rec.get("markdown_path", "")
        if not mdp or not os.path.exists(mdp):
            continue
        try:
            with open(mdp, "r", encoding="utf-8", errors="ignore") as f:
                content = f.read()
        except (OSError, RuntimeError, TypeError, ValueError, UnicodeError):
            continue
        item = dict(rec)
        item["_content"] = content
        item["_bytes"] = len(content.encode("utf-8"))
        valid.append(item)

    if not valid:
        return []

    chunks = []
    current = []
    current_bytes = 0
    for rec in valid:
        rec_bytes = int(rec.get("_bytes", 0) or 0)
        if current and (current_bytes + rec_bytes > max_size_bytes):
            chunks.append(current)
            current = []
            current_bytes = 0
        current.append(rec)
        current_bytes += rec_bytes
    if current:
        chunks.append(current)

    ts = now_fn().strftime("%Y%m%d_%H%M%S")
    outputs = []
    export_docx = bool(config.get("enable_mshelp_output_docx", False))
    export_pdf = bool(config.get("enable_mshelp_output_pdf", False))

    for idx, chunk in enumerate(chunks, 1):
        md_path = os.path.join(out_dir, f"MSHelp_API_Merged_{ts}_{idx:03d}.md")
        lines = [
            f"# MSHelp API Merged Package {idx}/{len(chunks)}",
            "",
            f"- generated_at: {now_fn().isoformat(timespec='seconds')}",
            f"- source_root: {config.get('source_folder', '')}",
            f"- document_count: {len(chunk)}",
            "",
            "## Source Map",
            "",
            "| No. | CAB Source | MSHelpViewer Dir | Markdown Path | Topic Count |",
            "| --- | --- | --- | --- | ---: |",
        ]
        for j, rec in enumerate(chunk, 1):
            lines.append(
                "| {0} | {1} | {2} | {3} | {4} |".format(
                    j,
                    str(rec.get("source_cab", "")).replace("|", "\\|"),
                    str(rec.get("mshelpviewer_dir", "")).replace("|", "\\|"),
                    str(rec.get("markdown_path", "")).replace("|", "\\|"),
                    int(rec.get("topic_count", 0) or 0),
                )
            )
        lines.append("")
        lines.append("## Documents")
        lines.append("")

        for j, rec in enumerate(chunk, 1):
            title = os.path.basename(str(rec.get("source_cab", "") or f"doc_{j}"))
            lines.extend(
                [
                    f"### [{j}] {title}",
                    "",
                    f"- source_cab: {rec.get('source_cab', '')}",
                    f"- source_markdown: {rec.get('markdown_path', '')}",
                    "",
                    "---",
                    "",
                    rec.get("_content", ""),
                    "",
                ]
            )

        with open(md_path, "w", encoding="utf-8") as f:
            f.write("\n".join(lines).rstrip() + "\n")
        outputs.append(md_path)
        if isinstance(generated_outputs, list):
            generated_outputs.append(md_path)
        if callable(log_info):
            log_info(f"MSHelp merged markdown generated: {md_path}")

        if export_docx:
            docx_path = os.path.splitext(md_path)[0] + ".docx"
            try:
                if callable(export_markdown_to_docx_fn):
                    export_markdown_to_docx_fn(md_path, docx_path)
                outputs.append(docx_path)
                if isinstance(generated_outputs, list):
                    generated_outputs.append(docx_path)
                if callable(log_info):
                    log_info(f"MSHelp merged DOCX generated: {docx_path}")
            except (OSError, RuntimeError, TypeError, ValueError, AttributeError) as e:
                if callable(log_warning):
                    log_warning(f"MSHelp merged DOCX skipped: {e}")

        if export_pdf:
            pdf_path = os.path.splitext(md_path)[0] + ".pdf"
            try:
                if callable(export_markdown_to_pdf_fn):
                    export_markdown_to_pdf_fn(md_path, pdf_path)
                outputs.append(pdf_path)
                if isinstance(generated_outputs, list):
                    generated_outputs.append(pdf_path)
                if callable(log_info):
                    log_info(f"MSHelp merged PDF generated: {pdf_path}")
            except (OSError, RuntimeError, TypeError, ValueError, AttributeError) as e:
                if callable(log_warning):
                    log_warning(f"MSHelp merged PDF skipped: {e}")

    return outputs


def merge_mshelp_markdowns_for_converter(
    converter,
    *,
    now_fn,
    perf_counter_fn,
    log_info,
    log_warning,
):
    merge_start = perf_counter_fn()
    outputs = merge_mshelp_markdowns(
        converter.mshelp_records,
        converter.config,
        generated_outputs=converter.generated_mshelp_outputs,
        export_markdown_to_docx_fn=converter._export_markdown_to_docx,
        export_markdown_to_pdf_fn=converter._export_markdown_to_pdf,
        now_fn=now_fn,
        log_info=log_info,
        log_warning=log_warning,
    )
    if outputs:
        converter._add_perf_seconds("mshelp_merge_seconds", perf_counter_fn() - merge_start)
    return outputs
