# -*- coding: utf-8 -*-
"""Merge-mode pipeline extracted from office_converter.py."""

import logging
import os

from converter.constants import MERGE_CONVERT_SUBMODE_PDF_TO_MD


def run_merge_mode_pipeline(converter, batch_results):
    pref = converter._get_output_pref()
    merged_outputs = []
    submode = converter._get_merge_convert_submode()

    if submode == MERGE_CONVERT_SUBMODE_PDF_TO_MD:
        pdf_files = converter._scan_merge_candidates_by_ext(".pdf")
        need_markdown = pref["md"] and (pref["independent"] or pref["merged"])
        if need_markdown:
            for path in pdf_files:
                md_path = converter._export_pdf_markdown(path, source_path_hint=path)
                batch_results.append(
                    {
                        "source_path": os.path.abspath(path),
                        "status": "success_non_pdf" if md_path else "failed",
                        "detail": "pdf_to_md" if md_path else "pdf_to_md_failed",
                        "final_path": md_path or "",
                        "elapsed": 0.0,
                    }
                )

        if pref["merged"]:
            if pref["pdf"] and converter.config.get("enable_merge", True):
                merged_outputs.extend(converter.merge_pdfs() or [])
            if pref["md"]:
                md_candidates = [
                    path
                    for path in converter.generated_markdown_outputs
                    if os.path.exists(path)
                ] or converter._scan_merge_candidates_by_ext(".md")
                if not md_candidates:
                    if not converter._confirm_continue_missing_md_merge():
                        raise RuntimeError("Markdown merge canceled by user.")
                else:
                    merged_outputs.extend(converter.merge_markdowns(md_candidates) or [])
        return merged_outputs

    if pref["merged"]:
        if pref["pdf"] and converter.config.get("enable_merge", True):
            merged_outputs.extend(converter.merge_pdfs() or [])
        if pref["md"]:
            md_candidates = converter._scan_merge_candidates_by_ext(".md")
            if not md_candidates:
                if not converter._confirm_continue_missing_md_merge():
                    raise RuntimeError("Markdown merge canceled by user.")
                logging.info("Markdown merge skipped: no .md files found.")
            else:
                merged_outputs.extend(converter.merge_markdowns(md_candidates) or [])
    return merged_outputs
