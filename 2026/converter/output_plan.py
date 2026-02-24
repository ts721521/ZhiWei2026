# -*- coding: utf-8 -*-
"""Output planning helpers extracted from office_converter.py."""

from converter.constants import MODE_CONVERT_THEN_MERGE


def compute_convert_output_plan(run_mode, cfg):
    cfg = cfg or {}
    want_pdf = bool(cfg.get("output_enable_pdf", True))
    want_md = bool(cfg.get("output_enable_md", True))
    want_merged = bool(cfg.get("output_enable_merged", True))
    want_independent = bool(cfg.get("output_enable_independent", False))
    enable_merge = bool(cfg.get("enable_merge", True))

    merge_in_convert_phase = (
        run_mode == MODE_CONVERT_THEN_MERGE and want_merged and enable_merge
    )
    need_pdf_for_merge = merge_in_convert_phase and want_pdf
    need_pdf_independent = want_independent and want_pdf
    need_markdown_for_merge = merge_in_convert_phase and want_md
    need_markdown_independent = want_independent and want_md

    return {
        "want_pdf": want_pdf,
        "want_md": want_md,
        "want_merged": want_merged,
        "want_independent": want_independent,
        "need_pdf_for_merge": need_pdf_for_merge,
        "need_pdf_independent": need_pdf_independent,
        "need_final_pdf": need_pdf_for_merge or need_pdf_independent,
        "need_markdown_for_merge": need_markdown_for_merge,
        "need_markdown_independent": need_markdown_independent,
        "need_markdown": need_markdown_for_merge or need_markdown_independent,
    }
