# -*- coding: utf-8 -*-
"""Runtime preference helpers extracted from office_converter.py."""

from converter.constants import (
    MERGE_CONVERT_SUBMODE_MERGE_ONLY,
    MERGE_CONVERT_SUBMODE_PDF_TO_MD,
)


def get_output_pref(cfg):
    cfg = cfg or {}
    return {
        "pdf": bool(cfg.get("output_enable_pdf", True)),
        "md": bool(cfg.get("output_enable_md", True)),
        "merged": bool(cfg.get("output_enable_merged", True)),
        "independent": bool(cfg.get("output_enable_independent", False)),
    }


def get_merge_convert_submode(cfg):
    cfg = cfg or {}
    raw = str(
        cfg.get("merge_convert_submode", MERGE_CONVERT_SUBMODE_MERGE_ONLY)
        or MERGE_CONVERT_SUBMODE_MERGE_ONLY
    ).strip()
    if raw not in (
        MERGE_CONVERT_SUBMODE_MERGE_ONLY,
        MERGE_CONVERT_SUBMODE_PDF_TO_MD,
    ):
        return MERGE_CONVERT_SUBMODE_MERGE_ONLY
    return raw
