# -*- coding: utf-8 -*-
"""Failure-stage inference helpers extracted from office_converter.py."""

import os
import re


def sanitize_failure_log_stem(value):
    stem = re.sub(r"[^\w\-.]+", "_", str(value or "").strip())
    stem = stem.strip("._")
    return stem or "failed_file"


def get_failure_output_expectation(run_mode, config, compute_convert_output_plan):
    try:
        plan = compute_convert_output_plan(run_mode, config)
    except Exception:
        return {"need_final_pdf": None, "need_markdown": None}
    return {
        "need_final_pdf": bool(plan.get("need_final_pdf")),
        "need_markdown": bool(plan.get("need_markdown")),
    }


def infer_failure_stage(
    source_path,
    raw_error="",
    context=None,
    cab_extensions=None,
    expected_outputs_getter=None,
):
    ctx = context or {}
    phase = str(ctx.get("phase", "")).strip().lower()
    if phase == "scan":
        scope = str(ctx.get("scan_scope", "")).strip().lower()
        return f"scan_{scope}" if scope else "scan_access"

    err = str(raw_error or "").lower()
    if "markdown export failed" in err or (
        "markdown" in err and ("failed" in err or "error" in err)
    ):
        return "markdown_export"

    ext = os.path.splitext(str(source_path or ""))[1].lower()
    cab_exts = {str(e).lower() for e in (cab_extensions or [])}
    if ext in cab_exts:
        return "cab_to_markdown"
    if ext == ".pdf":
        expected = {}
        if callable(expected_outputs_getter):
            try:
                expected = expected_outputs_getter() or {}
            except Exception:
                expected = {}
        if expected.get("need_markdown") and not expected.get("need_final_pdf"):
            return "pdf_to_markdown"
        return "pdf_pipeline"
    return "office_to_pdf"
