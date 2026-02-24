# -*- coding: utf-8 -*-
"""Naming and extension-bucket helpers extracted from office_converter.py."""

import os
import re
from datetime import datetime


def format_merge_filename(pattern, category="All", idx=1, now=None):
    """
    Format merge output filename from pattern.
    Placeholders: {category}, {timestamp}, {date}, {time}, {idx}
    """
    if now is None:
        now = datetime.now()
    timestamp = now.strftime("%Y%m%d_%H%M%S")
    date_part = now.strftime("%Y%m%d")
    time_part = now.strftime("%H%M%S")
    name = (
        pattern.replace("{category}", str(category))
        .replace("{timestamp}", timestamp)
        .replace("{date}", date_part)
        .replace("{time}", time_part)
        .replace("{idx}", str(idx))
    )
    # Sanitize: only allow safe filename chars (alphanumeric, dash, underscore, dot)
    safe = re.sub(r"[^\w\-.]", "_", name)
    safe = re.sub(r"_+", "_", safe).strip("_")
    if not safe:
        safe = f"Merged_{category}_{timestamp}_{idx}"
    if not safe.lower().endswith(".pdf"):
        safe = f"{safe}.pdf"
    return safe


def ext_bucket(path, allowed_extensions):
    ext = os.path.splitext(path)[1].lower()
    if ext in [e.lower() for e in allowed_extensions.get("word", [])]:
        return "word"
    if ext in [e.lower() for e in allowed_extensions.get("excel", [])]:
        return "excel"
    if ext in [e.lower() for e in allowed_extensions.get("powerpoint", [])]:
        return "powerpoint"
    if ext == ".pdf":
        return "pdf"
    return ext or "unknown"
