# -*- coding: utf-8 -*-
"""General text/file helpers extracted from office_converter.py."""

import os
import re
import zipfile


def find_files_recursive(root_dir, exts):
    results = []
    if not root_dir or not os.path.isdir(root_dir):
        return results
    ext_set = tuple(str(e).lower() for e in (exts or []))
    for current_root, _, files in os.walk(root_dir):
        for name in files:
            if str(name).lower().endswith(ext_set):
                results.append(os.path.join(current_root, name))
    return results


def extract_mshc_payload(mshc_path, content_dir):
    os.makedirs(content_dir, exist_ok=True)
    with zipfile.ZipFile(mshc_path, "r") as zf:
        zf.extractall(content_dir)


def meta_content_by_names(soup, names):
    if not soup:
        return ""
    name_set = {str(n).strip().lower() for n in names}
    for meta in soup.find_all("meta"):
        meta_name = str(meta.get("name", "")).strip().lower()
        if meta_name in name_set:
            return str(meta.get("content", "") or "").strip()
    return ""


def normalize_md_line(text):
    return re.sub(r"\s+", " ", str(text or "")).strip()


def wrap_plain_text_for_pdf(text, width=100):
    words = str(text or "").split()
    if not words:
        return [""]
    lines = []
    current = []
    current_len = 0
    for w in words:
        if current_len + len(w) + (1 if current else 0) > width:
            lines.append(" ".join(current))
            current = [w]
            current_len = len(w)
        else:
            current.append(w)
            current_len += len(w) + (1 if current_len > 0 else 0)
    if current:
        lines.append(" ".join(current))
    return lines or [""]
