# -*- coding: utf-8 -*-
"""Markdown text cleaning/render helpers extracted from office_converter.py."""

import re


def normalize_extracted_text(text):
    text = (text or "").replace("\r\n", "\n").replace("\r", "\n")
    lines = [ln.rstrip() for ln in text.split("\n")]
    cleaned = []
    last_blank = False
    for ln in lines:
        if ln.strip():
            cleaned.append(ln)
            last_blank = False
        else:
            if not last_blank:
                cleaned.append("")
            last_blank = True
    return "\n".join(cleaned).strip()


def normalize_margin_line(line):
    s = str(line or "").strip().lower()
    s = re.sub(r"\s+", " ", s)
    # Ignore punctuation-only differences for repeated margin detection.
    s = re.sub(r"[\W_]+", "", s, flags=re.UNICODE)
    return s


def looks_like_page_number_line(line):
    s = str(line or "").strip().lower()
    if not s:
        return False
    if re.fullmatch(r"[#\-\s]*\d+[#\-\s]*", s):
        return True
    if re.fullmatch(r"page\s+\d+(\s*/\s*\d+)?", s):
        return True
    if re.fullmatch(r"page\s+\d+\s+of\s+\d+", s):
        return True
    if re.fullmatch(r"\d+\s*/\s*\d+", s):
        return True
    if re.fullmatch(r"第\s*\d+\s*页(\s*/\s*共?\s*\d+\s*页)?", s):
        return True
    return False


def collect_margin_candidates(page_raw_texts):
    """Collect repeated top/bottom lines across pages as header/footer candidates."""
    header_counts = {}
    footer_counts = {}
    page_count = len(page_raw_texts)
    threshold = max(2, (page_count + 1) // 2)
    if page_count < 2:
        return set(), set()

    for raw in page_raw_texts:
        lines = normalize_extracted_text(raw).splitlines()
        non_empty = [ln.strip() for ln in lines if ln.strip()]
        if not non_empty:
            continue
        top = non_empty[0]
        bottom = non_empty[-1]
        top_key = normalize_margin_line(top)
        bottom_key = normalize_margin_line(bottom)
        if top_key and not looks_like_page_number_line(top):
            header_counts[top_key] = header_counts.get(top_key, 0) + 1
        if bottom_key and not looks_like_page_number_line(bottom):
            footer_counts[bottom_key] = footer_counts.get(bottom_key, 0) + 1

    header_keys = {k for k, v in header_counts.items() if v >= threshold}
    footer_keys = {k for k, v in footer_counts.items() if v >= threshold}
    return header_keys, footer_keys


def clean_markdown_page_lines(raw_text, header_keys, footer_keys):
    lines = normalize_extracted_text(raw_text).splitlines()
    kept = [ln.strip() for ln in lines]

    removed_header = 0
    removed_footer = 0
    removed_page_no = 0

    while kept:
        top = kept[0]
        top_key = normalize_margin_line(top)
        if looks_like_page_number_line(top):
            kept.pop(0)
            removed_page_no += 1
            continue
        if top_key and top_key in header_keys:
            kept.pop(0)
            removed_header += 1
            continue
        break

    while kept:
        bottom = kept[-1]
        bottom_key = normalize_margin_line(bottom)
        if looks_like_page_number_line(bottom):
            kept.pop()
            removed_page_no += 1
            continue
        if bottom_key and bottom_key in footer_keys:
            kept.pop()
            removed_footer += 1
            continue
        break

    return kept, {
        "removed_header_lines": removed_header,
        "removed_footer_lines": removed_footer,
        "removed_page_number_lines": removed_page_no,
        "remaining_lines": len(kept),
    }


def looks_like_heading_line(line):
    s = str(line or "").strip()
    if not s:
        return False
    if len(s) > 90:
        return False
    if re.match(r"^(\d+(\.\d+){0,3}|[一二三四五六七八九十]+)[\.\、\)]\s*\S+", s):
        return True
    if re.match(r"^(chapter|section)\s+\d+", s, flags=re.IGNORECASE):
        return True
    if re.match(r"^[A-Z0-9][A-Z0-9 \-_/]{3,}$", s):
        return True
    if s.endswith(":") and len(s) <= 40:
        return True
    return False


def render_markdown_blocks(lines, structured_headings=True):
    blocks = []
    buf = []
    heading_count = 0

    def flush_para():
        if not buf:
            return
        merged = ""
        for ln in buf:
            if not merged:
                merged = ln
                continue
            if merged.endswith("-") and ln:
                merged = merged[:-1] + ln
            elif merged.endswith(("。", "！", "？", ".", "!", "?", "；", ";", "：", ":")):
                merged += "\n" + ln
            else:
                merged += " " + ln
        blocks.append(merged.strip())
        buf.clear()

    for raw in lines:
        ln = str(raw or "").strip()
        if not ln:
            flush_para()
            continue
        if structured_headings and looks_like_heading_line(ln):
            flush_para()
            blocks.append(f"### {ln}")
            heading_count += 1
        else:
            buf.append(ln)
    flush_para()
    return blocks, heading_count
