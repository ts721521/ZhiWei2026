# -*- coding: utf-8 -*-
"""Markdown image remap and manifest helpers for merged markdown outputs."""

import hashlib
import os
import re
import shutil
from urllib.parse import unquote


_IMAGE_LINK_RE = re.compile(r"!\[(?P<alt>[^\]]*)\]\((?P<target>[^)\n]+)\)")
_FRONT_MATTER_SOURCE_RE = re.compile(
    r"^---\s*\n(?P<body>.*?)\n---\s*(?:\n|$)",
    re.DOTALL,
)


def extract_frontmatter_source_file(markdown_text):
    text = str(markdown_text or "")
    match = _FRONT_MATTER_SOURCE_RE.match(text)
    if not match:
        return ""
    body = match.group("body") or ""
    for line in body.splitlines():
        raw = str(line or "").strip()
        if not raw or ":" not in raw:
            continue
        key, value = raw.split(":", 1)
        if key.strip().lower() != "source_file":
            continue
        return value.strip().strip("'\"")
    return ""


def _parse_markdown_link_target(raw_target):
    raw = str(raw_target or "").strip()
    if not raw:
        return "", ""
    if raw.startswith("<"):
        end = raw.find(">")
        if end > 0:
            url = raw[1:end].strip()
            rest = raw[end + 1 :].strip()
            return url, rest
    parts = raw.split(maxsplit=1)
    url = parts[0].strip()
    title = parts[1].strip() if len(parts) > 1 else ""
    return url, title


def _is_external_or_anchor_url(url):
    check = str(url or "").strip().lower()
    return check.startswith(
        ("http://", "https://", "data:", "file://", "mailto:", "#")
    )


def _safe_doc_id(doc_index):
    return f"doc_{int(doc_index):03d}"


def _build_unique_image_name(
    *,
    doc_id,
    start_page,
    image_index,
    source_image_abs,
    original_url,
):
    ext = os.path.splitext(str(original_url or ""))[1] or os.path.splitext(source_image_abs)[1]
    ext = (ext or ".bin").lower()
    digest_src = f"{source_image_abs}|{original_url}|{image_index}"
    digest = hashlib.md5(digest_src.encode("utf-8")).hexdigest()[:8]
    page_num = int(start_page or 0)
    return f"{doc_id}_p{page_num:04d}_i{int(image_index):03d}_{digest}{ext}"


def _pick_merge_record(source_md_path, source_file_path, merge_index_records):
    if not merge_index_records:
        return None
    source_md_abs = os.path.normcase(os.path.abspath(str(source_md_path or "")))
    source_md_name = os.path.basename(source_md_abs)
    source_md_stem = os.path.splitext(source_md_name)[0].lower()

    src_file_abs = ""
    src_file_name = ""
    src_file_stem = ""
    if source_file_path:
        src_file_abs = os.path.normcase(os.path.abspath(source_file_path))
        src_file_name = os.path.basename(src_file_abs)
        src_file_stem = os.path.splitext(src_file_name)[0].lower()

    for rec in merge_index_records:
        rec_abs = os.path.normcase(os.path.abspath(str(rec.get("source_abspath", "") or "")))
        rec_name = os.path.basename(rec_abs or str(rec.get("source_filename", "") or ""))
        rec_stem = os.path.splitext(rec_name)[0].lower()
        if src_file_abs and rec_abs and src_file_abs == rec_abs:
            return rec
        if src_file_name and rec_name and src_file_name.lower() == rec_name.lower():
            return rec
        if src_file_stem and rec_stem and src_file_stem == rec_stem:
            return rec
        if source_md_stem and rec_stem and source_md_stem == rec_stem:
            return rec
    return None


def remap_markdown_images_for_merge(
    *,
    markdown_text,
    source_markdown_path,
    merge_output_dir,
    merged_markdown_name,
    doc_index,
    merge_index_records,
    open_fn=open,
    exists_fn=os.path.exists,
    makedirs_fn=os.makedirs,
    copy2_fn=shutil.copy2,
):
    content = str(markdown_text or "")
    source_md_abs = os.path.abspath(str(source_markdown_path or ""))
    source_md_dir = os.path.dirname(source_md_abs)
    source_file = extract_frontmatter_source_file(content)
    merge_rec = _pick_merge_record(source_md_abs, source_file, merge_index_records)
    merged_pdf_name = str((merge_rec or {}).get("merged_pdf_name", "") or "")
    merged_pdf_path = str((merge_rec or {}).get("merged_pdf_path", "") or "")
    start_page = int((merge_rec or {}).get("start_page_1based", 0) or 0)
    end_page = int((merge_rec or {}).get("end_page_1based", 0) or 0)
    doc_id = _safe_doc_id(doc_index)
    stem = os.path.splitext(str(merged_markdown_name or "Merged"))[0]
    assets_dir = os.path.join(merge_output_dir, f"{stem}_assets")

    image_entries = []
    image_count = {"v": 0}
    used_names = set()

    def _replace(match):
        image_count["v"] += 1
        local_idx = image_count["v"]
        alt = match.group("alt")
        raw_target = match.group("target")
        url, title = _parse_markdown_link_target(raw_target)
        ref_id = f"{doc_id}_img_{local_idx:03d}"
        replaced_url = url
        source_image_abs = ""
        source_found = False
        copied_rel = ""

        if url and not _is_external_or_anchor_url(url):
            decoded_url = unquote(url)
            source_image_abs = os.path.normpath(os.path.join(source_md_dir, decoded_url))
            source_found = bool(exists_fn(source_image_abs))
            if source_found:
                makedirs_fn(assets_dir, exist_ok=True)
                unique_name = _build_unique_image_name(
                    doc_id=doc_id,
                    start_page=start_page,
                    image_index=local_idx,
                    source_image_abs=source_image_abs,
                    original_url=url,
                )
                candidate = unique_name
                suffix = 2
                while candidate in used_names:
                    base, ext = os.path.splitext(unique_name)
                    candidate = f"{base}_v{suffix}{ext}"
                    suffix += 1
                used_names.add(candidate)
                target_abs = os.path.join(assets_dir, candidate)
                copy2_fn(source_image_abs, target_abs)
                copied_rel = os.path.relpath(target_abs, merge_output_dir).replace("\\", "/")
                replaced_url = copied_rel

        rebuilt_target = replaced_url
        if title:
            rebuilt_target = f"{rebuilt_target} {title}"

        image_entries.append(
            {
                "doc_id": doc_id,
                "source_markdown": source_md_abs,
                "source_file": source_file or "",
                "image_ref_id": ref_id,
                "image_index_in_doc": local_idx,
                "original_markdown_url": url,
                "original_markdown_target": raw_target.strip(),
                "source_image_abspath": source_image_abs if source_found else "",
                "source_image_exists": source_found,
                "merged_markdown_image_url": replaced_url,
                "copied_asset_relpath": copied_rel,
                "merged_pdf_name": merged_pdf_name,
                "merged_pdf_path": merged_pdf_path,
                "merged_pdf_start_page_1based": start_page,
                "merged_pdf_end_page_1based": end_page,
                "merged_pdf_page_hint_1based": start_page,
            }
        )

        return (
            f"<!-- IMG_REF:{ref_id} pdf={merged_pdf_name or 'N/A'} "
            f"pages={start_page}-{end_page} -->\n"
            f"![{alt}]({rebuilt_target})"
        )

    rewritten = _IMAGE_LINK_RE.sub(_replace, content)
    return rewritten, image_entries

