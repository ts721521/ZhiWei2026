# -*- coding: utf-8 -*-
"""ChromaDB helper functions extracted from office_converter.py."""

import os
import re


def sanitize_chromadb_collection_name(name):
    s = re.sub(r"[^A-Za-z0-9._-]+", "_", str(name or "").strip())
    s = s.strip("._-")
    if len(s) < 3:
        s = f"corpus_{s}" if s else "office_corpus"
    if len(s) > 63:
        s = s[:63].rstrip("._-")
    if not s:
        s = "office_corpus"
    return s


def resolve_chromadb_persist_dir(config):
    configured = str(config.get("chromadb_persist_dir", "") or "").strip()
    if configured:
        if os.path.isabs(configured):
            return configured
        return os.path.abspath(os.path.join(config.get("target_folder", ""), configured))
    return os.path.join(config.get("target_folder", ""), "_AI", "ChromaDB", "db")


def chunk_text_for_vector(text, max_chars=1800, overlap=200):
    content = str(text or "").strip()
    if not content:
        return []
    max_chars = max(200, int(max_chars or 1800))
    overlap = max(0, int(overlap or 0))
    if overlap >= max_chars:
        overlap = max(0, max_chars // 5)

    chunks = []
    start = 0
    n = len(content)
    while start < n:
        end = min(n, start + max_chars)
        piece = content[start:end].strip()
        if piece:
            chunks.append(piece)
        if end >= n:
            break
        start = max(0, end - overlap)
    return chunks
