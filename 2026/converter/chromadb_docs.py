# -*- coding: utf-8 -*-
"""ChromaDB document collection helper extracted from office_converter.py."""

import hashlib
import os


def collect_chromadb_documents(
    *,
    generated_markdown_outputs,
    markdown_quality_records,
    config,
    chunk_text_for_vector_fn,
):
    docs = []
    md_paths = []
    seen = set()
    for p in generated_markdown_outputs or []:
        abs_p = os.path.abspath(str(p))
        if abs_p in seen or not os.path.exists(abs_p):
            continue
        seen.add(abs_p)
        md_paths.append(abs_p)

    md_to_pdf = {}
    for rec in markdown_quality_records or []:
        mdp = os.path.abspath(str(rec.get("markdown_path", "") or ""))
        if mdp:
            md_to_pdf[mdp] = str(rec.get("source_pdf", "") or "")

    max_chars = int(config.get("chromadb_max_chars_per_chunk", 1800) or 1800)
    overlap = int(config.get("chromadb_chunk_overlap", 200) or 200)

    for md_path in md_paths:
        try:
            with open(md_path, "r", encoding="utf-8") as f:
                raw = f.read()
        except Exception:
            continue
        chunks = chunk_text_for_vector_fn(raw, max_chars=max_chars, overlap=overlap)
        if not chunks:
            continue
        source_pdf = md_to_pdf.get(md_path, "")
        path_hash = hashlib.sha1(md_path.encode("utf-8", errors="ignore")).hexdigest()[:16]
        for idx, chunk in enumerate(chunks, 1):
            doc_id = f"md_{path_hash}_{idx:05d}"
            docs.append(
                {
                    "id": doc_id,
                    "document": chunk,
                    "metadata": {
                        "kind": "markdown",
                        "source_markdown_path": md_path,
                        "source_pdf_path": source_pdf,
                        "chunk_index": idx,
                        "chunk_count": len(chunks),
                        "char_count": len(chunk),
                    },
                }
            )
    return docs
