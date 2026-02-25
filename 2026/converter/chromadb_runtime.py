# -*- coding: utf-8 -*-
"""Runtime wrapper for ChromaDB export extracted from office_converter.py."""


def write_chromadb_export_for_converter(
    converter,
    *,
    write_chromadb_export_fn,
    has_chromadb,
    chromadb_module,
    now_fn,
    log_info_fn,
):
    if not converter.config.get("enable_chromadb_export", False):
        return None

    target_root = converter.config.get("target_folder", "")
    if not target_root:
        return None

    docs = converter._collect_chromadb_documents()
    if not docs:
        converter.generated_chromadb_outputs = []
        converter.chromadb_export_manifest_path = None
        log_info_fn("ChromaDB export skipped: no Markdown chunks available.")
        return None

    manifest_path, outputs = write_chromadb_export_fn(
        docs,
        config=converter.config,
        target_root=target_root,
        has_chromadb=has_chromadb,
        chromadb_module=chromadb_module if has_chromadb else None,
        sanitize_collection_name_fn=converter._sanitize_chromadb_collection_name,
        resolve_persist_dir_fn=converter._resolve_chromadb_persist_dir,
        now_fn=now_fn,
        log_info=log_info_fn,
    )
    converter.generated_chromadb_outputs = outputs
    converter.chromadb_export_manifest_path = manifest_path
    return manifest_path
