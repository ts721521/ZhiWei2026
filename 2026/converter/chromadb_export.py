# -*- coding: utf-8 -*-
"""ChromaDB export helper extracted from office_converter.py."""

import json
import os


def write_chromadb_export(
    docs,
    *,
    config,
    target_root,
    has_chromadb,
    chromadb_module=None,
    sanitize_collection_name_fn=None,
    resolve_persist_dir_fn=None,
    now_fn=None,
    log_info=None,
):
    if now_fn is None:
        from datetime import datetime as _dt  # noqa: PLC0415

        now_fn = _dt.now

    if not target_root:
        return None, []
    os.makedirs(target_root, exist_ok=True)

    out_dir = os.path.join(target_root, "_AI", "ChromaDB")
    os.makedirs(out_dir, exist_ok=True)
    ts = now_fn().strftime("%Y%m%d_%H%M%S")
    manifest_path = os.path.join(out_dir, f"chroma_export_{ts}.json")
    jsonl_path = os.path.join(out_dir, f"chroma_docs_{ts}.jsonl")

    write_jsonl = bool(config.get("chromadb_write_jsonl_fallback", True))
    if write_jsonl:
        with open(jsonl_path, "w", encoding="utf-8") as f:
            for item in docs:
                f.write(json.dumps(item, ensure_ascii=False) + "\n")

    raw_name = config.get("chromadb_collection_name", "office_corpus")
    collection_name = (
        sanitize_collection_name_fn(raw_name)
        if callable(sanitize_collection_name_fn)
        else str(raw_name)
    )
    persist_dir = (
        resolve_persist_dir_fn() if callable(resolve_persist_dir_fn) else os.path.join(out_dir, "db")
    )
    status = "empty"
    error = ""
    collection_count = 0

    if docs and has_chromadb and chromadb_module is not None:
        try:
            os.makedirs(persist_dir, exist_ok=True)
            client = chromadb_module.PersistentClient(path=persist_dir)
            collection = client.get_or_create_collection(name=collection_name)

            batch_size = 200
            for i in range(0, len(docs), batch_size):
                batch = docs[i : i + batch_size]
                ids = [x["id"] for x in batch]
                documents = [x["document"] for x in batch]
                metadatas = []
                for x in batch:
                    md = {}
                    for k, v in (x.get("metadata", {}) or {}).items():
                        if isinstance(v, (str, int, float, bool)):
                            md[k] = v
                        elif v is None:
                            md[k] = ""
                        else:
                            md[k] = str(v)
                    metadatas.append(md)
                collection.upsert(ids=ids, documents=documents, metadatas=metadatas)
            collection_count = int(collection.count() or 0)
            status = "ok"
        except Exception as e:
            status = "failed"
            error = str(e)
    elif docs and not has_chromadb:
        status = "chromadb_missing"
        error = "chromadb not installed"

    payload = {
        "version": 1,
        "generated_at": now_fn().isoformat(timespec="seconds"),
        "status": status,
        "error": error,
        "record_count": len(docs),
        "chromadb_available": has_chromadb,
        "persist_dir": os.path.abspath(persist_dir),
        "collection_name": collection_name,
        "collection_count": collection_count,
        "jsonl_path": os.path.abspath(jsonl_path) if write_jsonl else "",
    }
    with open(manifest_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)

    outputs = [manifest_path]
    if write_jsonl:
        outputs.append(jsonl_path)
    if callable(log_info):
        log_info(f"ChromaDB export status={status}, records={len(docs)}: {manifest_path}")
    return manifest_path, outputs
