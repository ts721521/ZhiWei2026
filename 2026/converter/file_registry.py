# -*- coding: utf-8 -*-
"""Incremental registry model extracted from office_converter.py."""

import json
import os
from datetime import datetime

from converter.platform_utils import is_win


class FileRegistry:
    """Incremental registry persisted to JSON."""

    def __init__(self, path, base_root=""):
        self.path = path
        self.base_root = os.path.abspath(base_root) if base_root else ""
        self.entries = {}
        self.version = 1
        self.loaded = False

    @staticmethod
    def _legacy_abs_key(path):
        p = os.path.abspath(path)
        if is_win():
            return p.lower()
        return p

    def _is_within_base(self, abs_path):
        if not self.base_root:
            return False
        base_norm = os.path.normcase(os.path.normpath(self.base_root))
        path_norm = os.path.normcase(os.path.normpath(abs_path))
        return path_norm == base_norm or path_norm.startswith(base_norm + os.sep)

    def normalize_path(self, path):
        if not path:
            return ""
        abs_path = os.path.abspath(path)
        if self._is_within_base(abs_path):
            rel = os.path.relpath(abs_path, self.base_root)
            key = rel.replace("\\", "/")
        else:
            key = abs_path.replace("\\", "/")
        if is_win():
            key = key.lower()
        return key

    def load(self):
        self.entries = {}
        if not self.path or not os.path.exists(self.path):
            self.loaded = True
            return
        try:
            with open(self.path, "r", encoding="utf-8") as f:
                data = json.load(f)
            raw_entries = data.get("entries", {})
            if isinstance(raw_entries, dict):
                migrated = {}
                for old_key, old_entry in raw_entries.items():
                    if isinstance(old_entry, dict):
                        source_ref = old_entry.get("source_path") or old_key
                        entry = dict(old_entry)
                    else:
                        source_ref = old_key
                        entry = {"source_path": old_key}

                    new_key = self.normalize_path(source_ref)
                    if not new_key:
                        continue

                    if not entry.get("source_path"):
                        source_ref_str = str(source_ref)
                        if os.path.isabs(source_ref_str):
                            entry["source_path"] = os.path.abspath(source_ref_str)
                        elif self.base_root:
                            entry["source_path"] = os.path.abspath(
                                os.path.join(
                                    self.base_root,
                                    source_ref_str.replace("/", os.sep),
                                )
                            )
                    migrated[new_key] = entry
                self.entries = migrated
            self.version = int(data.get("version", 1) or 1)
        except Exception:
            self.entries = {}
            self.version = 1
        self.loaded = True

    def save(self, run_summary=None):
        folder = os.path.dirname(self.path)
        if folder:
            os.makedirs(folder, exist_ok=True)
        payload = {
            "version": self.version,
            "updated_at": datetime.now().isoformat(timespec="seconds"),
            "key_strategy": "source_rel_forward_slash",
            "entry_count": len(self.entries),
            "entries": self.entries,
        }
        if run_summary:
            payload["last_run"] = run_summary
        with open(self.path, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)

    def get(self, source_path):
        key = self.normalize_path(source_path)
        hit = self.entries.get(key)
        if hit is not None:
            return hit
        # Backward compatibility for legacy absolute-path key style.
        return self.entries.get(self._legacy_abs_key(source_path))

    def set(self, source_path, entry):
        key = self.normalize_path(source_path)
        self.entries[key] = entry

    def keys(self):
        return list(self.entries.keys())
