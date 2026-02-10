# -*- coding: utf-8 -*-
"""Search adapters for local desktop file retrieval tools."""

from __future__ import annotations

import os
import shutil
import subprocess
from dataclasses import dataclass
from typing import Optional


@dataclass
class SearchAdapterResult:
    ok: bool
    engine: str
    command: list[str]
    stderr: str = ""


class EverythingAdapter:
    COMMON_PATHS = [
        r"C:\\Program Files\\Everything\\es.exe",
        r"C:\\Program Files (x86)\\Everything\\es.exe",
    ]

    def __init__(self, es_path: str = "", timeout_ms: int = 1500):
        self.timeout_ms = timeout_ms
        self.es_path = self._resolve_es_path(es_path)

    def _resolve_es_path(self, es_path: str) -> Optional[str]:
        if es_path:
            if os.path.isfile(es_path):
                return es_path
            return None

        in_path = shutil.which("es.exe") or shutil.which("es")
        if in_path:
            return in_path

        for p in self.COMMON_PATHS:
            if os.path.isfile(p):
                return p

        return None

    def is_available(self) -> bool:
        return bool(self.es_path)

    def build_query_command(self, filename: str, directory: str = "") -> list[str]:
        if not self.es_path:
            return []
        cmd = [self.es_path]
        if directory:
            cmd.extend(["-path", directory])
        if filename:
            cmd.extend(["-name", filename])
        cmd.extend(["-sort-path", "-n", "20"])
        return cmd

    def run_query(self, filename: str, directory: str = "") -> SearchAdapterResult:
        cmd = self.build_query_command(filename=filename, directory=directory)
        if not cmd:
            return SearchAdapterResult(
                ok=False,
                engine="everything",
                command=[],
                stderr="es.exe not found",
            )
        try:
            subprocess.run(
                cmd,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.PIPE,
                text=True,
                timeout=max(self.timeout_ms / 1000.0, 0.2),
                check=False,
            )
            return SearchAdapterResult(ok=True, engine="everything", command=cmd)
        except Exception as e:
            return SearchAdapterResult(
                ok=False,
                engine="everything",
                command=cmd,
                stderr=str(e),
            )


def build_listary_query(short_id: str, md5_value: str, filename: str, path: str) -> str:
    parts = []
    if short_id:
        parts.append(f"id:{short_id}")
    if md5_value:
        parts.append(f"md5:{md5_value}")
    if filename:
        parts.append(f'file:"{filename}"')
    if path:
        parts.append(f'path:"{path}"')
    return " ".join(parts)
