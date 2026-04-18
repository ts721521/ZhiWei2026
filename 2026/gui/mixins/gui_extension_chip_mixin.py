# -*- coding: utf-8 -*-
"""扩展名 chip 编辑器：把 allowed_extensions 桶结构以可点击 chip 暴露给用户。

桶顺序与默认值见 converter/default_config.py：
    word / excel / powerpoint / pdf / cab
"""

import tkinter as tk
from tkinter.constants import LEFT, X, Y, W


EXTENSION_BUCKETS = ("word", "excel", "powerpoint", "pdf", "cab")
_BUCKET_LABELS = {
    "word": "Word",
    "excel": "Excel",
    "powerpoint": "PowerPoint",
    "pdf": "PDF",
    "cab": "CAB",
}


class ExtensionChipEditorMixin:
    def _create_extension_chip_editor(self, parent, initial=None, on_change=None):
        """构建 {bucket: [.ext, ...]} 编辑器。

        Returns: (frame, getter() -> dict, setter(dict))
        """
        # ttk widgets do not expose -background; fall back to the root style lookup.
        bg = "#FFFFFF"
        try:
            bg = parent.cget("background") or bg
        except Exception:
            try:
                from tkinter import ttk
                style = ttk.Style()
                bg = style.lookup("TFrame", "background") or bg
            except Exception:
                pass
        wrap = tk.Frame(parent, bg=bg)

        state = {b: [] for b in EXTENSION_BUCKETS}
        chip_holders = {}

        def _normalize_ext(raw):
            s = (raw or "").strip().lower()
            if not s:
                return ""
            if not s.startswith("."):
                s = "." + s
            # drop spaces and stray characters
            s = "".join(c for c in s if c.isalnum() or c == ".")
            if s == ".":
                return ""
            return s

        def _emit_change():
            if callable(on_change):
                try:
                    on_change(getter())
                except Exception:
                    pass

        def _redraw_bucket(bucket):
            holder = chip_holders[bucket]
            for w in holder.winfo_children():
                w.destroy()
            for ext in state[bucket]:
                chip = tk.Frame(
                    holder,
                    bg="#E0E7FF",
                    highlightthickness=1,
                    highlightbackground="#A5B4FC",
                )
                tk.Label(
                    chip,
                    text=ext,
                    bg="#E0E7FF",
                    fg="#1E3A8A",
                    font=("System", 8),
                    padx=4,
                ).pack(side=LEFT)
                tk.Button(
                    chip,
                    text="\u00d7",
                    bg="#E0E7FF",
                    fg="#7F1D1D",
                    bd=0,
                    relief="flat",
                    padx=2,
                    font=("System", 8, "bold"),
                    command=lambda b=bucket, e=ext: _remove(b, e),
                ).pack(side=LEFT)
                chip.pack(side=LEFT, padx=2, pady=2)

        def _add(bucket, raw):
            ext = _normalize_ext(raw)
            if not ext:
                return False
            if ext in state[bucket]:
                return False
            state[bucket].append(ext)
            state[bucket].sort()
            _redraw_bucket(bucket)
            _emit_change()
            return True

        def _remove(bucket, ext):
            if ext in state[bucket]:
                state[bucket].remove(ext)
                _redraw_bucket(bucket)
                _emit_change()

        for bucket in EXTENSION_BUCKETS:
            row = tk.Frame(wrap, bg=bg)
            row.pack(fill=X, pady=1)
            tk.Label(
                row,
                text=_BUCKET_LABELS[bucket] + ":",
                bg=bg,
                width=10,
                anchor=W,
                font=("System", 9, "bold"),
            ).pack(side=LEFT)
            chips_holder = tk.Frame(row, bg=bg)
            chips_holder.pack(side=LEFT, fill=X, expand=True)
            chip_holders[bucket] = chips_holder
            ent = tk.Entry(row, width=8, font=("Consolas", 9))
            ent.pack(side=LEFT, padx=(4, 2))

            def _commit(event=None, b=bucket, e=ent):
                if _add(b, e.get()):
                    e.delete(0, tk.END)
                return "break"

            ent.bind("<Return>", _commit)
            tk.Button(
                row,
                text="+",
                width=2,
                command=lambda b=bucket, e=ent: _commit(b=b, e=e),
            ).pack(side=LEFT)

        def setter(value):
            value = value or {}
            for bucket in EXTENSION_BUCKETS:
                vals = value.get(bucket) or []
                normalized = []
                for v in vals:
                    n = _normalize_ext(v)
                    if n and n not in normalized:
                        normalized.append(n)
                normalized.sort()
                state[bucket] = normalized
                _redraw_bucket(bucket)

        def getter():
            return {b: list(state[b]) for b in EXTENSION_BUCKETS}

        if initial is not None:
            setter(initial)

        return wrap, getter, setter
