# -*- coding: utf-8 -*-
"""Profile management methods extracted from OfficeGUI."""

import json
import os
import re
from datetime import datetime
import tkinter as tk
from tkinter import messagebox, simpledialog, ttk
from tkinter.constants import *

try:
    import ttkbootstrap as tb
except ModuleNotFoundError:
    class _BootstyleMixin:
        def __init__(self, *args, **kwargs):
            kwargs.pop("bootstyle", None)
            super().__init__(*args, **kwargs)

        def configure(self, cnf=None, **kwargs):
            kwargs.pop("bootstyle", None)
            if isinstance(cnf, dict) and "bootstyle" in cnf:
                cnf = dict(cnf)
                cnf.pop("bootstyle", None)
            return super().configure(cnf, **kwargs)

        config = configure

    class _FallbackFrame(_BootstyleMixin, ttk.Frame):
        pass

    class _FallbackLabel(_BootstyleMixin, ttk.Label):
        pass

    class _FallbackButton(_BootstyleMixin, ttk.Button):
        pass

    class _FallbackEntry(_BootstyleMixin, ttk.Entry):
        pass

    class _TBNamespace:
        Frame = _FallbackFrame
        Label = _FallbackLabel
        Button = _FallbackButton
        Entry = _FallbackEntry

    tb = _TBNamespace()


class ProfileManagementMixin:
    def _profiles_dir(self):
        return os.path.join(self.script_dir, "config_profiles")

    def _profiles_index_path(self):
        return os.path.join(self._profiles_dir(), "profiles_index.json")

    def _profile_abs_path(self, file_name):
        safe_file = os.path.basename(str(file_name or "").strip())
        return os.path.join(self._profiles_dir(), safe_file)

    def _profile_timestamp(self):
        return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    def _profile_file_mtime(self, file_path):
        try:
            dt = datetime.fromtimestamp(os.path.getmtime(file_path))
            return dt.strftime("%Y-%m-%d %H:%M:%S")
        except Exception:
            return ""

    def _sanitize_profile_stem(self, name):
        stem = str(name or "").strip()
        stem = re.sub(r'[\\/:*?"<>|]+', "_", stem)
        stem = stem.strip().strip(".")
        return stem[:80] or "profile"

    def _next_profile_filename(self, name, records, exclude_file=None):
        stem = self._sanitize_profile_stem(name)
        existing = {
            str(rec.get("file", "")).strip().lower()
            for rec in (records or [])
            if isinstance(rec, dict)
        }
        exclude_lower = str(exclude_file or "").strip().lower()
        idx = 1
        while True:
            suffix = "" if idx == 1 else f"_{idx}"
            candidate = f"{stem}{suffix}.json"
            lower = candidate.lower()
            if lower == exclude_lower or lower not in existing:
                return candidate
            idx += 1

    def _load_profile_records(self):
        os.makedirs(self._profiles_dir(), exist_ok=True)
        index_path = self._profiles_index_path()
        index_name = os.path.basename(index_path).lower()
        records = []
        payload = {}
        if os.path.exists(index_path):
            try:
                with open(index_path, "r", encoding="utf-8") as f:
                    payload = json.load(f)
            except Exception:
                payload = {}

        raw_records = payload.get("profiles", []) if isinstance(payload, dict) else []
        for rec in raw_records:
            if not isinstance(rec, dict):
                continue
            file_name = os.path.basename(str(rec.get("file", "")).strip())
            if not file_name or file_name.lower() == index_name:
                continue
            abs_path = self._profile_abs_path(file_name)
            if not os.path.isfile(abs_path):
                continue
            name = str(rec.get("name", "")).strip() or os.path.splitext(file_name)[0]
            note = str(rec.get("note", "")).strip()
            updated_at = str(
                rec.get("updated_at", "")
            ).strip() or self._profile_file_mtime(abs_path)
            records.append(
                {
                    "id": str(rec.get("id", "")).strip()
                    or datetime.now().strftime("%Y%m%d%H%M%S%f"),
                    "name": name,
                    "file": file_name,
                    "note": note,
                    "updated_at": updated_at,
                }
            )

        known_files = {str(rec.get("file", "")).strip().lower() for rec in records}
        for entry in sorted(os.listdir(self._profiles_dir())):
            lower = entry.lower()
            if not lower.endswith(".json"):
                continue
            if lower == index_name or lower in known_files:
                continue
            abs_path = self._profile_abs_path(entry)
            if not os.path.isfile(abs_path):
                continue
            records.append(
                {
                    "id": datetime.now().strftime("%Y%m%d%H%M%S%f"),
                    "name": os.path.splitext(entry)[0],
                    "file": entry,
                    "note": "",
                    "updated_at": self._profile_file_mtime(abs_path),
                }
            )
        return records

    def _load_builtin_config_records(self):
        records = []
        root_dirs = [
            ("preset", os.path.join(self.script_dir, "configs", "presets")),
            ("scenario", os.path.join(self.script_dir, "configs", "scenarios")),
        ]
        for kind, root in root_dirs:
            if not os.path.isdir(root):
                continue
            for base, _dirs, files in os.walk(root):
                for name in sorted(files):
                    if not str(name).lower().endswith(".json"):
                        continue
                    abs_path = os.path.join(base, name)
                    rel_path = os.path.relpath(abs_path, self.script_dir).replace(
                        "\\", "/"
                    )
                    stem = os.path.splitext(name)[0]
                    note = self._builtin_config_note_from_file(abs_path)
                    records.append(
                        {
                            "id": f"builtin:{kind}:{rel_path}",
                            "name": f"[{kind}] {stem}",
                            "file": rel_path,
                            "note": note,
                            "updated_at": self._profile_file_mtime(abs_path),
                            "abs_path": abs_path,
                            "is_builtin": True,
                        }
                    )
        return records

    def _builtin_config_note_from_file(self, abs_path, max_len=200):
        """Read _meta.description and _meta.notes from a config JSON for display."""
        try:
            with open(abs_path, "r", encoding="utf-8") as f:
                cfg = json.load(f)
            _meta = cfg.get("_meta") or {}
            description = str(_meta.get("description", "") or "").strip()
            notes = _meta.get("notes")
            if isinstance(notes, list) and notes:
                first_note = str(notes[0] or "").strip()
                if first_note:
                    note = f"{description} | {first_note}" if description else first_note
                else:
                    note = description
            else:
                note = description
            if not note:
                return "Built-in config"
            return note[:max_len] if len(note) > max_len else note
        except (OSError, TypeError, ValueError, json.JSONDecodeError):
            return "Built-in config"

    def _save_profile_records(self, records):
        os.makedirs(self._profiles_dir(), exist_ok=True)
        index_path = self._profiles_index_path()
        out_records = []
        for rec in records or []:
            if not isinstance(rec, dict):
                continue
            file_name = os.path.basename(str(rec.get("file", "")).strip())
            if not file_name:
                continue
            out_records.append(
                {
                    "id": str(rec.get("id", "")).strip()
                    or datetime.now().strftime("%Y%m%d%H%M%S%f"),
                    "name": str(rec.get("name", "")).strip()
                    or os.path.splitext(file_name)[0],
                    "file": file_name,
                    "note": str(rec.get("note", "")).strip(),
                    "updated_at": str(rec.get("updated_at", "")).strip(),
                }
            )
        with open(index_path, "w", encoding="utf-8") as f:
            json.dump(
                {"version": 1, "profiles": out_records}, f, indent=4, ensure_ascii=False
            )

    def _get_selected_profile_record(self):
        if self.profile_tree is None or not self.profile_tree.winfo_exists():
            return None
        selection = self.profile_tree.selection()
        if not selection:
            return None
        return self._profile_tree_rows.get(selection[0])

    def _update_profile_manager_controls(self):
        if (
            self.profile_manager_win is None
            or not self.profile_manager_win.winfo_exists()
        ):
            return
        has_selected = self._get_selected_profile_record() is not None
        base_state = "disabled" if self._ui_running else "normal"
        selected_state = (
            "disabled" if (self._ui_running or not has_selected) else "normal"
        )
        for btn_name in (
            "btn_profile_new",
            "btn_profile_refresh",
            "btn_profile_open_dir",
        ):
            btn = getattr(self, btn_name, None)
            if btn is not None and btn.winfo_exists():
                btn.configure(state=base_state)
        for btn_name in (
            "btn_profile_load",
            "btn_profile_save",
            "btn_profile_rename",
            "btn_profile_note",
            "btn_profile_delete",
        ):
            btn = getattr(self, btn_name, None)
            if btn is not None and btn.winfo_exists():
                btn.configure(state=selected_state)

    def _on_profile_tree_select(self, _event=None):
        self._update_profile_manager_controls()

    def _refresh_profile_tree(self, select_file=None):
        if self.profile_tree is None or not self.profile_tree.winfo_exists():
            return
        records = self._load_profile_records()
        self._profile_tree_rows = {}
        self.profile_tree.delete(*self.profile_tree.get_children())
        active_path = os.path.abspath(self.config_path)
        target_file = str(select_file or "").strip()
        target_iid = None
        for rec in records:
            abs_path = os.path.abspath(self._profile_abs_path(rec.get("file", "")))
            iid = self.profile_tree.insert(
                "",
                "end",
                values=(
                    rec.get("name", ""),
                    rec.get("file", ""),
                    rec.get("note", ""),
                    rec.get("updated_at", ""),
                ),
            )
            row = dict(rec)
            row["abs_path"] = abs_path
            self._profile_tree_rows[iid] = row
            if target_file and rec.get("file", "") == target_file:
                target_iid = iid
            elif not target_file and abs_path == active_path:
                target_iid = iid
        if target_iid is not None:
            self.profile_tree.selection_set(target_iid)
            self.profile_tree.focus(target_iid)
            self.profile_tree.see(target_iid)
        self.var_profile_active_path.set(self.config_path)
        self._update_profile_manager_controls()

    def _close_profile_manager_window(self):
        if (
            self.profile_manager_win is not None
            and self.profile_manager_win.winfo_exists()
        ):
            try:
                self.profile_manager_win.destroy()
            except Exception:
                pass
        self.profile_manager_win = None
        self.profile_tree = None
        self._profile_tree_rows = {}

    def open_profile_manager_window(self):
        if (
            self.profile_manager_win is not None
            and self.profile_manager_win.winfo_exists()
        ):
            self.profile_manager_win.lift()
            self.profile_manager_win.focus_force()
            self._refresh_profile_tree()
            return
        os.makedirs(self._profiles_dir(), exist_ok=True)
        self.profile_manager_win = tk.Toplevel(self)
        self.profile_manager_win.title(self.tr("win_profile_manager"))
        self.profile_manager_win.geometry("980x520")
        self.profile_manager_win.minsize(840, 420)
        self.profile_manager_win.protocol(
            "WM_DELETE_WINDOW", self._close_profile_manager_window
        )

        root = tb.Frame(self.profile_manager_win, padding=10)
        root.pack(fill=BOTH, expand=YES)

        top_row = tb.Frame(root)
        top_row.pack(fill=X, pady=(0, 8))
        tb.Label(
            top_row,
            text=self.tr("lbl_profile_active_config"),
            font=("System", 9, "bold"),
        ).pack(side=LEFT)
        tb.Label(
            top_row,
            textvariable=self.var_profile_active_path,
            bootstyle="secondary",
        ).pack(side=LEFT, fill=X, expand=YES, padx=(8, 0))

        tree_frame = tb.Frame(root)
        tree_frame.pack(fill=BOTH, expand=YES)
        cols = ("name", "file", "note", "updated")
        self.profile_tree = ttk.Treeview(
            tree_frame, columns=cols, show="headings", selectmode="browse"
        )
        self.profile_tree.heading("name", text=self.tr("col_profile_name"))
        self.profile_tree.heading("file", text=self.tr("col_profile_file"))
        self.profile_tree.heading("note", text=self.tr("col_profile_note"))
        self.profile_tree.heading("updated", text=self.tr("col_profile_updated"))
        self.profile_tree.column("name", width=180, anchor="w")
        self.profile_tree.column("file", width=220, anchor="w")
        self.profile_tree.column("note", width=360, anchor="w")
        self.profile_tree.column("updated", width=160, anchor="center")
        scr_y = tb.Scrollbar(
            tree_frame, orient=VERTICAL, command=self.profile_tree.yview
        )
        scr_x = tb.Scrollbar(
            tree_frame, orient=HORIZONTAL, command=self.profile_tree.xview
        )
        self.profile_tree.configure(yscrollcommand=scr_y.set, xscrollcommand=scr_x.set)
        self.profile_tree.pack(side=LEFT, fill=BOTH, expand=YES)
        scr_y.pack(side=RIGHT, fill=Y)
        scr_x.pack(side=BOTTOM, fill=X)
        self.profile_tree.bind("<<TreeviewSelect>>", self._on_profile_tree_select)

        btn_row = tb.Frame(root)
        btn_row.pack(fill=X, pady=(8, 0))
        self.btn_profile_new = tb.Button(
            btn_row,
            text=self.tr("btn_profile_new"),
            command=self._create_profile_from_current,
            bootstyle="success-outline",
            width=14,
        )
        self.btn_profile_new.pack(side=LEFT)
        self.btn_profile_load = tb.Button(
            btn_row,
            text=self.tr("btn_profile_load"),
            command=self._load_selected_profile,
            bootstyle="primary",
            width=12,
        )
        self.btn_profile_load.pack(side=LEFT, padx=(6, 0))
        self.btn_profile_save = tb.Button(
            btn_row,
            text=self.tr("btn_profile_save"),
            command=self._save_current_to_selected_profile,
            bootstyle="success",
            width=14,
        )
        self.btn_profile_save.pack(side=LEFT, padx=(6, 0))
        self.btn_profile_rename = tb.Button(
            btn_row,
            text=self.tr("btn_profile_rename"),
            command=self._rename_selected_profile,
            bootstyle="secondary-outline",
            width=10,
        )
        self.btn_profile_rename.pack(side=LEFT, padx=(6, 0))
        self.btn_profile_note = tb.Button(
            btn_row,
            text=self.tr("btn_profile_note"),
            command=self._edit_selected_profile_note,
            bootstyle="secondary-outline",
            width=10,
        )
        self.btn_profile_note.pack(side=LEFT, padx=(6, 0))
        self.btn_profile_delete = tb.Button(
            btn_row,
            text=self.tr("btn_profile_delete"),
            command=self._delete_selected_profile,
            bootstyle="danger-outline",
            width=10,
        )
        self.btn_profile_delete.pack(side=LEFT, padx=(6, 0))
        self.btn_profile_open_dir = tb.Button(
            btn_row,
            text=self.tr("btn_profile_open_dir"),
            command=self._open_profiles_folder,
            bootstyle="info-outline",
            width=10,
        )
        self.btn_profile_open_dir.pack(side=RIGHT)
        self.btn_profile_refresh = tb.Button(
            btn_row,
            text=self.tr("btn_profile_refresh"),
            command=self._refresh_profile_tree,
            bootstyle="info-outline",
            width=10,
        )
        self.btn_profile_refresh.pack(side=RIGHT, padx=(0, 6))

        self._refresh_profile_tree()
        self._update_profile_manager_controls()

    def _build_default_profile_name(self):
        return datetime.now().strftime("config_%Y%m%d_%H%M%S")

    def _save_profile_with_meta(self, name, note="", show_msg=True):
        profile_name = str(name or "").strip()
        if not profile_name:
            if show_msg:
                messagebox.showwarning(
                    self.tr("btn_save_cfg"), self.tr("msg_profile_name_required")
                )
            return False
        profile_note = str(note or "").strip()
        records = self._load_profile_records()
        file_name = self._next_profile_filename(profile_name, records)
        payload = self._compose_config_from_ui(
            self._load_config_for_write(), scope="all"
        )
        profile_path = self._profile_abs_path(file_name)
        try:
            with open(profile_path, "w", encoding="utf-8") as f:
                json.dump(payload, f, indent=4, ensure_ascii=False)
            records.insert(
                0,
                {
                    "id": datetime.now().strftime("%Y%m%d%H%M%S%f"),
                    "name": profile_name,
                    "file": file_name,
                    "note": profile_note,
                    "updated_at": self._profile_timestamp(),
                },
            )
            self._save_profile_records(records)
            if (
                self.profile_manager_win is not None
                and self.profile_manager_win.winfo_exists()
            ):
                self._refresh_profile_tree(select_file=file_name)
            if (
                self.load_profile_dialog is not None
                and self.load_profile_dialog.winfo_exists()
            ):
                self._refresh_load_profile_tree(select_file=file_name)
            msg = self.tr("msg_profile_create_ok").format(profile_name)
            if show_msg:
                messagebox.showinfo(self.tr("btn_save_cfg"), msg)
            if hasattr(self, "var_status"):
                self.var_status.set(msg)
            if hasattr(self, "var_locator_result"):
                self.var_locator_result.set(msg)
            return True
        except Exception as e:
            if show_msg:
                messagebox.showerror(
                    self.tr("btn_save_cfg"), self.tr("msg_save_fail").format(e)
                )
            return False

    def _load_profile_record(self, rec, confirm_dirty=True, show_msg=False):
        if not isinstance(rec, dict):
            return False
        if confirm_dirty and self.cfg_dirty:
            confirm = messagebox.askyesno(
                self.tr("btn_load_cfg"),
                self.tr("msg_confirm_load_config_dirty"),
            )
            if not confirm:
                return False
        profile_path = str(rec.get("abs_path", "")).strip()
        if not os.path.isfile(profile_path):
            messagebox.showerror(
                self.tr("btn_load_cfg"),
                self.tr("msg_profile_file_missing").format(profile_path),
            )
            return False
        try:
            with open(profile_path, "r", encoding="utf-8") as f:
                profile_cfg = json.load(f)
            if not isinstance(profile_cfg, dict):
                raise ValueError("profile config must be an object")
            display_name = (
                str(rec.get("name", "")).strip()
                or str(rec.get("file", "")).strip()
                or os.path.basename(profile_path)
            )
            self._active_config_label = display_name
            self._active_config_origin = profile_path
            is_builtin = bool(rec.get("is_builtin", False))
            if is_builtin:
                # Built-in presets/scenarios should not hijack active config path.
                # Instead, load by replacing current active config content.
                with open(self.config_path, "w", encoding="utf-8") as f:
                    json.dump(profile_cfg, f, indent=4, ensure_ascii=False)
            else:
                self.config_path = profile_path
                if hasattr(self, "var_config_path"):
                    self.var_config_path.set(self.config_path)
                self.var_profile_active_path.set(self.config_path)
            self._load_config_to_ui()
            if (
                self.profile_manager_win is not None
                and self.profile_manager_win.winfo_exists()
            ):
                self._refresh_profile_tree(select_file=rec.get("file", ""))
            if (
                self.load_profile_dialog is not None
                and self.load_profile_dialog.winfo_exists()
            ):
                self._refresh_load_profile_tree(select_abs_path=profile_path)
            msg = self.tr("msg_profile_load_ok").format(rec.get("name", ""))
            if show_msg:
                messagebox.showinfo(self.tr("btn_load_cfg"), msg)
            if hasattr(self, "var_status"):
                self.var_status.set(msg)
            if hasattr(self, "var_locator_result"):
                self.var_locator_result.set(msg)
            return True
        except Exception as e:
            messagebox.showerror(
                self.tr("btn_load_cfg"), self.tr("msg_save_fail").format(e)
            )
            return False

    def _create_profile_from_current(self):
        parent = (
            self.profile_manager_win
            if self.profile_manager_win is not None
            and self.profile_manager_win.winfo_exists()
            else self
        )
        default_name = self._build_default_profile_name()
        name = simpledialog.askstring(
            self.tr("btn_profile_new"),
            self.tr("msg_profile_name_prompt"),
            parent=parent,
            initialvalue=default_name,
        )
        if name is None:
            return
        name = str(name).strip()
        if not name:
            messagebox.showwarning(
                self.tr("btn_profile_new"), self.tr("msg_profile_name_required")
            )
            return
        note = simpledialog.askstring(
            self.tr("btn_profile_note"),
            self.tr("msg_profile_note_prompt"),
            parent=parent,
            initialvalue="",
        )
        note = "" if note is None else note
        self._save_profile_with_meta(name, note, show_msg=False)

    def _load_selected_profile(self):
        rec = self._get_selected_profile_record()
        if rec is None:
            messagebox.showwarning(
                self.tr("btn_profile_load"), self.tr("msg_profile_select_required")
            )
            return
        self._load_profile_record(rec, confirm_dirty=True, show_msg=False)

    def _save_current_to_selected_profile(self):
        rec = self._get_selected_profile_record()
        if rec is None:
            messagebox.showwarning(
                self.tr("btn_profile_save"), self.tr("msg_profile_select_required")
            )
            return
        profile_path = rec.get("abs_path", "")
        if not profile_path:
            return
        payload = self._compose_config_from_ui(
            self._load_config_for_write(), scope="all"
        )
        records = self._load_profile_records()
        try:
            with open(profile_path, "w", encoding="utf-8") as f:
                json.dump(payload, f, indent=4, ensure_ascii=False)
            for item in records:
                if item.get("file", "") == rec.get("file", ""):
                    item["updated_at"] = self._profile_timestamp()
                    break
            self._save_profile_records(records)
            self._refresh_profile_tree(select_file=rec.get("file", ""))
            msg = self.tr("msg_profile_save_ok").format(rec.get("name", ""))
            if hasattr(self, "var_status"):
                self.var_status.set(msg)
            if hasattr(self, "var_locator_result"):
                self.var_locator_result.set(msg)
        except Exception as e:
            messagebox.showerror(
                self.tr("btn_profile_save"), self.tr("msg_save_fail").format(e)
            )

    def _rename_selected_profile(self):
        rec = self._get_selected_profile_record()
        if rec is None:
            messagebox.showwarning(
                self.tr("btn_profile_rename"), self.tr("msg_profile_select_required")
            )
            return
        parent = (
            self.profile_manager_win
            if self.profile_manager_win is not None
            and self.profile_manager_win.winfo_exists()
            else self
        )
        new_name = simpledialog.askstring(
            self.tr("btn_profile_rename"),
            self.tr("msg_profile_rename_prompt"),
            parent=parent,
            initialvalue=rec.get("name", ""),
        )
        if new_name is None:
            return
        new_name = str(new_name).strip()
        if not new_name:
            messagebox.showwarning(
                self.tr("btn_profile_rename"), self.tr("msg_profile_name_required")
            )
            return
        records = self._load_profile_records()
        old_file = rec.get("file", "")
        new_file = self._next_profile_filename(new_name, records, exclude_file=old_file)
        old_path = self._profile_abs_path(old_file)
        new_path = self._profile_abs_path(new_file)
        try:
            if new_file != old_file:
                os.replace(old_path, new_path)
            for item in records:
                if item.get("file", "") == old_file:
                    item["name"] = new_name
                    item["file"] = new_file
                    item["updated_at"] = self._profile_timestamp()
                    break
            self._save_profile_records(records)
            if os.path.abspath(self.config_path) == os.path.abspath(old_path):
                self.config_path = new_path
                self.var_profile_active_path.set(self.config_path)
                if hasattr(self, "var_config_path"):
                    self.var_config_path.set(self.config_path)
            self._refresh_profile_tree(select_file=new_file)
            msg = self.tr("msg_profile_rename_ok").format(new_name)
            if hasattr(self, "var_status"):
                self.var_status.set(msg)
            if hasattr(self, "var_locator_result"):
                self.var_locator_result.set(msg)
        except Exception as e:
            messagebox.showerror(
                self.tr("btn_profile_rename"), self.tr("msg_save_fail").format(e)
            )

    def _edit_selected_profile_note(self):
        rec = self._get_selected_profile_record()
        if rec is None:
            messagebox.showwarning(
                self.tr("btn_profile_note"), self.tr("msg_profile_select_required")
            )
            return
        parent = (
            self.profile_manager_win
            if self.profile_manager_win is not None
            and self.profile_manager_win.winfo_exists()
            else self
        )
        note = simpledialog.askstring(
            self.tr("btn_profile_note"),
            self.tr("msg_profile_note_prompt"),
            parent=parent,
            initialvalue=rec.get("note", ""),
        )
        if note is None:
            return
        records = self._load_profile_records()
        for item in records:
            if item.get("file", "") == rec.get("file", ""):
                item["note"] = str(note).strip()
                item["updated_at"] = self._profile_timestamp()
                break
        try:
            self._save_profile_records(records)
            self._refresh_profile_tree(select_file=rec.get("file", ""))
            msg = self.tr("msg_profile_note_ok")
            if hasattr(self, "var_status"):
                self.var_status.set(msg)
            if hasattr(self, "var_locator_result"):
                self.var_locator_result.set(msg)
        except Exception as e:
            messagebox.showerror(
                self.tr("btn_profile_note"), self.tr("msg_save_fail").format(e)
            )

    def _delete_selected_profile(self):
        rec = self._get_selected_profile_record()
        if rec is None:
            messagebox.showwarning(
                self.tr("btn_profile_delete"), self.tr("msg_profile_select_required")
            )
            return
        profile_path = rec.get("abs_path", "")
        if os.path.abspath(profile_path) == os.path.abspath(self.config_path):
            messagebox.showwarning(
                self.tr("btn_profile_delete"),
                self.tr("msg_profile_delete_active_blocked"),
            )
            return
        confirm = messagebox.askyesno(
            self.tr("btn_profile_delete"),
            self.tr("msg_profile_delete_confirm").format(rec.get("name", "")),
        )
        if not confirm:
            return
        records = self._load_profile_records()
        try:
            if os.path.exists(profile_path):
                os.remove(profile_path)
            records = [
                item for item in records if item.get("file", "") != rec.get("file", "")
            ]
            self._save_profile_records(records)
            self._refresh_profile_tree()
            msg = self.tr("msg_profile_delete_ok").format(rec.get("name", ""))
            if hasattr(self, "var_status"):
                self.var_status.set(msg)
            if hasattr(self, "var_locator_result"):
                self.var_locator_result.set(msg)
        except Exception as e:
            messagebox.showerror(
                self.tr("btn_profile_delete"), self.tr("msg_save_fail").format(e)
            )

    def _open_profiles_folder(self):
        os.makedirs(self._profiles_dir(), exist_ok=True)
        self._open_path(self._profiles_dir())

    def _close_save_profile_dialog(self):
        if (
            self.save_profile_dialog is not None
            and self.save_profile_dialog.winfo_exists()
        ):
            try:
                if self.save_profile_dialog.grab_current() == self.save_profile_dialog:
                    self.save_profile_dialog.grab_release()
            except Exception:
                pass
            try:
                self.save_profile_dialog.destroy()
            except Exception:
                pass
        self.save_profile_dialog = None

    def _place_dialog_in_main(self, dialog, width, height):
        try:
            self.update_idletasks()
            base_x = int(self.winfo_rootx())
            base_y = int(self.winfo_rooty())
            base_w = max(int(self.winfo_width()), 200)
            base_h = max(int(self.winfo_height()), 200)
            x = base_x + max((base_w - int(width)) // 2, 20)
            y = base_y + max((base_h - int(height)) // 2, 20)
            dialog.geometry(f"{int(width)}x{int(height)}+{x}+{y}")
        except Exception:
            pass

    def open_save_profile_dialog(self):
        if self._ui_running:
            return
        if (
            self.save_profile_dialog is not None
            and self.save_profile_dialog.winfo_exists()
        ):
            self._place_dialog_in_main(self.save_profile_dialog, 520, 220)
            self.save_profile_dialog.lift()
            self.save_profile_dialog.focus_force()
            return
        dlg = tk.Toplevel(self)
        dlg.title(self.tr("win_save_config"))
        self._place_dialog_in_main(dlg, 520, 220)
        dlg.minsize(460, 210)
        dlg.transient(self)
        dlg.grab_set()
        dlg.lift()
        dlg.protocol("WM_DELETE_WINDOW", self._close_save_profile_dialog)
        self.save_profile_dialog = dlg

        frame = tb.Frame(dlg, padding=12)
        frame.pack(fill=BOTH, expand=YES)
        tb.Label(
            frame, text=self.tr("lbl_profile_name"), font=("System", 9, "bold")
        ).pack(anchor="w")
        self.var_save_profile_name = tk.StringVar(
            value=self._build_default_profile_name()
        )
        ent_name = tb.Entry(frame, textvariable=self.var_save_profile_name)
        ent_name.pack(fill=X, pady=(2, 10))
        ent_name.focus_set()
        ent_name.selection_range(0, END)

        tb.Label(
            frame, text=self.tr("lbl_profile_note"), font=("System", 9, "bold")
        ).pack(anchor="w")
        self.var_save_profile_note = tk.StringVar(value="")
        ent_note = tb.Entry(frame, textvariable=self.var_save_profile_note)
        ent_note.pack(fill=X, pady=(2, 12))

        btn_row = tb.Frame(frame)
        btn_row.pack(fill=X)
        self.btn_save_profile_confirm = tb.Button(
            btn_row,
            text=self.tr("btn_confirm_save_profile"),
            command=self._confirm_save_profile_dialog,
            bootstyle="success",
            width=14,
        )
        self.btn_save_profile_confirm.pack(side=LEFT)
        self.btn_save_profile_cancel = tb.Button(
            btn_row,
            text=self.tr("btn_cancel"),
            command=self._close_save_profile_dialog,
            bootstyle="secondary-outline",
            width=10,
        )
        self.btn_save_profile_cancel.pack(side=LEFT, padx=(8, 0))
        self._update_profile_dialog_controls()

    def _confirm_save_profile_dialog(self):
        name = (
            self.var_save_profile_name.get().strip()
            if hasattr(self, "var_save_profile_name")
            else ""
        )
        note = (
            self.var_save_profile_note.get().strip()
            if hasattr(self, "var_save_profile_note")
            else ""
        )
        if not name:
            messagebox.showwarning(
                self.tr("btn_save_cfg"), self.tr("msg_profile_name_required")
            )
            return
        if self._save_profile_with_meta(name, note, show_msg=True):
            self._close_save_profile_dialog()

    def _close_load_profile_dialog(self):
        if (
            self.load_profile_dialog is not None
            and self.load_profile_dialog.winfo_exists()
        ):
            try:
                if self.load_profile_dialog.grab_current() == self.load_profile_dialog:
                    self.load_profile_dialog.grab_release()
            except Exception:
                pass
            try:
                self.load_profile_dialog.destroy()
            except Exception:
                pass
        self.load_profile_dialog = None
        self.load_profile_tree = None
        self._load_profile_tree_rows = {}

    def _get_selected_load_profile_record(self):
        if self.load_profile_tree is None or not self.load_profile_tree.winfo_exists():
            return None
        selection = self.load_profile_tree.selection()
        if not selection:
            return None
        return self._load_profile_tree_rows.get(selection[0])

    def _refresh_load_profile_tree(self, select_file=None, select_abs_path=None):
        if self.load_profile_tree is None or not self.load_profile_tree.winfo_exists():
            return
        records = self._load_profile_records() + self._load_builtin_config_records()
        self._load_profile_tree_rows = {}
        self.load_profile_tree.delete(*self.load_profile_tree.get_children())
        target_file = str(select_file or "").strip()
        target_abs = os.path.abspath(str(select_abs_path or "").strip())
        active_path = os.path.abspath(self.config_path)
        target_iid = None
        for rec in records:
            rec_abs = str(rec.get("abs_path", "")).strip()
            if rec_abs:
                abs_path = os.path.abspath(rec_abs)
            else:
                abs_path = os.path.abspath(self._profile_abs_path(rec.get("file", "")))
            iid = self.load_profile_tree.insert(
                "",
                "end",
                values=(
                    rec.get("name", ""),
                    rec.get("file", ""),
                    rec.get("note", ""),
                    rec.get("updated_at", ""),
                ),
            )
            row = dict(rec)
            row["abs_path"] = abs_path
            self._load_profile_tree_rows[iid] = row
            if target_abs and abs_path == target_abs:
                target_iid = iid
            elif target_file and rec.get("file", "") == target_file:
                target_iid = iid
            elif not target_file and abs_path == active_path:
                target_iid = iid
        if target_iid is not None:
            self.load_profile_tree.selection_set(target_iid)
            self.load_profile_tree.focus(target_iid)
            self.load_profile_tree.see(target_iid)
        self._update_profile_dialog_controls()

    def open_load_profile_dialog(self):
        if self._ui_running:
            return
        if (
            self.load_profile_dialog is not None
            and self.load_profile_dialog.winfo_exists()
        ):
            self._place_dialog_in_main(self.load_profile_dialog, 860, 460)
            self.load_profile_dialog.lift()
            self.load_profile_dialog.focus_force()
            self._refresh_load_profile_tree()
            return
        dlg = tk.Toplevel(self)
        dlg.title(self.tr("win_load_config"))
        self._place_dialog_in_main(dlg, 860, 460)
        dlg.minsize(760, 380)
        dlg.transient(self)
        dlg.grab_set()
        dlg.lift()
        dlg.protocol("WM_DELETE_WINDOW", self._close_load_profile_dialog)
        self.load_profile_dialog = dlg

        root = tb.Frame(dlg, padding=10)
        root.pack(fill=BOTH, expand=YES)
        tb.Label(
            root, text=self.tr("lbl_profile_select"), font=("System", 9, "bold")
        ).pack(anchor="w", pady=(0, 6))

        tree_frame = tb.Frame(root)
        tree_frame.pack(fill=BOTH, expand=YES)
        cols = ("name", "file", "note", "updated")
        self.load_profile_tree = ttk.Treeview(
            tree_frame, columns=cols, show="headings", selectmode="browse"
        )
        self.load_profile_tree.heading("name", text=self.tr("col_profile_name"))
        self.load_profile_tree.heading("file", text=self.tr("col_profile_file"))
        self.load_profile_tree.heading("note", text=self.tr("col_profile_note"))
        self.load_profile_tree.heading("updated", text=self.tr("col_profile_updated"))
        self.load_profile_tree.column("name", width=180, anchor="w")
        self.load_profile_tree.column("file", width=220, anchor="w")
        self.load_profile_tree.column("note", width=300, anchor="w")
        self.load_profile_tree.column("updated", width=140, anchor="center")
        scr_y = tb.Scrollbar(
            tree_frame, orient=VERTICAL, command=self.load_profile_tree.yview
        )
        scr_x = tb.Scrollbar(
            tree_frame, orient=HORIZONTAL, command=self.load_profile_tree.xview
        )
        self.load_profile_tree.configure(
            yscrollcommand=scr_y.set, xscrollcommand=scr_x.set
        )
        self.load_profile_tree.pack(side=LEFT, fill=BOTH, expand=YES)
        scr_y.pack(side=RIGHT, fill=Y)
        scr_x.pack(side=BOTTOM, fill=X)
        self.load_profile_tree.bind(
            "<<TreeviewSelect>>", lambda _e: self._update_profile_dialog_controls()
        )

        btn_row = tb.Frame(root)
        btn_row.pack(fill=X, pady=(8, 0))
        self.btn_load_profile_confirm = tb.Button(
            btn_row,
            text=self.tr("btn_confirm_load_profile"),
            command=self._confirm_load_profile_dialog,
            bootstyle="primary",
            width=14,
        )
        self.btn_load_profile_confirm.pack(side=LEFT)
        self.btn_load_profile_refresh = tb.Button(
            btn_row,
            text=self.tr("btn_profile_refresh"),
            command=self._refresh_load_profile_tree,
            bootstyle="info-outline",
            width=10,
        )
        self.btn_load_profile_refresh.pack(side=LEFT, padx=(8, 0))
        self.btn_load_profile_cancel = tb.Button(
            btn_row,
            text=self.tr("btn_cancel"),
            command=self._close_load_profile_dialog,
            bootstyle="secondary-outline",
            width=10,
        )
        self.btn_load_profile_cancel.pack(side=RIGHT)

        self._refresh_load_profile_tree()
        self._update_profile_dialog_controls()

    def _confirm_load_profile_dialog(self):
        rec = self._get_selected_load_profile_record()
        if rec is None:
            messagebox.showwarning(
                self.tr("btn_load_cfg"),
                self.tr("msg_profile_select_required"),
            )
            return
        if self._load_profile_record(rec, confirm_dirty=True, show_msg=True):
            self._close_load_profile_dialog()

    def _update_profile_dialog_controls(self):
        state_base = "disabled" if self._ui_running else "normal"
        if (
            self.save_profile_dialog is not None
            and self.save_profile_dialog.winfo_exists()
        ):
            for btn_name in ("btn_save_profile_confirm", "btn_save_profile_cancel"):
                btn = getattr(self, btn_name, None)
                if btn is not None and btn.winfo_exists():
                    btn.configure(state=state_base)
        if (
            self.load_profile_dialog is not None
            and self.load_profile_dialog.winfo_exists()
        ):
            selected = self._get_selected_load_profile_record() is not None
            confirm_state = (
                "disabled" if (self._ui_running or not selected) else "normal"
            )
            btn = getattr(self, "btn_load_profile_confirm", None)
            if btn is not None and btn.winfo_exists():
                btn.configure(state=confirm_state)
            for btn_name in ("btn_load_profile_refresh", "btn_load_profile_cancel"):
                btn = getattr(self, btn_name, None)
                if btn is not None and btn.winfo_exists():
                    btn.configure(state=state_base)

