# -*- coding: utf-8 -*-
"""Task workflow methods extracted from OfficeGUI to keep office_gui.py maintainable."""

import json
import os
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk
from tkinter.constants import *

from office_converter import (
    MODE_CONVERT_ONLY,
    MODE_MERGE_ONLY,
    MODE_CONVERT_THEN_MERGE,
    MODE_COLLECT_ONLY,
    MODE_MSHELP_ONLY,
    COLLECT_MODE_COPY_AND_INDEX,
    MERGE_MODE_CATEGORY,
    MERGE_MODE_ALL_IN_ONE,
    ENGINE_WPS,
    KILL_MODE_AUTO,
)
from task_manager import (
    TASK_BINDING_ACTIVE,
    TASK_BINDING_PROFILE,
    TASK_BINDING_SNAPSHOT,
    build_task_runtime_config,
    normalize_task_binding_mode,
)


class TaskWorkflowMixin:
    def _resolve_text_core_widget(self, widget):
        if widget is None:
            return None
        core = getattr(widget, "text", None)
        return core if core is not None else widget

    def _set_text_widget_content(self, widget, text):
        core = self._resolve_text_core_widget(widget)
        if core is None:
            return
        try:
            core.configure(state="normal")
        except (tk.TclError, AttributeError, RuntimeError):
            pass
        try:
            core.delete("1.0", tk.END)
            core.insert(tk.END, text)
        except (tk.TclError, AttributeError, RuntimeError):
            return
        try:
            core.configure(state="disabled")
        except (tk.TclError, AttributeError, RuntimeError):
            pass

    def _get_selected_task_id(self):
        if not hasattr(self, "tree_tasks"):
            return None
        sel = self.tree_tasks.selection()
        if not sel:
            return None
        return str(sel[0])

    def _get_selected_task_ids(self):
        """Return list of selected task IDs in tree order (for batch run)."""
        if not hasattr(self, "tree_tasks"):
            return []
        sel = self.tree_tasks.selection()
        return [str(iid) for iid in sel]

    def _short_path(self, path, max_len=36):
        p = str(path or "").strip()
        if len(p) <= max_len:
            return p
        return "..." + p[-(max_len - 3) :]

    def _report_nonfatal_ui_error(self, scope, exc=None, detail=""):
        scope_text = str(scope or "ui").strip() or "ui"
        message = str(detail or "").strip()
        if not message and exc is not None:
            message = str(exc)
        if not message:
            message = "(no detail)"

        record = {
            "at": datetime.now().isoformat(timespec="seconds"),
            "scope": scope_text,
            "message": message,
        }
        errors = getattr(self, "_ui_nonfatal_errors", None)
        if not isinstance(errors, list):
            errors = []
            self._ui_nonfatal_errors = errors
        errors.append(record)
        if len(errors) > 50:
            del errors[:-50]

        try:
            q = getattr(self, "log_queue", None)
            if q is not None:
                q.put(f"[UI-WARN] {scope_text}: {message}")
        except (AttributeError, TypeError, ValueError, RuntimeError):
            pass
        return record

    def _task_list_filter_text(self):
        try:
            var = getattr(self, "var_task_filter_text", None)
            return str(var.get() if var is not None else "").strip().lower()
        except (AttributeError, TypeError, ValueError, tk.TclError):
            return ""

    def _task_list_status_filter(self):
        try:
            var = getattr(self, "var_task_status_filter", None)
            value = str(var.get() if var is not None else "").strip().lower()
            return value or "all"
        except (AttributeError, TypeError, ValueError, tk.TclError):
            return "all"

    def _task_list_sort_by(self):
        try:
            var = getattr(self, "var_task_sort_by", None)
            value = str(var.get() if var is not None else "").strip().lower()
            return value or "updated_desc"
        except (AttributeError, TypeError, ValueError, tk.TclError):
            return "updated_desc"

    def _task_scope_current_config_only(self):
        try:
            var = getattr(self, "var_task_scope_current_config_only", None)
            if var is None:
                return False
            return bool(int(var.get()))
        except (AttributeError, TypeError, ValueError, tk.TclError):
            return False

    def _task_matches_current_config_scope(self, task):
        if not self._task_scope_current_config_only():
            return True
        if not isinstance(task, dict):
            return False

        current_config = self._safe_abs_path(getattr(self, "config_path", ""))
        if not current_config:
            return True

        task_cfg_path = self._safe_abs_path(task.get("config_snapshot_path", ""))
        if task_cfg_path:
            return task_cfg_path == current_config

        mode = normalize_task_binding_mode(task.get("config_binding_mode"))
        if mode == TASK_BINDING_ACTIVE:
            return True

        task_snapshot_sig = self._normalize_config_for_compare(
            task.get("project_config_snapshot")
        )
        if not task_snapshot_sig:
            return False
        loader = getattr(self, "_load_config_for_write", None)
        if not callable(loader):
            return False
        try:
            current_cfg_sig = self._normalize_config_for_compare(loader())
        except (AttributeError, TypeError, ValueError, OSError, RuntimeError) as e:
            self._report_nonfatal_ui_error("task.scope.current_cfg", exc=e)
            return False
        return bool(current_cfg_sig and current_cfg_sig == task_snapshot_sig)

    def _refresh_task_status_filter_values(self, tasks):
        cb = getattr(self, "cb_task_status_filter", None)
        if cb is None:
            return
        statuses = []
        for task in tasks or []:
            if not isinstance(task, dict):
                continue
            status = str(task.get("status", "")).strip().lower()
            if status and status not in statuses:
                statuses.append(status)
        values = ["all"] + sorted(statuses)
        try:
            current = tuple(cb.cget("values"))
        except (tk.TclError, AttributeError, TypeError, ValueError):
            current = tuple()
        if current != tuple(values):
            try:
                cb.configure(values=values)
            except (tk.TclError, AttributeError, TypeError, ValueError):
                try:
                    cb["values"] = values
                except (tk.TclError, KeyError, TypeError, ValueError):
                    return
        var = getattr(self, "var_task_status_filter", None)
        if var is None:
            return
        try:
            cur = str(var.get() or "").strip().lower()
        except (AttributeError, TypeError, ValueError, tk.TclError):
            cur = ""
        if cur not in values:
            try:
                var.set("all")
            except (AttributeError, TypeError, ValueError, tk.TclError):
                pass

    def _filter_task_list_rows(self, tasks):
        rows = [t for t in (tasks or []) if isinstance(t, dict)]
        query = self._task_list_filter_text()
        status_filter = self._task_list_status_filter()
        out = []
        for task in rows:
            if not self._task_matches_current_config_scope(task):
                continue
            if status_filter != "all":
                row_status = str(task.get("status", "idle")).strip().lower()
                if row_status != status_filter:
                    continue
            if query:
                haystack = " ".join(
                    [
                        str(task.get("id", "")),
                        str(task.get("name", "")),
                        str(task.get("source_folder", "")),
                        str(task.get("target_folder", "")),
                        str(task.get("status", "")),
                    ]
                ).lower()
                if query not in haystack:
                    continue
            out.append(task)
        return out

    def _sort_task_list_rows(self, tasks):
        rows = [dict(t) for t in (tasks or []) if isinstance(t, dict)]
        sort_by = self._task_list_sort_by()
        if sort_by == "name_asc":
            rows.sort(key=lambda t: str(t.get("name", "")).lower())
            return rows
        if sort_by == "name_desc":
            rows.sort(key=lambda t: str(t.get("name", "")).lower(), reverse=True)
            return rows
        if sort_by == "last_run_desc":
            rows.sort(key=lambda t: str(t.get("last_run_at", "") or ""), reverse=True)
            return rows
        if sort_by == "status_name":
            rows.sort(
                key=lambda t: (
                    str(t.get("status", "idle")).lower(),
                    str(t.get("name", "")).lower(),
                )
            )
            return rows
        rows.sort(
            key=lambda t: str(
                t.get("updated_at") or t.get("last_run_at") or t.get("created_at") or ""
            ),
            reverse=True,
        )
        return rows

    def _reset_task_list_filters(self):
        try:
            if hasattr(self, "var_task_filter_text"):
                self.var_task_filter_text.set("")
            if hasattr(self, "var_task_status_filter"):
                self.var_task_status_filter.set("all")
            if hasattr(self, "var_task_sort_by"):
                self.var_task_sort_by.set("updated_desc")
        except (AttributeError, TypeError, ValueError, tk.TclError) as e:
            self._report_nonfatal_ui_error("task.reset_filters", exc=e)
        self._refresh_task_list_ui()

    def _refresh_task_list_ui(self):
        if not hasattr(self, "tree_tasks"):
            return
        selected_id = self._get_selected_task_id()
        for iid in self.tree_tasks.get_children():
            self.tree_tasks.delete(iid)
        tasks = []
        try:
            tasks = self.task_store.list_tasks()
        except (AttributeError, TypeError, ValueError, OSError, RuntimeError) as e:
            self._report_nonfatal_ui_error("task.list_tasks", exc=e)
            tasks = []
        scoped_tasks = [t for t in tasks if self._task_matches_current_config_scope(t)]
        self._refresh_task_status_filter_values(scoped_tasks)
        tasks = self._filter_task_list_rows(tasks)
        tasks = self._sort_task_list_rows(tasks)
        for task in tasks:
            task_id = str(task.get("id", ""))
            name = str(task.get("name", ""))[:48]
            source = self._short_path(task.get("source_folder", ""))
            target = self._short_path(task.get("target_folder", ""))
            try:
                full_task = self.task_store.get_task(task_id) or task
            except (AttributeError, TypeError, ValueError, OSError, RuntimeError) as e:
                self._report_nonfatal_ui_error("task.get_task", exc=e, detail=task_id)
                full_task = task
            binding = self._summarize_task_config_binding(full_task, runtime_preview={})
            binding_name = str(binding.get("display_name", "") or "-")
            binding_relation = str(binding.get("relation_label", "") or "-")
            status = str(task.get("status", "idle"))
            last_run = (task.get("last_run_at") or "")[:16]
            try:
                self.tree_tasks.insert(
                    "",
                    END,
                    iid=task_id,
                    values=(
                        name,
                        source,
                        target,
                        binding_name,
                        binding_relation,
                        status,
                        last_run,
                    ),
                )
            except (tk.TclError, AttributeError, TypeError, ValueError, RuntimeError) as e:
                self._report_nonfatal_ui_error("task.tree_insert", exc=e, detail=task_id)
        if selected_id and self.tree_tasks.exists(selected_id):
            self.tree_tasks.selection_set(selected_id)
            self.tree_tasks.focus(selected_id)
        try:
            self._on_task_select()
        except (AttributeError, TypeError, ValueError, OSError, RuntimeError) as e:
            self._report_nonfatal_ui_error("task.on_select", exc=e)

    def _strip_task_runtime_meta(self, cfg):
        if not isinstance(cfg, dict):
            return {}
        out = {}
        for k, v in cfg.items():
            if str(k).startswith("_task_"):
                continue
            out[k] = v
        return out

    def _normalize_config_for_compare(self, cfg):
        if not isinstance(cfg, dict):
            return ""
        try:
            return json.dumps(cfg, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
        except (TypeError, ValueError):
            return ""

    def _find_profile_record_by_path(self, config_path):
        target = str(config_path or "").strip()
        if not target:
            return None
        target_abs = os.path.abspath(target)
        try:
            records = self._load_profile_records()
        except (AttributeError, TypeError, ValueError, OSError):
            return None
        for rec in records or []:
            try:
                rec_abs = self._profile_record_abs_path(rec)
            except (AttributeError, TypeError, ValueError, OSError):
                continue
            if rec_abs and rec_abs == target_abs:
                row = dict(rec)
                row["abs_path"] = rec_abs
                return row
        return None

    def _build_task_config_binding_meta(self):
        meta = {
            "config_binding_mode": TASK_BINDING_ACTIVE,
            "config_snapshot_path": os.path.abspath(str(self.config_path or "").strip()),
        }
        rec = self._find_profile_record_by_path(self.config_path)
        if isinstance(rec, dict):
            meta["config_snapshot_profile_name"] = str(rec.get("name", "")).strip()
            meta["config_snapshot_profile_file"] = str(rec.get("file", "")).strip()
        return meta

    def _task_runtime_source_desc(self, runtime_preview):
        cfg = runtime_preview if isinstance(runtime_preview, dict) else {}
        config_source = str(cfg.get("_task_config_source", "unknown"))
        return {
            "project_config(active)": "当前活动配置",
            "task.profile_config": "绑定配置档",
            "task.project_config_snapshot(fallback)": "任务配置快照(回退)",
            "task.project_config_snapshot": "任务配置快照",
            "task.runtime_config_snapshot(fallback)": "任务运行快照(回退)",
            "task.runtime_config_snapshot": "任务运行快照(回退)",
        }.get(config_source, config_source)

    def _safe_abs_path(self, path):
        text = str(path or "").strip()
        if not text:
            return ""
        try:
            return os.path.abspath(text)
        except (TypeError, ValueError, OSError):
            return text

    def _profile_record_abs_path(self, record):
        if not isinstance(record, dict):
            return ""
        direct = self._safe_abs_path(record.get("abs_path", ""))
        if direct:
            return direct
        file_name = str(record.get("file", "")).strip()
        if not file_name:
            return ""
        try:
            if hasattr(self, "_profile_abs_path"):
                return self._safe_abs_path(self._profile_abs_path(file_name))
        except (AttributeError, TypeError, ValueError, OSError):
            pass
        script_dir = str(getattr(self, "script_dir", "")).strip()
        if not script_dir:
            return ""
        return self._safe_abs_path(
            os.path.join(script_dir, "config_profiles", os.path.basename(file_name))
        )

    def _summarize_task_config_binding(self, task, runtime_preview=None):
        task = task if isinstance(task, dict) else {}
        binding_mode = normalize_task_binding_mode(task.get("config_binding_mode"))
        task_bound = self._resolve_task_bound_profile(task)
        config_path = ""
        profile_name = ""
        profile_file = ""
        match_mode = str(task_bound.get("match_mode", "unknown")).strip() or "unknown"
        active_config_path = self._safe_abs_path(getattr(self, "config_path", ""))

        if binding_mode == TASK_BINDING_ACTIVE:
            config_path = active_config_path
            rec = self._find_profile_record_by_path(config_path)
            active_label = str(getattr(self, "_active_config_label", "")).strip()
            if isinstance(rec, dict):
                profile_name = str(rec.get("name", "")).strip()
                profile_file = str(rec.get("file", "")).strip()
                match_mode = "active_config"
            if profile_file:
                display_name = profile_file
            elif active_label:
                display_name = active_label
            elif config_path:
                display_name = os.path.basename(config_path)
            elif profile_name:
                display_name = profile_name
            else:
                display_name = "(当前活动配置)"
            relation_label = "跟随当前活动配置"
        elif binding_mode == TASK_BINDING_PROFILE:
            config_path = str(task_bound.get("config_path", "")).strip() or str(
                task.get("config_snapshot_path", "")
            ).strip()
            profile_name = str(task_bound.get("profile_name", "")).strip() or str(
                task.get("config_snapshot_profile_name", "")
            ).strip()
            profile_file = str(task_bound.get("profile_file", "")).strip() or str(
                task.get("config_snapshot_profile_file", "")
            ).strip()
            if profile_name:
                display_name = profile_name
            elif profile_file:
                display_name = profile_file
            elif config_path:
                display_name = os.path.basename(config_path)
            else:
                display_name = "(未绑定配置档)"
            relation_label = "绑定指定配置"
        else:
            config_path = str(task_bound.get("config_path", "")).strip() or str(
                task.get("config_snapshot_path", "")
            ).strip()
            profile_name = str(task_bound.get("profile_name", "")).strip() or str(
                task.get("config_snapshot_profile_name", "")
            ).strip()
            profile_file = str(task_bound.get("profile_file", "")).strip() or str(
                task.get("config_snapshot_profile_file", "")
            ).strip()
            if isinstance(task.get("project_config_snapshot"), dict):
                display_name = "(任务快照)"
            elif profile_name:
                display_name = profile_name
            elif profile_file:
                display_name = profile_file
            elif config_path:
                display_name = os.path.basename(config_path)
            else:
                display_name = "(任务快照缺失)"
            relation_label = "使用任务快照"

        return {
            "binding_mode": binding_mode,
            "display_name": display_name,
            "config_path": config_path,
            "profile_name": profile_name,
            "profile_file": profile_file,
            "relation_label": relation_label,
            "runtime_source_desc": self._task_runtime_source_desc(runtime_preview),
            "match_mode": match_mode,
        }

    def _resolve_task_bound_profile(self, task):
        info = {
            "config_path": "",
            "profile_name": "",
            "profile_file": "",
            "match_mode": "unknown",
        }
        if not isinstance(task, dict):
            return info

        task_cfg_path = str(task.get("config_snapshot_path", "")).strip()
        if task_cfg_path:
            info["config_path"] = task_cfg_path
            rec = self._find_profile_record_by_path(task_cfg_path)
            if isinstance(rec, dict):
                info["profile_name"] = str(rec.get("name", "")).strip()
                info["profile_file"] = str(rec.get("file", "")).strip()
                info["match_mode"] = "task.config_snapshot_path"
                return info
            info["profile_name"] = str(task.get("config_snapshot_profile_name", "")).strip()
            info["profile_file"] = str(task.get("config_snapshot_profile_file", "")).strip()
            info["match_mode"] = "task.config_snapshot_path"
            return info

        snapshot = task.get("project_config_snapshot")
        snapshot_sig = self._normalize_config_for_compare(snapshot)
        if not snapshot_sig:
            return info
        matches = []
        try:
            records = self._load_profile_records()
        except (AttributeError, TypeError, ValueError, OSError):
            records = []
        for rec in records or []:
            abs_path = self._profile_record_abs_path(rec)
            if not abs_path or not os.path.isfile(abs_path):
                continue
            try:
                with open(abs_path, "r", encoding="utf-8") as f:
                    profile_cfg = json.load(f)
            except (OSError, UnicodeDecodeError, ValueError):
                continue
            if self._normalize_config_for_compare(profile_cfg) == snapshot_sig:
                row = dict(rec)
                row["abs_path"] = abs_path
                matches.append(row)
        if len(matches) == 1:
            rec = matches[0]
            info["config_path"] = str(rec.get("abs_path", "")).strip()
            info["profile_name"] = str(rec.get("name", "")).strip()
            info["profile_file"] = str(rec.get("file", "")).strip()
            info["match_mode"] = "project_config_snapshot==profile"
        elif len(matches) > 1:
            info["match_mode"] = "project_config_snapshot==multiple_profiles"
        return info

    def _ensure_task_config_snapshots(self, task, project_cfg=None, persist=True):
        if not isinstance(task, dict):
            return task, {}
        changed = False
        base_cfg = (
            project_cfg if isinstance(project_cfg, dict) else self._load_config_for_write()
        )

        if not isinstance(task.get("project_config_snapshot"), dict):
            task["project_config_snapshot"] = dict(base_cfg)
            changed = True

        runtime_cfg = build_task_runtime_config(
            base_cfg,
            task,
            force_full_rebuild=False,
            prefer_runtime_snapshot=False,
        )
        runtime_snapshot = self._strip_task_runtime_meta(runtime_cfg)
        if task.get("runtime_config_snapshot") != runtime_snapshot:
            task["runtime_config_snapshot"] = runtime_snapshot
            task["runtime_config_snapshot_updated_at"] = datetime.now().isoformat(
                timespec="seconds"
            )
            changed = True

        if changed and persist:
            try:
                saved = self.task_store.save_task(task)
                if isinstance(saved, dict):
                    task = saved
            except (AttributeError, TypeError, ValueError, OSError, RuntimeError) as e:
                self._report_nonfatal_ui_error("task.save_runtime_snapshot", exc=e)

        runtime_cfg = build_task_runtime_config(
            base_cfg,
            task,
            force_full_rebuild=False,
            prefer_runtime_snapshot=False,
        )
        return task, runtime_cfg

    def _on_task_select(self):
        task_id = self._get_selected_task_id()
        task = self.task_store.get_task(task_id) if task_id else None
        if not task:
            if hasattr(self, "txt_task_detail"):
                self._set_text_widget_content(
                    self.txt_task_detail, self.tr("msg_task_none_selected")
                )
            if hasattr(self, "btn_task_resume"):
                self.btn_task_resume.configure(state="disabled")
            self._update_task_tab_for_app_mode()
            return

        project_cfg = self._load_config_for_write()
        task, runtime_preview = self._ensure_task_config_snapshots(
            task, project_cfg=project_cfg, persist=True
        )
        cp = self.task_store.load_checkpoint(task_id)
        planned = len((cp or {}).get("planned_files", []) or [])
        done = len((cp or {}).get("completed_files", []) or [])
        run_mode = ""
        try:
            run_mode = str(runtime_preview.get("run_mode", ""))
        except (AttributeError, TypeError, ValueError):
            run_mode = str(task.get("config_overrides", {}).get("run_mode", ""))

        display_source = runtime_preview.get("source_folders") or task.get(
            "source_folders", []
        ) or [task.get("source_folder", "")]
        display_target = runtime_preview.get(
            "target_folder", task.get("target_folder", "")
        )
        binding = self._summarize_task_config_binding(task, runtime_preview)
        active_config_path = self._safe_abs_path(self.config_path)

        # 下方展示任务完整配置（非仅摘要）
        full_lines = [
            self.tr("msg_task_detail").format(
                task.get("name", ""),
                (display_source[0] if display_source else ""),
                display_target,
                run_mode or "-",
                "ON" if task.get("run_incremental", True) else "OFF",
                task.get("status", "idle"),
                done,
                planned,
            ),
            "",
            "--- 任务-配置映射 ---",
            "绑定模式: "
            + self._task_binding_mode_text(binding.get("binding_mode", "")),
            "任务绑定配置: " + str(binding.get("display_name", "")),
            "绑定配置路径: " + str(binding.get("config_path", "") or "-"),
            "与当前活动配置关系: " + str(binding.get("relation_label", "")),
            "任务运行配置来源: " + str(binding.get("runtime_source_desc", "")),
            "映射匹配方式: " + str(binding.get("match_mode", "")),
            "",
            "--- " + (self.tr("lbl_task_full_config") or "Task full config") + " ---",
            "name: " + str(task.get("name", "")),
            "description: "
            + str((task.get("description") or "").replace("\n", " ")[:200]),
            "source_folders: " + str(display_source),
            "target_folder: " + str(display_target),
            "run_incremental: " + str(task.get("run_incremental", True)),
            "has_project_config_snapshot: "
            + str(isinstance(task.get("project_config_snapshot"), dict)),
            "has_runtime_config_snapshot: "
            + str(isinstance(task.get("runtime_config_snapshot"), dict)),
            "config_source: " + str(binding.get("runtime_source_desc", "")),
            "active_config_path_now: " + str(active_config_path),
            "task_bound_config_path: " + str(binding.get("config_path", "")),
            "task_bound_profile_name: " + str(binding.get("profile_name", "")),
            "task_bound_profile_file: " + str(binding.get("profile_file", "")),
            "task_bound_profile_match: " + str(binding.get("match_mode", "")),
            "runtime_config_snapshot_updated_at: "
            + str(task.get("runtime_config_snapshot_updated_at") or ""),
            "task_file: " + str(self.task_store.task_path(task_id)),
            "status: " + str(task.get("status", "idle")),
            "last_run_at: " + str(task.get("last_run_at") or ""),
            "checkpoint: " + str(done) + " / " + str(planned),
        ]
        overrides = task.get("config_overrides") or {}
        if overrides:
            full_lines.append("")
            full_lines.append("config_overrides:")
            for k, v in sorted(overrides.items()):
                full_lines.append("  " + str(k) + ": " + str(v))
        if runtime_preview:
            full_lines.append("")
            full_lines.append("effective_runtime_config:")
            runtime_keys = [
                "_task_config_source",
                "run_mode",
                "collect_mode",
                "content_strategy",
                "default_engine",
                "enable_incremental_mode",
                "enable_merge",
                "output_enable_pdf",
                "output_enable_md",
                "output_enable_merged",
                "output_enable_independent",
                "enable_fast_md_engine",
                "enable_traceability_anchor_and_map",
                "enable_markdown_image_manifest",
                "enable_prompt_wrapper",
                "prompt_template_type",
                "short_id_prefix",
                "merge_mode",
                "merge_source",
                "max_merge_size_mb",
                "source_folders",
                "target_folder",
            ]
            for key in runtime_keys:
                full_lines.append("  " + key + ": " + str(runtime_preview.get(key)))
        full_text = "\n".join(full_lines)

        if hasattr(self, "txt_task_detail"):
            self._set_text_widget_content(self.txt_task_detail, full_text)
        if (
            runtime_preview
            and getattr(self, "var_app_mode", None)
            and self.var_app_mode.get() == "task"
        ):
            try:
                self._apply_task_runtime_to_ui(
                    runtime_preview, preserve_current_run_tab=True
                )
            except (AttributeError, TypeError, ValueError, tk.TclError, RuntimeError) as e:
                self._report_nonfatal_ui_error("task.apply_runtime_to_ui", exc=e)
        can_resume = bool(cp and done < planned)
        self.btn_task_resume.configure(state="normal" if can_resume else "disabled")
        self._update_task_tab_for_app_mode()

    def _update_task_tab_for_app_mode(self):
        """In classic mode hide task tab entirely; in task mode show it and enable Run/Resume (design 7.1)."""
        if not hasattr(self, "var_app_mode"):
            return
        is_classic = self.var_app_mode.get() == "classic"
        if hasattr(self, "main_notebook") and hasattr(self, "tab_run_tasks"):
            self._set_run_tab_state(
                self.tab_run_tasks, "hidden" if is_classic else "normal"
            )
        # 打开任务模式后任务管理标签页显示为绿色
        self._set_task_tab_highlight(not is_classic)
        if not hasattr(self, "btn_task_run"):
            return
        if is_classic:
            self.btn_task_run.configure(state="disabled")
            if hasattr(self, "btn_task_resume"):
                self.btn_task_resume.configure(state="disabled")
            if hasattr(self, "btn_task_batch_run"):
                self.btn_task_batch_run.configure(state="disabled")
            if hasattr(self, "btn_task_schedule"):
                self.btn_task_schedule.configure(state="disabled")
        else:
            self.btn_task_run.configure(state="normal")
            if hasattr(self, "btn_task_batch_run"):
                self.btn_task_batch_run.configure(state="normal")
            if hasattr(self, "btn_task_schedule"):
                self.btn_task_schedule.configure(state="normal")
            # resume state left as set by _on_task_select

    def _on_app_mode_change_for_task_tab(self):
        self._update_task_tab_for_app_mode()
        if getattr(self, "var_app_mode", None) and self.var_app_mode.get() == "task":
            self._on_task_select()

    def _normalize_task_cfg_for_compare(self, cfg):
        normalized = dict(cfg if isinstance(cfg, dict) else {})
        if "output_enable_md" not in normalized:
            normalized["output_enable_md"] = bool(
                normalized.get("enable_markdown", True)
            )
        normalized["enable_markdown"] = bool(normalized.get("output_enable_md", True))
        if normalized.get("run_mode") == MODE_CONVERT_THEN_MERGE:
            normalized["merge_source"] = "target"
        return normalized

    def _build_task_overrides_from_ui(self, project_cfg=None, only_diff=True):
        mode = self.var_run_mode.get()
        merge_source = (
            "target" if mode == MODE_CONVERT_THEN_MERGE else self.var_merge_source.get()
        )
        current = {
            "run_mode": mode,
            "collect_mode": self.var_collect_mode.get(),
            "content_strategy": self.var_strategy.get(),
            "default_engine": self.var_engine.get(),
            "kill_process_mode": self.var_kill_mode.get(),
            "enable_merge": bool(self.var_enable_merge.get()),
            "output_enable_pdf": bool(self.var_output_enable_pdf.get()),
            "output_enable_md": bool(self.var_output_enable_md.get()),
            "output_enable_merged": bool(self.var_output_enable_merged.get()),
            "output_enable_independent": bool(self.var_output_enable_independent.get()),
            "merge_convert_submode": self.var_merge_convert_submode.get(),
            "merge_mode": self.var_merge_mode.get(),
            "merge_source": merge_source,
            "enable_merge_index": bool(self.var_enable_merge_index.get()),
            "enable_merge_excel": bool(self.var_enable_merge_excel.get()),
            "enable_sandbox": bool(self.var_enable_sandbox.get()),
            "sandbox_min_free_gb": self._safe_positive_int(
                self.var_sandbox_min_free_gb.get(), 10
            ),
            "sandbox_low_space_policy": self.var_sandbox_low_space_policy.get()
            or "block",
            "enable_markdown_quality_report": bool(
                self.var_enable_markdown_quality_report.get()
            ),
            "enable_excel_json": bool(self.var_enable_excel_json.get()),
            "enable_chromadb_export": bool(self.var_enable_chromadb_export.get()),
            "enable_fast_md_engine": bool(self.var_enable_fast_md_engine.get()),
            "enable_traceability_anchor_and_map": bool(
                self.var_enable_traceability_anchor_and_map.get()
            ),
            "enable_markdown_image_manifest": bool(
                self.var_enable_markdown_image_manifest.get()
            ),
            "enable_prompt_wrapper": bool(self.var_enable_prompt_wrapper.get()),
            "prompt_template_type": str(
                self.var_prompt_template_type.get() or "new_solution"
            ),
            "short_id_prefix": str(self.var_short_id_prefix.get() or "ZW-")
            .strip()
            .upper()
            or "ZW-",
            "incremental_verify_hash": bool(self.var_incremental_verify_hash.get()),
            "incremental_reprocess_renamed": bool(
                self.var_incremental_reprocess_renamed.get()
            ),
            "source_priority_skip_same_name_pdf": bool(
                self.var_source_priority_skip_same_name_pdf.get()
            ),
            "global_md5_dedup": bool(self.var_global_md5_dedup.get()),
            "enable_update_package": bool(self.var_enable_update_package.get()),
            "enable_parallel_conversion": bool(
                self.var_enable_parallel_conversion.get()
            ),
            "parallel_workers": self._safe_positive_int(
                self.var_parallel_workers.get(), 4
            ),
            "enable_checkpoint": bool(self.var_enable_checkpoint.get()),
            "checkpoint_auto_resume": bool(self.var_checkpoint_auto_resume.get()),
        }
        current = self._normalize_task_cfg_for_compare(current)
        if not only_diff:
            return current

        base = self._normalize_task_cfg_for_compare(
            project_cfg
            if isinstance(project_cfg, dict)
            else self._load_config_for_write()
        )
        overrides = {}
        for key, value in current.items():
            if base.get(key) != value:
                overrides[key] = value
        return overrides

    def _new_task_id(self):
        return datetime.now().strftime("task_%Y%m%d_%H%M%S_%f")

    def _on_click_task_create(self):
        self._open_task_wizard()

    def _open_task_wizard(self):
        win = tk.Toplevel(self)
        win.title(self.tr("win_task_wizard"))
        win.minsize(680, 560)
        win.geometry("700x580")
        win.transient(self)
        try:
            win.configure(bg="#f0f0f0")
        except (tk.TclError, RuntimeError):
            pass
        data = {
            "name": "",
            "description": "",
            "source_folder": "",
            "source_folders": [],
            "target_folder": "",
            "run_incremental": True,
            "run_mode": MODE_CONVERT_THEN_MERGE,
            "output_enable_pdf": True,
            "output_enable_md": True,
            "output_enable_merged": True,
            "output_enable_independent": False,
            "output_choice": "both",
            "merge_how": MERGE_MODE_CATEGORY,
            "max_merge_size_mb": 80,
            "merge_filename_pattern": (
                getattr(self, "var_merge_filename_pattern", None)
                and self.var_merge_filename_pattern.get().strip()
                or "Merged_{category}_{timestamp}_{idx}"
            ),
        }
        win._wizard_data = data
        win._wizard_step = 1
        # 向导窗口内全部使用纯 tk 控件，避免 ttk/ttkbootstrap 在 Toplevel 上不渲染导致空白
        _bg = "#f0f0f0"
        nav = tk.Frame(win, bg=_bg, padx=10, pady=10)
        nav.pack(side=BOTTOM, fill=X)
        content_holder = tk.Frame(win, bg=_bg, padx=15, pady=15)
        content_holder.pack(fill=BOTH, expand=True)
        # 可滚动区域，避免第 3 步配置项被挡住
        canvas = tk.Canvas(content_holder, bg=_bg, highlightthickness=0)
        scrollbar = tk.Scrollbar(
            content_holder, orient="vertical", command=canvas.yview
        )
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        container = tk.Frame(canvas, bg=_bg)
        canvas_window = canvas.create_window(0, 0, window=container, anchor=tk.NW)

        def _on_container_configure(_ev=None):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def _on_canvas_configure(ev):
            if container.winfo_reqwidth() != ev.width:
                canvas.itemconfig(canvas_window, width=ev.width)

        container.bind("<Configure>", _on_container_configure)
        canvas.bind("<Configure>", _on_canvas_configure)

        def _on_mousewheel(ev):
            canvas.yview_scroll(int(-1 * (ev.delta / 120)), "units")

        def _unbind_mousewheel(e):
            if e.widget == win:
                try:
                    canvas.unbind_all("<MouseWheel>")
                except (tk.TclError, RuntimeError):
                    pass

        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        win.bind("<Destroy>", _unbind_mousewheel)
        f1 = tk.Frame(container, bg=_bg)
        f2 = tk.Frame(container, bg=_bg)
        f3 = tk.Frame(container, bg=_bg)
        f4 = tk.Frame(container, bg=_bg)
        tk.Label(
            f1, text=self.tr("wizard_step1"), font=("System", 11, "bold"), bg=_bg
        ).pack(anchor=W)
        tk.Label(f1, text=self.tr("msg_task_input_name"), bg=_bg).pack(anchor=W)
        ent_name = tk.Entry(f1, width=50)
        ent_name.pack(fill=X, pady=(0, 8))
        tk.Label(f1, text=self.tr("lbl_task_description"), bg=_bg).pack(anchor=W)
        txt_desc = tk.Text(f1, height=3, width=50)
        txt_desc.pack(fill=X, pady=(0, 8))
        tk.Label(
            f2, text=self.tr("wizard_step2"), font=("System", 11, "bold"), bg=_bg
        ).pack(anchor=W)
        tk.Label(f2, text=self.tr("lbl_source"), bg=_bg).pack(anchor=W)
        f2_src = tk.Frame(f2, bg=_bg)
        f2_src.pack(fill=X)
        lst_src = tk.Listbox(f2_src, height=4, selectmode=SINGLE, font=("Consolas", 9))
        lst_src.pack(side=LEFT, fill=BOTH, expand=True, padx=(0, 4))
        scr_src = tk.Scrollbar(f2_src, orient="vertical", command=lst_src.yview)
        scr_src.pack(side=LEFT, fill=Y)
        lst_src.configure(yscrollcommand=scr_src.set)
        f2_src_btns = tk.Frame(f2_src, bg=_bg)
        f2_src_btns.pack(side=LEFT, fill=Y)

        def add_src():
            # Ask user if they want to select multiple folders
            result = messagebox.askyesno(
                self.tr("msg_task_pick_source"),
                self.tr("msg_multi_select_folders"),
                icon="question",
                parent=win,
            )
            if result:
                # Multi-select mode - open multi-folder dialog
                self._open_task_multi_folder_dialog(win, lst_src)
            else:
                # Single-select mode
                p = filedialog.askdirectory(
                    title=self.tr("msg_task_pick_source"), parent=win
                )
                if p and p not in lst_src.get(0, END):
                    lst_src.insert(END, p)

        def remove_src():
            sel = lst_src.curselection()
            if sel:
                lst_src.delete(sel[0])

        tk.Button(f2_src_btns, text="+", command=add_src, width=3).pack(pady=2)
        tk.Button(f2_src_btns, text="-", command=remove_src, width=3).pack(pady=2)
        tk.Label(f2, text=self.tr("lbl_target"), bg=_bg).pack(anchor=W, pady=(8, 0))
        f2_tgt = tk.Frame(f2, bg=_bg)
        f2_tgt.pack(fill=X)
        ent_tgt = tk.Entry(f2_tgt, width=40)
        ent_tgt.pack(side=LEFT, fill=X, expand=True, padx=(0, 4))

        def pick_tgt():
            p = filedialog.askdirectory(
                title=self.tr("msg_task_pick_target"), parent=win
            )
            if p:
                ent_tgt.delete(0, END)
                ent_tgt.insert(0, p)

        tk.Button(f2_tgt, text=self.tr("btn_browse"), command=pick_tgt, width=4).pack(
            side=LEFT
        )
        var_inc = tk.IntVar(value=1)
        tk.Checkbutton(
            f2, text=self.tr("chk_incremental_mode"), variable=var_inc, bg=_bg
        ).pack(anchor=W, pady=(8, 0))
        tk.Label(
            f3, text=self.tr("wizard_step3"), font=("System", 11, "bold"), bg=_bg
        ).pack(anchor=W)
        var_mode = tk.StringVar(value=MODE_CONVERT_THEN_MERGE)
        for val, key in [
            (MODE_CONVERT_ONLY, "mode_convert"),
            (MODE_CONVERT_THEN_MERGE, "mode_convert_merge"),
            (MODE_MERGE_ONLY, "mode_merge"),
            (MODE_COLLECT_ONLY, "mode_collect"),
            (MODE_MSHELP_ONLY, "mode_mshelp"),
        ]:
            tk.Radiobutton(
                f3, text=self.tr(key), variable=var_mode, value=val, bg=_bg
            ).pack(anchor=W)
        tk.Label(
            f3, text=self.tr("grp_output_controls"), font=("System", 9, "bold"), bg=_bg
        ).pack(anchor=W, pady=(8, 0))
        var_output_choice = tk.StringVar(value="both")
        tk.Radiobutton(
            f3,
            text=self.tr("wizard_output_only_independent"),
            variable=var_output_choice,
            value="only_independent",
            bg=_bg,
        ).pack(anchor=W)
        tk.Radiobutton(
            f3,
            text=self.tr("wizard_output_only_merged"),
            variable=var_output_choice,
            value="only_merged",
            bg=_bg,
        ).pack(anchor=W)
        tk.Radiobutton(
            f3,
            text=self.tr("wizard_output_both"),
            variable=var_output_choice,
            value="both",
            bg=_bg,
        ).pack(anchor=W)
        f3_merge = tk.Frame(f3, bg=_bg)
        f3_merge.pack(fill=X, pady=(8, 0))
        var_merge_how = tk.StringVar(value=MERGE_MODE_CATEGORY)
        tk.Radiobutton(
            f3_merge,
            text=self.tr("wizard_merge_single_file"),
            variable=var_merge_how,
            value=MERGE_MODE_ALL_IN_ONE,
            bg=_bg,
        ).pack(anchor=W)
        tk.Radiobutton(
            f3_merge,
            text=self.tr("wizard_merge_by_size"),
            variable=var_merge_how,
            value=MERGE_MODE_CATEGORY,
            bg=_bg,
        ).pack(anchor=W)
        f3_mb = tk.Frame(f3_merge, bg=_bg)
        f3_mb.pack(fill=X, pady=(2, 0))
        tk.Label(f3_mb, text=self.tr("wizard_merge_size_mb"), bg=_bg).pack(
            side=LEFT, padx=(0, 4)
        )
        var_mb = tk.StringVar(value="80")
        ent_mb = tk.Entry(f3_mb, width=6, textvariable=var_mb)
        ent_mb.pack(side=LEFT)
        lbl_mb_tip = tk.Label(f3_merge, text="", bg=_bg, fg="#666")
        lbl_mb_tip.pack(anchor=W, pady=(2, 0))

        def _wizard_update_merge_ui():
            need_merge = var_output_choice.get() in ("only_merged", "both")
            if need_merge:
                f3_merge.pack(fill=X, pady=(8, 0))
                is_split = var_merge_how.get() == MERGE_MODE_CATEGORY
                if is_split:
                    ent_mb.configure(state="normal")
                    lbl_mb_tip.configure(text="")
                else:
                    ent_mb.configure(state="disabled")
                    lbl_mb_tip.configure(text=self.tr("tip_wizard_merge_size_disabled"))
            else:
                f3_merge.pack_forget()

        var_output_choice.trace_add("write", lambda *a: _wizard_update_merge_ui())
        var_merge_how.trace_add("write", lambda *a: _wizard_update_merge_ui())
        _wizard_update_merge_ui()
        var_pdf = tk.IntVar(value=1)
        var_md = tk.IntVar(value=1)
        tk.Checkbutton(
            f3, text=self.tr("chk_output_pdf"), variable=var_pdf, bg=_bg
        ).pack(anchor=W, pady=(4, 0))
        tk.Checkbutton(f3, text=self.tr("chk_output_md"), variable=var_md, bg=_bg).pack(
            anchor=W
        )
        tk.Label(
            f4, text=self.tr("wizard_step4"), font=("System", 11, "bold"), bg=_bg
        ).pack(anchor=W)
        lbl_summary = tk.Label(f4, text="", justify=LEFT, wraplength=450, bg=_bg)
        lbl_summary.pack(anchor=W, pady=(8, 0))
        var_run_after_save = tk.IntVar(value=0)
        chk_run_after = tk.Checkbutton(
            f4,
            text=self.tr("chk_wizard_run_after_save"),
            variable=var_run_after_save,
            bg=_bg,
        )
        chk_run_after.pack(anchor=W, pady=(8, 0))

        def _show(step):
            win._wizard_step = step
            f1.pack_forget()
            f2.pack_forget()
            f3.pack_forget()
            f4.pack_forget()
            (f1, f2, f3, f4)[step - 1].pack(fill=BOTH, expand=True)
            btn_back.configure(state="normal" if step > 1 else "disabled")
            btn_next.configure(state="normal" if step < 4 else "disabled")
            btn_save.pack_forget()
            btn_next.pack_forget()
            if step == 4:
                btn_save.pack(side=LEFT, padx=4)
                _refresh_summary()
            else:
                btn_next.pack(side=LEFT, padx=4)
            try:
                canvas.yview_moveto(0)
                _on_container_configure()
            except (tk.TclError, RuntimeError, AttributeError):
                pass

        def _merge_summary_text(d):
            if not d.get("output_enable_merged"):
                return self.tr("wizard_merge_summary_none")
            if d.get("merge_mode") == MERGE_MODE_ALL_IN_ONE:
                return self.tr("wizard_merge_summary_single")
            return self.tr("wizard_merge_summary_split").format(
                d.get("max_merge_size_mb", 80)
            )

        def _refresh_summary():
            d = win._wizard_data
            src_display = (
                d["source_folder"]
                if len(d.get("source_folders", [])) <= 1
                else f"{d['source_folder']} (+{len(d['source_folders']) - 1})"
            )
            lbl_summary.configure(
                text=f"{self.tr('msg_task_input_name')} {d['name']}\n"
                f"{self.tr('lbl_source')} {src_display}\n"
                f"{self.tr('lbl_target')} {d['target_folder']}\n"
                f"{self.tr('chk_incremental_mode')}: {'Y' if d['run_incremental'] else 'N'}\n"
                f"Run mode: {d['run_mode']}\n"
                f"{_merge_summary_text(d)}"
            )

        def _collect():
            data["name"] = ent_name.get().strip()
            data["description"] = txt_desc.get("1.0", END).strip()
            data["source_folders"] = list(lst_src.get(0, END))
            data["source_folder"] = (data["source_folders"] or [""])[0]
            data["target_folder"] = ent_tgt.get().strip()
            data["run_incremental"] = bool(var_inc.get())
            data["run_mode"] = var_mode.get()
            data["output_enable_pdf"] = bool(var_pdf.get())
            data["output_enable_md"] = bool(var_md.get())
            choice = var_output_choice.get()
            data["output_choice"] = choice
            if choice == "only_independent":
                data["output_enable_merged"] = False
                data["output_enable_independent"] = True
            elif choice == "only_merged":
                data["output_enable_merged"] = True
                data["output_enable_independent"] = False
            else:
                data["output_enable_merged"] = True
                data["output_enable_independent"] = True
            data["merge_how"] = var_merge_how.get()
            data["max_merge_size_mb"] = self._safe_positive_int(var_mb.get(), 80)
            data["merge_mode"] = (
                data["merge_how"]
                if data["output_enable_merged"]
                else MERGE_MODE_ALL_IN_ONE
            )

        def _go(delta):
            _collect()
            step = win._wizard_step + delta
            step = max(1, min(4, step))
            _show(step)

        def _save():
            _collect()
            d = win._wizard_data
            if not d["name"]:
                messagebox.showwarning(
                    self.tr("win_task_wizard"),
                    self.tr("msg_task_input_name"),
                    parent=win,
                )
                return
            name_trim = d["name"].strip()
            for t in self.task_store.list_tasks() or []:
                if isinstance(t, dict) and (t.get("name") or "").strip() == name_trim:
                    messagebox.showwarning(
                        self.tr("win_task_wizard"),
                        self.tr("msg_task_name_duplicate"),
                        parent=win,
                    )
                    return
            source_folders = [
                p.strip()
                for p in (d.get("source_folders") or [])
                if p and str(p).strip()
            ]
            if not source_folders:
                messagebox.showwarning(
                    self.tr("win_task_wizard"),
                    self.tr("msg_source_folder_required"),
                    parent=win,
                )
                return
            for p in source_folders:
                if not os.path.isdir(p):
                    messagebox.showwarning(
                        self.tr("win_task_wizard"),
                        self.tr("msg_source_folder_required"),
                        parent=win,
                    )
                    return
            if not d["target_folder"] or not os.path.isdir(d["target_folder"]):
                messagebox.showwarning(
                    self.tr("win_task_wizard"),
                    self.tr("msg_target_folder_required"),
                    parent=win,
                )
                return
            project_cfg = self._load_config_for_write()
            overrides = self._build_task_overrides_from_ui(
                project_cfg=project_cfg, only_diff=True
            )
            overrides["run_mode"] = d["run_mode"]
            overrides["output_enable_pdf"] = d["output_enable_pdf"]
            overrides["output_enable_md"] = d["output_enable_md"]
            overrides["output_enable_merged"] = d["output_enable_merged"]
            overrides["output_enable_independent"] = d["output_enable_independent"]
            overrides["merge_mode"] = d.get("merge_mode", MERGE_MODE_CATEGORY)
            overrides["max_merge_size_mb"] = d.get("max_merge_size_mb", 80)
            overrides["merge_filename_pattern"] = (
                d.get("merge_filename_pattern") or "Merged_{category}_{timestamp}_{idx}"
            )
            task = {
                "id": self._new_task_id(),
                "name": d["name"],
                "description": d.get("description", ""),
                "source_folders": source_folders,
                "source_folder": (source_folders or [""])[0],
                "target_folder": d["target_folder"],
                "run_incremental": d["run_incremental"],
                "project_config_snapshot": dict(project_cfg),
                **self._build_task_config_binding_meta(),
                "config_overrides": overrides,
                "status": "idle",
            }
            task, _ = self._ensure_task_config_snapshots(
                task, project_cfg=project_cfg, persist=False
            )
            try:
                self.task_store.save_task(task)
            except (AttributeError, TypeError, ValueError, OSError, RuntimeError) as e:
                messagebox.showerror(self.tr("win_task_create"), str(e), parent=win)
                return
            saved_id = task["id"]
            run_after_save = bool(var_run_after_save.get())
            self._refresh_task_list_ui()
            win.destroy()
            if (
                run_after_save
                and saved_id
                and hasattr(self, "tree_tasks")
                and self.tree_tasks.exists(saved_id)
            ):
                self.tree_tasks.selection_set(saved_id)
                self.tree_tasks.focus(saved_id)
                self.after(200, lambda: self._on_click_task_run(False))

        # 在 _go/_save 定义之后再创建按钮，避免 Python 3.12 闭包 "free variable not associated" 错误
        btn_back = tk.Button(
            nav,
            text=self.tr("btn_wizard_back"),
            state="disabled",
            command=lambda: _go(-1),
        )
        btn_back.pack(side=LEFT, padx=4)
        btn_next = tk.Button(
            nav, text=self.tr("btn_wizard_next"), command=lambda: _go(1)
        )
        btn_next.pack(side=LEFT, padx=4)
        btn_save = tk.Button(nav, text=self.tr("btn_wizard_save"), command=_save)
        btn_save.pack(side=LEFT, padx=4)

        # 直接显示第一步（纯 tk 控件在 Toplevel 上无需等待主题/布局，可立即渲染）
        _show(1)
        try:
            win.update_idletasks()
            win.update()
        except (tk.TclError, RuntimeError):
            pass

    def _task_binding_mode_text(self, mode):
        mode = normalize_task_binding_mode(mode)
        return {
            TASK_BINDING_ACTIVE: "跟随当前活动配置",
            TASK_BINDING_PROFILE: "绑定指定配置档",
            TASK_BINDING_SNAPSHOT: "使用任务快照",
        }.get(mode, mode)

    def _edit_task_binding_in_dialog(self, task, parent=None):
        task = task if isinstance(task, dict) else {}
        parent = parent or self
        current_mode = normalize_task_binding_mode(task.get("config_binding_mode"))
        bound_path = self._safe_abs_path(task.get("config_snapshot_path", ""))
        profile_records = []
        try:
            for rec in self._load_profile_records() or []:
                if not isinstance(rec, dict):
                    continue
                abs_path = self._profile_record_abs_path(rec)
                if not abs_path:
                    continue
                row = dict(rec)
                row["abs_path"] = abs_path
                profile_records.append(row)
        except (AttributeError, TypeError, ValueError, OSError):
            profile_records = []

        profile_labels = []
        profile_map = {}
        selected_label = ""
        for rec in profile_records:
            name = str(rec.get("name", "")).strip()
            file_name = str(rec.get("file", "")).strip()
            abs_path = str(rec.get("abs_path", "")).strip()
            label = f"{name} ({file_name})"
            profile_labels.append(label)
            profile_map[label] = rec
            if not selected_label and bound_path and self._safe_abs_path(abs_path) == bound_path:
                selected_label = label
        if not selected_label and profile_labels:
            selected_label = profile_labels[0]

        win = tk.Toplevel(parent)
        win.title("任务配置绑定")
        win.geometry("640x320")
        win.minsize(560, 280)
        win.transient(parent)
        win.grab_set()

        var_mode = tk.StringVar(value=current_mode)
        var_profile = tk.StringVar(value=selected_label)
        result = {"updates": None}

        tk.Label(
            win,
            text="选择任务运行时使用的配置来源：",
            anchor="w",
            font=("Microsoft YaHei", 10, "bold"),
        ).pack(fill=X, padx=12, pady=(12, 8))

        mode_frame = tk.LabelFrame(win, text="绑定模式", padx=8, pady=8)
        mode_frame.pack(fill=X, padx=12, pady=(0, 8))
        for mode_value in (
            TASK_BINDING_ACTIVE,
            TASK_BINDING_PROFILE,
            TASK_BINDING_SNAPSHOT,
        ):
            tk.Radiobutton(
                mode_frame,
                text=self._task_binding_mode_text(mode_value),
                variable=var_mode,
                value=mode_value,
                anchor="w",
            ).pack(fill=X, anchor="w")

        profile_frame = tk.LabelFrame(win, text="指定配置档", padx=8, pady=8)
        profile_frame.pack(fill=X, padx=12, pady=(0, 8))
        cmb_profiles = ttk.Combobox(
            profile_frame,
            textvariable=var_profile,
            state="readonly",
            values=profile_labels,
        )
        cmb_profiles.pack(fill=X)
        if selected_label:
            cmb_profiles.set(selected_label)
        hint_text = "提示：仅在“绑定指定配置档”模式下生效。"
        if not profile_labels:
            hint_text = "未找到配置档，请先在“配置管理”中保存配置档。"
        tk.Label(profile_frame, text=hint_text, anchor="w").pack(
            fill=X, pady=(6, 0)
        )

        def _refresh_mode_ui(*_args):
            if var_mode.get() == TASK_BINDING_PROFILE and profile_labels:
                cmb_profiles.configure(state="readonly")
            else:
                cmb_profiles.configure(state="disabled")

        var_mode.trace_add("write", _refresh_mode_ui)
        _refresh_mode_ui()

        def _confirm():
            chosen_mode = normalize_task_binding_mode(var_mode.get())
            updates = {"config_binding_mode": chosen_mode}
            if chosen_mode == TASK_BINDING_ACTIVE:
                updates.update(self._build_task_config_binding_meta())
            elif chosen_mode == TASK_BINDING_PROFILE:
                rec = profile_map.get(var_profile.get())
                if not rec:
                    messagebox.showwarning(
                        "任务配置绑定",
                        "请先选择一个可用的配置档。",
                        parent=win,
                    )
                    return
                updates["config_snapshot_path"] = str(rec.get("abs_path", "")).strip()
                updates["config_snapshot_profile_name"] = str(rec.get("name", "")).strip()
                updates["config_snapshot_profile_file"] = str(rec.get("file", "")).strip()
            else:
                updates["config_snapshot_path"] = ""
                updates["config_snapshot_profile_name"] = ""
                updates["config_snapshot_profile_file"] = ""
            result["updates"] = updates
            win.destroy()

        def _cancel():
            win.destroy()

        btn_row = tk.Frame(win)
        btn_row.pack(fill=X, padx=12, pady=(8, 12))
        tk.Button(btn_row, text="确认", command=_confirm).pack(side=LEFT)
        tk.Button(btn_row, text="取消", command=_cancel).pack(side=LEFT, padx=(8, 0))
        win.protocol("WM_DELETE_WINDOW", _cancel)
        self.wait_window(win)
        return result["updates"]

    def _open_task_edit_form(self, task, parent=None):
        task = task if isinstance(task, dict) else {}
        parent = parent or self
        records = []
        try:
            for rec in self._load_profile_records() or []:
                if not isinstance(rec, dict):
                    continue
                abs_path = self._profile_record_abs_path(rec)
                if not abs_path:
                    continue
                row = dict(rec)
                row["abs_path"] = abs_path
                records.append(row)
        except (AttributeError, TypeError, ValueError, OSError):
            records = []

        source_folders = task.get("source_folders") or []
        source_default = str(
            (source_folders[0] if source_folders else task.get("source_folder", "")) or ""
        ).strip()
        target_default = str(task.get("target_folder", "") or "").strip()
        mode_default = normalize_task_binding_mode(task.get("config_binding_mode"))
        task_cfg_path = self._safe_abs_path(task.get("config_snapshot_path", ""))
        selected_profile = None
        for rec in records:
            if self._safe_abs_path(rec.get("abs_path", "")) == task_cfg_path:
                selected_profile = rec
                break

        win = tk.Toplevel(parent)
        win.title(self.tr("win_task_edit"))
        win.geometry("760x520")
        win.minsize(700, 480)
        win.transient(parent)
        win.grab_set()

        frm = ttk.Frame(win, padding=12)
        frm.pack(fill=BOTH, expand=YES)

        var_name = tk.StringVar(value=str(task.get("name", "")))
        var_source = tk.StringVar(value=source_default)
        var_target = tk.StringVar(value=target_default)
        var_incremental = tk.IntVar(value=1 if task.get("run_incremental", True) else 0)
        var_mode = tk.StringVar(value=mode_default)
        var_profile_path = tk.StringVar(value=task_cfg_path)

        profile_labels = []
        profile_by_label = {}
        label_by_path = {}
        for rec in records:
            label = f"{rec.get('name', '')} ({rec.get('file', '')})"
            profile_labels.append(label)
            profile_by_label[label] = rec
            label_by_path[self._safe_abs_path(rec.get("abs_path", ""))] = label
        var_profile = tk.StringVar(
            value=label_by_path.get(task_cfg_path, profile_labels[0] if profile_labels else "")
        )
        if selected_profile is not None:
            var_profile_path.set(str(selected_profile.get("abs_path", "")).strip())

        ttk.Label(frm, text="任务名称").pack(anchor=W, pady=(0, 4))
        ttk.Entry(frm, textvariable=var_name).pack(fill=X)

        ttk.Label(frm, text="任务描述（可选）").pack(anchor=W, pady=(8, 4))
        txt_desc = tk.Text(frm, height=3, wrap="word")
        txt_desc.pack(fill=X)
        txt_desc.insert("1.0", str(task.get("description", "") or ""))

        row_src = ttk.Frame(frm)
        row_src.pack(fill=X, pady=(8, 0))
        ttk.Label(row_src, text="源目录").pack(side=LEFT)
        ttk.Entry(row_src, textvariable=var_source).pack(side=LEFT, fill=X, expand=YES, padx=(8, 6))
        ttk.Button(
            row_src,
            text="浏览",
            command=lambda: (
                (lambda p: var_source.set(p) if p else None)(
                    filedialog.askdirectory(title=self.tr("msg_task_pick_source"), parent=win)
                )
            ),
        ).pack(side=LEFT)

        row_tgt = ttk.Frame(frm)
        row_tgt.pack(fill=X, pady=(6, 0))
        ttk.Label(row_tgt, text="目标目录").pack(side=LEFT)
        ttk.Entry(row_tgt, textvariable=var_target).pack(side=LEFT, fill=X, expand=YES, padx=(8, 6))
        ttk.Button(
            row_tgt,
            text="浏览",
            command=lambda: (
                (lambda p: var_target.set(p) if p else None)(
                    filedialog.askdirectory(title=self.tr("msg_task_pick_target"), parent=win)
                )
            ),
        ).pack(side=LEFT)

        ttk.Checkbutton(
            frm,
            text=self.tr("msg_task_incremental_prompt"),
            variable=var_incremental,
        ).pack(anchor=W, pady=(10, 0))

        lf_binding = ttk.LabelFrame(frm, text="任务配置绑定", padding=8)
        lf_binding.pack(fill=X, pady=(10, 0))

        for mv in (TASK_BINDING_ACTIVE, TASK_BINDING_PROFILE, TASK_BINDING_SNAPSHOT):
            ttk.Radiobutton(
                lf_binding,
                text=self._task_binding_mode_text(mv),
                variable=var_mode,
                value=mv,
            ).pack(anchor=W)

        row_profile = ttk.Frame(lf_binding)
        row_profile.pack(fill=X, pady=(8, 0))
        ttk.Label(row_profile, text="配置档").pack(side=LEFT)
        cmb_profile = ttk.Combobox(
            row_profile,
            textvariable=var_profile,
            values=profile_labels,
            state="readonly",
        )
        cmb_profile.pack(side=LEFT, fill=X, expand=YES, padx=(8, 6))

        def _on_pick_profile(_event=None):
            rec = profile_by_label.get(var_profile.get())
            if isinstance(rec, dict):
                var_profile_path.set(str(rec.get("abs_path", "")).strip())

        cmb_profile.bind("<<ComboboxSelected>>", _on_pick_profile)

        row_profile_path = ttk.Frame(lf_binding)
        row_profile_path.pack(fill=X, pady=(6, 0))
        ttk.Label(row_profile_path, text="绑定配置文件").pack(side=LEFT)
        ent_profile_path = ttk.Entry(row_profile_path, textvariable=var_profile_path)
        ent_profile_path.pack(side=LEFT, fill=X, expand=YES, padx=(8, 6))

        def _pick_profile_json():
            picked = filedialog.askopenfilename(
                title="选择配置 JSON 文件",
                parent=win,
                filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
            )
            if not picked:
                return
            abs_picked = self._safe_abs_path(picked)
            var_profile_path.set(abs_picked)
            if abs_picked in label_by_path:
                var_profile.set(label_by_path[abs_picked])

        btn_pick_json = ttk.Button(
            row_profile_path, text="选择文件", command=_pick_profile_json
        )
        btn_pick_json.pack(side=LEFT)

        lbl_hint = ttk.Label(
            lf_binding,
            text="提示：'绑定指定配置档' 可选择已保存配置档或任意 JSON 配置文件。",
        )
        lbl_hint.pack(anchor=W, pady=(6, 0))

        def _refresh_binding_controls(*_args):
            is_profile = var_mode.get() == TASK_BINDING_PROFILE
            state = "readonly" if is_profile and profile_labels else "disabled"
            cmb_profile.configure(state=state)
            ent_profile_path.configure(state="normal" if is_profile else "disabled")
            btn_pick_json.configure(state="normal" if is_profile else "disabled")

        var_mode.trace_add("write", _refresh_binding_controls)
        _refresh_binding_controls()

        result = {"updates": None}

        def _save():
            name = var_name.get().strip()
            source = var_source.get().strip()
            target = var_target.get().strip()
            if not name:
                messagebox.showwarning(self.tr("win_task_edit"), self.tr("msg_task_input_name"), parent=win)
                return
            if not source or not os.path.isdir(source):
                messagebox.showwarning(self.tr("win_task_edit"), self.tr("msg_source_folder_required"), parent=win)
                return
            if not target or not os.path.isdir(target):
                messagebox.showwarning(self.tr("win_task_edit"), self.tr("msg_target_folder_required"), parent=win)
                return

            updates = {
                "name": name,
                "description": txt_desc.get("1.0", END).strip(),
                "source_folder": source,
                "source_folders": [source],
                "target_folder": target,
                "run_incremental": bool(var_incremental.get()),
            }

            mode = normalize_task_binding_mode(var_mode.get())
            updates["config_binding_mode"] = mode
            if mode == TASK_BINDING_ACTIVE:
                updates.update(self._build_task_config_binding_meta())
            elif mode == TASK_BINDING_PROFILE:
                chosen_path = self._safe_abs_path(var_profile_path.get())
                if not chosen_path or not os.path.isfile(chosen_path):
                    messagebox.showwarning(
                        self.tr("win_task_edit"),
                        "请先选择一个存在的配置 JSON 文件。",
                        parent=win,
                    )
                    return
                updates["config_snapshot_path"] = chosen_path
                rec = self._find_profile_record_by_path(chosen_path)
                if isinstance(rec, dict):
                    updates["config_snapshot_profile_name"] = str(rec.get("name", "")).strip()
                    updates["config_snapshot_profile_file"] = str(rec.get("file", "")).strip()
                else:
                    updates["config_snapshot_profile_name"] = ""
                    updates["config_snapshot_profile_file"] = os.path.basename(chosen_path)
            else:
                updates["config_snapshot_path"] = ""
                updates["config_snapshot_profile_name"] = ""
                updates["config_snapshot_profile_file"] = ""

            result["updates"] = updates
            win.destroy()

        def _cancel():
            win.destroy()

        row_btn = ttk.Frame(frm)
        row_btn.pack(fill=X, pady=(12, 0))
        ttk.Button(row_btn, text="保存", command=_save).pack(side=LEFT)
        ttk.Button(row_btn, text="取消", command=_cancel).pack(side=LEFT, padx=(8, 0))
        win.protocol("WM_DELETE_WINDOW", _cancel)
        self.wait_window(win)
        return result["updates"]

    def _on_click_task_edit(self):
        task_id = self._get_selected_task_id()
        if not task_id:
            messagebox.showinfo(
                self.tr("grp_task_runtime"), self.tr("msg_task_select_required")
            )
            return
        task = self.task_store.get_task(task_id)
        if not task:
            return
        updates = self._open_task_edit_form(task, parent=self)
        if not isinstance(updates, dict):
            return
        task.update(updates)
        # Keep per-task config independent from current UI state.
        # Edit dialog only updates basic task metadata, not advanced runtime config.
        task["config_overrides"] = (
            task.get("config_overrides")
            if isinstance(task.get("config_overrides"), dict)
            else {}
        )
        task, _ = self._ensure_task_config_snapshots(
            task, project_cfg=self._load_config_for_write(), persist=False
        )
        try:
            self.task_store.save_task(task)
        except (AttributeError, TypeError, ValueError, OSError, RuntimeError) as e:
            messagebox.showerror(self.tr("win_task_edit"), str(e))
            return
        self._refresh_task_list_ui()

    def _on_click_task_delete(self):
        task_id = self._get_selected_task_id()
        if not task_id:
            messagebox.showinfo(
                self.tr("grp_task_runtime"), self.tr("msg_task_select_required")
            )
            return
        task = self.task_store.get_task(task_id) or {}
        if not messagebox.askyesno(
            self.tr("btn_task_delete"),
            self.tr("msg_task_delete_confirm").format(task.get("name", task_id)),
            parent=self,
        ):
            return
        self.task_store.delete_task(task_id)
        self._refresh_task_list_ui()

    def _on_click_task_load_to_ui(self):
        """Load selected task's effective config into main UI (design 7.1: use task as template in classic mode)."""
        task_id = self._get_selected_task_id()
        if not task_id:
            messagebox.showinfo(
                self.tr("grp_task_runtime"), self.tr("msg_task_select_required")
            )
            return
        task = self.task_store.get_task(task_id)
        if not task:
            return
        try:
            task, cfg = self._ensure_task_config_snapshots(
                task, project_cfg=self._load_config_for_write(), persist=True
            )
        except (AttributeError, TypeError, ValueError, OSError, RuntimeError) as e:
            messagebox.showerror(self.tr("btn_task_load_to_ui"), str(e))
            return
        self._apply_task_runtime_to_ui(cfg)
        messagebox.showinfo(
            self.tr("btn_task_load_to_ui"),
            self.tr("msg_task_load_to_ui_done"),
            parent=self,
        )

    def _on_click_task_save_to_task(self):
        """Save current UI config binding to the selected task (one-click 'update task with active config')."""
        task_id = self._get_selected_task_id()
        if not task_id:
            messagebox.showinfo(
                self.tr("grp_task_runtime"), self.tr("msg_task_select_required")
            )
            return
        task = self.task_store.get_task(task_id)
        if not task:
            return
        try:
            task = dict(task)
            meta = self._build_task_config_binding_meta()
            task["config_binding_mode"] = meta.get("config_binding_mode", TASK_BINDING_ACTIVE)
            task["config_snapshot_path"] = meta.get("config_snapshot_path", "")
            task["config_snapshot_profile_name"] = meta.get("config_snapshot_profile_name", "")
            task["config_snapshot_profile_file"] = meta.get("config_snapshot_profile_file", "")
            self.task_store.save_task(task)
        except (AttributeError, TypeError, ValueError, OSError, RuntimeError) as e:
            messagebox.showerror(self.tr("btn_task_save_to_task"), str(e))
            return
        messagebox.showinfo(
            self.tr("btn_task_save_to_task"),
            self.tr("msg_task_save_to_task_done"),
            parent=self,
        )
        self._refresh_task_list_ui()

    def _apply_task_runtime_to_ui(self, cfg, preserve_current_run_tab=False):
        self._suspend_cfg_dirty = True
        try:
            src_list = cfg.get("source_folders") or []
            if not src_list and cfg.get("source_folder"):
                src_list = [cfg.get("source_folder")]
            self.source_folders_list = list(src_list)
            self.lst_source_folders.delete(0, END)
            for p in self.source_folders_list:
                self.lst_source_folders.insert(END, p)
            self.var_source_folder.set(
                self.source_folders_list[0] if self.source_folders_list else ""
            )
            self.var_target_folder.set(cfg.get("target_folder") or "")
            self.var_run_mode.set(cfg.get("run_mode", MODE_CONVERT_THEN_MERGE))
            self.var_output_enable_pdf.set(
                1 if cfg.get("output_enable_pdf", True) else 0
            )
            self.var_output_enable_md.set(1 if cfg.get("output_enable_md", True) else 0)
            self.var_output_enable_merged.set(
                1 if cfg.get("output_enable_merged", True) else 0
            )
            self.var_output_enable_independent.set(
                1 if cfg.get("output_enable_independent", False) else 0
            )
            self.var_enable_fast_md_engine.set(
                1 if cfg.get("enable_fast_md_engine", False) else 0
            )
            self.var_enable_traceability_anchor_and_map.set(
                1 if cfg.get("enable_traceability_anchor_and_map", True) else 0
            )
            self.var_enable_markdown_image_manifest.set(
                1 if cfg.get("enable_markdown_image_manifest", True) else 0
            )
            self.var_enable_prompt_wrapper.set(
                1 if cfg.get("enable_prompt_wrapper", False) else 0
            )
            self.var_prompt_template_type.set(
                str(cfg.get("prompt_template_type", "new_solution") or "new_solution")
            )
            self.var_short_id_prefix.set(
                str(cfg.get("short_id_prefix", "ZW-") or "ZW-")
            )
            self.var_merge_mode.set(cfg.get("merge_mode", MERGE_MODE_CATEGORY))
            self.var_max_merge_size_mb.set(str(cfg.get("max_merge_size_mb", 80)))
            self.var_merge_filename_pattern.set(
                cfg.get("merge_filename_pattern")
                or "Merged_{category}_{timestamp}_{idx}"
            )
            self.var_collect_mode.set(
                cfg.get("collect_mode", COLLECT_MODE_COPY_AND_INDEX)
            )
            self.var_strategy.set(cfg.get("content_strategy", STRATEGY_STANDARD))
            self.var_engine.set(cfg.get("default_engine", ENGINE_WPS))
            self.var_kill_mode.set(cfg.get("kill_process_mode", KILL_MODE_AUTO))
            self.var_enable_incremental_mode.set(
                1 if cfg.get("enable_incremental_mode", False) else 0
            )
            self.var_enable_merge.set(1 if cfg.get("enable_merge", True) else 0)
            self.var_merge_source.set(cfg.get("merge_source", "source"))
            self.var_enable_merge_index.set(
                1 if cfg.get("enable_merge_index", False) else 0
            )
            self.var_enable_merge_excel.set(
                1 if cfg.get("enable_merge_excel", False) else 0
            )
        finally:
            self._suspend_cfg_dirty = False
        prev_suppress = bool(getattr(self, "_suppress_run_tab_autoselect", False))
        if preserve_current_run_tab:
            self._suppress_run_tab_autoselect = True
        try:
            self._on_run_mode_change()
        finally:
            self._suppress_run_tab_autoselect = prev_suppress
        self._on_toggle_incremental_mode()
        self.validate_runtime_inputs(silent=True, scope="all")

