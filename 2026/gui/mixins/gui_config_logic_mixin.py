# -*- coding: utf-8 -*-
"""Config logic methods extracted from OfficeGUI for maintainability."""

import json
import os
import re
from datetime import datetime
import tkinter as tk
from tkinter.constants import *

from office_converter import (
    MODE_CONVERT_ONLY,
    MODE_MERGE_ONLY,
    MODE_CONVERT_THEN_MERGE,
    MODE_COLLECT_ONLY,
    MODE_MSHELP_ONLY,
    MERGE_CONVERT_SUBMODE_MERGE_ONLY,
    MERGE_MODE_CATEGORY,
    KILL_MODE_AUTO,
)


class ConfigLogicMixin:
    @staticmethod
    def _is_valid_hex_color(text):
        return bool(re.fullmatch(r"#(?:[0-9A-Fa-f]{6})", str(text).strip()))

    def _set_entry_valid_state(self, entry, valid=True):
        if entry is None:
            return
        try:
            entry.configure(bootstyle="default" if valid else "danger")
        except Exception:
            pass

    def _bind_var_validation(self, var, callback):
        if var is None:
            return
        try:
            var.trace_add("write", lambda *_: callback())
        except Exception:
            pass

    @staticmethod
    def _safe_positive_int(raw, default):
        try:
            value = int(str(raw).strip())
            return value if value > 0 else default
        except Exception:
            return default

    @staticmethod

    def _normalize_lines(lines):
        return [str(x).strip() for x in (lines or []) if str(x).strip()]

    def _read_text_lines(self, widget):
        if widget is None:
            return []
        try:
            raw = widget.get("1.0", "end").strip()
        except Exception:
            return []
        return self._normalize_lines(raw.splitlines())

    def _build_config_snapshot_from_ui(self):
        return {
            "kill_process_mode": self.var_kill_mode.get(),
            "log_folder": self.var_log_folder.get().strip() or "./logs",
            "timeout_seconds": self._safe_positive_int(
                self.var_timeout_seconds.get(), 60
            ),
            "pdf_wait_seconds": self._safe_positive_int(
                self.var_pdf_wait_seconds.get(), 15
            ),
            "ppt_timeout_seconds": self._safe_positive_int(
                self.var_ppt_timeout_seconds.get(), 180
            ),
            "ppt_pdf_wait_seconds": self._safe_positive_int(
                self.var_ppt_pdf_wait_seconds.get(), 30
            ),
            "office_reuse_app": bool(self.var_office_reuse_app.get()),
            "office_restart_every_n_files": self._safe_positive_int(
                self.var_office_restart_every_n_files.get(), 25
            ),
            "enable_sandbox": bool(self.var_enable_sandbox.get()),
            "temp_sandbox_root": self.var_temp_sandbox_root.get().strip(),
            "sandbox_min_free_gb": self._safe_positive_int(
                self.var_sandbox_min_free_gb.get(), 10
            ),
            "sandbox_low_space_policy": self.var_sandbox_low_space_policy.get()
            or "block",
            "enable_merge": bool(self.var_enable_merge.get()),
            "output_enable_pdf": bool(self.var_output_enable_pdf.get()),
            "output_enable_md": bool(self.var_output_enable_md.get()),
            "output_enable_merged": bool(self.var_output_enable_merged.get()),
            "output_enable_independent": bool(self.var_output_enable_independent.get()),
            "merge_convert_submode": self.var_merge_convert_submode.get(),
            "merge_mode": self.var_merge_mode.get(),
            "merge_source": self.var_merge_source.get(),
            "enable_merge_index": bool(self.var_enable_merge_index.get()),
            "enable_merge_excel": bool(self.var_enable_merge_excel.get()),
            "max_merge_size_mb": self._safe_positive_int(
                self.var_max_merge_size_mb.get(), 80
            ),
            "merge_filename_pattern": self.var_merge_filename_pattern.get().strip()
            or "Merged_{category}_{timestamp}_{idx}",
            "enable_corpus_manifest": bool(self.var_enable_corpus_manifest.get()),
            "enable_markdown": bool(self.var_output_enable_md.get()),
            "markdown_strip_header_footer": bool(
                self.var_markdown_strip_header_footer.get()
            ),
            "markdown_structured_headings": bool(
                self.var_markdown_structured_headings.get()
            ),
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
            "enable_llm_delivery_hub": bool(self.var_enable_llm_delivery_hub.get()),
            "llm_delivery_root": self.var_llm_delivery_root.get().strip(),
            "llm_delivery_flatten": bool(self.var_llm_delivery_flatten.get()),
            "llm_delivery_include_pdf": bool(self.var_llm_delivery_include_pdf.get()),
            "enable_gdrive_upload": bool(self.var_enable_gdrive_upload.get()),
            "gdrive_client_secrets_path": self.var_gdrive_client_secrets_path.get().strip(),
            "gdrive_folder_id": self.var_gdrive_folder_id.get().strip(),
            "mshelpviewer_folder_name": str(
                self.var_mshelpviewer_folder_name.get()
            ).strip()
            or "MSHelpViewer",
            "enable_mshelp_merge_output": bool(
                self.var_enable_mshelp_merge_output.get()
            ),
            "enable_mshelp_output_docx": bool(self.var_enable_mshelp_output_docx.get()),
            "enable_mshelp_output_pdf": bool(self.var_enable_mshelp_output_pdf.get()),
            "enable_incremental_mode": bool(self.var_enable_incremental_mode.get()),
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
            "excluded_folders": self._read_text_lines(self.txt_excluded_folders),
            "price_keywords": self._read_text_lines(self.txt_price_keywords),
            "ui": {
                "tooltip_delay_ms": self._safe_positive_int(
                    self.var_tooltip_delay_ms.get(),
                    self.TOOLTIP_DEFAULTS["tooltip_delay_ms"],
                ),
                "tooltip_bg": str(self.var_tooltip_bg.get()).strip().upper()
                or self.TOOLTIP_DEFAULTS["tooltip_bg"],
                "tooltip_fg": str(self.var_tooltip_fg.get()).strip().upper()
                or self.TOOLTIP_DEFAULTS["tooltip_fg"],
                "tooltip_font_size": self._safe_positive_int(
                    self.var_tooltip_font_size.get(),
                    self.TOOLTIP_DEFAULTS["tooltip_font_size"],
                ),
                "tooltip_auto_theme": bool(self.var_tooltip_auto_theme.get()),
                "confirm_revert_dirty": bool(self.var_confirm_revert_dirty.get()),
                "task_current_config_only": (
                    bool(self.var_task_scope_current_config_only.get())
                    if hasattr(self, "var_task_scope_current_config_only")
                    else True
                ),
            },
        }

    def _build_config_snapshot_from_cfg(self, cfg):
        ui_cfg = cfg.get("ui", {}) if isinstance(cfg.get("ui"), dict) else {}
        return {
            "kill_process_mode": cfg.get("kill_process_mode", KILL_MODE_AUTO),
            "log_folder": str(cfg.get("log_folder", "./logs")).strip() or "./logs",
            "timeout_seconds": self._safe_positive_int(
                cfg.get("timeout_seconds", 60), 60
            ),
            "pdf_wait_seconds": self._safe_positive_int(
                cfg.get("pdf_wait_seconds", 15), 15
            ),
            "ppt_timeout_seconds": self._safe_positive_int(
                cfg.get("ppt_timeout_seconds", 180), 180
            ),
            "ppt_pdf_wait_seconds": self._safe_positive_int(
                cfg.get("ppt_pdf_wait_seconds", 30), 30
            ),
            "office_reuse_app": bool(cfg.get("office_reuse_app", True)),
            "office_restart_every_n_files": self._safe_positive_int(
                cfg.get("office_restart_every_n_files", 25), 25
            ),
            "enable_sandbox": bool(cfg.get("enable_sandbox", True)),
            "temp_sandbox_root": str(cfg.get("temp_sandbox_root", "")).strip(),
            "sandbox_min_free_gb": self._safe_positive_int(
                cfg.get("sandbox_min_free_gb", 10), 10
            ),
            "sandbox_low_space_policy": str(
                cfg.get("sandbox_low_space_policy", "block")
            ).strip()
            or "block",
            "enable_merge": bool(cfg.get("enable_merge", True)),
            "output_enable_pdf": bool(cfg.get("output_enable_pdf", True)),
            "output_enable_md": bool(cfg.get("output_enable_md", True)),
            "output_enable_merged": bool(cfg.get("output_enable_merged", True)),
            "output_enable_independent": bool(
                cfg.get("output_enable_independent", False)
            ),
            "merge_convert_submode": str(
                cfg.get(
                    "merge_convert_submode",
                    MERGE_CONVERT_SUBMODE_MERGE_ONLY,
                )
            ),
            "merge_mode": cfg.get("merge_mode", MERGE_MODE_CATEGORY),
            "merge_source": cfg.get("merge_source", "source"),
            "enable_merge_index": bool(cfg.get("enable_merge_index", False)),
            "enable_merge_excel": bool(cfg.get("enable_merge_excel", False)),
            "max_merge_size_mb": self._safe_positive_int(
                cfg.get("max_merge_size_mb", 80), 80
            ),
            "merge_filename_pattern": cfg.get("merge_filename_pattern")
            or "Merged_{category}_{timestamp}_{idx}",
            "enable_corpus_manifest": bool(cfg.get("enable_corpus_manifest", True)),
            "enable_markdown": bool(
                cfg.get("output_enable_md", cfg.get("enable_markdown", True))
            ),
            "markdown_strip_header_footer": bool(
                cfg.get("markdown_strip_header_footer", True)
            ),
            "markdown_structured_headings": bool(
                cfg.get("markdown_structured_headings", True)
            ),
            "enable_markdown_quality_report": bool(
                cfg.get("enable_markdown_quality_report", True)
            ),
            "enable_excel_json": bool(cfg.get("enable_excel_json", False)),
            "enable_chromadb_export": bool(cfg.get("enable_chromadb_export", False)),
            "enable_fast_md_engine": bool(cfg.get("enable_fast_md_engine", False)),
            "enable_traceability_anchor_and_map": bool(
                cfg.get("enable_traceability_anchor_and_map", True)
            ),
            "enable_markdown_image_manifest": bool(
                cfg.get("enable_markdown_image_manifest", True)
            ),
            "enable_prompt_wrapper": bool(cfg.get("enable_prompt_wrapper", False)),
            "prompt_template_type": str(
                cfg.get("prompt_template_type", "new_solution") or "new_solution"
            ),
            "short_id_prefix": str(cfg.get("short_id_prefix", "ZW-") or "ZW-")
            .strip()
            .upper()
            or "ZW-",
            "enable_llm_delivery_hub": bool(cfg.get("enable_llm_delivery_hub", True)),
            "llm_delivery_root": cfg.get("llm_delivery_root", "") or "",
            "llm_delivery_flatten": bool(cfg.get("llm_delivery_flatten", False)),
            "llm_delivery_include_pdf": bool(
                cfg.get("llm_delivery_include_pdf", False)
            ),
            "enable_gdrive_upload": bool(cfg.get("enable_gdrive_upload", False)),
            "gdrive_client_secrets_path": cfg.get("gdrive_client_secrets_path", "")
            or "",
            "gdrive_folder_id": cfg.get("gdrive_folder_id", "") or "",
            "mshelpviewer_folder_name": str(
                cfg.get("mshelpviewer_folder_name", "MSHelpViewer")
            ).strip()
            or "MSHelpViewer",
            "enable_mshelp_merge_output": bool(
                cfg.get("enable_mshelp_merge_output", True)
            ),
            "enable_mshelp_output_docx": bool(
                cfg.get("enable_mshelp_output_docx", False)
            ),
            "enable_mshelp_output_pdf": bool(
                cfg.get("enable_mshelp_output_pdf", False)
            ),
            "enable_incremental_mode": bool(cfg.get("enable_incremental_mode", False)),
            "incremental_verify_hash": bool(cfg.get("incremental_verify_hash", False)),
            "incremental_reprocess_renamed": bool(
                cfg.get("incremental_reprocess_renamed", False)
            ),
            "source_priority_skip_same_name_pdf": bool(
                cfg.get("source_priority_skip_same_name_pdf", False)
            ),
            "global_md5_dedup": bool(cfg.get("global_md5_dedup", False)),
            "enable_update_package": bool(cfg.get("enable_update_package", True)),
            "enable_parallel_conversion": bool(
                cfg.get("enable_parallel_conversion", False)
            ),
            "parallel_workers": self._safe_positive_int(
                cfg.get("parallel_workers", 4), 4
            ),
            "enable_checkpoint": bool(cfg.get("enable_checkpoint", True)),
            "checkpoint_auto_resume": bool(cfg.get("checkpoint_auto_resume", True)),
            "excluded_folders": self._normalize_lines(cfg.get("excluded_folders", [])),
            "price_keywords": self._normalize_lines(cfg.get("price_keywords", [])),
            "ui": {
                "tooltip_delay_ms": self._safe_positive_int(
                    ui_cfg.get(
                        "tooltip_delay_ms", self.TOOLTIP_DEFAULTS["tooltip_delay_ms"]
                    ),
                    self.TOOLTIP_DEFAULTS["tooltip_delay_ms"],
                ),
                "tooltip_bg": str(
                    ui_cfg.get("tooltip_bg", self.TOOLTIP_DEFAULTS["tooltip_bg"])
                )
                .strip()
                .upper(),
                "tooltip_fg": str(
                    ui_cfg.get("tooltip_fg", self.TOOLTIP_DEFAULTS["tooltip_fg"])
                )
                .strip()
                .upper(),
                "tooltip_font_size": self._safe_positive_int(
                    ui_cfg.get(
                        "tooltip_font_size", self.TOOLTIP_DEFAULTS["tooltip_font_size"]
                    ),
                    self.TOOLTIP_DEFAULTS["tooltip_font_size"],
                ),
                "tooltip_auto_theme": bool(
                    ui_cfg.get(
                        "tooltip_auto_theme",
                        self.TOOLTIP_DEFAULTS["tooltip_auto_theme"],
                    )
                ),
                "confirm_revert_dirty": bool(ui_cfg.get("confirm_revert_dirty", True)),
                "task_current_config_only": bool(
                    ui_cfg.get("task_current_config_only", True)
                ),
            },
        }

    def _refresh_config_dirty_from_file(self):
        if self._suspend_cfg_dirty:
            return
        if not os.path.exists(self.config_path):
            return
        try:
            with open(self.config_path, "r", encoding="utf-8") as f:
                cfg = json.load(f)
        except Exception:
            return
        self._baseline_config_snapshot = self._build_config_snapshot_from_cfg(cfg)
        self._refresh_config_dirty_state()

    def _to_halfwidth(self, text):
        if text is None:
            return ""
        out = []
        for ch in str(text):
            code = ord(ch)
            if code == 0x3000:
                out.append(" ")
            elif 0xFF01 <= code <= 0xFF5E:
                out.append(chr(code - 0xFEE0))
            else:
                out.append(ch)
        return "".join(out)

    def _normalize_numeric_var(self, var):
        raw = var.get()
        norm = self._to_halfwidth(raw).strip()
        if norm != raw:
            var.set(norm)

    def _normalize_short_id_var(self, var):
        raw = var.get()
        norm = self._to_halfwidth(raw).strip().replace(" ", "").upper()
        if norm != raw:
            var.set(norm)

    def _normalize_date_var(self, var):
        raw = var.get()
        norm = self._to_halfwidth(raw).strip()
        norm = norm.replace("/", "-").replace(".", "-")
        norm = re.sub(r"-{2,}", "-", norm)
        m = re.fullmatch(r"(\d{4})-(\d{1,2})-(\d{1,2})", norm)
        if m:
            norm = f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
        if norm != raw:
            var.set(norm)

    def _normalize_then_validate(self, var, normalizer, scope):
        if self._normalizing_inputs:
            return
        self._normalizing_inputs = True
        try:
            normalizer(var)
        finally:
            self._normalizing_inputs = False
        self.validate_runtime_inputs(silent=False, scope=scope)

    def _set_status_validation_error(self, message):
        if hasattr(self, "var_status"):
            self.var_status.set(message)
        if hasattr(self, "var_locator_result"):
            self.var_locator_result.set(message)

    @staticmethod

    def _is_valid_short_id(text):
        return bool(re.fullmatch(r"[0-9A-Za-z-]{4,40}", str(text).strip()))

    def validate_runtime_inputs(self, silent=True, scope="all"):
        first_error = None

        def _mark(entry, ok, message_key, label_key):
            nonlocal first_error
            self._set_entry_valid_state(entry, ok)
            if (not ok) and first_error is None:
                first_error = self.tr(message_key).format(self.tr(label_key))

        # Locator quick check
        if scope in ("all", "locator"):
            page_raw = (
                self.var_locator_page.get().strip()
                if hasattr(self, "var_locator_page")
                else ""
            )
            page_ok = True
            if page_raw:
                page_ok = page_raw.isdigit() and int(page_raw) > 0
            _mark(
                getattr(self, "ent_locator_page", None),
                page_ok,
                "msg_validation_invalid_number",
                "lbl_locator_page",
            )

            short_id = (
                self.var_locator_short_id.get().strip()
                if hasattr(self, "var_locator_short_id")
                else ""
            )
            sid_ok = True if not short_id else self._is_valid_short_id(short_id)
            _mark(
                getattr(self, "ent_locator_short_id", None),
                sid_ok,
                "msg_validation_invalid_short_id",
                "lbl_locator_id",
            )

        # Runtime date filter check
        if scope in ("all", "run"):
            date_entry_widget = None
            if hasattr(self, "ent_date"):
                date_entry_widget = getattr(self.ent_date, "entry", self.ent_date)
            date_ok = True
            if (
                hasattr(self, "var_enable_date_filter")
                and self.var_enable_date_filter.get()
            ):
                date_str = (
                    self.var_date_str.get().strip()
                    if hasattr(self, "var_date_str")
                    else ""
                )
                try:
                    datetime.strptime(date_str, "%Y-%m-%d")
                except Exception:
                    date_ok = False
            _mark(
                date_entry_widget,
                date_ok,
                "msg_validation_invalid_date",
                "lbl_filter_date",
            )

        # Config numeric defaults
        if scope in ("all", "config"):
            numeric_fields = [
                ("var_timeout_seconds", "ent_timeout_seconds", "lbl_gen_timeout"),
                ("var_pdf_wait_seconds", "ent_pdf_wait_seconds", "lbl_pdf_wait"),
                (
                    "var_ppt_timeout_seconds",
                    "ent_ppt_timeout_seconds",
                    "lbl_ppt_timeout",
                ),
                (
                    "var_ppt_pdf_wait_seconds",
                    "ent_ppt_pdf_wait_seconds",
                    "lbl_ppt_wait",
                ),
                (
                    "var_office_restart_every_n_files",
                    "ent_office_restart_every_n_files",
                    "lbl_office_restart_every",
                ),
                ("var_max_merge_size_mb", "ent_max_merge_size_mb", "lbl_max_mb"),
            ]
            for var_name, ent_name, label_key in numeric_fields:
                raw = (
                    getattr(self, var_name).get().strip()
                    if hasattr(self, var_name)
                    else ""
                )
                ok = raw.isdigit() and int(raw) > 0
                _mark(
                    getattr(self, ent_name, None),
                    ok,
                    "msg_validation_invalid_number",
                    label_key,
                )

        if first_error and not silent:
            self._set_status_validation_error(first_error)
        elif not first_error and not silent and hasattr(self, "var_status"):
            self.var_status.set(self.tr("status_ready"))

        return first_error is None

    def _set_text_widget_lines(self, widget, lines):
        if widget is None:
            return
        widget.delete("1.0", "end")
        if lines:
            widget.insert("1.0", "\n".join(lines))

    def _reset_config_section_defaults(self, section_name):
        section = str(section_name or "").strip().lower()
        section_title_key = None

        if section == "shared":
            self.var_kill_mode.set(KILL_MODE_AUTO)
            self.var_log_folder.set("./logs")
            section_title_key = "grp_cfg_shared"
        elif section == "convert":
            self.var_timeout_seconds.set("60")
            self.var_pdf_wait_seconds.set("15")
            self.var_ppt_timeout_seconds.set("180")
            self.var_ppt_pdf_wait_seconds.set("30")
            self.var_office_reuse_app.set(1)
            self.var_office_restart_every_n_files.set("25")
            self.var_enable_sandbox.set(1)
            self.var_temp_sandbox_root.set("")
            self.var_sandbox_min_free_gb.set("10")
            self.var_sandbox_low_space_policy.set("block")
            self._on_toggle_sandbox()
            section_title_key = "grp_cfg_convert"
        elif section == "ai":
            self.var_enable_corpus_manifest.set(1)
            self.var_enable_markdown.set(1)
            self.var_markdown_strip_header_footer.set(1)
            self.var_markdown_structured_headings.set(1)
            self.var_enable_markdown_quality_report.set(1)
            self.var_enable_excel_json.set(0)
            self.var_enable_chromadb_export.set(0)
            self.var_enable_markdown_image_manifest.set(1)
            self.var_enable_llm_delivery_hub.set(1)
            self.var_llm_delivery_root.set("")
            self.var_llm_delivery_flatten.set(0)
            self.var_llm_delivery_include_pdf.set(0)
            self.var_enable_upload_readme.set(1)
            self.var_enable_upload_json_manifest.set(1)
            self.var_upload_dedup_merged.set(1)
            self.var_enable_gdrive_upload.set(0)
            self.var_gdrive_client_secrets_path.set("")
            self.var_gdrive_folder_id.set("")
            self.var_mshelpviewer_folder_name.set("MSHelpViewer")
            self.var_enable_mshelp_merge_output.set(1)
            self.var_enable_mshelp_output_docx.set(0)
            self.var_enable_mshelp_output_pdf.set(0)
            section_title_key = "grp_cfg_ai"
        elif section == "incremental":
            self.var_enable_incremental_mode.set(0)
            self.var_incremental_verify_hash.set(0)
            self.var_incremental_reprocess_renamed.set(0)
            self.var_source_priority_skip_same_name_pdf.set(0)
            self.var_global_md5_dedup.set(0)
            self.var_enable_update_package.set(1)
            section_title_key = "grp_cfg_incremental"
        elif section == "merge":
            self.var_enable_merge.set(1)
            self.var_output_enable_pdf.set(1)
            self.var_output_enable_md.set(1)
            self.var_output_enable_merged.set(1)
            self.var_output_enable_independent.set(0)
            self.var_merge_convert_submode.set(MERGE_CONVERT_SUBMODE_MERGE_ONLY)
            self.var_merge_mode.set(MERGE_MODE_CATEGORY)
            self.var_merge_source.set("source")
            self.var_enable_merge_index.set(0)
            self.var_enable_merge_excel.set(0)
            self.var_max_merge_size_mb.set("80")
            section_title_key = "grp_cfg_merge"
        elif section == "ui":
            self.reset_tooltip_settings()
            section_title_key = "grp_cfg_ui"
        elif section == "rules":
            self._set_text_widget_lines(
                self.txt_excluded_folders, ["temp", "backup", "archive"]
            )
            self._set_text_widget_lines(
                self.txt_price_keywords, ["报价", "价格表", "Price", "Quotation"]
            )
            section_title_key = "grp_cfg_rules"
        else:
            return

        self._on_run_mode_change()
        self._on_toggle_sandbox()
        self._on_toggle_incremental_mode()
        self.validate_runtime_inputs(silent=True, scope="all")
        self._refresh_config_dirty_state()

        section_title = self.tr(section_title_key) if section_title_key else section
        reset_message = self.tr("msg_cfg_section_reset").format(section_title)
        if hasattr(self, "var_status"):
            self.var_status.set(reset_message)
        if hasattr(self, "var_locator_result"):
            self.var_locator_result.set(reset_message)

