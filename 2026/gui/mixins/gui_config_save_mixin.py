# -*- coding: utf-8 -*-
"""Config write/save methods extracted from config IO mixin."""

import json
import os
from tkinter import messagebox

from office_converter import (
    MODE_CONVERT_THEN_MERGE,
    MERGE_CONVERT_SUBMODE_MERGE_ONLY,
    MERGE_MODE_CATEGORY,
)


class ConfigSaveMixin:
    def _load_config_for_write(self):
        cfg = {}
        if os.path.exists(self.config_path):
            try:
                with open(self.config_path, "r", encoding="utf-8") as f:
                    cfg = json.load(f)
            except Exception:
                cfg = {}
        return cfg

    def _write_config_sections_to_cfg(self, cfg, section_names):
        sections = set(section_names or [])
        if "shared" in sections:
            cfg["kill_process_mode"] = self.var_kill_mode.get()
            cfg["log_folder"] = self.var_log_folder.get().strip() or "./logs"

        if "convert" in sections:
            cfg["timeout_seconds"] = self._safe_positive_int(
                self.var_timeout_seconds.get(), 60
            )
            cfg["pdf_wait_seconds"] = self._safe_positive_int(
                self.var_pdf_wait_seconds.get(), 15
            )
            cfg["ppt_timeout_seconds"] = self._safe_positive_int(
                self.var_ppt_timeout_seconds.get(), 180
            )
            cfg["ppt_pdf_wait_seconds"] = self._safe_positive_int(
                self.var_ppt_pdf_wait_seconds.get(), 30
            )
            cfg["office_reuse_app"] = bool(self.var_office_reuse_app.get())
            cfg["office_restart_every_n_files"] = self._safe_positive_int(
                self.var_office_restart_every_n_files.get(), 25
            )
            cfg["enable_sandbox"] = bool(self.var_enable_sandbox.get())
            cfg["temp_sandbox_root"] = self.var_temp_sandbox_root.get().strip()
            cfg["sandbox_min_free_gb"] = self._safe_positive_int(
                self.var_sandbox_min_free_gb.get(), 10
            )
            cfg["sandbox_low_space_policy"] = (
                self.var_sandbox_low_space_policy.get() or "block"
            )

        if "ai" in sections:
            cfg["enable_corpus_manifest"] = bool(self.var_enable_corpus_manifest.get())
            cfg.pop("enable_markdown", None)
            cfg["markdown_strip_header_footer"] = bool(
                self.var_markdown_strip_header_footer.get()
            )
            cfg["markdown_structured_headings"] = bool(
                self.var_markdown_structured_headings.get()
            )
            cfg["enable_markdown_quality_report"] = bool(
                self.var_enable_markdown_quality_report.get()
            )
            cfg["enable_excel_json"] = bool(self.var_enable_excel_json.get())
            cfg["enable_chromadb_export"] = bool(self.var_enable_chromadb_export.get())
            cfg["mshelpviewer_folder_name"] = (
                self.var_mshelpviewer_folder_name.get().strip() or "MSHelpViewer"
            )
            cfg["enable_mshelp_merge_output"] = bool(
                self.var_enable_mshelp_merge_output.get()
            )
            cfg["enable_mshelp_output_docx"] = bool(
                self.var_enable_mshelp_output_docx.get()
            )
            cfg["enable_mshelp_output_pdf"] = bool(
                self.var_enable_mshelp_output_pdf.get()
            )

        if "incremental" in sections:
            cfg["enable_incremental_mode"] = bool(
                self.var_enable_incremental_mode.get()
            )
            cfg["incremental_verify_hash"] = bool(
                self.var_incremental_verify_hash.get()
            )
            cfg["incremental_reprocess_renamed"] = bool(
                self.var_incremental_reprocess_renamed.get()
            )
            cfg["source_priority_skip_same_name_pdf"] = bool(
                self.var_source_priority_skip_same_name_pdf.get()
            )
            cfg["global_md5_dedup"] = bool(self.var_global_md5_dedup.get())
            cfg["enable_update_package"] = bool(self.var_enable_update_package.get())
            cfg["enable_parallel_conversion"] = bool(
                self.var_enable_parallel_conversion.get()
            )
            cfg["parallel_workers"] = self._safe_positive_int(
                self.var_parallel_workers.get(), 4
            )
            cfg["enable_checkpoint"] = bool(self.var_enable_checkpoint.get())
            cfg["checkpoint_auto_resume"] = bool(self.var_checkpoint_auto_resume.get())

        if "merge" in sections:
            cfg["enable_merge"] = bool(self.var_enable_merge.get())
            cfg["output_enable_pdf"] = bool(self.var_output_enable_pdf.get())
            cfg["output_enable_md"] = bool(self.var_output_enable_md.get())
            cfg["output_enable_merged"] = bool(self.var_output_enable_merged.get())
            cfg["output_enable_independent"] = bool(
                self.var_output_enable_independent.get()
            )
            cfg["merge_convert_submode"] = self.var_merge_convert_submode.get()
            cfg["merge_mode"] = self.var_merge_mode.get()
            cfg["merge_source"] = (
                "target"
                if self.var_run_mode.get() == MODE_CONVERT_THEN_MERGE
                else self.var_merge_source.get()
            )
            cfg["enable_merge_index"] = bool(self.var_enable_merge_index.get())
            cfg["enable_merge_excel"] = bool(self.var_enable_merge_excel.get())
            cfg["max_merge_size_mb"] = self._safe_positive_int(
                self.var_max_merge_size_mb.get(), 80
            )
            cfg["merge_filename_pattern"] = (
                self.var_merge_filename_pattern.get().strip()
                or "Merged_{category}_{timestamp}_{idx}"
            )

        if "rules" in sections:
            cfg["excluded_folders"] = self._read_text_lines(self.txt_excluded_folders)
            cfg["price_keywords"] = self._read_text_lines(self.txt_price_keywords)

        if "ui" in sections:
            self.apply_tooltip_settings(silent=True)
            cfg["ui"] = {
                "tooltip_delay_ms": self.tooltip_delay_ms,
                "tooltip_bg": self.tooltip_bg,
                "tooltip_fg": self.tooltip_fg,
                "tooltip_font_family": self.tooltip_font_family,
                "tooltip_font_size": self.tooltip_font_size,
                "tooltip_auto_theme": self.tooltip_auto_theme,
                "confirm_revert_dirty": bool(self.var_confirm_revert_dirty.get()),
                "task_current_config_only": (
                    bool(self.var_task_scope_current_config_only.get())
                    if hasattr(self, "var_task_scope_current_config_only")
                    else True
                ),
                # 鐠佹澘绻傜粣妤€褰涚亸鍝勵嚟娑撳簼缍呯純?
                "window_geometry": self.geometry(),
                "window_state": self.state(),
            }
        return cfg

    def _save_specific_config_section(self, section_name, show_msg=True):
        section = str(section_name or "").strip().lower()
        valid_sections = {
            "shared",
            "convert",
            "ai",
            "incremental",
            "merge",
            "ui",
            "rules",
        }
        if section not in valid_sections:
            return
        cfg = self._load_config_for_write()
        self._write_config_sections_to_cfg(cfg, [section])
        try:
            with open(self.config_path, "w", encoding="utf-8") as f:
                json.dump(cfg, f, indent=4, ensure_ascii=False)
            self._baseline_config_snapshot = self._build_config_snapshot_from_cfg(cfg)
            self._refresh_config_dirty_state()
            section_title = ", ".join(self._get_cfg_section_titles([section]))
            msg = self.tr("msg_cfg_section_saved").format(section_title)
            if show_msg:
                messagebox.showinfo(self.tr("btn_save_cfg_section"), msg)
            if hasattr(self, "var_status"):
                self.var_status.set(msg)
            if hasattr(self, "var_locator_result"):
                self.var_locator_result.set(msg)
        except Exception as e:
            if show_msg:
                messagebox.showerror(
                    self.tr("btn_save_cfg_section"),
                    self.tr("msg_save_fail").format(e),
                )

    def _on_close_main_window(self):
        """Persist UI geometry settings and close the window."""
        try:
            cfg = self._load_config_for_write()
            self._write_config_sections_to_cfg(cfg, ["ui"])
            with open(self.config_path, "w", encoding="utf-8") as f:
                json.dump(cfg, f, indent=4, ensure_ascii=False)
        except Exception:
            pass
        try:
            poll_id = getattr(self, "_after_poll_log_id", None)
            if poll_id:
                self.after_cancel(poll_id)
        except Exception:
            pass
        try:
            for aid in list(getattr(self, "_after_force_refresh_ids", []) or []):
                try:
                    self.after_cancel(aid)
                except Exception:
                    pass
        except Exception:
            pass
        try:
            self.destroy()
        except Exception:
            pass

    def _save_dirty_config_sections(self, show_msg=True):
        dirty_sections = []
        for section_name, _, _ in self._cfg_tab_meta:
            if self._last_section_dirty.get(section_name, False):
                dirty_sections.append(section_name)
        if not dirty_sections:
            self._refresh_config_dirty_state()
            msg = self.tr("msg_save_dirty_none")
            if show_msg:
                messagebox.showinfo(self.tr("btn_save_cfg_dirty"), msg)
            if hasattr(self, "var_status"):
                self.var_status.set(msg)
            if hasattr(self, "var_locator_result"):
                self.var_locator_result.set(msg)
            return

        cfg = self._load_config_for_write()
        self._write_config_sections_to_cfg(cfg, dirty_sections)

        try:
            with open(self.config_path, "w", encoding="utf-8") as f:
                json.dump(cfg, f, indent=4, ensure_ascii=False)
            self._baseline_config_snapshot = self._build_config_snapshot_from_cfg(cfg)
            self._refresh_config_dirty_state()
            saved_sections_text = ", ".join(
                self._get_cfg_section_titles(dirty_sections)
            )
            msg = self.tr("msg_save_dirty_sections").format(saved_sections_text)
            if show_msg:
                messagebox.showinfo(self.tr("btn_save_cfg_dirty"), msg)
            if hasattr(self, "var_status"):
                self.var_status.set(msg)
            if hasattr(self, "var_locator_result"):
                self.var_locator_result.set(msg)
        except Exception as e:
            if show_msg:
                messagebox.showerror(
                    self.tr("btn_save_cfg_dirty"),
                    self.tr("msg_save_fail").format(e),
                )

    def _save_settings_to_file(self, show_msg=True, scope="all"):
        """Save current UI values into the active config file."""
        cfg = self._compose_config_from_ui(self._load_config_for_write(), scope=scope)
        try:
            with open(self.config_path, "w", encoding="utf-8") as f:
                json.dump(cfg, f, indent=4, ensure_ascii=False)
            self._baseline_config_snapshot = self._build_config_snapshot_from_cfg(cfg)
            self._refresh_config_dirty_state()
            if show_msg:
                messagebox.showinfo(self.tr("btn_save_cfg"), self.tr("msg_save_ok"))
        except Exception as e:
            if show_msg:
                messagebox.showerror(
                    self.tr("btn_save_cfg"),
                    self.tr("msg_save_fail").format(e),
                )

