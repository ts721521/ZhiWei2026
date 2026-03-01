# -*- coding: utf-8 -*-
"""Config dirty-state methods extracted from config logic mixin."""

from tkinter import messagebox

from office_converter import (
    MERGE_CONVERT_SUBMODE_MERGE_ONLY,
    MERGE_MODE_CATEGORY,
    KILL_MODE_AUTO,
)


class ConfigDirtyStateMixin:
    def _set_config_dirty(self, dirty=True):
        self.cfg_dirty = bool(dirty)
        if hasattr(self, "lbl_cfg_dirty_state"):
            if self.cfg_dirty:
                self.lbl_cfg_dirty_state.configure(
                    text=self.tr("lbl_cfg_dirty_unsaved"),
                    bootstyle="warning",
                )
            else:
                self.lbl_cfg_dirty_state.configure(
                    text=self.tr("lbl_cfg_dirty_clean"),
                    bootstyle="success",
                )
        if not self.cfg_dirty:
            self._update_config_dirty_summary({})

    def _update_config_tab_dirty_markers(self, section_dirty=None):
        if not hasattr(self, "main_notebook"):
            return
        if not self._cfg_tab_meta:
            return
        section_dirty = section_dirty or {}
        # 婢舵矮閲?section 閸欘垵鍏橀弰鐘茬殸閸掓澘鎮撴稉鈧稉顏嗗⒖閻?tab閿涘矁浠涢崥?dirty 閻樿埖鈧?
        tab_dirty = {}
        tab_labels = {}
        for section_name, tab_widget, label_key in self._cfg_tab_meta:
            if id(tab_widget) not in tab_dirty:
                tab_dirty[id(tab_widget)] = False
                tab_labels[id(tab_widget)] = (tab_widget, label_key)
            if section_dirty.get(section_name, False):
                tab_dirty[id(tab_widget)] = True
        for tid, is_dirty in tab_dirty.items():
            tab_widget, label_key = tab_labels[tid]
            try:
                title = self.tr(label_key)
                if is_dirty:
                    title = f"{title} *"
                self.main_notebook.tab(tab_widget, text=title)
            except Exception:
                pass

    def _compute_section_dirty(self, ui_snapshot, base_snapshot):
        section_fields = {
            "shared": ["kill_process_mode", "log_folder"],
            "convert": [
                "timeout_seconds",
                "pdf_wait_seconds",
                "ppt_timeout_seconds",
                "ppt_pdf_wait_seconds",
                "office_reuse_app",
                "office_restart_every_n_files",
                "enable_sandbox",
                "temp_sandbox_root",
                "sandbox_min_free_gb",
                "sandbox_low_space_policy",
            ],
            "ai": [
                "enable_corpus_manifest",
                "enable_markdown",
                "markdown_strip_header_footer",
                "markdown_structured_headings",
                "enable_markdown_quality_report",
                "enable_excel_json",
                "enable_chromadb_export",
                "enable_fast_md_engine",
                "enable_traceability_anchor_and_map",
                "enable_markdown_image_manifest",
                "enable_prompt_wrapper",
                "prompt_template_type",
                "short_id_prefix",
                "enable_llm_delivery_hub",
                "llm_delivery_root",
                "llm_delivery_flatten",
                "llm_delivery_include_pdf",
                "enable_gdrive_upload",
                "gdrive_client_secrets_path",
                "gdrive_folder_id",
                "mshelpviewer_folder_name",
                "enable_mshelp_merge_output",
                "enable_mshelp_output_docx",
                "enable_mshelp_output_pdf",
            ],
            "incremental": [
                "enable_incremental_mode",
                "incremental_verify_hash",
                "incremental_reprocess_renamed",
                "source_priority_skip_same_name_pdf",
                "global_md5_dedup",
                "enable_update_package",
            ],
            "merge": [
                "enable_merge",
                "output_enable_pdf",
                "output_enable_md",
                "output_enable_merged",
                "output_enable_independent",
                "merge_convert_submode",
                "merge_mode",
                "merge_source",
                "enable_merge_index",
                "enable_merge_excel",
                "max_merge_size_mb",
            ],
            "rules": ["excluded_folders", "price_keywords"],
        }
        section_dirty = {}
        for section_name, keys in section_fields.items():
            section_dirty[section_name] = any(
                ui_snapshot.get(k) != base_snapshot.get(k) for k in keys
            )
        section_dirty["ui"] = ui_snapshot.get("ui", {}) != base_snapshot.get("ui", {})
        return section_dirty

    def _update_config_dirty_summary(self, section_dirty):
        if not hasattr(self, "lbl_cfg_dirty_sections"):
            return
        section_dirty = section_dirty or {}
        self._last_section_dirty = dict(section_dirty)
        dirty_names = []
        for section_name, _, label_key in self._cfg_tab_meta:
            if section_dirty.get(section_name, False):
                dirty_names.append(self.tr(label_key))
        if dirty_names:
            self.lbl_cfg_dirty_sections.configure(
                text=self.tr("lbl_cfg_dirty_sections").format(", ".join(dirty_names)),
                bootstyle="warning",
            )
        else:
            self.lbl_cfg_dirty_sections.configure(
                text=self.tr("lbl_cfg_dirty_none"),
                bootstyle="secondary",
            )
        can_act = bool(dirty_names) and (not self._ui_running)
        if hasattr(self, "btn_save_cfg_dirty"):
            count = len(dirty_names)
            if count > 0:
                self.btn_save_cfg_dirty.configure(
                    text=self.tr("btn_save_cfg_dirty_count").format(count)
                )
            else:
                self.btn_save_cfg_dirty.configure(text=self.tr("btn_save_cfg_dirty"))
            self.btn_save_cfg_dirty.configure(
                state=("normal" if can_act else "disabled")
            )

    def _focus_dirty_section(self, section_name):
        if not hasattr(self, "main_notebook"):
            return
        if not self._cfg_tab_meta:
            return
        for section, tab_widget, _ in self._cfg_tab_meta:
            if section == section_name:
                try:
                    self.main_notebook.select(tab_widget)
                except Exception:
                    pass
                return

    def _get_cfg_section_titles(self, section_names):
        titles = []
        section_set = set(section_names or [])
        for section_name, _, label_key in self._cfg_tab_meta:
            if section_name in section_set:
                titles.append(self.tr(label_key))
        return titles

    def _apply_snapshot_sections_to_ui(self, snapshot, section_names):
        snapshot = snapshot or {}
        sections = set(section_names or [])
        if "shared" in sections:
            self.var_kill_mode.set(snapshot.get("kill_process_mode", KILL_MODE_AUTO))
            self.var_log_folder.set(snapshot.get("log_folder", "./logs"))
        if "convert" in sections:
            self.var_timeout_seconds.set(str(snapshot.get("timeout_seconds", 60)))
            self.var_pdf_wait_seconds.set(str(snapshot.get("pdf_wait_seconds", 15)))
            self.var_ppt_timeout_seconds.set(
                str(snapshot.get("ppt_timeout_seconds", 180))
            )
            self.var_ppt_pdf_wait_seconds.set(
                str(snapshot.get("ppt_pdf_wait_seconds", 30))
            )
            self.var_office_reuse_app.set(
                1 if snapshot.get("office_reuse_app", True) else 0
            )
            self.var_office_restart_every_n_files.set(
                str(snapshot.get("office_restart_every_n_files", 25))
            )
            self.var_enable_sandbox.set(
                1 if snapshot.get("enable_sandbox", True) else 0
            )
            self.var_temp_sandbox_root.set(snapshot.get("temp_sandbox_root", "") or "")
            self.var_sandbox_min_free_gb.set(
                str(snapshot.get("sandbox_min_free_gb", 10))
            )
            self.var_sandbox_low_space_policy.set(
                snapshot.get("sandbox_low_space_policy", "block") or "block"
            )
            self._on_toggle_sandbox()
        if "ai" in sections:
            self.var_enable_corpus_manifest.set(
                1 if snapshot.get("enable_corpus_manifest", True) else 0
            )
            self.var_enable_markdown.set(
                1 if snapshot.get("enable_markdown", True) else 0
            )
            self.var_markdown_strip_header_footer.set(
                1 if snapshot.get("markdown_strip_header_footer", True) else 0
            )
            self.var_markdown_structured_headings.set(
                1 if snapshot.get("markdown_structured_headings", True) else 0
            )
            self.var_enable_markdown_quality_report.set(
                1 if snapshot.get("enable_markdown_quality_report", True) else 0
            )
            self.var_enable_excel_json.set(
                1 if snapshot.get("enable_excel_json", False) else 0
            )
            self.var_enable_chromadb_export.set(
                1 if snapshot.get("enable_chromadb_export", False) else 0
            )
            self.var_enable_fast_md_engine.set(
                1 if snapshot.get("enable_fast_md_engine", False) else 0
            )
            self.var_enable_traceability_anchor_and_map.set(
                1 if snapshot.get("enable_traceability_anchor_and_map", True) else 0
            )
            self.var_enable_markdown_image_manifest.set(
                1 if snapshot.get("enable_markdown_image_manifest", True) else 0
            )
            self.var_enable_prompt_wrapper.set(
                1 if snapshot.get("enable_prompt_wrapper", False) else 0
            )
            self.var_prompt_template_type.set(
                str(snapshot.get("prompt_template_type", "new_solution") or "new_solution")
            )
            self.var_short_id_prefix.set(
                str(snapshot.get("short_id_prefix", "ZW-") or "ZW-")
            )
            self.var_enable_llm_delivery_hub.set(
                1 if snapshot.get("enable_llm_delivery_hub", True) else 0
            )
            self.var_llm_delivery_root.set(snapshot.get("llm_delivery_root", "") or "")
            self.var_llm_delivery_flatten.set(
                1 if snapshot.get("llm_delivery_flatten", False) else 0
            )
            self.var_llm_delivery_include_pdf.set(
                1 if snapshot.get("llm_delivery_include_pdf", False) else 0
            )
            self.var_enable_upload_readme.set(
                1 if snapshot.get("enable_upload_readme", True) else 0
            )
            self.var_enable_upload_json_manifest.set(
                1 if snapshot.get("enable_upload_json_manifest", True) else 0
            )
            self.var_upload_dedup_merged.set(
                1 if snapshot.get("upload_dedup_merged", True) else 0
            )
            self.var_enable_gdrive_upload.set(
                1 if snapshot.get("enable_gdrive_upload", False) else 0
            )
            self.var_gdrive_client_secrets_path.set(
                snapshot.get("gdrive_client_secrets_path", "") or ""
            )
            self.var_gdrive_folder_id.set(snapshot.get("gdrive_folder_id", "") or "")
            self.var_mshelpviewer_folder_name.set(
                str(
                    snapshot.get("mshelpviewer_folder_name", "MSHelpViewer")
                    or "MSHelpViewer"
                )
            )
            self.var_enable_mshelp_merge_output.set(
                1 if snapshot.get("enable_mshelp_merge_output", True) else 0
            )
            self.var_enable_mshelp_output_docx.set(
                1 if snapshot.get("enable_mshelp_output_docx", False) else 0
            )
            self.var_enable_mshelp_output_pdf.set(
                1 if snapshot.get("enable_mshelp_output_pdf", False) else 0
            )
        if "incremental" in sections:
            self.var_enable_incremental_mode.set(
                1 if snapshot.get("enable_incremental_mode", False) else 0
            )
            self.var_incremental_verify_hash.set(
                1 if snapshot.get("incremental_verify_hash", False) else 0
            )
            self.var_incremental_reprocess_renamed.set(
                1 if snapshot.get("incremental_reprocess_renamed", False) else 0
            )
            self.var_source_priority_skip_same_name_pdf.set(
                1 if snapshot.get("source_priority_skip_same_name_pdf", False) else 0
            )
            self.var_global_md5_dedup.set(
                1 if snapshot.get("global_md5_dedup", False) else 0
            )
            self.var_enable_update_package.set(
                1 if snapshot.get("enable_update_package", True) else 0
            )
            self.var_enable_parallel_conversion.set(
                1 if snapshot.get("enable_parallel_conversion", False) else 0
            )
            self.var_parallel_workers.set(str(snapshot.get("parallel_workers", 4)))
            self.var_enable_checkpoint.set(
                1 if snapshot.get("enable_checkpoint", True) else 0
            )
            self.var_checkpoint_auto_resume.set(
                1 if snapshot.get("checkpoint_auto_resume", True) else 0
            )
            self._on_toggle_parallel_conversion()
        if "merge" in sections:
            self.var_enable_merge.set(1 if snapshot.get("enable_merge", True) else 0)
            self.var_output_enable_pdf.set(
                1 if snapshot.get("output_enable_pdf", True) else 0
            )
            self.var_output_enable_md.set(
                1 if snapshot.get("output_enable_md", True) else 0
            )
            self.var_output_enable_merged.set(
                1 if snapshot.get("output_enable_merged", True) else 0
            )
            self.var_output_enable_independent.set(
                1 if snapshot.get("output_enable_independent", False) else 0
            )
            self.var_merge_convert_submode.set(
                snapshot.get(
                    "merge_convert_submode",
                    MERGE_CONVERT_SUBMODE_MERGE_ONLY,
                )
            )
            self.var_merge_mode.set(snapshot.get("merge_mode", MERGE_MODE_CATEGORY))
            self.var_merge_source.set(snapshot.get("merge_source", "source"))
            self.var_enable_merge_index.set(
                1 if snapshot.get("enable_merge_index", False) else 0
            )
            self.var_enable_merge_excel.set(
                1 if snapshot.get("enable_merge_excel", False) else 0
            )
            self.var_max_merge_size_mb.set(str(snapshot.get("max_merge_size_mb", 80)))
            self.var_merge_filename_pattern.set(
                snapshot.get("merge_filename_pattern")
                or "Merged_{category}_{timestamp}_{idx}"
            )
        if "rules" in sections:
            self._set_text_widget_lines(
                self.txt_excluded_folders, snapshot.get("excluded_folders", [])
            )
            self._set_text_widget_lines(
                self.txt_price_keywords, snapshot.get("price_keywords", [])
            )
        if "ui" in sections:
            ui_snapshot = (
                snapshot.get("ui", {}) if isinstance(snapshot.get("ui"), dict) else {}
            )
            self.var_tooltip_delay_ms.set(
                str(
                    ui_snapshot.get(
                        "tooltip_delay_ms", self.TOOLTIP_DEFAULTS["tooltip_delay_ms"]
                    )
                )
            )
            self.var_tooltip_font_size.set(
                str(
                    ui_snapshot.get(
                        "tooltip_font_size", self.TOOLTIP_DEFAULTS["tooltip_font_size"]
                    )
                )
            )
            self.var_tooltip_bg.set(
                ui_snapshot.get("tooltip_bg", self.TOOLTIP_DEFAULTS["tooltip_bg"])
            )
            self.var_tooltip_fg.set(
                ui_snapshot.get("tooltip_fg", self.TOOLTIP_DEFAULTS["tooltip_fg"])
            )
            self.var_tooltip_auto_theme.set(
                1
                if ui_snapshot.get(
                    "tooltip_auto_theme", self.TOOLTIP_DEFAULTS["tooltip_auto_theme"]
                )
                else 0
            )
            self.var_confirm_revert_dirty.set(
                1 if ui_snapshot.get("confirm_revert_dirty", True) else 0
            )
            self.apply_tooltip_settings(silent=True)
            self.validate_tooltip_inputs(silent=True)
        self._sync_markdown_master_with_global_output()

    def _revert_dirty_config_sections(self, show_msg=True):
        self._refresh_config_dirty_from_file()
        dirty_sections = []
        for section_name, _, _ in self._cfg_tab_meta:
            if self._last_section_dirty.get(section_name, False):
                dirty_sections.append(section_name)
        if not dirty_sections:
            msg = self.tr("msg_revert_dirty_none")
            if show_msg:
                messagebox.showinfo(self.tr("btn_revert_cfg_dirty"), msg)
            if hasattr(self, "var_status"):
                self.var_status.set(msg)
            if hasattr(self, "var_locator_result"):
                self.var_locator_result.set(msg)
            return
        section_titles = self._get_cfg_section_titles(dirty_sections)
        section_text = ", ".join(section_titles)
        need_confirm = hasattr(self, "var_confirm_revert_dirty") and bool(
            self.var_confirm_revert_dirty.get()
        )
        if show_msg and need_confirm:
            confirm = messagebox.askyesno(
                self.tr("btn_revert_cfg_dirty"),
                self.tr("msg_confirm_revert_dirty").format(section_text),
            )
            if not confirm:
                return
        self._suspend_cfg_dirty = True
        try:
            self._apply_snapshot_sections_to_ui(
                self._baseline_config_snapshot, dirty_sections
            )
            self._on_toggle_incremental_mode()
            self.validate_runtime_inputs(silent=True, scope="config")
        finally:
            self._suspend_cfg_dirty = False
        self._refresh_config_dirty_state()
        msg = self.tr("msg_revert_dirty_sections").format(section_text)
        if show_msg:
            messagebox.showinfo(self.tr("btn_revert_cfg_dirty"), msg)
        if hasattr(self, "var_status"):
            self.var_status.set(msg)
        if hasattr(self, "var_locator_result"):
            self.var_locator_result.set(msg)

    def _focus_first_dirty_section(self):
        if not hasattr(self, "main_notebook"):
            return
        if not self._cfg_tab_meta:
            return
        for section_name, tab_widget, _ in self._cfg_tab_meta:
            if self._last_section_dirty.get(section_name, False):
                self._focus_dirty_section(section_name)
                return

    def _refresh_config_dirty_state(self):
        if self._suspend_cfg_dirty:
            return
        if not isinstance(self._baseline_config_snapshot, dict):
            self._baseline_config_snapshot = {}
        ui_snapshot = self._build_config_snapshot_from_ui()
        section_dirty = self._compute_section_dirty(
            ui_snapshot, self._baseline_config_snapshot
        )
        self._set_config_dirty(any(section_dirty.values()))
        self._update_config_tab_dirty_markers(section_dirty)
        self._update_config_dirty_summary(section_dirty)

    def _mark_config_dirty(self):
        if self._suspend_cfg_dirty:
            return
        self._refresh_config_dirty_state()

    def _bind_config_dirty_var(self, var):
        if var is None:
            return
        try:
            var.trace_add("write", lambda *_: self._mark_config_dirty())
        except Exception:
            pass

    def _bind_config_dirty_text(self, widget):
        if widget is None:
            return
        for event_name in ("<KeyRelease>", "<<Paste>>", "<<Cut>>"):
            try:
                widget.bind(event_name, lambda *_: self._mark_config_dirty(), add="+")
            except Exception:
                pass

