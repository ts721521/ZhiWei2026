# -*- coding: utf-8 -*-
"""Config compose methods extracted from config save mixin."""

import sys

from office_converter import (
    MODE_CONVERT_ONLY,
    MODE_MERGE_ONLY,
    MODE_CONVERT_THEN_MERGE,
    MODE_COLLECT_ONLY,
    MODE_MSHELP_ONLY,
)


class ConfigComposeMixin:
    def _compose_config_from_ui(self, cfg, scope="all"):
        cfg = cfg if isinstance(cfg, dict) else {}
        scope = "mode" if str(scope).lower() == "mode" else "all"
        mode = self.var_run_mode.get()
        write_convert = scope == "all" or mode in (
            MODE_CONVERT_ONLY,
            MODE_CONVERT_THEN_MERGE,
        )
        write_merge = scope == "all" or mode in (
            MODE_CONVERT_THEN_MERGE,
            MODE_MERGE_ONLY,
        )
        write_collect = scope == "all" or mode == MODE_COLLECT_ONLY
        write_mshelp = scope == "all" or mode == MODE_MSHELP_ONLY
        write_rules = scope == "all" or mode in (
            MODE_CONVERT_ONLY,
            MODE_CONVERT_THEN_MERGE,
            MODE_COLLECT_ONLY,
        )

        is_win = sys.platform == "win32"
        is_mac = sys.platform == "darwin"

        if is_win:
            cfg["source_folders_win"] = self.source_folders_list
            cfg["source_folder_win"] = (
                self.source_folders_list[0] if self.source_folders_list else ""
            )
            cfg["target_folder_win"] = self.var_target_folder.get().strip()
            if write_convert:
                cfg["temp_sandbox_root_win"] = self.var_temp_sandbox_root.get().strip()
                cfg["llm_delivery_root_win"] = self.var_llm_delivery_root.get().strip()
        elif is_mac:
            cfg["source_folders_mac"] = self.source_folders_list
            cfg["source_folder_mac"] = (
                self.source_folders_list[0] if self.source_folders_list else ""
            )
            cfg["target_folder_mac"] = self.var_target_folder.get().strip()
            if write_convert:
                cfg["temp_sandbox_root_mac"] = self.var_temp_sandbox_root.get().strip()
                cfg["llm_delivery_root_mac"] = self.var_llm_delivery_root.get().strip()

        if hasattr(self, "var_app_mode"):
            cfg["app_mode"] = self.var_app_mode.get()
        else:
            cfg["app_mode"] = "classic"
        cfg["source_folders"] = self.source_folders_list
        cfg["source_folder"] = (
            self.source_folders_list[0] if self.source_folders_list else ""
        )
        cfg["target_folder"] = self.var_target_folder.get().strip()
        cfg["output_enable_pdf"] = bool(self.var_output_enable_pdf.get())
        cfg["output_enable_md"] = bool(self.var_output_enable_md.get())
        cfg["output_enable_merged"] = bool(self.var_output_enable_merged.get())
        cfg["output_enable_independent"] = bool(
            self.var_output_enable_independent.get()
        )
        cfg["merge_convert_submode"] = self.var_merge_convert_submode.get()
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
        cfg["enable_fast_md_engine"] = bool(self.var_enable_fast_md_engine.get())
        cfg["enable_traceability_anchor_and_map"] = bool(
            self.var_enable_traceability_anchor_and_map.get()
        )
        cfg["enable_prompt_wrapper"] = bool(self.var_enable_prompt_wrapper.get())
        cfg["prompt_template_type"] = (
            self.var_prompt_template_type.get().strip() or "new_solution"
        )
        cfg["short_id_prefix"] = self.var_short_id_prefix.get().strip().upper() or "ZW-"
        if write_mshelp:
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
        if write_convert:
            cfg["enable_sandbox"] = bool(self.var_enable_sandbox.get())
            cfg["temp_sandbox_root"] = self.var_temp_sandbox_root.get().strip()
            cfg["sandbox_min_free_gb"] = self._safe_positive_int(
                self.var_sandbox_min_free_gb.get(), 10
            )
            cfg["sandbox_low_space_policy"] = (
                self.var_sandbox_low_space_policy.get() or "block"
            )
            cfg["enable_llm_delivery_hub"] = bool(
                self.var_enable_llm_delivery_hub.get()
            )
            cfg["llm_delivery_root"] = self.var_llm_delivery_root.get().strip()
            cfg["llm_delivery_flatten"] = bool(self.var_llm_delivery_flatten.get())
            cfg["llm_delivery_include_pdf"] = bool(
                self.var_llm_delivery_include_pdf.get()
            )
            cfg["enable_upload_readme"] = bool(self.var_enable_upload_readme.get())
            cfg["enable_upload_json_manifest"] = bool(
                self.var_enable_upload_json_manifest.get()
            )
            cfg["upload_dedup_merged"] = bool(self.var_upload_dedup_merged.get())
            cfg["enable_gdrive_upload"] = bool(self.var_enable_gdrive_upload.get())
            cfg["gdrive_client_secrets_path"] = (
                self.var_gdrive_client_secrets_path.get().strip()
            )
            cfg["gdrive_folder_id"] = self.var_gdrive_folder_id.get().strip()
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

        if write_merge:
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
                if mode == MODE_CONVERT_THEN_MERGE
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

        cfg["run_mode"] = self.var_run_mode.get()
        if write_collect:
            cfg["collect_mode"] = self.var_collect_mode.get()
        if write_convert:
            cfg["content_strategy"] = self.var_strategy.get()

        if write_convert:
            cfg["default_engine"] = self.var_engine.get()
        cfg["kill_process_mode"] = self.var_kill_mode.get()
        cfg["log_folder"] = self.var_log_folder.get().strip() or "./logs"

        if write_rules:
            excluded_text = self.txt_excluded_folders.get("1.0", "end").strip()
            cfg["excluded_folders"] = [
                line.strip() for line in excluded_text.splitlines() if line.strip()
            ]

            kw_text = self.txt_price_keywords.get("1.0", "end").strip()
            cfg["price_keywords"] = [
                line.strip() for line in kw_text.splitlines() if line.strip()
            ]

        def _to_int(var, default):
            try:
                v = int(var.get().strip())
                return v if v > 0 else default
            except Exception:
                return default

        if write_convert:
            cfg["timeout_seconds"] = _to_int(self.var_timeout_seconds, 60)
            cfg["pdf_wait_seconds"] = _to_int(self.var_pdf_wait_seconds, 15)
            cfg["ppt_timeout_seconds"] = _to_int(self.var_ppt_timeout_seconds, 180)
            cfg["ppt_pdf_wait_seconds"] = _to_int(self.var_ppt_pdf_wait_seconds, 30)
            cfg["office_reuse_app"] = bool(self.var_office_reuse_app.get())
            cfg["office_restart_every_n_files"] = _to_int(
                self.var_office_restart_every_n_files, 25
            )
        if write_merge:
            cfg["max_merge_size_mb"] = _to_int(self.var_max_merge_size_mb, 80)
        cfg["enable_merge_map"] = cfg.get("enable_merge_map", True)
        cfg["bookmark_with_short_id"] = cfg.get("bookmark_with_short_id", True)
        cfg["everything"] = cfg.get(
            "everything",
            {
                "enabled": True,
                "es_path": "",
                "prefer_path_exact": True,
                "timeout_ms": 1500,
            },
        )
        cfg["listary"] = cfg.get(
            "listary",
            {
                "enabled": True,
                "copy_query_on_locate": True,
            },
        )
        cfg["privacy"] = cfg.get("privacy", {"mask_md5_in_logs": True})
        self.apply_tooltip_settings(silent=True)
        task_scope_current_only = True
        if hasattr(self, "var_task_scope_current_config_only"):
            try:
                task_scope_current_only = bool(
                    self.var_task_scope_current_config_only.get()
                )
            except Exception:
                task_scope_current_only = True
        cfg["ui"] = {
            "tooltip_delay_ms": self.tooltip_delay_ms,
            "tooltip_bg": self.tooltip_bg,
            "tooltip_fg": self.tooltip_fg,
            "tooltip_font_family": self.tooltip_font_family,
            "tooltip_font_size": self.tooltip_font_size,
            "tooltip_auto_theme": self.tooltip_auto_theme,
            "confirm_revert_dirty": bool(self.var_confirm_revert_dirty.get()),
            "task_current_config_only": task_scope_current_only,
        }
        return cfg

