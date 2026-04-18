# -*- coding: utf-8 -*-
"""Config file load/save methods extracted from config logic mixin."""

import json
import os
import sys
from tkinter import messagebox
from tkinter.constants import *

from office_converter import (
    MODE_CONVERT_ONLY,
    MODE_MERGE_ONLY,
    MODE_CONVERT_THEN_MERGE,
    MODE_COLLECT_ONLY,
    MODE_MSHELP_ONLY,
    MERGE_CONVERT_SUBMODE_MERGE_ONLY,
    COLLECT_MODE_COPY_AND_INDEX,
    MERGE_MODE_CATEGORY,
    ENGINE_WPS,
    ENGINE_MS,
    KILL_MODE_AUTO,
    KILL_MODE_KEEP,
)


class ConfigIOMixin:
    def _load_config_to_ui(self):
        """Load values from config.json into UI controls."""
        self._suspend_cfg_dirty = True
        if hasattr(self, "var_profile_active_path"):
            self.var_profile_active_path.set(self.config_path)
        if not str(getattr(self, "_active_config_label", "")).strip():
            self._active_config_label = os.path.basename(str(self.config_path or ""))
            self._active_config_origin = str(self.config_path or "")
        if not os.path.exists(self.config_path):
            self._suspend_cfg_dirty = False
            return
        try:
            with open(self.config_path, "r", encoding="utf-8") as f:
                cfg = json.load(f)
        except Exception:
            self._suspend_cfg_dirty = False
            return

        ui_cfg = cfg.get("ui", {}) if isinstance(cfg.get("ui", {}), dict) else {}
        self.tooltip_delay_ms = int(ui_cfg.get("tooltip_delay_ms", 300) or 300)
        self.tooltip_bg = str(ui_cfg.get("tooltip_bg", "#FFF7D6") or "#FFF7D6")
        self.tooltip_fg = str(ui_cfg.get("tooltip_fg", "#202124") or "#202124")
        self.tooltip_font_family = str(
            ui_cfg.get("tooltip_font_family", "System") or "System"
        )
        self.tooltip_font_size = int(ui_cfg.get("tooltip_font_size", 9) or 9)
        self.tooltip_auto_theme = bool(ui_cfg.get("tooltip_auto_theme", True))
        if hasattr(self, "var_tooltip_auto_theme"):
            self.var_tooltip_auto_theme.set(1 if self.tooltip_auto_theme else 0)
        if hasattr(self, "var_confirm_revert_dirty"):
            self.var_confirm_revert_dirty.set(
                1 if ui_cfg.get("confirm_revert_dirty", True) else 0
            )
        if hasattr(self, "var_tooltip_delay_ms"):
            self.var_tooltip_delay_ms.set(str(self.tooltip_delay_ms))
        if hasattr(self, "var_tooltip_font_size"):
            self.var_tooltip_font_size.set(str(self.tooltip_font_size))
        if hasattr(self, "var_tooltip_bg"):
            self.var_tooltip_bg.set(self.tooltip_bg)
        if hasattr(self, "var_tooltip_fg"):
            self.var_tooltip_fg.set(self.tooltip_fg)
        self.apply_tooltip_settings(silent=True)
        self.validate_tooltip_inputs(silent=True)

        # 閹垹顦茬粣妤€褰涚亸鍝勵嚟娑撳簼缍呯純顕嗙礄婵″倹鐏夊韫箽鐎涙﹫绱?
        try:
            win_geo = ui_cfg.get("window_geometry")
            if isinstance(win_geo, str) and win_geo:
                self.geometry(win_geo)
            win_state = ui_cfg.get("window_state")
            if isinstance(win_state, str) and win_state:
                self.state(win_state)
        except Exception:
            pass

        if hasattr(self, "var_app_mode"):
            # 自 2026-03 起，仅保留任务模式作为运行入口；配置中的 app_mode 统一归一为 task。
            try:
                _ = self.var_app_mode.get()  # ensure variable exists
            except Exception:
                pass
            self.var_app_mode.set("task")

        # Runtime parameters

        is_win = sys.platform == "win32"
        is_mac = sys.platform == "darwin"

        def _get_os_path(key_base):
            if is_win:
                val = cfg.get(f"{key_base}_win")
            elif is_mac:
                val = cfg.get(f"{key_base}_mac")
            else:
                val = None
            if not val:
                val = cfg.get(key_base)
            return val

        src_val = _get_os_path("source_folder")
        src_list_raw = _get_os_path("source_folders")  # Assume list if present

        src_list = []
        if src_list_raw and isinstance(src_list_raw, list):
            src_list = src_list_raw
        elif src_val:
            src_list = [src_val]
        else:
            # Fallback to generic if OS specific failed
            src_list = cfg.get("source_folders", [])
            if not src_list:
                single = cfg.get("source_folder", "")
                if single:
                    src_list = [single]

        self.source_folders_list = src_list
        self.lst_source_folders.delete(0, END)
        for p in self.source_folders_list:
            self.lst_source_folders.insert(END, p)

        if self.source_folders_list:
            self.var_source_folder.set(self.source_folders_list[0])
        else:
            self.var_source_folder.set("")

        self.var_target_folder.set(_get_os_path("target_folder") or "")
        self.var_enable_sandbox.set(1 if cfg.get("enable_sandbox", True) else 0)
        self.var_temp_sandbox_root.set(_get_os_path("temp_sandbox_root") or "")
        self.var_sandbox_min_free_gb.set(str(cfg.get("sandbox_min_free_gb", 10)))
        self.var_sandbox_low_space_policy.set(
            cfg.get("sandbox_low_space_policy", "block")
        )
        self.var_enable_llm_delivery_hub.set(
            1 if cfg.get("enable_llm_delivery_hub", True) else 0
        )
        self.var_llm_delivery_root.set(_get_os_path("llm_delivery_root") or "")
        self.var_llm_delivery_flatten.set(
            1 if cfg.get("llm_delivery_flatten", False) else 0
        )
        self.var_llm_delivery_include_pdf.set(
            1 if cfg.get("llm_delivery_include_pdf", False) else 0
        )
        self.var_enable_gdrive_upload.set(
            1 if cfg.get("enable_gdrive_upload", False) else 0
        )
        self.var_gdrive_client_secrets_path.set(
            cfg.get("gdrive_client_secrets_path", "") or ""
        )
        self.var_gdrive_folder_id.set(cfg.get("gdrive_folder_id", "") or "")
        self.var_enable_upload_readme.set(
            1 if cfg.get("enable_upload_readme", True) else 0
        )
        self.var_enable_upload_json_manifest.set(
            1 if cfg.get("enable_upload_json_manifest", True) else 0
        )
        self.var_upload_dedup_merged.set(
            1 if cfg.get("upload_dedup_merged", True) else 0
        )
        self.var_enable_corpus_manifest.set(
            1 if cfg.get("enable_corpus_manifest", True) else 0
        )
        self.var_enable_markdown.set(
            1 if cfg.get("output_enable_md", cfg.get("enable_markdown", True)) else 0
        )
        self.var_markdown_strip_header_footer.set(
            1 if cfg.get("markdown_strip_header_footer", True) else 0
        )
        self.var_markdown_structured_headings.set(
            1 if cfg.get("markdown_structured_headings", True) else 0
        )
        self.var_enable_markdown_quality_report.set(
            1 if cfg.get("enable_markdown_quality_report", True) else 0
        )
        self.var_enable_excel_json.set(1 if cfg.get("enable_excel_json", False) else 0)
        self.var_enable_chromadb_export.set(
            1 if cfg.get("enable_chromadb_export", False) else 0
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
        self.var_short_id_prefix.set(str(cfg.get("short_id_prefix", "ZW-") or "ZW-"))
        self._sync_markdown_master_with_global_output()
        self.var_mshelpviewer_folder_name.set(
            str(cfg.get("mshelpviewer_folder_name", "MSHelpViewer") or "MSHelpViewer")
        )
        self.var_enable_mshelp_merge_output.set(
            1 if cfg.get("enable_mshelp_merge_output", True) else 0
        )
        self.var_enable_mshelp_output_docx.set(
            1 if cfg.get("enable_mshelp_output_docx", False) else 0
        )
        self.var_enable_mshelp_output_pdf.set(
            1 if cfg.get("enable_mshelp_output_pdf", False) else 0
        )
        self.var_enable_incremental_mode.set(
            1 if cfg.get("enable_incremental_mode", False) else 0
        )
        self.var_incremental_verify_hash.set(
            1 if cfg.get("incremental_verify_hash", False) else 0
        )
        self.var_incremental_reprocess_renamed.set(
            1 if cfg.get("incremental_reprocess_renamed", False) else 0
        )
        self.var_source_priority_skip_same_name_pdf.set(
            1 if cfg.get("source_priority_skip_same_name_pdf", False) else 0
        )
        self.var_global_md5_dedup.set(1 if cfg.get("global_md5_dedup", False) else 0)
        self.var_enable_update_package.set(
            1 if cfg.get("enable_update_package", True) else 0
        )
        self.var_enable_parallel_conversion.set(
            1 if cfg.get("enable_parallel_conversion", False) else 0
        )
        self.var_parallel_workers.set(str(cfg.get("parallel_workers", 4)))
        self.var_enable_checkpoint.set(1 if cfg.get("enable_checkpoint", True) else 0)
        self.var_checkpoint_auto_resume.set(
            1 if cfg.get("checkpoint_auto_resume", True) else 0
        )
        self._on_toggle_parallel_conversion()

        self.var_enable_merge.set(1 if cfg.get("enable_merge", True) else 0)
        self.var_output_enable_pdf.set(1 if cfg.get("output_enable_pdf", True) else 0)
        self.var_output_enable_md.set(1 if cfg.get("output_enable_md", True) else 0)
        self.var_output_enable_merged.set(
            1 if cfg.get("output_enable_merged", True) else 0
        )
        self.var_output_enable_independent.set(
            1 if cfg.get("output_enable_independent", False) else 0
        )
        self.var_merge_convert_submode.set(
            cfg.get("merge_convert_submode", MERGE_CONVERT_SUBMODE_MERGE_ONLY)
        )
        self.var_merge_mode.set(cfg.get("merge_mode", MERGE_MODE_CATEGORY))
        self.var_merge_source.set(cfg.get("merge_source", "source"))
        self.var_enable_merge_index.set(
            1 if cfg.get("enable_merge_index", False) else 0
        )
        self.var_enable_merge_excel.set(
            1 if cfg.get("enable_merge_excel", False) else 0
        )

        # 鏉╂劘顢戝Ο鈥崇础 / 鐎涙劖膩瀵?/ 缁涙牜鏆愰敍鍫滅稊娑撴椽绮拋銈忕礆
        self.var_run_mode.set(cfg.get("run_mode", MODE_CONVERT_THEN_MERGE))
        self.var_collect_mode.set(cfg.get("collect_mode", COLLECT_MODE_COPY_AND_INDEX))
        self.var_strategy.set(cfg.get("content_strategy", "standard"))

        # 瀵洘鎼?& 鏉╂稓鈻肩粵鏍殣
        default_engine = cfg.get("default_engine", ENGINE_WPS)
        if default_engine not in (ENGINE_WPS, ENGINE_MS):
            default_engine = ENGINE_WPS
        self.var_engine.set(default_engine)

        kill_mode = cfg.get("kill_process_mode", KILL_MODE_AUTO)
        if kill_mode not in (KILL_MODE_AUTO, KILL_MODE_KEEP):
            kill_mode = KILL_MODE_AUTO
        self.var_kill_mode.set(kill_mode)

        # 闁板秶鐤嗙粻锛勬倞妞?
        self.var_log_folder.set(cfg.get("log_folder", "./logs"))

        excluded = cfg.get("excluded_folders", [])
        self.txt_excluded_folders.delete("1.0", "end")
        if isinstance(excluded, list):
            self.txt_excluded_folders.insert("end", "\n".join(excluded))

        price_kws = cfg.get("price_keywords", [])
        self.txt_price_keywords.delete("1.0", "end")
        if isinstance(price_kws, list):
            self.txt_price_keywords.insert("end", "\n".join(price_kws))

        # 扩展名 chip 编辑器：从 cfg 同步桶到 UI（如果还没构建则跳过）
        if hasattr(self, "_cfg_set_allowed_extensions"):
            ext_cfg = cfg.get("allowed_extensions") or {}
            try:
                self._cfg_set_allowed_extensions(ext_cfg)
            except Exception:
                pass

        self.var_timeout_seconds.set(str(cfg.get("timeout_seconds", 60)))
        self.var_pdf_wait_seconds.set(str(cfg.get("pdf_wait_seconds", 15)))
        self.var_ppt_timeout_seconds.set(str(cfg.get("ppt_timeout_seconds", 180)))
        self.var_ppt_pdf_wait_seconds.set(str(cfg.get("ppt_pdf_wait_seconds", 30)))
        self.var_office_reuse_app.set(1 if cfg.get("office_reuse_app", True) else 0)
        self.var_office_restart_every_n_files.set(
            str(cfg.get("office_restart_every_n_files", 25))
        )
        self.var_max_merge_size_mb.set(str(cfg.get("max_merge_size_mb", 80)))
        self.var_merge_filename_pattern.set(
            cfg.get("merge_filename_pattern") or "Merged_{category}_{timestamp}_{idx}"
        )

        # 閼辨柨濮╅崚閿嬫煀
        self._on_run_mode_change()
        self._on_toggle_sandbox()
        self._on_toggle_incremental_mode()
        self.refresh_locator_maps()
        self._normalize_numeric_var(self.var_timeout_seconds)
        self._normalize_numeric_var(self.var_pdf_wait_seconds)
        self._normalize_numeric_var(self.var_ppt_timeout_seconds)
        self._normalize_numeric_var(self.var_ppt_pdf_wait_seconds)
        self._normalize_numeric_var(self.var_office_restart_every_n_files)
        self._normalize_numeric_var(self.var_max_merge_size_mb)
        self._normalize_short_id_var(self.var_locator_short_id)
        self._normalize_date_var(self.var_date_str)
        self.validate_runtime_inputs(silent=True, scope="all")
        self._suspend_cfg_dirty = False
        self._refresh_config_dirty_from_file()
        if (
            self.profile_manager_win is not None
            and self.profile_manager_win.winfo_exists()
        ):
            self._refresh_profile_tree()
        if (
            self.load_profile_dialog is not None
            and self.load_profile_dialog.winfo_exists()
        ):
            self._refresh_load_profile_tree()
        if hasattr(self, "_update_task_tab_for_app_mode"):
            self._update_task_tab_for_app_mode()
        if hasattr(self, "_refresh_task_list_ui"):
            try:
                self._refresh_task_list_ui()
            except Exception:
                pass

    def _reload_config_from_file(self, show_msg=True, confirm_dirty=True):
        if confirm_dirty and self.cfg_dirty:
            confirm = messagebox.askyesno(
                self.tr("btn_load_cfg"),
                self.tr("msg_confirm_load_config_dirty"),
            )
            if not confirm:
                return
        if not os.path.exists(self.config_path):
            if show_msg:
                messagebox.showerror(
                    self.tr("btn_load_cfg"),
                    self.tr("msg_load_missing").format(self.config_path),
                )
            return
        try:
            with open(self.config_path, "r", encoding="utf-8") as f:
                json.load(f)
            self._active_config_label = os.path.basename(str(self.config_path or ""))
            self._active_config_origin = str(self.config_path or "")
            self._load_config_to_ui()
            if show_msg:
                messagebox.showinfo(self.tr("btn_load_cfg"), self.tr("msg_load_ok"))
            if hasattr(self, "var_status"):
                self.var_status.set(self.tr("msg_load_ok"))
            if hasattr(self, "var_locator_result"):
                self.var_locator_result.set(self.tr("msg_load_ok"))
        except Exception as e:
            if show_msg:
                messagebox.showerror(
                    self.tr("btn_load_cfg"),
                    self.tr("msg_load_fail").format(e),
                )

