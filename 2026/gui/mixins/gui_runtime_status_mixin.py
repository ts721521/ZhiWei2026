# -*- coding: utf-8 -*-
"""Runtime status/helper methods extracted from OfficeGUI."""

import os
from tkinter import messagebox

from office_converter import (
    MODE_CONVERT_THEN_MERGE,
    MODE_MERGE_ONLY,
    MERGE_CONVERT_SUBMODE_MERGE_ONLY,
    MERGE_MODE_ALL_IN_ONE,
    MERGE_MODE_CATEGORY,
)


class RuntimeStatusMixin:
    def _set_running_ui_state(self, running: bool):
        self._ui_running = bool(running)
        if running:
            if hasattr(self, "btn_start"):
                self.btn_start.configure(state="disabled")
            if hasattr(self, "btn_stop"):
                self.btn_stop.configure(state="normal")
            if hasattr(self, "btn_task_stop"):
                self.btn_task_stop.configure(state="normal")
            if hasattr(self, "btn_save_cfg"):
                self.btn_save_cfg.configure(state="disabled")
            if hasattr(self, "btn_load_cfg"):
                self.btn_load_cfg.configure(state="disabled")
            if hasattr(self, "btn_manage_profiles"):
                self.btn_manage_profiles.configure(state="disabled")
            if hasattr(self, "btn_manage_profiles"):
                self.btn_manage_profiles.configure(state="disabled")
            for btn_name in (
                "btn_task_create",
                "btn_task_edit",
                "btn_task_delete",
                "btn_task_run",
                "btn_task_resume",
                "btn_task_refresh",
                "btn_task_load_to_ui",
            ):
                if hasattr(self, btn_name):
                    getattr(self, btn_name).configure(state="disabled")
            if hasattr(self, "chk_task_force_full_rebuild"):
                self.chk_task_force_full_rebuild.configure(state="disabled")
            for btn_name in (
                "btn_save_cfg_shared",
                "btn_save_cfg_convert",
                "btn_save_cfg_ai",
                "btn_save_cfg_merge",
                "btn_save_cfg_ui",
                "btn_save_cfg_rules",
                "btn_reset_cfg_shared",
                "btn_reset_cfg_convert",
                "btn_reset_cfg_ai",
                "btn_reset_cfg_merge",
                "btn_reset_cfg_ui",
                "btn_reset_cfg_rules",
                "btn_save_cfg_dirty",
            ):
                if hasattr(self, btn_name):
                    getattr(self, btn_name).configure(state="disabled")
            self.progress["mode"] = "determinate"
            self.progress["value"] = 0
            self.var_status.set(
                self.tr("status_init") if hasattr(self, "tr") else "Initializing..."
            )
        else:
            if hasattr(self, "btn_start"):
                self.btn_start.configure(state="normal")
            if hasattr(self, "btn_stop"):
                self.btn_stop.configure(state="disabled")
            if hasattr(self, "btn_task_stop"):
                self.btn_task_stop.configure(state="disabled")
            if hasattr(self, "btn_save_cfg"):
                self.btn_save_cfg.configure(state="normal")
            if hasattr(self, "btn_load_cfg"):
                self.btn_load_cfg.configure(state="normal")
            if hasattr(self, "btn_manage_profiles"):
                self.btn_manage_profiles.configure(state="normal")
            for btn_name in (
                "btn_task_create",
                "btn_task_edit",
                "btn_task_delete",
                "btn_task_run",
                "btn_task_refresh",
                "btn_task_load_to_ui",
            ):
                if hasattr(self, btn_name):
                    getattr(self, btn_name).configure(state="normal")
            if hasattr(self, "chk_task_force_full_rebuild"):
                self.chk_task_force_full_rebuild.configure(state="normal")
            for btn_name in (
                "btn_save_cfg_shared",
                "btn_save_cfg_convert",
                "btn_save_cfg_ai",
                "btn_save_cfg_merge",
                "btn_save_cfg_ui",
                "btn_save_cfg_rules",
                "btn_reset_cfg_shared",
                "btn_reset_cfg_convert",
                "btn_reset_cfg_ai",
                "btn_reset_cfg_merge",
                "btn_reset_cfg_ui",
                "btn_reset_cfg_rules",
            ):
                if hasattr(self, btn_name):
                    getattr(self, btn_name).configure(state="normal")
            self._update_config_dirty_summary(getattr(self, "_last_section_dirty", {}))
            if hasattr(self, "_on_task_select"):
                self._on_task_select()
            if hasattr(self, "_update_task_tab_for_app_mode"):
                self._update_task_tab_for_app_mode()
            self.progress.stop()
            self.progress["value"] = 100
            self.var_status.set(
                self.tr("status_ready") if hasattr(self, "tr") else "Ready"
            )
        self._update_profile_manager_controls()
        self._update_profile_dialog_controls()

    def on_progress_update(self, current, total):
        """Thread-safe callback invoked from converter worker thread."""

        def _update():
            if total > 0:
                pct = (current / total) * 100
                self.progress["value"] = pct
                self.var_status.set(
                    self.tr("status_processing").format(current, total, pct)
                )
            else:
                self.progress["mode"] = "indeterminate"
                self.progress.start(20)
                self.var_status.set(
                    self.tr("status_processing_unknown").format(current)
                )

        # Thread-safe marshal to main UI loop.
        self.after(0, _update)

    def _build_artifact_summary_text(self, converter, step_index, total_steps):
        if converter is None:
            return ""

        converted_count = len(getattr(converter, "generated_pdfs", []) or [])
        merged_count = len(getattr(converter, "generated_merge_outputs", []) or [])
        merged_md_count = len(
            getattr(converter, "generated_merge_markdown_outputs", []) or []
        )
        map_count = len(getattr(converter, "generated_map_outputs", []) or [])
        markdown_count = len(getattr(converter, "generated_markdown_outputs", []) or [])
        markdown_quality_count = len(
            getattr(converter, "generated_markdown_quality_outputs", []) or []
        )
        excel_json_count = len(
            getattr(converter, "generated_excel_json_outputs", []) or []
        )
        records_json_count = len(
            getattr(converter, "generated_records_json_outputs", []) or []
        )
        chromadb_count = len(getattr(converter, "generated_chromadb_outputs", []) or [])
        mshelp_count = len(getattr(converter, "generated_mshelp_outputs", []) or [])

        lines = [
            self.tr("log_artifacts_title").format(step_index, total_steps),
            self.tr("log_artifacts_counts").format(
                converted_count, merged_count + merged_md_count, map_count
            ),
            self.tr("log_artifacts_ai_counts").format(
                markdown_count, excel_json_count, records_json_count
            ),
            self.tr("log_artifacts_ai_quality").format(markdown_quality_count),
            self.tr("log_artifacts_ai_vector").format(chromadb_count),
        ]
        if mshelp_count:
            lines.append(self.tr("log_artifacts_mshelp").format(mshelp_count))

        manifest_path = getattr(converter, "corpus_manifest_path", "")
        if manifest_path:
            lines.append(self.tr("log_artifacts_manifest").format(manifest_path))

        convert_index = getattr(converter, "convert_index_path", "")
        if convert_index:
            lines.append(self.tr("log_artifacts_convert_index").format(convert_index))

        collect_index = getattr(converter, "collect_index_path", "")
        if collect_index:
            lines.append(self.tr("log_artifacts_collect_index").format(collect_index))

        merge_excel = getattr(converter, "merge_excel_path", "")
        if merge_excel:
            lines.append(self.tr("log_artifacts_merge_excel").format(merge_excel))

        for md_path in (getattr(converter, "generated_markdown_outputs", []) or [])[:2]:
            lines.append(self.tr("log_artifacts_markdown").format(md_path))
        for q_path in (
            getattr(converter, "generated_markdown_quality_outputs", []) or []
        )[:2]:
            lines.append(self.tr("log_artifacts_markdown_quality").format(q_path))
        for excel_json_path in (
            getattr(converter, "generated_excel_json_outputs", []) or []
        )[:2]:
            lines.append(self.tr("log_artifacts_excel_json").format(excel_json_path))
        for js_path in (getattr(converter, "generated_records_json_outputs", []) or [])[
            :2
        ]:
            lines.append(self.tr("log_artifacts_records_json").format(js_path))
        for vec_path in (getattr(converter, "generated_chromadb_outputs", []) or [])[
            :2
        ]:
            lines.append(self.tr("log_artifacts_chromadb").format(vec_path))
        for mshelp_path in (getattr(converter, "generated_mshelp_outputs", []) or [])[
            :2
        ]:
            lines.append(self.tr("log_artifacts_markdown").format(mshelp_path))
        for merged_md_path in (
            getattr(converter, "generated_merge_markdown_outputs", []) or []
        )[:2]:
            lines.append(self.tr("log_artifacts_markdown").format(merged_md_path))
        update_manifest = getattr(converter, "update_package_manifest_path", "")
        if update_manifest:
            lines.append(
                self.tr("log_artifacts_update_package").format(update_manifest)
            )
        llm_hub_root = getattr(converter, "llm_hub_root", "")
        if llm_hub_root:
            lines.append(self.tr("log_artifacts_llm_hub").format(llm_hub_root))
        inc_ctx = getattr(converter, "_incremental_context", None) or {}
        if inc_ctx.get("enabled"):
            lines.append(
                self.tr("log_artifacts_incremental").format(
                    inc_ctx.get("added_count", 0),
                    inc_ctx.get("modified_count", 0),
                    inc_ctx.get("renamed_count", 0),
                    inc_ctx.get("unchanged_count", 0),
                    inc_ctx.get("deleted_count", 0),
                )
            )

        # 添加失败文件摘要
        detailed_errors = getattr(converter, "detailed_error_records", []) or []
        if detailed_errors:
            lines.append(self.tr("log_artifacts_failed_title"))

            # 按错误类型分组统计
            error_type_counts = {}
            for err in detailed_errors:
                et = err.get("error_type", "unknown")
                error_type_counts[et] = error_type_counts.get(et, 0) + 1

            for et, count in sorted(error_type_counts.items(), key=lambda x: -x[1]):
                suggestion_key = f"log_error_suggestion_{et}"
                suggestion = self.tr(suggestion_key) if hasattr(self, "tr") else ""
                if suggestion == suggestion_key:  # 没有找到翻译
                    suggestion = ""
                lines.append(self.tr("log_artifacts_failed_type").format(et, count))

            # 报告路径
            failed_report_path = getattr(converter, "failed_report_path", "")
            if failed_report_path:
                lines.append(
                    self.tr("log_artifacts_failed_report").format(failed_report_path)
                )

            # 提示用户处理建议
            retryable_count = sum(1 for e in detailed_errors if e.get("is_retryable"))
            manual_count = sum(
                1 for e in detailed_errors if e.get("requires_manual_action")
            )
            lines.append(
                self.tr("log_artifacts_failed_summary").format(
                    len(detailed_errors), retryable_count, manual_count
                )
            )

        return "\n".join(lines)

    def _scan_first_file_with_ext(self, roots, ext):
        ext = str(ext or "").lower()
        for root in roots or []:
            if not root or not os.path.isdir(root):
                continue
            for cur, _, files in os.walk(root):
                for name in files:
                    if name.lower().endswith(ext):
                        return os.path.join(cur, name)
        return ""

    def _should_continue_when_md_merge_missing(self, clean_sources, target):
        if self.var_run_mode.get() != MODE_MERGE_ONLY:
            return True
        if self.var_merge_convert_submode.get() != MERGE_CONVERT_SUBMODE_MERGE_ONLY:
            return True
        if not bool(self.var_output_enable_merged.get()):
            return True
        if not bool(self.var_output_enable_md.get()):
            return True
        roots = (
            [target]
            if self.var_merge_source.get() == "target"
            else list(clean_sources or [])
        )
        if self._scan_first_file_with_ext(roots, ".md"):
            return True
        msg = (
            "已选择“仅合并 + MD合并”，但未找到可合并的 .md 文件。\n\n"
            "选择“是”将继续执行并跳过MD合并；选择“否”将退出本次任务。"
        )
        return bool(messagebox.askyesno(self.tr("btn_start"), msg))

    def _sanitize_runtime_config_for_mode(self, cfg, run_mode):
        """Apply mode-driven coercion rules in place; return list of message strings."""
        messages = []
        if run_mode == MODE_CONVERT_THEN_MERGE:
            prev = cfg.get("merge_source", "source")
            if prev != "target":
                cfg["merge_source"] = "target"
                messages.append(self.tr("msg_coercion_merge_source").format(prev))
        merge_mode = cfg.get("merge_mode", MERGE_MODE_CATEGORY)
        if merge_mode == MERGE_MODE_ALL_IN_ONE:
            messages.append(self.tr("msg_coercion_max_mb_ignored"))
        elif merge_mode == MERGE_MODE_CATEGORY:
            try:
                mb = int(cfg.get("max_merge_size_mb", 80))
                if mb < 1:
                    cfg["max_merge_size_mb"] = 80
                    messages.append(self.tr("msg_coercion_max_mb_default").format(80))
            except (TypeError, ValueError):
                cfg["max_merge_size_mb"] = 80
                messages.append(self.tr("msg_coercion_max_mb_default").format(80))
        if bool(cfg.get("enable_fast_md_engine", False)):
            if bool(cfg.get("output_enable_pdf", True)):
                cfg["output_enable_pdf"] = False
                messages.append("output_enable_pdf: true -> false (Fast MD)")
            if bool(cfg.get("output_enable_merged", True)):
                cfg["output_enable_merged"] = False
                messages.append("output_enable_merged: true -> false (Fast MD)")
            if not bool(cfg.get("output_enable_md", True)):
                cfg["output_enable_md"] = True
                messages.append("output_enable_md: false -> true (Fast MD)")
            if not bool(cfg.get("output_enable_independent", False)):
                cfg["output_enable_independent"] = True
                messages.append("output_enable_independent: false -> true (Fast MD)")
            if bool(cfg.get("enable_parallel_conversion", False)):
                cfg["enable_parallel_conversion"] = False
                messages.append("enable_parallel_conversion: true -> false (Fast MD)")
        return messages

    def _log_coercion_summary(self, coercion_messages, show_dialog=False):
        """Write coercion summary to log and optionally show dialog."""
        if not coercion_messages:
            return
        header = self.tr("log_coercion_header")
        block = "\n".join(f"  - {m}" for m in coercion_messages)
        self.txt_log.insert("end", f"{header}\n{block}\n")
        self.txt_log.see("end")
        if show_dialog:
            body = self.tr("msg_coercion_body").format(block)
            messagebox.showinfo(self.tr("msg_coercion_title"), body)

    # ===================== 任务控制 =====================

