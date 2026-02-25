# -*- coding: utf-8 -*-
"""Execution flow methods extracted from OfficeGUI."""

import os
import tempfile
import threading
import traceback
from datetime import datetime
from tkinter import messagebox

from office_converter import (
    MODE_CONVERT_ONLY,
    MODE_MERGE_ONLY,
    MODE_CONVERT_THEN_MERGE,
    MODE_COLLECT_ONLY,
    MODE_MSHELP_ONLY,
    get_app_path,
)
from task_manager import (
    build_task_runtime_config,
    create_checkpoint,
    mark_checkpoint_file_done,
    remove_task_registry_if_exists,
)


class ExecutionFlowMixin:
    def _new_gui_converter(self):
        converter_cls = getattr(self, "_converter_cls", None)
        if converter_cls is None:
            raise RuntimeError("converter class is not configured")
        return converter_cls(self.config_path)

    def _apply_runtime_cfg_to_converter(self, converter, runtime_cfg):
        converter.config = dict(runtime_cfg)
        converter._apply_config_defaults()
        converter._init_paths_from_config()
        converter.run_mode = converter.config.get("run_mode", converter.run_mode)
        converter.collect_mode = converter.config.get(
            "collect_mode", converter.collect_mode
        )
        converter.content_strategy = converter.config.get(
            "content_strategy", converter.content_strategy
        )
        converter.merge_mode = converter.config.get("merge_mode", converter.merge_mode)
        converter.enable_merge_index = bool(
            converter.config.get("enable_merge_index", converter.enable_merge_index)
        )
        converter.enable_merge_excel = bool(
            converter.config.get("enable_merge_excel", converter.enable_merge_excel)
        )
        converter.engine_type = converter.config.get(
            "default_engine", self.var_engine.get()
        )

    def _on_click_task_run(self, resume=False):
        if self.worker_thread and self.worker_thread.is_alive():
            messagebox.showinfo(
                self.tr("btn_start"), self.tr("msg_task_already_running")
            )
            return
        task_id = self._get_selected_task_id()
        if not task_id:
            messagebox.showinfo(
                self.tr("grp_task_runtime"), self.tr("msg_task_select_required")
            )
            return
        task = self.task_store.get_task(task_id)
        if not task:
            return

        checkpoint = self.task_store.load_checkpoint(task_id)
        resume_list = None
        if resume:
            if not checkpoint:
                messagebox.showinfo(
                    self.tr("btn_task_resume"), self.tr("msg_task_resume_not_ready")
                )
                return
            planned = checkpoint.get("planned_files", []) or []
            completed = set(checkpoint.get("completed_files", []) or [])
            resume_list = [p for p in planned if p not in completed]
            if not resume_list:
                self.task_store.clear_checkpoint(task_id)
                self.task_store.update_task_runtime(task_id, status="idle")
                self._refresh_task_list_ui()
                messagebox.showinfo(
                    self.tr("btn_task_resume"), self.tr("msg_task_resume_empty")
                )
                return

        project_cfg = self._load_config_for_write()
        task, _ = self._ensure_task_config_snapshots(
            task, project_cfg=project_cfg, persist=True
        )
        force_full_rebuild = bool(self.var_task_force_full_rebuild.get()) and not resume
        runtime_cfg = build_task_runtime_config(
            project_cfg, task, force_full_rebuild=force_full_rebuild
        )
        base_mode = runtime_cfg.get("run_mode", "")
        binding = self._summarize_task_config_binding(task, runtime_cfg)
        coercion_msgs = self._sanitize_runtime_config_for_mode(runtime_cfg, base_mode)
        if coercion_msgs:
            self._log_coercion_summary(coercion_msgs, show_dialog=True)
        mapping_lines = [
            "",
            "[TASK] 任务配置映射",
            f"  - task_id: {task_id}",
            f"  - task_name: {task.get('name', '')}",
            f"  - 任务绑定配置: {binding.get('display_name', '')}",
            f"  - 绑定配置路径: {binding.get('config_path', '') or '-'}",
            f"  - 与当前活动配置关系: {binding.get('relation_label', '')}",
            f"  - 任务运行配置来源: {binding.get('runtime_source_desc', '')}",
        ]
        for line in mapping_lines:
            print(line)
            try:
                self.txt_log.insert("end", f"{line}\n")
                self.txt_log.see("end")
            except Exception:
                pass
        if force_full_rebuild:
            remove_task_registry_if_exists(
                task_id, runtime_cfg.get("target_folder", "")
            )
            self.task_store.clear_checkpoint(task_id)

        self.stop_requested = False
        self.current_task_id = task_id
        self.current_run_context = "task"
        self.task_store.update_task_runtime(task_id, status="running", last_error="")

        def worker():
            try:
                converter = self._new_gui_converter()
                converter.progress_callback = self.on_progress_update
                self._apply_runtime_cfg_to_converter(converter, runtime_cfg)

                def on_plan(files):
                    if resume:
                        return
                    cp = create_checkpoint(task_id, files)
                    self.task_store.save_checkpoint(task_id, cp)

                def on_done(record):
                    cp = self.task_store.load_checkpoint(task_id)
                    if not cp:
                        seed = resume_list if resume else []
                        cp = create_checkpoint(task_id, seed)
                    cp = mark_checkpoint_file_done(cp, record.get("source_path", ""))
                    self.task_store.save_checkpoint(task_id, cp)

                converter.file_plan_callback = on_plan
                converter.file_done_callback = on_done
                self.current_converter = converter
                converter.run(resume_file_list=resume_list)

                now = datetime.now().isoformat(timespec="seconds")
                if self.stop_requested:
                    self.task_store.update_task_runtime(task_id, status="paused")
                else:
                    self.task_store.clear_checkpoint(task_id)
                    self.task_store.update_task_runtime(
                        task_id, status="idle", last_run_at=now, last_error=""
                    )
            except Exception as e:
                self.task_store.update_task_runtime(
                    task_id,
                    status="paused" if self.stop_requested else "error",
                    last_error=str(e),
                )
                self.after(
                    0,
                    lambda err=str(e): messagebox.showerror(
                        self.tr("msg_runtime_error_title"),
                        self.tr("msg_runtime_error_body").format(err),
                    ),
                )
            finally:
                self.current_converter = None
                self.current_task_id = None
                self.current_run_context = "manual"
                self.stop_requested = False
                self.after(0, lambda: self._set_running_ui_state(False))
                self.after(0, self._refresh_task_list_ui)

        self.worker_thread = threading.Thread(target=worker, daemon=True)
        self.worker_thread.start()
        self._set_running_ui_state(True)

    def _on_click_start(self):
        if self.worker_thread and self.worker_thread.is_alive():
            messagebox.showinfo(
                self.tr("btn_start"), self.tr("msg_task_already_running")
            )
            return
        if getattr(self, "var_app_mode", None) and self.var_app_mode.get() == "task":
            task_id = self._get_selected_task_id()
            if not task_id:
                messagebox.showinfo(
                    self.tr("btn_start"), self.tr("msg_task_select_required")
                )
                return
            self._on_click_task_run(resume=False)
            return
        if not self.validate_runtime_inputs(silent=False, scope="all"):
            messagebox.showerror(
                self.tr("btn_start"), self.tr("msg_validation_fix_before_run")
            )
            return

        clean_sources = []
        for s in self.source_folders_list:
            s = s.strip().strip('"').strip("'")
            if os.path.isdir(s):
                clean_sources.append(s)

        if not clean_sources:
            fallback = self.var_source_folder.get().strip().strip('"').strip("'")
            if fallback and os.path.isdir(fallback):
                clean_sources.append(fallback)

        target = self.var_target_folder.get().strip().strip('"').strip("'")
        self.var_target_folder.set(target)

        if not clean_sources:
            messagebox.showerror(
                self.tr("btn_start"), self.tr("msg_source_folder_required")
            )
            return
        if not target:
            messagebox.showerror(
                self.tr("btn_start"), self.tr("msg_target_folder_required")
            )
            return
        if not self._should_continue_when_md_merge_missing(clean_sources, target):
            return

        self.stop_requested = False
        self.current_task_id = None
        self.current_run_context = "manual"
        self.txt_log.insert("end", f"\n========== {self.tr('log_start')} ==========\n")
        self.txt_log.see("end")

        def worker():
            try:
                base_mode = self.var_run_mode.get()
                steps = []

                if base_mode == MODE_COLLECT_ONLY:
                    for src in clean_sources:
                        steps.append(
                            {
                                "source": src,
                                "mode": MODE_COLLECT_ONLY,
                                "desc": f"Collect: {src}",
                            }
                        )

                elif base_mode == MODE_MSHELP_ONLY:
                    for src in clean_sources:
                        steps.append(
                            {
                                "source": src,
                                "mode": MODE_MSHELP_ONLY,
                                "desc": f"MSHelp: {src}",
                            }
                        )

                elif base_mode == MODE_MERGE_ONLY:
                    m_src = self.var_merge_source.get()
                    if m_src == "target":
                        steps.append(
                            {
                                "source": clean_sources[0],
                                "mode": MODE_MERGE_ONLY,
                                "desc": "Merge (target-based)",
                            }
                        )
                    else:
                        for src in clean_sources:
                            steps.append(
                                {
                                    "source": src,
                                    "mode": MODE_MERGE_ONLY,
                                    "desc": f"Merge (source-based: {src})",
                                }
                            )

                else:
                    for src in clean_sources:
                        steps.append(
                            {
                                "source": src,
                                "mode": MODE_CONVERT_ONLY,
                                "desc": f"Convert: {src}",
                            }
                        )

                    if (
                        base_mode == MODE_CONVERT_THEN_MERGE
                        and self.var_enable_merge.get()
                        and self.var_output_enable_merged.get()
                    ):
                        steps.append(
                            {
                                "source": clean_sources[0],
                                "mode": MODE_MERGE_ONLY,
                                "desc": "Merge (target-based)",
                            }
                        )

                total_steps = len(steps)
                print(f"[GUI] total steps: {total_steps}")

                for idx, step in enumerate(steps, 1):
                    if self.stop_requested:
                        print("[GUI] stop request accepted; remaining steps skipped.")
                        break

                    step_desc = step["desc"]
                    print(f"\n[GUI] >>> step {idx}/{total_steps}: {step_desc}")
                    self.txt_log.insert(
                        "end", f"\n>>> step {idx}/{total_steps}: {step_desc}\n"
                    )
                    self.txt_log.see("end")

                    print(f"[GUI] using config file: {self.config_path}")
                    converter = self._new_gui_converter()
                    converter.progress_callback = self.on_progress_update
                    self.current_converter = converter

                    cfg = converter.config
                    cfg["source_folder"] = step["source"]
                    cfg["source_folders"] = [step["source"]]
                    cfg["target_folder"] = target
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
                    cfg["llm_delivery_flatten"] = bool(
                        self.var_llm_delivery_flatten.get()
                    )
                    cfg["llm_delivery_include_pdf"] = bool(
                        self.var_llm_delivery_include_pdf.get()
                    )
                    cfg["enable_upload_readme"] = bool(
                        self.var_enable_upload_readme.get()
                    )
                    cfg["enable_upload_json_manifest"] = bool(
                        self.var_enable_upload_json_manifest.get()
                    )
                    cfg["upload_dedup_merged"] = bool(
                        self.var_upload_dedup_merged.get()
                    )
                    cfg["enable_gdrive_upload"] = bool(
                        self.var_enable_gdrive_upload.get()
                    )
                    cfg["gdrive_client_secrets_path"] = (
                        self.var_gdrive_client_secrets_path.get().strip()
                    )
                    cfg["gdrive_folder_id"] = self.var_gdrive_folder_id.get().strip()
                    cfg["enable_merge"] = bool(self.var_enable_merge.get())
                    cfg["output_enable_pdf"] = bool(self.var_output_enable_pdf.get())
                    cfg["output_enable_md"] = bool(self.var_output_enable_md.get())
                    cfg["output_enable_merged"] = bool(
                        self.var_output_enable_merged.get()
                    )
                    cfg["output_enable_independent"] = bool(
                        self.var_output_enable_independent.get()
                    )
                    cfg["merge_convert_submode"] = self.var_merge_convert_submode.get()
                    cfg["merge_mode"] = self.var_merge_mode.get()
                    cfg["merge_source"] = (
                        "target"
                        if base_mode == MODE_CONVERT_THEN_MERGE
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
                    cfg["enable_corpus_manifest"] = bool(
                        self.var_enable_corpus_manifest.get()
                    )
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
                    cfg["enable_chromadb_export"] = bool(
                        self.var_enable_chromadb_export.get()
                    )
                    cfg["mshelpviewer_folder_name"] = (
                        self.var_mshelpviewer_folder_name.get().strip()
                        or "MSHelpViewer"
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
                    cfg["enable_update_package"] = bool(
                        self.var_enable_update_package.get()
                    )
                    cfg["enable_parallel_conversion"] = bool(
                        self.var_enable_parallel_conversion.get()
                    )
                    cfg["parallel_workers"] = self._safe_positive_int(
                        self.var_parallel_workers.get(), 4
                    )
                    cfg["enable_checkpoint"] = bool(self.var_enable_checkpoint.get())
                    cfg["checkpoint_auto_resume"] = bool(
                        self.var_checkpoint_auto_resume.get()
                    )
                    cfg["kill_process_mode"] = self.var_kill_mode.get()
                    cfg["default_engine"] = self.var_engine.get()
                    cfg["office_reuse_app"] = bool(self.var_office_reuse_app.get())
                    cfg["office_restart_every_n_files"] = self._safe_positive_int(
                        self.var_office_restart_every_n_files.get(), 25
                    )
                    coercion_msgs = self._sanitize_runtime_config_for_mode(
                        cfg, base_mode
                    )
                    if idx == 1 and coercion_msgs:
                        self.after(
                            0,
                            lambda m=coercion_msgs: self._log_coercion_summary(
                                m, show_dialog=True
                            ),
                        )

                    converter.run_mode = step["mode"]
                    converter.collect_mode = self.var_collect_mode.get()
                    converter.content_strategy = self.var_strategy.get()
                    converter.merge_mode = self.var_merge_mode.get()
                    converter.engine_type = self.var_engine.get()
                    converter.enable_merge_index = bool(
                        self.var_enable_merge_index.get()
                    )
                    converter.enable_merge_excel = bool(
                        self.var_enable_merge_excel.get()
                    )

                    if self.var_enable_date_filter.get():
                        date_str = self.var_date_str.get().strip()
                        try:
                            converter.filter_date = datetime.strptime(
                                date_str, "%Y-%m-%d"
                            )
                            converter.filter_mode = self.var_filter_mode.get()
                        except ValueError:
                            pass

                    temp_root = cfg.get("temp_sandbox_root", "").strip()
                    if temp_root:
                        if not os.path.isabs(temp_root):
                            temp_root = os.path.abspath(
                                os.path.join(get_app_path(), temp_root)
                            )
                    else:
                        temp_root = tempfile.gettempdir()
                    converter.temp_sandbox_root = temp_root
                    converter.temp_sandbox = os.path.join(
                        temp_root, "OfficeToPDF_Sandbox"
                    )
                    os.makedirs(converter.temp_sandbox, exist_ok=True)

                    converter.failed_dir = os.path.join(
                        cfg["target_folder"], "_FAILED_FILES"
                    )
                    os.makedirs(converter.failed_dir, exist_ok=True)
                    converter.merge_output_dir = os.path.join(
                        cfg["target_folder"], "_MERGED"
                    )
                    os.makedirs(converter.merge_output_dir, exist_ok=True)

                    converter.run()
                    artifact_summary = self._build_artifact_summary_text(
                        converter, idx, total_steps
                    )
                    if artifact_summary:
                        self.txt_log.insert("end", f"{artifact_summary}\n")
                        self.txt_log.see("end")

                print("[GUI] all tasks completed.")
                self.txt_log.insert(
                    "end", f"\n========== {self.tr('log_stop')} ==========\n"
                )
                self.txt_log.see("end")

            except Exception as e:
                print(f"[GUI] runtime error: {e}")
                print(traceback.format_exc())
                messagebox.showerror(
                    self.tr("msg_runtime_error_title"),
                    self.tr("msg_runtime_error_body").format(e),
                )
            finally:
                self.current_converter = None
                self.stop_requested = False
                self.after(0, lambda: self._set_running_ui_state(False))
                self.after(0, self.refresh_locator_maps)

        self.worker_thread = threading.Thread(target=worker, daemon=True)
        self.worker_thread.start()
        self._set_running_ui_state(True)

    def _on_click_stop(self):
        if self.current_converter is None:
            return
        if messagebox.askyesno(self.tr("btn_stop"), self.tr("msg_confirm_stop")):
            self.stop_requested = True
            self.current_converter.is_running = False
            print("[GUI] stop requested; waiting for current step to finish...")
            self.var_status.set(self.tr("status_stop_wait"))

    # ===================== 程序入口 =====================
