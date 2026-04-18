# -*- coding: utf-8 -*-
"""Execution flow methods extracted from OfficeGUI."""

import threading
from datetime import datetime
from tkinter import messagebox

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
        self._run_single_task(task_id, resume)

    def _run_single_task(self, task_id, resume=False):
        """Run one task by id; used by single Run and by batch queue. On finish, may start next from _task_run_queue."""
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
                completed_id = self.current_task_id
                self.current_converter = None
                self.current_task_id = None
                self.current_run_context = "manual"
                self.stop_requested = False
                self.after(0, lambda: self._set_running_ui_state(False))
                self.after(0, self._refresh_task_list_ui)
                self.after(0, lambda cid=completed_id: self._maybe_run_next_queued_task(cid))

        self.worker_thread = threading.Thread(target=worker, daemon=True)
        self.worker_thread.start()
        # 保证 Start/Stop、保存、任务管理按钮立即切到"运行中"状态；
        # finally 分支会在 worker 结束后调用 _set_running_ui_state(False) 复原。
        self._set_running_ui_state(True)


class TaskOnlyStartMixin:
    """Start / 批量运行 / 停止 按钮的统一入口：始终走任务运行路径。"""

    def _on_click_start(self):
        if getattr(self, "worker_thread", None) and self.worker_thread.is_alive():
            messagebox.showinfo(
                self.tr("btn_start"), self.tr("msg_task_already_running")
            )
            return
        task_id = getattr(self, "_get_selected_task_id", lambda: None)()
        if not task_id:
            messagebox.showinfo(
                self.tr("btn_start"), self.tr("msg_task_select_required")
            )
            return
        self._on_click_task_run(resume=False)

    def _maybe_run_next_queued_task(self, completed_id):
        """After a single task run finishes, if batch queue exists and completed_id was head, run next."""
        queue = getattr(self, "_task_run_queue", None)
        if not queue or queue[0] != completed_id:
            return
        queue.pop(0)
        if queue:
            next_id = queue[0]
            self.after(100, lambda: self._run_single_task(next_id, False))

    def _on_click_task_batch_run(self):
        """Queue selected tasks and run them one by one."""
        if self.worker_thread and self.worker_thread.is_alive():
            messagebox.showinfo(
                self.tr("btn_start"), self.tr("msg_task_already_running")
            )
            return
        ids = self._get_selected_task_ids()
        if len(ids) < 2:
            messagebox.showinfo(
                self.tr("btn_task_batch_run"),
                self.tr("msg_task_batch_select_at_least_two"),
                parent=self,
            )
            return
        self._task_run_queue = list(ids)
        self._run_single_task(self._task_run_queue[0], False)

    def _on_click_stop(self):
        if self.current_converter is None:
            return
        if messagebox.askyesno(self.tr("btn_stop"), self.tr("msg_confirm_stop")):
            self.stop_requested = True
            if hasattr(self, "_task_run_queue"):
                self._task_run_queue.clear()
            self.current_converter.is_running = False
            print("[GUI] stop requested; waiting for current step to finish...")
            self.var_status.set(self.tr("status_stop_wait"))

    # ===================== 程序入口 =====================
