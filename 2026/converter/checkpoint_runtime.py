# -*- coding: utf-8 -*-
"""Checkpoint runtime initialization helper extracted from office_converter.py."""

import json


def init_checkpoint(
    file_list,
    *,
    config,
    get_checkpoint_path_fn,
    checkpoint_resume_callback,
    save_checkpoint_fn,
    now_fn,
    exists_fn,
    remove_fn,
    open_fn,
    print_fn=print,
    log_warning_fn=None,
):
    if not config.get("enable_checkpoint", True):
        return None, file_list

    checkpoint_path = get_checkpoint_path_fn()
    if not checkpoint_path:
        return None, file_list

    if exists_fn(checkpoint_path):
        try:
            with open_fn(checkpoint_path, "r", encoding="utf-8") as f:
                checkpoint = json.load(f)

            if checkpoint.get("status") == "running":
                completed = set(checkpoint.get("completed_files", []))
                pending = [f for f in file_list if f not in completed]

                if pending and config.get("checkpoint_auto_resume", True):
                    completed_count = len(completed)
                    total_count = len(file_list)
                    print_fn(
                        f"\n[checkpoint] detected unfinished task: {completed_count}/{total_count} completed"
                    )
                    print_fn(f"[checkpoint] pending files: {len(pending)}")

                    if callable(checkpoint_resume_callback):
                        should_resume = checkpoint_resume_callback(
                            completed_count, total_count
                        )
                        if not should_resume:
                            print_fn(
                                "[checkpoint] user declined resume; restart from scratch"
                            )
                            remove_fn(checkpoint_path)
                            return None, file_list

                    print_fn("[checkpoint] resuming from checkpoint...")
                    return checkpoint, pending
        except (OSError, RuntimeError, TypeError, ValueError, json.JSONDecodeError) as e:
            if log_warning_fn:
                log_warning_fn(f"Failed to load checkpoint: {e}")

    now = now_fn().isoformat(timespec="seconds")
    checkpoint = {
        "version": 1,
        "created_at": now,
        "updated_at": now,
        "planned_files": list(file_list),
        "completed_files": [],
        "status": "running",
    }
    save_checkpoint_fn(checkpoint)
    return checkpoint, file_list
