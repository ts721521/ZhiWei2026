# -*- coding: utf-8 -*-
"""
Task persistence and runtime config helpers.
"""

import copy
import json
import os
from datetime import datetime

TASK_BINDING_ACTIVE = "active"
TASK_BINDING_PROFILE = "profile"
TASK_BINDING_SNAPSHOT = "snapshot"
VALID_TASK_BINDING_MODES = {
    TASK_BINDING_ACTIVE,
    TASK_BINDING_PROFILE,
    TASK_BINDING_SNAPSHOT,
}


def normalize_task_binding_mode(mode):
    value = str(mode or "").strip().lower()
    return value if value in VALID_TASK_BINDING_MODES else TASK_BINDING_ACTIVE


def _now_iso():
    return datetime.now().isoformat(timespec="seconds")


def _ensure_dir(path):
    os.makedirs(path, exist_ok=True)


def _read_json(path, default):
    if not os.path.isfile(path):
        return copy.deepcopy(default)
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(default, dict) and isinstance(data, dict):
            return data
        if isinstance(default, list) and isinstance(data, list):
            return data
    except Exception:
        pass
    return copy.deepcopy(default)


def _write_json(path, payload):
    parent = os.path.dirname(path)
    if parent:
        _ensure_dir(parent)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)


def _deep_merge(base, override):
    out = copy.deepcopy(base if isinstance(base, dict) else {})
    for key, value in (override or {}).items():
        if isinstance(value, dict) and isinstance(out.get(key), dict):
            out[key] = _deep_merge(out.get(key), value)
        else:
            out[key] = copy.deepcopy(value)
    return out


def create_checkpoint(task_id, planned_files, run_id=None):
    now = _now_iso()
    return {
        "version": 1,
        "task_id": str(task_id),
        "run_id": str(run_id or f"run_{datetime.now().strftime('%Y%m%d_%H%M%S')}"),
        "planned_files": list(planned_files or []),
        "completed_files": [],
        "status": "running",
        "created_at": now,
        "updated_at": now,
    }


def mark_checkpoint_file_done(checkpoint, file_path):
    cp = copy.deepcopy(checkpoint if isinstance(checkpoint, dict) else {})
    completed = cp.setdefault("completed_files", [])
    path_value = str(file_path or "").strip()
    if path_value and path_value not in completed:
        completed.append(path_value)
    cp["updated_at"] = _now_iso()
    planned_count = len(cp.get("planned_files") or [])
    if planned_count > 0 and len(completed) >= planned_count:
        cp["status"] = "completed"
    return cp


def task_registry_path(task_id, target_folder):
    target = os.path.abspath(str(target_folder or ""))
    return os.path.join(
        target, "_AI", "registry", f"task_{task_id}_incremental_registry.json"
    )


def remove_task_registry_if_exists(task_id, target_folder):
    reg_path = task_registry_path(task_id, target_folder)
    try:
        if os.path.isfile(reg_path):
            os.remove(reg_path)
            return True
    except Exception:
        pass
    return False


def build_task_runtime_config(
    project_config,
    task,
    force_full_rebuild=False,
    prefer_runtime_snapshot=True,
):
    task = task if isinstance(task, dict) else {}
    binding_mode = normalize_task_binding_mode(task.get("config_binding_mode"))
    base_cfg = project_config if isinstance(project_config, dict) else {}
    cfg_source = "project_config(active)"
    runtime_snapshot = task.get("runtime_config_snapshot")
    snapshot = task.get("project_config_snapshot")

    if binding_mode == TASK_BINDING_PROFILE:
        profile_path = str(task.get("config_snapshot_path", "")).strip()
        profile_cfg = _read_json(profile_path, {}) if profile_path else {}
        if isinstance(profile_cfg, dict) and profile_cfg:
            base_cfg = profile_cfg
            cfg_source = "task.profile_config"
        elif isinstance(snapshot, dict) and snapshot:
            base_cfg = snapshot
            cfg_source = "task.project_config_snapshot(fallback)"
        elif prefer_runtime_snapshot and isinstance(runtime_snapshot, dict) and runtime_snapshot:
            base_cfg = runtime_snapshot
            cfg_source = "task.runtime_config_snapshot(fallback)"
    elif binding_mode == TASK_BINDING_SNAPSHOT:
        if isinstance(snapshot, dict) and snapshot:
            base_cfg = snapshot
            cfg_source = "task.project_config_snapshot"
        elif prefer_runtime_snapshot and isinstance(runtime_snapshot, dict) and runtime_snapshot:
            base_cfg = runtime_snapshot
            cfg_source = "task.runtime_config_snapshot"
    elif (
        prefer_runtime_snapshot
        and isinstance(runtime_snapshot, dict)
        and runtime_snapshot
        and not isinstance(project_config, dict)
    ):
        base_cfg = runtime_snapshot
        cfg_source = "task.runtime_config_snapshot(fallback)"

    merged = copy.deepcopy(base_cfg)
    overrides = task.get("config_overrides", {})
    if not isinstance(overrides, dict):
        overrides = {}
    merged = _deep_merge(merged, overrides)

    source_folders = task.get("source_folders")
    if isinstance(source_folders, list) and source_folders:
        source_folders = [str(p).strip() for p in source_folders if str(p).strip()]
    if not source_folders:
        fallback = str(task.get("source_folder", merged.get("source_folder", "") or "")).strip()
        source_folders = [fallback] if fallback else []
    source_folder = (source_folders or [""])[0]
    target_folder = task.get("target_folder", merged.get("target_folder", ""))
    task_id = str(task.get("id", "") or "task")

    merged["source_folder"] = source_folder
    merged["target_folder"] = target_folder
    merged["source_folders"] = source_folders if source_folders else []

    run_incremental = bool(task.get("run_incremental", True)) and not bool(
        force_full_rebuild
    )
    merged["enable_incremental_mode"] = run_incremental
    merged["incremental_registry_path"] = task_registry_path(task_id, target_folder)

    if "output_enable_md" not in merged:
        merged["output_enable_md"] = bool(merged.get("enable_markdown", True))
    merged["enable_markdown"] = bool(merged.get("output_enable_md", True))

    if merged.get("run_mode") == "convert_then_merge":
        merged["merge_source"] = "target"

    # Diagnostic metadata for UI and logs.
    merged["_task_config_source"] = cfg_source
    merged["_task_config_binding_mode"] = binding_mode
    merged["_task_id"] = task_id

    return merged


class TaskStore:
    def __init__(self, root_dir):
        self.root_dir = os.path.abspath(root_dir)
        self.tasks_dir = os.path.join(self.root_dir, "tasks")
        self.index_path = os.path.join(self.tasks_dir, "tasks_index.json")
        _ensure_dir(self.tasks_dir)

    def _default_index(self):
        return {"version": 1, "tasks": []}

    def task_path(self, task_id):
        return os.path.join(self.tasks_dir, f"{task_id}.json")

    def checkpoint_path(self, task_id):
        return os.path.join(self.tasks_dir, f"{task_id}_checkpoint.json")

    def load_index(self):
        index = _read_json(self.index_path, self._default_index())
        if not isinstance(index.get("tasks"), list):
            index["tasks"] = []
        return index

    def save_index(self, index):
        payload = index if isinstance(index, dict) else self._default_index()
        payload.setdefault("version", 1)
        payload.setdefault("tasks", [])
        _write_json(self.index_path, payload)
        return payload

    def list_tasks(self):
        index = self.load_index()
        out = []
        rebuilt_rows = []
        seen_ids = set()
        for item in index.get("tasks", []):
            if not isinstance(item, dict):
                continue
            task_id = str(item.get("id", "")).strip()
            if not task_id:
                continue
            seen_ids.add(task_id)
            task = self.get_task(task_id)
            row = self._build_summary(task) if task else copy.deepcopy(item)
            row["has_checkpoint"] = os.path.isfile(self.checkpoint_path(task_id))
            out.append(row)
            rebuilt_rows.append(self._build_summary(task) if task else copy.deepcopy(item))

        for task_id in self._iter_task_ids_from_disk():
            if task_id in seen_ids:
                continue
            task = self.get_task(task_id)
            if not task:
                continue
            row = self._build_summary(task)
            row["has_checkpoint"] = os.path.isfile(self.checkpoint_path(task_id))
            out.append(row)
            rebuilt_rows.append(self._build_summary(task))

        if rebuilt_rows != index.get("tasks", []):
            self.save_index({"version": index.get("version", 1), "tasks": rebuilt_rows})
        return out

    def get_task(self, task_id):
        path = self.task_path(task_id)
        task = _read_json(path, {})
        if not task:
            return None
        mode = normalize_task_binding_mode(task.get("config_binding_mode"))
        if task.get("config_binding_mode") != mode:
            task["config_binding_mode"] = mode
            _write_json(path, task)
        return task

    def _iter_task_ids_from_disk(self):
        try:
            names = os.listdir(self.tasks_dir)
        except Exception:
            return []
        out = []
        for name in names:
            lower = name.lower()
            if not lower.endswith(".json"):
                continue
            if lower == "tasks_index.json" or lower.endswith("_checkpoint.json"):
                continue
            out.append(name[:-5])
        return out

    def _build_summary(self, task):
        src = task.get("source_folder") or (task.get("source_folders") or [""])[0]
        config_snapshot_path = str(task.get("config_snapshot_path", "")).strip()
        config_snapshot_profile_name = str(
            task.get("config_snapshot_profile_name", "")
        ).strip()
        config_snapshot_profile_file = str(
            task.get("config_snapshot_profile_file", "")
        ).strip()
        config_binding_mode = normalize_task_binding_mode(
            task.get("config_binding_mode")
        )
        return {
            "id": task["id"],
            "name": task["name"],
            "source_folder": src,
            "target_folder": task.get("target_folder", ""),
            "run_incremental": bool(task.get("run_incremental", True)),
            "status": task.get("status", "idle"),
            "last_run_at": task.get("last_run_at", ""),
            "updated_at": task.get("updated_at", ""),
            "created_at": task.get("created_at", ""),
            "config_snapshot_path": config_snapshot_path,
            "config_snapshot_profile_name": config_snapshot_profile_name,
            "config_snapshot_profile_file": config_snapshot_profile_file,
            "config_binding_mode": config_binding_mode,
        }

    def save_task(self, task):
        if not isinstance(task, dict):
            raise ValueError("task must be dict")
        task_id = str(task.get("id", "")).strip()
        name = str(task.get("name", "")).strip()
        source_folders = task.get("source_folders")
        if isinstance(source_folders, list):
            source_folders = [str(p).strip() for p in source_folders if str(p).strip()]
        if not source_folders:
            source_folders = [str(task.get("source_folder", "")).strip()]
        source_folder = (source_folders or [""])[0]
        target_folder = str(task.get("target_folder", "")).strip()
        if not task_id or not name:
            raise ValueError("task id/name required")
        if not source_folder or not target_folder:
            raise ValueError("source_folder/target_folder required")

        existing = self.get_task(task_id) or {}
        now = _now_iso()
        payload = copy.deepcopy(existing)
        payload.update(task)
        payload["id"] = task_id
        payload["name"] = name
        payload["source_folders"] = source_folders
        payload["source_folder"] = source_folder
        payload["target_folder"] = target_folder
        payload["run_incremental"] = bool(payload.get("run_incremental", True))
        payload["config_binding_mode"] = normalize_task_binding_mode(
            payload.get("config_binding_mode")
        )
        overrides = payload.get("config_overrides", {})
        payload["config_overrides"] = overrides if isinstance(overrides, dict) else {}
        payload["created_at"] = payload.get("created_at") or now
        payload["updated_at"] = now
        payload["status"] = payload.get("status", "idle")
        payload["last_run_at"] = payload.get("last_run_at", "")
        _write_json(self.task_path(task_id), payload)

        index = self.load_index()
        summaries = []
        replaced = False
        for item in index.get("tasks", []):
            if isinstance(item, dict) and item.get("id") == task_id:
                summaries.append(self._build_summary(payload))
                replaced = True
            elif isinstance(item, dict):
                summaries.append(item)
        if not replaced:
            summaries.append(self._build_summary(payload))
        index["tasks"] = summaries
        self.save_index(index)
        return payload

    def migrate_legacy_tasks(self):
        migrated = 0
        for task_id in self._iter_task_ids_from_disk():
            path = self.task_path(task_id)
            payload = _read_json(path, {})
            if not payload:
                continue
            mode = normalize_task_binding_mode(payload.get("config_binding_mode"))
            if payload.get("config_binding_mode") != mode:
                payload["config_binding_mode"] = mode
                _write_json(path, payload)
                migrated += 1
        self.list_tasks()
        return migrated

    def delete_task(self, task_id):
        task_id = str(task_id or "").strip()
        if not task_id:
            return
        try:
            if os.path.isfile(self.task_path(task_id)):
                os.remove(self.task_path(task_id))
        except Exception:
            pass
        self.clear_checkpoint(task_id)

        index = self.load_index()
        kept = []
        for item in index.get("tasks", []):
            if not isinstance(item, dict):
                continue
            if item.get("id") != task_id:
                kept.append(item)
        index["tasks"] = kept
        self.save_index(index)

    def update_task_runtime(self, task_id, status=None, last_run_at=None, last_error=None):
        task = self.get_task(task_id)
        if not task:
            return None
        if status is not None:
            task["status"] = str(status)
        if last_run_at is not None:
            task["last_run_at"] = str(last_run_at)
        if last_error is not None:
            task["last_error"] = str(last_error)
        return self.save_task(task)

    def load_checkpoint(self, task_id):
        data = _read_json(self.checkpoint_path(task_id), {})
        return data if data else None

    def save_checkpoint(self, task_id, checkpoint):
        payload = checkpoint if isinstance(checkpoint, dict) else {}
        payload["task_id"] = str(task_id)
        payload.setdefault("version", 1)
        payload.setdefault("planned_files", [])
        payload.setdefault("completed_files", [])
        payload.setdefault("status", "running")
        payload.setdefault("created_at", _now_iso())
        payload["updated_at"] = _now_iso()
        _write_json(self.checkpoint_path(task_id), payload)
        return payload

    def clear_checkpoint(self, task_id):
        try:
            cp_path = self.checkpoint_path(task_id)
            if os.path.isfile(cp_path):
                os.remove(cp_path)
        except Exception:
            pass
