# -*- coding: utf-8 -*-
"""
基于 zhishi 任务与 NotebookLM 预设的 E2E 回归脚本。

使用 task_manager.build_task_runtime_config 构建与 GUI 一致的运行时配置，
无交互执行一次完整转换+合并，检查 _LLM_UPLOAD 行为并写入测试记录。

用法（须在 2026 目录下运行）：
  python scripts/run_zhishi_task_e2e.py [--max-files N] [--dry-run]
  python scripts/run_zhishi_task_e2e.py --task-name zhishi

退出码：0=通过，1=失败，2=成功但需优化。
"""

from __future__ import annotations

import argparse
import json
import os
import sys
from datetime import datetime

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
_PROJECT_ROOT = os.path.dirname(_SCRIPT_DIR)
if _PROJECT_ROOT not in sys.path:
    sys.path.insert(0, _PROJECT_ROOT)
os.chdir(_PROJECT_ROOT)

REPORTS_DIR = os.path.join(_PROJECT_ROOT, "docs", "test-reports")
TASKS_DIR = os.path.join(_PROJECT_ROOT, "tasks")
DEFAULT_RESULT_JSON = os.path.join(REPORTS_DIR, "notebooklm_e2e_zhishi_result.json")
DEFAULT_RUNTIME_CONFIG_JSON = os.path.join(REPORTS_DIR, "e2e_zhishi_runtime.json")


def _ensure_dir(path: str) -> None:
    d = os.path.dirname(path)
    if d:
        os.makedirs(d, exist_ok=True)


def _load_json(path: str, default: dict | list):
    if not os.path.isfile(path):
        return default
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, (dict, list)) else default
    except Exception:
        return default


def find_task_by_name(task_name: str) -> dict | None:
    index_path = os.path.join(TASKS_DIR, "tasks_index.json")
    index = _load_json(index_path, {"tasks": []})
    tasks = index.get("tasks") or []
    for t in tasks:
        if str(t.get("name", "")).strip().lower() == task_name.strip().lower():
            task_id = t.get("id")
            if not task_id:
                return t
            full_path = os.path.join(TASKS_DIR, f"{task_id}.json")
            full_task = _load_json(full_path, t)
            if isinstance(full_task, dict):
                return full_task
            return t
    return None


def apply_e2e_overrides(cfg: dict) -> dict:
    """无交互、稳定性相关覆盖。"""
    cfg = dict(cfg)
    cfg["kill_process_mode"] = "auto"
    cfg["default_engine"] = "wps"
    cfg["auto_open_output_dir"] = False
    cfg.setdefault("enable_llm_delivery_hub", True)
    cfg.setdefault("run_mode", "convert_then_merge")
    if cfg.get("run_mode") == "convert_then_merge":
        cfg["merge_source"] = "target"
    return cfg


def check_llm_upload(target_folder: str) -> tuple[bool, int, float, bool]:
    """返回 (exists, file_count, max_file_size_mb, manifest_valid)。"""
    llm_root = os.path.join(target_folder, "_LLM_UPLOAD")
    if not os.path.isdir(llm_root):
        return False, 0, 0.0, False
    manifest_path = os.path.join(llm_root, "llm_upload_manifest.json")
    manifest_valid = False
    if os.path.isfile(manifest_path):
        try:
            with open(manifest_path, "r", encoding="utf-8") as f:
                json.load(f)
            manifest_valid = True
        except (json.JSONDecodeError, OSError):
            pass
    count = 0
    max_mb = 0.0
    for name in os.listdir(llm_root):
        p = os.path.join(llm_root, name)
        if os.path.isfile(p):
            count += 1
            size_mb = os.path.getsize(p) / (1024 * 1024)
            if size_mb > max_mb:
                max_mb = size_mb
    return True, count, max_mb, manifest_valid


def main() -> int:
    parser = argparse.ArgumentParser(
        description="NotebookLM E2E: run zhishi task headless and check _LLM_UPLOAD."
    )
    parser.add_argument(
        "--task-name",
        default="zhishi",
        help="Task name in tasks_index (default: zhishi)",
    )
    parser.add_argument(
        "--output",
        default=DEFAULT_RESULT_JSON,
        help="Result JSON path",
    )
    parser.add_argument(
        "--runtime-config",
        default=DEFAULT_RUNTIME_CONFIG_JSON,
        dest="runtime_config",
        help="Runtime config output path",
    )
    parser.add_argument(
        "--max-files",
        type=int,
        default=None,
        dest="max_files",
        help="Limit to first N files (quick run)",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Only build config and write to --runtime-config, do not run converter",
    )
    args = parser.parse_args()

    task = find_task_by_name(args.task_name)
    if not task:
        print(f"[E2E] Task not found: {args.task_name}", file=sys.stderr)
        result = {
            "success": False,
            "task_name": args.task_name,
            "message": "Task not found",
            "errors": [f"Task '{args.task_name}' not in tasks_index or missing task file"],
        }
        _ensure_dir(args.output)
        with open(args.output, "w", encoding="utf-8") as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        return 1

    project_config_path = (
        task.get("config_snapshot_path")
        or os.path.join(_PROJECT_ROOT, "config.json")
    )
    if not os.path.isabs(project_config_path):
        project_config_path = os.path.join(_PROJECT_ROOT, project_config_path)
    project_cfg = _load_json(project_config_path, {})
    if not project_cfg:
        print(f"[E2E] Project config empty or missing: {project_config_path}", file=sys.stderr)
        result = {
            "success": False,
            "task_name": args.task_name,
            "message": "Project config missing or empty",
            "errors": [f"Could not load {project_config_path}"],
        }
        _ensure_dir(args.output)
        with open(args.output, "w", encoding="utf-8") as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        return 1

    from task_manager import build_task_runtime_config

    runtime_cfg = build_task_runtime_config(
        project_cfg, task, force_full_rebuild=False, prefer_runtime_snapshot=True
    )
    runtime_cfg = apply_e2e_overrides(runtime_cfg)
    source_folder = runtime_cfg.get("source_folder", "")
    target_folder = runtime_cfg.get("target_folder", "")

    _ensure_dir(args.runtime_config)
    with open(args.runtime_config, "w", encoding="utf-8") as f:
        json.dump(runtime_cfg, f, ensure_ascii=False, indent=2)
    print(f"[E2E] Runtime config written: {args.runtime_config}")
    print(f"[E2E] source_folder={source_folder}")
    print(f"[E2E] target_folder={target_folder}")

    if args.dry_run:
        result = {
            "success": True,
            "dry_run": True,
            "task_name": args.task_name,
            "task_id": task.get("id"),
            "runtime_config_path": os.path.abspath(args.runtime_config),
            "source_folder": source_folder,
            "target_folder": target_folder,
            "message": "Dry run: config built only.",
        }
        _ensure_dir(args.output)
        with open(args.output, "w", encoding="utf-8") as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        return 0

    converter = None
    log_path = None
    run_exc = None
    try:
        from office_converter import OfficeConverter
        converter = OfficeConverter(args.runtime_config, interactive=False)
        if args.max_files is not None and args.max_files > 0:
            files = converter._scan_convert_candidates()
            limited = files[: args.max_files]
            print(f"[E2E] Limiting to first {len(limited)} of {len(files)} files")
            converter.run(resume_file_list=limited)
        else:
            converter.run()
        log_path = getattr(converter, "log_path", None) or ""
    except Exception as e:
        run_exc = e
        if converter is not None:
            log_path = getattr(converter, "log_path", None) or ""
        print(f"[E2E] Run failed: {e}", file=sys.stderr)

    target_folder = (
        (converter.config.get("target_folder", target_folder) if converter and getattr(converter, "config", None) else target_folder)
        or target_folder
    )
    llm_exists, file_count, max_mb, manifest_valid = check_llm_upload(target_folder)

    success = run_exc is None and llm_exists and (file_count > 0 or manifest_valid)
    need_optimization = success and (file_count > 50 or max_mb > 200.0)
    if need_optimization:
        exit_code = 2
        message = f"E2E passed but need optimization: file_count={file_count}, max_file_mb={max_mb:.2f}"
    elif success:
        exit_code = 0
        message = "E2E passed."
    else:
        exit_code = 1
        message = "E2E failed."

    errors = []
    if run_exc:
        errors.append(str(run_exc))
    if not llm_exists:
        errors.append("_LLM_UPLOAD missing or not a directory")
    elif file_count == 0 and not manifest_valid:
        errors.append("_LLM_UPLOAD empty and no valid manifest")

    llm_upload_dir = os.path.join(target_folder, "_LLM_UPLOAD") if target_folder else None
    result = {
        "success": success and not need_optimization,
        "task_name": args.task_name,
        "task_id": task.get("id"),
        "runtime_config_path": os.path.abspath(args.runtime_config),
        "source_folder": source_folder,
        "target_folder": target_folder,
        "log_path": os.path.abspath(log_path) if log_path else None,
        "llm_upload_dir": os.path.abspath(llm_upload_dir) if llm_upload_dir and os.path.isdir(llm_upload_dir) else None,
        "file_count": file_count,
        "max_file_size_mb": round(max_mb, 2),
        "manifest_valid": manifest_valid,
        "errors": errors,
        "message": message,
        "timestamp": datetime.now().isoformat(timespec="seconds"),
    }

    _ensure_dir(args.output)
    with open(args.output, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print(f"[E2E] Result: {message} (exit_code={exit_code})")
    print(f"[E2E] Result JSON: {args.output}")
    return exit_code


if __name__ == "__main__":
    sys.exit(main())
