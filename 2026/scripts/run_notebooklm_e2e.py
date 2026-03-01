# -*- coding: utf-8 -*-
"""
NotebookLM E2E 测试脚本：无交互运行一次转换，检查产物与 log，输出结果 JSON 与修复提示文件。

用法（须在 2026 目录下运行）：
  python scripts/run_notebooklm_e2e.py [--config CONFIG] [--output OUTPUT] [--repair-prompt PATH] [--source DIR] [--target DIR]

环境变量 ZW_E2E_SOURCE、ZW_E2E_TARGET 可覆盖源/目标路径。
退出码：0=通过，1=失败，2=成功但需优化（如文件数>50 或单文件>200MB）。
"""

from __future__ import annotations

import argparse
import json
import os
import sys

# 保证从 2026 目录可导入 office_converter
_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
_PROJECT_ROOT = os.path.dirname(_SCRIPT_DIR)
if _PROJECT_ROOT not in sys.path:
    sys.path.insert(0, _PROJECT_ROOT)
os.chdir(_PROJECT_ROOT)

# 默认路径（相对 2026）
DEFAULT_SOURCE = r"Z:\Schneider\5_投标"
DEFAULT_TARGET = r"D:\ZWPDFTSEST"
DEFAULT_CONFIG = "configs/scenarios/notebooklm/config.notebooklm_test.json"
DEFAULT_OUTPUT_JSON = "docs/test-reports/notebooklm_e2e_result.json"
DEFAULT_REPAIR_PROMPT = "docs/test-reports/notebooklm_e2e_repair_prompt.txt"


def _ensure_dir(path: str) -> None:
    d = os.path.dirname(path)
    if d:
        os.makedirs(d, exist_ok=True)


def _get_source_target_from_env_or_args(
    args_source: str | None,
    args_target: str | None,
) -> tuple[str, str]:
    source = args_source or os.environ.get("ZW_E2E_SOURCE", "").strip() or DEFAULT_SOURCE
    target = args_target or os.environ.get("ZW_E2E_TARGET", "").strip() or DEFAULT_TARGET
    return source, target


def apply_test_overrides(cfg: dict, source: str, target: str) -> dict:
    """将 NotebookLM E2E 推荐项写入 config（含 *_win 字段，避免 Windows 优先键覆盖）。"""
    cfg = dict(cfg or {})

    # 路径：Windows 下 get_path_from_config 会优先读取 *_win
    cfg["source_folder"] = source
    cfg["source_folders"] = [source]
    cfg["target_folder"] = target
    cfg["source_folder_win"] = source
    cfg["source_folders_win"] = [source]
    cfg["target_folder_win"] = target

    # NotebookLM 推荐 + 无交互
    cfg["run_mode"] = "convert_then_merge"
    cfg["output_enable_pdf"] = True
    cfg["output_enable_md"] = True
    cfg["output_enable_merged"] = True
    cfg["output_enable_independent"] = False

    cfg["enable_llm_delivery_hub"] = True
    cfg["llm_delivery_flatten"] = True
    cfg["upload_dedup_merged"] = True
    cfg["enable_upload_readme"] = True
    cfg["enable_upload_json_manifest"] = True

    # 避免 run 结束后打开资源管理器
    cfg["auto_open_output_dir"] = False

    # 非交互：避免 ask 模式卡住
    cfg["kill_process_mode"] = "auto"
    cfg["default_engine"] = "wps"

    # 大目录/长时间：断点续跑与稳定性
    cfg["enable_checkpoint"] = True
    cfg["checkpoint_auto_resume"] = True
    cfg["office_restart_every_n_files"] = int(cfg.get("office_restart_every_n_files") or 25)
    cfg["sandbox_min_free_gb"] = int(cfg.get("sandbox_min_free_gb") or 10)

    return cfg


def ensure_config(config_path: str, source: str, target: str) -> bool:
    """若 config 不存在则生成（以 default_config 为基 + 测试计划覆盖 + 无交互/长时间友好）。"""
    if os.path.isfile(config_path):
        return True
    # Backward compatibility with legacy root-level config location.
    legacy_path = os.path.join(_PROJECT_ROOT, "config.notebooklm_test.json")
    if os.path.isfile(legacy_path):
        try:
            os.makedirs(os.path.dirname(config_path), exist_ok=True)
            with open(legacy_path, "r", encoding="utf-8") as f:
                cfg = json.load(f)
            cfg = apply_test_overrides(cfg, source, target)
            with open(config_path, "w", encoding="utf-8") as f:
                json.dump(cfg, f, indent=4, ensure_ascii=False)
            print(f"[E2E] Migrated legacy config to: {config_path}")
            return True
        except (OSError, json.JSONDecodeError, TypeError, ValueError):
            pass
    try:
        from converter.default_config import create_default_config
        created = create_default_config(config_path)
        if not created:
            return False
    except Exception as e:
        print(f"[E2E] Failed to create default config: {e}", file=sys.stderr)
        return False
    with open(config_path, "r", encoding="utf-8") as f:
        cfg = json.load(f)
    cfg = apply_test_overrides(cfg, source, target)
    with open(config_path, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=4, ensure_ascii=False)
    print(f"[E2E] Generated config: {config_path}")
    return True


def infer_error_category(exc: BaseException | None, log_content: str, llm_empty: bool) -> str:
    """根据异常与 log 推断 error_category（与决策表一致）。"""
    if llm_empty:
        return "llm_hub_empty"
    exc_type = type(exc).__name__.lower() if exc is not None else ""
    exc_msg = str(exc).lower() if exc is not None else ""
    low_log = log_content.lower() if log_content else ""
    signal = "\n".join([exc_type, exc_msg, low_log])

    if (
        "stream has ended unexpectedly" in signal
        or "stream ended unexpectedly" in signal
        or "unexpected end of stream" in signal
    ):
        return "stream_unexpected_end"

    if (
        "filenotfounderror" in signal
        or "no such file" in signal
        or ("path" in signal and "not found" in signal)
    ):
        return "path_not_found"

    if (
        "[errno 22]" in signal
        or "invalid argument" in signal
        or "path too long" in signal
        or "filename, directory name, or volume label syntax is incorrect" in signal
    ):
        return "path_invalid_or_too_long"

    if "permission" in signal or "access denied" in signal:
        return "permission_denied"

    if any(
        kw in signal
        for kw in (
            "win32com not supported",
            "invalid class string",
            "class not registered",
            "activex component can't create object",
            "office not installed",
            "wps not installed",
        )
    ):
        return "office_not_installed"

    if any(
        kw in signal
        for kw in (
            "com_error",
            "com object",
            "rpc server is unavailable",
            "server busy",
            "call was rejected",
            "0x800",
            "conversion worker failed",
            "office",
            "wps",
        )
    ):
        return "office_com_busy_or_crash"

    if ("valueerror" in signal and "config" in signal) or "config" in signal:
        return "config"
    return "config"


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
    parser = argparse.ArgumentParser(description="NotebookLM E2E: run converter headless and check output.")
    parser.add_argument(
        "--config",
        default=DEFAULT_CONFIG,
        help="Config JSON path (default: configs/scenarios/notebooklm/config.notebooklm_test.json)",
    )
    parser.add_argument("--output", default=DEFAULT_OUTPUT_JSON, help="Result JSON path")
    parser.add_argument("--repair-prompt", default=DEFAULT_REPAIR_PROMPT, dest="repair_prompt", help="Repair prompt file path when exit code != 0")
    parser.add_argument("--source", default=None, help="Override source folder")
    parser.add_argument("--target", default=None, help="Override target folder")
    parser.add_argument("--max-files", type=int, default=None, dest="max_files", help="Limit to first N files (quick run); default=all")
    args = parser.parse_args()

    source, target = _get_source_target_from_env_or_args(args.source, args.target)
    config_path = os.path.join(_PROJECT_ROOT, args.config) if not os.path.isabs(args.config) else args.config
    output_path = os.path.join(_PROJECT_ROOT, args.output) if not os.path.isabs(args.output) else args.output
    repair_path = os.path.join(_PROJECT_ROOT, args.repair_prompt) if not os.path.isabs(args.repair_prompt) else args.repair_prompt

    if not ensure_config(config_path, source, target):
        result = {
            "success": False,
            "log_path": None,
            "llm_upload_dir": None,
            "file_count": 0,
            "max_file_size_mb": 0.0,
            "errors": ["Failed to create or load config"],
            "error_category": "config",
            "message": "Config creation or load failed.",
        }
        _ensure_dir(output_path)
        with open(output_path, "w", encoding="utf-8") as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        _write_repair_prompt(repair_path, result, config_path)
        return 1

    # 每次运行前都写回覆盖项（含 *_win），避免 Windows 优先键导致目标目录跑偏
    with open(config_path, "r", encoding="utf-8") as f:
        cfg = json.load(f)
    cfg = apply_test_overrides(cfg, source, target)
    with open(config_path, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=4, ensure_ascii=False)

    converter = None
    log_path = None
    run_exc = None
    try:
        from office_converter import OfficeConverter
        converter = OfficeConverter(config_path, interactive=False)
        if args.max_files is not None and args.max_files > 0:
            files = converter._scan_convert_candidates()
            limited = files[: args.max_files]
            print(f"[E2E] Limiting to first {len(limited)} of {len(files)} files (--max-files={args.max_files})")
            converter.run(resume_file_list=limited)
        else:
            converter.run()
        log_path = getattr(converter, "log_path", None) or ""
    except Exception as e:
        run_exc = e
        if converter is not None:
            log_path = getattr(converter, "log_path", None) or ""
        print(f"[E2E] Run failed: {e}", file=sys.stderr)

    log_content = ""
    if log_path and os.path.isfile(log_path):
        try:
            with open(log_path, "r", encoding="utf-8", errors="replace") as f:
                log_content = f.read()
        except OSError:
            pass

    target_folder = target
    if converter is not None and getattr(converter, "config", None):
        target_folder = converter.config.get("target_folder", target)
    llm_exists, file_count, max_mb, manifest_valid = check_llm_upload(target_folder)

    success = run_exc is None and llm_exists and (file_count > 0 or manifest_valid)
    need_optimization = success and (file_count > 50 or max_mb > 200.0)
    if need_optimization:
        error_category = "file_too_large_or_many"
    elif not success and llm_exists and file_count == 0 and not manifest_valid:
        error_category = infer_error_category(run_exc, log_content, llm_empty=True)
    elif not success:
        error_category = infer_error_category(run_exc, log_content, llm_empty=not llm_exists or file_count == 0)
    else:
        error_category = ""

    errors = []
    if run_exc:
        errors.append(str(run_exc))
    if not llm_exists:
        errors.append("_LLM_UPLOAD missing or not a directory")
    elif file_count == 0 and not manifest_valid:
        errors.append("_LLM_UPLOAD empty and no valid manifest")

    if success and not need_optimization:
        exit_code = 0
        message = "E2E passed."
    elif need_optimization:
        exit_code = 2
        message = f"E2E passed but need optimization: file_count={file_count}, max_file_mb={max_mb:.2f} (NotebookLM: ≤50 sources, ≤200MB/file)."
    else:
        exit_code = 1
        message = "E2E failed."

    llm_upload_dir = os.path.join(target_folder, "_LLM_UPLOAD") if target_folder else None
    result = {
        "success": success and not need_optimization,
        "log_path": os.path.abspath(log_path) if log_path else None,
        "llm_upload_dir": os.path.abspath(llm_upload_dir) if llm_upload_dir and os.path.isdir(llm_upload_dir) else None,
        "file_count": file_count,
        "max_file_size_mb": round(max_mb, 2),
        "errors": errors,
        "error_category": error_category,
        "message": message,
    }

    _ensure_dir(output_path)
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print(f"[E2E] Result: {message} (exit_code={exit_code})")
    print(f"[E2E] Result JSON: {output_path}")

    if exit_code != 0:
        _write_repair_prompt(repair_path, result, config_path, output_path)
        print(f"[E2E] Repair prompt: {repair_path}")

    return exit_code


def _write_repair_prompt(
    repair_path: str,
    result: dict,
    config_path: str,
    result_json_path: str | None = None,
) -> None:
    _ensure_dir(repair_path)
    result_json_path = result_json_path or result.get("result_json") or "docs/test-reports/notebooklm_e2e_result.json"
    lines = [
        "E2E 未通过，请根据下列信息修复后重新运行 E2E。",
        "",
        "log_path: " + (result.get("log_path") or ""),
        "result_json: " + result_json_path,
        "error_category: " + (result.get("error_category") or ""),
        "errors: " + str(result.get("errors", [])),
        "",
        "按 docs/plans/NotebookLM_知识库测试计划_5_投标_ZWPDFTSEST.md 第六节「结果→动作表」执行修复，然后运行：",
        "  python scripts/run_notebooklm_e2e.py",
        "",
    ]
    with open(repair_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))


if __name__ == "__main__":
    sys.exit(main())
