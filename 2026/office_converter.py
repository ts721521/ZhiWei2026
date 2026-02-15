# -*- coding: utf-8 -*-
"""
Office batch conversion and file organizing tool - core logic

Notes:
- Can run as standalone CLI and can also be invoked by GUI (office_gui.py).
- In GUI mode, use interactive=False to disable CLI input() prompts.
"""

import os
import sys
import time
import json
import csv
import shutil
import logging
import argparse
import uuid
import traceback
import tempfile
import subprocess
import threading
import signal
import random
import hashlib
import re
import zipfile
from datetime import datetime, date as dt_date, time as dt_time
from pathlib import Path

# Avoid UnicodeEncodeError on Windows consoles with legacy code pages.
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(errors="replace")
    except Exception:
        pass

# win32com is required on Windows; use mocks on non-Windows environments.
HAS_WIN32 = False
try:
    import win32com.client
    import pythoncom
    import pywintypes

    HAS_WIN32 = True
except ImportError:
    # Provide mock objects on non-Windows environments to avoid import failures.
    class MockComError(Exception):
        pass

    class MockPyWinTypes:
        com_error = MockComError

    class MockPythonCom:
        def CoInitialize(self):
            pass

        def CoUninitialize(self):
            pass

    class MockWin32ComClient:
        def Dispatch(self, *args, **kwargs):
            raise RuntimeError("Win32 COM not supported")

        def DispatchEx(self, *args, **kwargs):
            raise RuntimeError("Win32 COM not supported")

    class MockWin32Com:
        client = MockWin32ComClient()

    pywintypes = MockPyWinTypes()
    pythoncom = MockPythonCom()
    win32com = MockWin32Com()


# For timed keyboard input on Windows (retry prompt in CLI mode).
try:
    import msvcrt

    HAS_MSVCRT = True
except ImportError:
    HAS_MSVCRT = False

# pypdf is used for PDF merging and text extraction/scanning.
try:
    from pypdf import PdfWriter, PdfReader

    HAS_PYPDF = True
except ImportError:
    HAS_PYPDF = False

# openpyxl is used to generate Excel indexes in collect_only mode.
try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font

    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

try:
    import chromadb

    HAS_CHROMADB = True
except Exception:
    HAS_CHROMADB = False

try:
    from bs4 import BeautifulSoup

    HAS_BS4 = True
except Exception:
    HAS_BS4 = False

try:
    from docx import Document

    HAS_PYDOCX = True
except Exception:
    HAS_PYDOCX = False

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas

    HAS_REPORTLAB = True
except Exception:
    HAS_REPORTLAB = False

__version__ = "5.18.0"

# Office constants
wdFormatPDF = 17
xlTypePDF = 0
ppSaveAsPDF = 32
ppFixedFormatTypePDF = 2
xlPDF_SaveAs = 57
xlRepairFile = 1

# Engine types
ENGINE_WPS = "wps"
ENGINE_MS = "ms"
ENGINE_ASK = "ask"

# Process cleanup strategies
KILL_MODE_ASK = "ask"
KILL_MODE_AUTO = "auto"
KILL_MODE_KEEP = "keep"

# Main run modes
MODE_CONVERT_ONLY = "convert_only"
MODE_MERGE_ONLY = "merge_only"
MODE_CONVERT_THEN_MERGE = "convert_then_merge"
MODE_COLLECT_ONLY = "collect_only"  # collect and deduplicate mode
MODE_MSHELP_ONLY = "mshelp_only"  # dedicated mode for MSHelpViewer API docs

# Merge & convert sub-modes under MODE_MERGE_ONLY
MERGE_CONVERT_SUBMODE_MERGE_ONLY = "merge_only"
MERGE_CONVERT_SUBMODE_PDF_TO_MD = "pdf_to_md"

# collect_only sub-modes
COLLECT_MODE_COPY_AND_INDEX = "copy_and_index"  # dedup + copy + Excel
COLLECT_MODE_INDEX_ONLY = "index_only"  # Excel only, no copying

# Merge modes
MERGE_MODE_CATEGORY = "category_split"  # split by Price_/Word_/Excel_ categories
MERGE_MODE_ALL_IN_ONE = "all_in_one"  # merge all PDFs into one file

# Content processing strategy (conversion mode only)
STRATEGY_STANDARD = "standard"  # classify only by extension
STRATEGY_SMART_TAG = "smart_tag"  # filename/content keyword hit -> Price_
STRATEGY_PRICE_ONLY = "price_only"  # process only keyword-matching files

ERR_RPC_SERVER_BUSY = -2147417846
DEFAULT_SHORT_ID_LEN = 8

# =============== Error Classification ===============
# 错误类型枚举 - 用于分类和生成处理建议


class ConversionErrorType:
    """转换错误类型分类"""

    PERMISSION_DENIED = "permission_denied"  # 权限不足
    FILE_LOCKED = "file_locked"  # 文件被占用
    FILE_NOT_FOUND = "file_not_found"  # 文件不存在
    FILE_CORRUPTED = "file_corrupted"  # 文件损坏
    COM_ERROR = "com_error"  # Office COM 错误
    TIMEOUT = "timeout"  # 超时
    DISK_FULL = "disk_full"  # 磁盘空间不足
    INVALID_FORMAT = "invalid_format"  # 格式无效
    PASSWORD_PROTECTED = "password_protected"  # 密码保护
    UNSUPPORTED_FORMAT = "unsupported_format"  # 不支持的格式
    UNKNOWN = "unknown"  # 未知错误


def classify_conversion_error(exception, context=""):
    """
    根据异常信息分类错误类型，返回错误类型和处理建议。

    Args:
        exception: 异常对象或异常信息字符串
        context: 额外的上下文信息（如文件路径）

    Returns:
        dict: {
            "error_type": 错误类型,
            "error_category": 错误分类（可重试/不可重试/需人工）,
            "message": 用户友好消息,
            "suggestion": 处理建议,
            "is_retryable": 是否可自动重试,
            "requires_manual_action": 是否需要人工干预
        }
    """
    err_str = str(exception).lower() if exception else ""
    err_type = type(exception).__name__ if hasattr(exception, "__name__") else ""

    # 权限错误
    if any(
        kw in err_str
        for kw in ["permission", "access denied", "拒绝访问", "权限", "unauthorized"]
    ):
        return {
            "error_type": ConversionErrorType.PERMISSION_DENIED,
            "error_category": "needs_manual",
            "message": "文件访问权限不足",
            "suggestion": "1. 以管理员身份运行程序\n2. 检查文件属性，取消「只读」\n3. 检查文件夹权限设置",
            "is_retryable": False,
            "requires_manual_action": True,
        }

    # 文件被占用
    if any(
        kw in err_str
        for kw in ["being used", "locked", "占用", "正由另一程序", "sharing violation"]
    ):
        return {
            "error_type": ConversionErrorType.FILE_LOCKED,
            "error_category": "retryable",
            "message": "文件被其他程序占用",
            "suggestion": "1. 关闭 Word/Excel/WPS 等 Office 程序\n2. 关闭文件资源管理器中该文件的预览\n3. 等待几秒后重试",
            "is_retryable": True,
            "requires_manual_action": False,
        }

    # 文件不存在
    if any(
        kw in err_str
        for kw in ["not found", "does not exist", "找不到", "不存在", "filenotfound"]
    ):
        return {
            "error_type": ConversionErrorType.FILE_NOT_FOUND,
            "error_category": "unrecoverable",
            "message": "文件不存在",
            "suggestion": "文件可能已被移动或删除，请检查源目录",
            "is_retryable": False,
            "requires_manual_action": True,
        }

    # 文件损坏
    if any(
        kw in err_str
        for kw in [
            "corrupt",
            "damaged",
            "损坏",
            "repair",
            "无法读取",
            "unreadable",
            "invalid data",
        ]
    ):
        return {
            "error_type": ConversionErrorType.FILE_CORRUPTED,
            "error_category": "unrecoverable",
            "message": "文件可能已损坏",
            "suggestion": "1. 尝试用 Office 打开文件并另存为新文件\n2. 使用 Office 的「打开并修复」功能\n3. 从备份恢复文件",
            "is_retryable": False,
            "requires_manual_action": True,
        }

    # COM 错误（Office 相关）
    if any(
        kw in err_str
        for kw in [
            "com_error",
            "com object",
            "rpc",
            "server busy",
            "call was rejected",
            "0x800",
        ]
    ):
        return {
            "error_type": ConversionErrorType.COM_ERROR,
            "error_category": "retryable",
            "message": "Office 组件通信错误",
            "suggestion": "1. 重启 Office 程序\n2. 检查 Office 安装是否完整\n3. 尝试使用其他转换引擎（WPS/MS）",
            "is_retryable": True,
            "requires_manual_action": False,
        }

    # 超时
    if any(kw in err_str for kw in ["timeout", "超时", "timed out"]):
        return {
            "error_type": ConversionErrorType.TIMEOUT,
            "error_category": "retryable",
            "message": "转换超时",
            "suggestion": "1. 文件可能过大，尝试增加超时时间\n2. 关闭其他占用资源的程序\n3. 分批处理大文件",
            "is_retryable": True,
            "requires_manual_action": False,
        }

    # 磁盘空间不足
    if any(
        kw in err_str
        for kw in ["disk full", "no space", "磁盘已满", "空间不足", "storage"]
    ):
        return {
            "error_type": ConversionErrorType.DISK_FULL,
            "error_category": "needs_manual",
            "message": "磁盘空间不足",
            "suggestion": "1. 清理磁盘空间\n2. 更换输出目录到其他磁盘\n3. 在「高级设置」中调整沙盒路径",
            "is_retryable": False,
            "requires_manual_action": True,
        }

    # 密码保护
    if any(
        kw in err_str for kw in ["password", "protected", "密码", "encrypted", "加密"]
    ):
        return {
            "error_type": ConversionErrorType.PASSWORD_PROTECTED,
            "error_category": "needs_manual",
            "message": "文件受密码保护",
            "suggestion": "1. 先用 Office 打开文件并移除密码保护\n2. 将文件另存为无密码版本后再转换",
            "is_retryable": False,
            "requires_manual_action": True,
        }

    # 格式无效
    if any(
        kw in err_str
        for kw in ["invalid format", "format not supported", "格式无效", "格式错误"]
    ):
        return {
            "error_type": ConversionErrorType.INVALID_FORMAT,
            "error_category": "unrecoverable",
            "message": "文件格式无效",
            "suggestion": "1. 确认文件扩展名与实际内容匹配\n2. 尝试用对应 Office 程序打开并另存",
            "is_retryable": False,
            "requires_manual_action": True,
        }

    # 未知错误
    return {
        "error_type": ConversionErrorType.UNKNOWN,
        "error_category": "unknown",
        "message": f"未知错误: {str(exception)[:100] if exception else 'N/A'}",
        "suggestion": "1. 查看详细日志获取更多信息\n2. 尝试手动转换该文件\n3. 如问题持续，请联系开发者",
        "is_retryable": True,
        "requires_manual_action": True,
    }


def get_app_path():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))


def is_mac():
    return sys.platform == "darwin"


def is_win():
    return sys.platform == "win32"


def clear_console():
    try:
        if sys.stdout.isatty():
            os.system("cls" if os.name == "nt" else "clear")
    except Exception:
        pass


class FileRegistry:
    """Incremental registry persisted to JSON."""

    def __init__(self, path, base_root=""):
        self.path = path
        self.base_root = os.path.abspath(base_root) if base_root else ""
        self.entries = {}
        self.version = 1
        self.loaded = False

    @staticmethod
    def _legacy_abs_key(path):
        p = os.path.abspath(path)
        if is_win():
            return p.lower()
        return p

    def _is_within_base(self, abs_path):
        if not self.base_root:
            return False
        base_norm = os.path.normcase(os.path.normpath(self.base_root))
        path_norm = os.path.normcase(os.path.normpath(abs_path))
        return path_norm == base_norm or path_norm.startswith(base_norm + os.sep)

    def normalize_path(self, path):
        if not path:
            return ""
        abs_path = os.path.abspath(path)
        if self._is_within_base(abs_path):
            rel = os.path.relpath(abs_path, self.base_root)
            key = rel.replace("\\", "/")
        else:
            key = abs_path.replace("\\", "/")
        if is_win():
            key = key.lower()
        return key

    def load(self):
        self.entries = {}
        if not self.path or not os.path.exists(self.path):
            self.loaded = True
            return
        try:
            with open(self.path, "r", encoding="utf-8") as f:
                data = json.load(f)
            raw_entries = data.get("entries", {})
            if isinstance(raw_entries, dict):
                migrated = {}
                for old_key, old_entry in raw_entries.items():
                    if isinstance(old_entry, dict):
                        source_ref = old_entry.get("source_path") or old_key
                        entry = dict(old_entry)
                    else:
                        source_ref = old_key
                        entry = {"source_path": old_key}

                    new_key = self.normalize_path(source_ref)
                    if not new_key:
                        continue

                    if not entry.get("source_path"):
                        source_ref_str = str(source_ref)
                        if os.path.isabs(source_ref_str):
                            entry["source_path"] = os.path.abspath(source_ref_str)
                        elif self.base_root:
                            entry["source_path"] = os.path.abspath(
                                os.path.join(
                                    self.base_root,
                                    source_ref_str.replace("/", os.sep),
                                )
                            )
                    migrated[new_key] = entry
                self.entries = migrated
            self.version = int(data.get("version", 1) or 1)
        except Exception:
            self.entries = {}
            self.version = 1
        self.loaded = True

    def save(self, run_summary=None):
        folder = os.path.dirname(self.path)
        if folder:
            os.makedirs(folder, exist_ok=True)
        payload = {
            "version": self.version,
            "updated_at": datetime.now().isoformat(timespec="seconds"),
            "key_strategy": "source_rel_forward_slash",
            "entry_count": len(self.entries),
            "entries": self.entries,
        }
        if run_summary:
            payload["last_run"] = run_summary
        with open(self.path, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)

    def get(self, source_path):
        key = self.normalize_path(source_path)
        hit = self.entries.get(key)
        if hit is not None:
            return hit
        # Backward compatibility for legacy absolute-path key style.
        return self.entries.get(self._legacy_abs_key(source_path))

    def set(self, source_path, entry):
        key = self.normalize_path(source_path)
        self.entries[key] = entry

    def keys(self):
        return list(self.entries.keys())


class OfficeConverter:
    def __init__(self, config_path: str, interactive: bool = True):
        """Initialize converter."""
        self.config_path = config_path
        self.interactive = interactive

        self.temp_sandbox = None
        self.temp_sandbox_root = None
        self.failed_dir = None
        self.merge_output_dir = None

        self.engine_type = None
        self.is_running = True
        self.reuse_process = False

        # Runtime mode and strategy defaults (can be overridden by config/CLI/GUI).
        self.run_mode = MODE_CONVERT_THEN_MERGE
        self.collect_mode = COLLECT_MODE_COPY_AND_INDEX
        self.merge_mode = MERGE_MODE_CATEGORY
        self.content_strategy = STRATEGY_STANDARD

        self.price_keywords = []
        self.excluded_folders = []

        # Date filter.
        self.filter_date = None  # datetime object or None
        self.filter_mode = "after"  # "after" or "before"

        # Merge options.
        self.enable_merge_index = False
        self.enable_merge_excel = False

        self.progress_callback = None  # GUI callback hook: func(current, total)
        self.file_plan_callback = None  # Hook: func(file_list)
        self.file_done_callback = None  # Hook: func(result_record)
        self.generated_pdfs = []
        self.generated_merge_outputs = []
        self.generated_merge_markdown_outputs = []
        self.generated_map_outputs = []
        self.generated_markdown_outputs = []
        self.generated_markdown_quality_outputs = []
        self.generated_excel_json_outputs = []
        self.generated_records_json_outputs = []
        self.generated_chromadb_outputs = []
        self.generated_update_package_outputs = []
        self.generated_mshelp_outputs = []
        self.markdown_quality_records = []
        self.conversion_index_records = []
        self.merge_index_records = []
        self.mshelp_records = []
        self.collect_index_path = None
        self.convert_index_path = None
        self.merge_excel_path = None
        self.corpus_manifest_path = None
        self.update_package_manifest_path = None
        self.markdown_quality_report_path = None
        self.chromadb_export_manifest_path = None
        self.incremental_registry_path = ""
        self._incremental_context = None
        self._office_file_counter = 0
        self.perf_metrics = {
            "scan_seconds": 0.0,
            "batch_seconds": 0.0,
            "merge_seconds": 0.0,
            "postprocess_seconds": 0.0,
            "total_seconds": 0.0,
            "convert_core_seconds": 0.0,
            "pdf_wait_seconds": 0.0,
            "markdown_seconds": 0.0,
            "mshelp_merge_seconds": 0.0,
        }

        # Register signals only in main thread (GUI worker threads should not register).
        if threading.current_thread() is threading.main_thread():
            try:
                signal.signal(signal.SIGINT, self.signal_handler)
                signal.signal(signal.SIGTERM, self.signal_handler)
            except Exception:
                pass

        # 1) Load config (fills defaults into self.config).
        self.load_config(config_path)

        # 2) Initialize paths from current config.
        self._init_paths_from_config()

        # 3) Initialize statistics.
        self.stats = {
            "total": 0,
            "success": 0,
            "failed": 0,
            "timeout": 0,
            "skipped": 0,
            "permission_denied": 0,  # 权限错误计数
            "file_locked": 0,  # 文件锁定计数
            "file_corrupted": 0,  # 文件损坏计数
            "com_error": 0,  # COM错误计数
        }
        self.error_records = []  # 简单路径列表（兼容旧逻辑）
        self.detailed_error_records = []  # 结构化错误记录
        self.failed_report_path = None  # 失败报告路径

    def _reset_perf_metrics(self):
        self.perf_metrics = {
            "scan_seconds": 0.0,
            "batch_seconds": 0.0,
            "merge_seconds": 0.0,
            "postprocess_seconds": 0.0,
            "total_seconds": 0.0,
            "convert_core_seconds": 0.0,
            "pdf_wait_seconds": 0.0,
            "markdown_seconds": 0.0,
            "mshelp_merge_seconds": 0.0,
        }

    def _add_perf_seconds(self, key, seconds):
        if key not in self.perf_metrics:
            return
        try:
            value = float(seconds)
        except Exception:
            return
        if value < 0:
            return
        self.perf_metrics[key] += value

    # =============== Base initialization ===============

    def _init_paths_from_config(self):
        """Initialize temp/failed/merge directories from current config."""
        # Temp conversion sandbox root.
        temp_root = self.config.get("temp_sandbox_root", "").strip()
        if temp_root:
            if not os.path.isabs(temp_root):
                temp_root = os.path.abspath(os.path.join(get_app_path(), temp_root))
        else:
            temp_root = tempfile.gettempdir()

        self.temp_sandbox_root = temp_root
        self.temp_sandbox = os.path.join(temp_root, "OfficeToPDF_Sandbox")
        os.makedirs(self.temp_sandbox, exist_ok=True)

        # Failed files directory
        self.failed_dir = os.path.join(self.config["target_folder"], "_FAILED_FILES")
        os.makedirs(self.failed_dir, exist_ok=True)

        # Merge output directory
        self.merge_output_dir = os.path.join(self.config["target_folder"], "_MERGED")
        os.makedirs(self.merge_output_dir, exist_ok=True)

    # =============== Shared display helpers ===============

    def print_welcome(self):
        print("=" * 60)
        print(f" 知喂 ZhiWei · 知识投喂工具  v{__version__}")
        print(" Supports WPS / Microsoft Office, CLI / GUI dual mode")
        print("=" * 60)
        print(f"Config file: {self.config_path}\n")

    def print_step_title(self, text):
        print("\n" + "-" * 60)
        print(text)
        print("-" * 60)

    def get_readable_run_mode(self):
        m = {
            MODE_CONVERT_ONLY: "convert_only",
            MODE_MERGE_ONLY: "merge_only",
            MODE_CONVERT_THEN_MERGE: "convert_then_merge",
            MODE_COLLECT_ONLY: "collect_only",
            MODE_MSHELP_ONLY: "mshelp_only",
        }
        return m.get(self.run_mode, self.run_mode)

    def get_readable_collect_mode(self):
        m = {
            COLLECT_MODE_COPY_AND_INDEX: "copy_and_index",
            COLLECT_MODE_INDEX_ONLY: "index_only",
        }
        return m.get(self.collect_mode, self.collect_mode)

    def get_readable_content_strategy(self):
        m = {
            STRATEGY_STANDARD: "standard",
            STRATEGY_SMART_TAG: "smart_tag",
            STRATEGY_PRICE_ONLY: "price_only",
        }
        return m.get(self.content_strategy, self.content_strategy)

    def get_readable_engine_type(self):
        m = {
            ENGINE_WPS: "WPS Office",
            ENGINE_MS: "Microsoft Office",
            None: "not_used",
        }
        return m.get(self.engine_type, str(self.engine_type))

    def get_readable_merge_mode(self):
        m = {
            MERGE_MODE_CATEGORY: "category_split",
            MERGE_MODE_ALL_IN_ONE: "all_in_one",
        }
        return m.get(self.merge_mode, self.merge_mode)

    @staticmethod
    def compute_convert_output_plan(run_mode, cfg):
        cfg = cfg or {}
        want_pdf = bool(cfg.get("output_enable_pdf", True))
        want_md = bool(cfg.get("output_enable_md", True))
        want_merged = bool(cfg.get("output_enable_merged", True))
        want_independent = bool(cfg.get("output_enable_independent", False))
        enable_merge = bool(cfg.get("enable_merge", True))

        merge_in_convert_phase = (
            run_mode == MODE_CONVERT_THEN_MERGE and want_merged and enable_merge
        )
        need_pdf_for_merge = merge_in_convert_phase and want_pdf
        need_pdf_independent = want_independent and want_pdf
        need_markdown_for_merge = merge_in_convert_phase and want_md
        need_markdown_independent = want_independent and want_md

        return {
            "want_pdf": want_pdf,
            "want_md": want_md,
            "want_merged": want_merged,
            "want_independent": want_independent,
            "need_pdf_for_merge": need_pdf_for_merge,
            "need_pdf_independent": need_pdf_independent,
            "need_final_pdf": need_pdf_for_merge or need_pdf_independent,
            "need_markdown_for_merge": need_markdown_for_merge,
            "need_markdown_independent": need_markdown_independent,
            "need_markdown": need_markdown_for_merge or need_markdown_independent,
        }

    def _get_output_pref(self):
        return {
            "pdf": bool(self.config.get("output_enable_pdf", True)),
            "md": bool(self.config.get("output_enable_md", True)),
            "merged": bool(self.config.get("output_enable_merged", True)),
            "independent": bool(self.config.get("output_enable_independent", False)),
        }

    def _get_merge_convert_submode(self):
        raw = str(
            self.config.get("merge_convert_submode", MERGE_CONVERT_SUBMODE_MERGE_ONLY)
            or MERGE_CONVERT_SUBMODE_MERGE_ONLY
        ).strip()
        if raw not in (
            MERGE_CONVERT_SUBMODE_MERGE_ONLY,
            MERGE_CONVERT_SUBMODE_PDF_TO_MD,
        ):
            return MERGE_CONVERT_SUBMODE_MERGE_ONLY
        return raw

    def print_runtime_summary(self):
        print("\n" + "=" * 60)
        print(" Runtime Summary")
        print("=" * 60)
        print(f"  source_folder : {self.config.get('source_folder', '')}")
        print(f"  target_folder : {self.config.get('target_folder', '')}")
        print(f"  run_mode      : {self.run_mode}")
        print(f"  merge_mode    : {self.merge_mode}")
        pref = self._get_output_pref()
        print(
            f"  output        : pdf={pref['pdf']} md={pref['md']} merged={pref['merged']} independent={pref['independent']}"
        )
        if self.run_mode == MODE_MERGE_ONLY:
            print(f"  merge_submode : {self._get_merge_convert_submode()}")
        print(f"  strategy      : {self.content_strategy}")
        print(f"  reuse_office  : {self._should_reuse_office_app()}")
        print(f"  restart_every : {self._get_office_restart_every()}")
        print("=" * 60)

    def cleanup_all_processes(self):
        apps = (
            ["wps", "et", "wpp", "wpscenter", "wpscloudsvr"]
            if self.engine_type == ENGINE_WPS or self.engine_type is None
            else []
        )
        if self.engine_type == ENGINE_MS or self.engine_type is None:
            apps.extend(["winword", "excel", "powerpnt"])
        for app in apps:
            self._kill_process_by_name(app)

    def _kill_process_by_name(self, app_name):
        if not app_name or not HAS_WIN32:
            return
        try:
            cmd = f"taskkill /F /IM {app_name}.exe"
            subprocess.run(
                cmd, shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
            )
        except Exception:
            pass

    def _check_sandbox_free_space_or_raise(self):
        if not self.config.get("enable_sandbox", True):
            return

        threshold_gb = self.config.get("sandbox_min_free_gb", 10)
        try:
            threshold_gb = float(threshold_gb)
        except Exception:
            threshold_gb = 0
        if threshold_gb <= 0:
            return

        sandbox_root = self.config.get("temp_sandbox_root") or self.config.get(
            "target_folder", ""
        )
        if not sandbox_root:
            return

        probe_path = sandbox_root
        if not os.path.exists(probe_path):
            drive, _ = os.path.splitdrive(probe_path)
            if drive:
                probe_path = drive + os.sep
            else:
                probe_path = os.getcwd()

        try:
            usage = shutil.disk_usage(probe_path)
        except Exception as e:
            logging.warning(f"disk usage check failed for {probe_path}: {e}")
            return

        free_gb = usage.free / (1024 * 1024 * 1024)
        policy = (self.config.get("sandbox_low_space_policy") or "block").lower()

        msg = (
            f"Sandbox free space check: path={probe_path}, "
            f"free={free_gb:.2f} GB, threshold={threshold_gb:.2f} GB, policy={policy}"
        )
        logging.info(msg)

        if free_gb >= threshold_gb:
            return

        warn_text = (
            f"[WARN] Sandbox free space is below threshold: "
            f"{free_gb:.2f} GB < {threshold_gb:.2f} GB (policy={policy})"
        )
        print("\n" + warn_text)
        logging.warning(warn_text)

        if policy == "block":
            raise RuntimeError(
                "Sandbox free space below configured minimum; run blocked by policy."
            )
        # For now, 'confirm' is treated as 'warn' at converter layer;
        # GUI can add higher-level confirmation based on logs if needed.

    def load_config(self, path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                content = f.read().replace("\\", "/")
                self.config = json.loads(content)
        except Exception as e:
            print(f"[ERROR] Failed to load config file: {e}")
            sys.exit(1)

        self.config["source_folder"] = self._get_path_from_config("source_folder")
        self.config["target_folder"] = self._get_path_from_config("target_folder")
        # Normalize source_folders: ensure list and sync source_folder to first
        src_list = self.config.get("source_folders")
        if isinstance(src_list, list) and src_list:
            self.config["source_folders"] = [
                os.path.abspath(str(p).strip()) for p in src_list if str(p).strip()
            ]
            if self.config["source_folders"]:
                self.config["source_folder"] = self.config["source_folders"][0]
        else:
            self.config["source_folders"] = (
                [self.config["source_folder"]]
                if self.config.get("source_folder")
                else []
            )
        self._apply_config_defaults()

    def _apply_config_defaults(self):
        cfg = self.config
        cfg.setdefault("enable_corpus_manifest", True)
        if "output_enable_md" not in cfg and "enable_markdown" in cfg:
            cfg["output_enable_md"] = bool(cfg.get("enable_markdown", True))
        cfg.setdefault("markdown_strip_header_footer", True)
        cfg.setdefault("markdown_structured_headings", True)
        cfg.setdefault("enable_markdown_quality_report", True)
        cfg.setdefault("markdown_quality_sample_limit", 20)
        cfg.setdefault("enable_excel_json", False)
        cfg.setdefault("excel_json_max_rows", 2000)
        cfg.setdefault("excel_json_max_cols", 80)
        cfg.setdefault("excel_json_records_preview", 200)
        cfg.setdefault("excel_json_profile_rows", 500)
        cfg.setdefault("excel_json_include_formulas", True)
        cfg.setdefault("excel_json_extract_sheet_links", True)
        cfg.setdefault("excel_json_include_merged_ranges", True)
        cfg.setdefault("excel_json_formula_sample_limit", 200)
        cfg.setdefault("excel_json_merged_range_limit", 500)
        cfg.setdefault("enable_chromadb_export", False)
        cfg.setdefault("chromadb_persist_dir", "")
        cfg.setdefault("chromadb_collection_name", "office_corpus")
        cfg.setdefault("chromadb_max_chars_per_chunk", 1800)
        cfg.setdefault("chromadb_chunk_overlap", 200)
        cfg.setdefault("chromadb_write_jsonl_fallback", True)
        cfg.setdefault("timeout_seconds", 60)
        cfg.setdefault("enable_sandbox", True)
        cfg.setdefault("default_engine", ENGINE_ASK)
        cfg.setdefault("kill_process_mode", KILL_MODE_ASK)
        cfg.setdefault("auto_retry_failed", False)
        cfg.setdefault("office_reuse_app", True)
        cfg.setdefault("office_restart_every_n_files", 25)
        cfg.setdefault("pdf_wait_seconds", 15)
        cfg.setdefault("ppt_timeout_seconds", cfg.get("timeout_seconds", 60))
        cfg.setdefault("ppt_pdf_wait_seconds", cfg.get("pdf_wait_seconds", 15))
        cfg.setdefault("enable_merge", True)
        cfg.setdefault("max_merge_size_mb", 80)
        cfg.setdefault("output_enable_pdf", True)
        cfg.setdefault("output_enable_md", True)
        cfg.setdefault("output_enable_merged", True)
        cfg.setdefault("output_enable_independent", False)
        # Keep legacy key aligned to avoid split-brain behavior across versions.
        cfg["enable_markdown"] = bool(cfg.get("output_enable_md", True))
        cfg.setdefault("merge_convert_submode", MERGE_CONVERT_SUBMODE_MERGE_ONLY)
        cfg.setdefault("temp_sandbox_root", "")
        cfg.setdefault("sandbox_min_free_gb", 10)
        cfg.setdefault("sandbox_low_space_policy", "block")
        cfg.setdefault("enable_llm_delivery_hub", True)
        cfg.setdefault("llm_delivery_root", "")
        cfg.setdefault("llm_delivery_flatten", True)
        cfg.setdefault("llm_delivery_include_pdf", False)
        cfg.setdefault("enable_gdrive_upload", False)
        cfg.setdefault("gdrive_client_secrets_path", "")
        cfg.setdefault("gdrive_folder_id", "")
        cfg.setdefault("gdrive_token_path", "")
        cfg.setdefault("overwrite_same_size", True)
        cfg.setdefault("merge_mode", MERGE_MODE_CATEGORY)
        cfg.setdefault("merge_source", "source")
        cfg.setdefault("enable_merge_index", False)
        cfg.setdefault("enable_merge_excel", False)
        cfg.setdefault("enable_merge_map", True)
        cfg.setdefault("bookmark_with_short_id", True)
        cfg.setdefault("enable_incremental_mode", False)
        cfg.setdefault("incremental_verify_hash", False)
        cfg.setdefault("incremental_reprocess_renamed", False)
        cfg.setdefault("incremental_registry_path", "")
        cfg.setdefault("source_priority_skip_same_name_pdf", False)
        cfg.setdefault("global_md5_dedup", False)
        cfg.setdefault("enable_update_package", True)
        cfg.setdefault("update_package_root", "")
        cfg.setdefault("cab_7z_path", "")
        cfg.setdefault("mshelpviewer_folder_name", "MSHelpViewer")
        cfg.setdefault("enable_mshelp_merge_output", True)
        cfg.setdefault("enable_mshelp_output_docx", False)
        cfg.setdefault("enable_mshelp_output_pdf", False)

        if "privacy" not in cfg or not isinstance(cfg["privacy"], dict):
            cfg["privacy"] = {}
        cfg["privacy"].setdefault("mask_md5_in_logs", True)

        if "everything" not in cfg or not isinstance(cfg["everything"], dict):
            cfg["everything"] = {}
        cfg["everything"].setdefault("enabled", True)
        cfg["everything"].setdefault("es_path", "")
        cfg["everything"].setdefault("prefer_path_exact", True)
        cfg["everything"].setdefault("timeout_ms", 1500)

        if "listary" not in cfg or not isinstance(cfg["listary"], dict):
            cfg["listary"] = {}
        cfg["listary"].setdefault("enabled", True)
        cfg["listary"].setdefault("copy_query_on_locate", True)

        if "price_keywords" not in cfg or not isinstance(cfg["price_keywords"], list):
            cfg["price_keywords"] = ["报价", "价格表", "Price", "Quotation"]
        self.price_keywords = cfg["price_keywords"]

        if "excluded_folders" not in cfg or not isinstance(
            cfg["excluded_folders"], list
        ):
            cfg["excluded_folders"] = ["temp", "backup", "archive"]
        self.excluded_folders = cfg["excluded_folders"]

        exts = cfg.setdefault("allowed_extensions", {})
        exts.setdefault("word", [".doc", ".docx"])
        exts.setdefault("excel", [".xls", ".xlsx"])
        exts.setdefault("powerpoint", [".ppt", ".pptx"])
        exts.setdefault("pdf", [".pdf"])
        exts.setdefault("cab", [".cab"])

        self.merge_mode = cfg.get("merge_mode", MERGE_MODE_CATEGORY)
        self.run_mode = cfg.get("run_mode", self.run_mode)
        self.collect_mode = cfg.get("collect_mode", self.collect_mode)
        self.content_strategy = cfg.get("content_strategy", self.content_strategy)
        self.enable_merge_index = bool(
            cfg.get("enable_merge_index", self.enable_merge_index)
        )
        self.enable_merge_excel = bool(
            cfg.get("enable_merge_excel", self.enable_merge_excel)
        )

        # In convert-then-merge workflow, merge should always consume target outputs.
        if self.run_mode == MODE_CONVERT_THEN_MERGE:
            cfg["merge_source"] = "target"

    def _should_reuse_office_app(self):
        if not HAS_WIN32 or is_mac():
            return False
        return bool(self.config.get("office_reuse_app", True))

    def _get_office_restart_every(self):
        try:
            value = int(self.config.get("office_restart_every_n_files", 25))
        except Exception:
            value = 25
        return value if value > 0 else 0

    def _get_app_type_for_ext(self, ext):
        ext_lower = (ext or "").lower()
        if ext_lower in self.config.get("allowed_extensions", {}).get("word", []):
            return "word"
        if ext_lower in self.config.get("allowed_extensions", {}).get("excel", []):
            return "excel"
        if ext_lower in self.config.get("allowed_extensions", {}).get("powerpoint", []):
            return "ppt"
        return ""

    def _on_office_file_processed(self, ext):
        if not self._should_reuse_office_app():
            return
        if self.reuse_process:
            return
        restart_every = self._get_office_restart_every()
        if restart_every <= 0:
            return
        app_type = self._get_app_type_for_ext(ext)
        if not app_type:
            return

        self._office_file_counter += 1
        if self._office_file_counter % restart_every != 0:
            return

        logging.info(
            f"[perf] periodic office restart ({app_type}) at file #{self._office_file_counter}"
        )
        self._kill_current_app(app_type, force=True)

    def _get_path_from_config(self, key_base):
        val = None
        if is_win():
            val = self.config.get(f"{key_base}_win")
        elif is_mac():
            val = self.config.get(f"{key_base}_mac")
        if not val:
            val = self.config.get(key_base)
        if val:
            return os.path.abspath(val)
        return ""

    def save_config(self):
        try:
            with open(self.config_path, "w", encoding="utf-8") as f:
                json.dump(self.config, f, indent=4, ensure_ascii=False)
        except Exception:
            pass

    def cli_wizard(self):
        if not self.interactive:
            return
        self.print_welcome()
        self.confirm_config_in_terminal()
        self.ask_for_subfolder()
        self.select_run_mode()
        if self.run_mode == MODE_COLLECT_ONLY:
            self.select_collect_mode()
        elif self.run_mode in (MODE_CONVERT_ONLY, MODE_CONVERT_THEN_MERGE):
            self.select_content_strategy()
        if self.run_mode in (
            MODE_CONVERT_THEN_MERGE,
            MODE_MERGE_ONLY,
        ) and self.config.get("enable_merge", True):
            self.select_merge_mode()
        if self.run_mode not in (MODE_MERGE_ONLY, MODE_COLLECT_ONLY, MODE_MSHELP_ONLY):
            self.select_engine_mode()
            self.check_and_handle_running_processes()
        self._init_paths_from_config()

    def check_and_handle_running_processes(self):
        if self.run_mode in (MODE_MERGE_ONLY, MODE_COLLECT_ONLY, MODE_MSHELP_ONLY):
            return
        mode = self.config.get("kill_process_mode", KILL_MODE_ASK)
        if mode == KILL_MODE_KEEP:
            self.reuse_process = True
            return
        if mode == KILL_MODE_AUTO or not self.interactive:
            self.cleanup_all_processes()
            self.reuse_process = False
            return
        self.reuse_process = False

    def confirm_config_in_terminal(self):
        self.print_step_title("Step 1/4: Confirm Source and Target")
        print("Current paths:")
        print(f"  source: {self.config['source_folder']}")
        print(f"  target: {self.config['target_folder']}")
        print("-" * 60)

        while True:
            choice = input("Modify these paths? [y/N]: ").strip().lower()
            if choice in ("", "n"):
                break
            if choice == "y":
                print("\n=== Edit Paths ===")
                print(f"Current source: {self.config['source_folder']}")
                new_s = (
                    input("New source (Enter to keep): ")
                    .strip()
                    .replace('"', "")
                    .replace("'", "")
                )
                if new_s:
                    self.config["source_folder"] = os.path.abspath(new_s)

                print(f"\nCurrent target: {self.config['target_folder']}")
                new_t = (
                    input("New target (Enter to keep): ")
                    .strip()
                    .replace('"', "")
                    .replace("'", "")
                )
                if new_t:
                    self.config["target_folder"] = os.path.abspath(new_t)

                self.save_config()
                print("Config saved.")
                print("-" * 60)
                print("Updated paths:")
                print(f"  source: {self.config['source_folder']}")
                print(f"  target: {self.config['target_folder']}")
                print("-" * 60)
                continue
            print("Invalid input. Please enter Y or N.")
        print("--> Path confirmation complete.\n")

    def ask_for_subfolder(self):
        self.print_step_title("Step 2/4: Optional Output Subfolder")
        print("You can create a subfolder for this run under target directory.")
        print("-" * 60)
        sub = input("Subfolder name (Enter to skip): ").strip()
        if sub:
            for char in '<>:"/\\|?*':
                sub = sub.replace(char, "")
            self.config["target_folder"] = os.path.abspath(
                os.path.join(self.config["target_folder"], sub)
            )
            print(f"--> Run target folder: {self.config['target_folder']}")
        else:
            print("--> Keep original target folder from config.")

    def select_run_mode(self):
        self.print_step_title("Step 3/4: Select Run Mode")
        print("  [1] Convert only")
        print("  [2] Merge & convert")
        print("  [3] Convert then merge (recommended)")
        print("  [4] Collect / deduplicate only")
        print("  [5] MSHelp API docs (CAB->MD, merged package)")
        print("-" * 60)
        choice = input("Choose (1/2/3/4/5, default 3): ").strip()
        if choice == "1":
            self.run_mode = MODE_CONVERT_ONLY
        elif choice == "2":
            self.run_mode = MODE_MERGE_ONLY
        elif choice == "4":
            self.run_mode = MODE_COLLECT_ONLY
        elif choice == "5":
            self.run_mode = MODE_MSHELP_ONLY
        else:
            self.run_mode = MODE_CONVERT_THEN_MERGE
        print(f"--> Run mode: {self.get_readable_run_mode()} ({self.run_mode})")

    def select_collect_mode(self):
        self.print_step_title("Select Collect Sub-Mode")
        print("  [1] Deduplicate + copy + Excel index")
        print("  [2] Generate Excel index only (no copy)")
        print("-" * 60)
        choice = input("Choose (1/2, default 1): ").strip()
        if choice == "2":
            self.collect_mode = COLLECT_MODE_INDEX_ONLY
        else:
            self.collect_mode = COLLECT_MODE_COPY_AND_INDEX
        print(
            f"--> Collect mode: {self.get_readable_collect_mode()} ({self.collect_mode})"
        )

    def select_merge_mode(self):
        if not self.config.get("enable_merge", True):
            self.merge_mode = MERGE_MODE_CATEGORY
            return

        cfg_mode = self.config.get("merge_mode", MERGE_MODE_CATEGORY)
        if cfg_mode in (MERGE_MODE_ALL_IN_ONE, MERGE_MODE_CATEGORY):
            self.merge_mode = cfg_mode
            print(
                f"--> Merge mode from config: {self.get_readable_merge_mode()} ({self.merge_mode})"
            )
            return

        self.print_step_title("Select Merge Mode")
        print("  [1] Category split (Price/Word/Excel/PPT/PDF)")
        print("  [2] All in one PDF")
        print("-" * 60)
        choice = input("Choose (1/2, default 1): ").strip()
        if choice == "2":
            self.merge_mode = MERGE_MODE_ALL_IN_ONE
        else:
            self.merge_mode = MERGE_MODE_CATEGORY
        print(f"--> Merge mode: {self.get_readable_merge_mode()} ({self.merge_mode})")

    def select_content_strategy(self):
        self.print_step_title("Step 4/4: Select Content Strategy")
        print("  [1] Standard classification")
        print("  [2] Smart tag (price keyword hit)")
        print("  [3] Price only")
        print("-" * 60)
        print(f"Current keywords: {self.price_keywords}")
        choice = input("Choose (1/2/3, default 1): ").strip()
        if choice == "2":
            self.content_strategy = STRATEGY_SMART_TAG
        elif choice == "3":
            self.content_strategy = STRATEGY_PRICE_ONLY
        else:
            self.content_strategy = STRATEGY_STANDARD
        print(
            f"--> Strategy: {self.get_readable_content_strategy()} ({self.content_strategy})\n"
        )

    def select_engine_mode(self):
        default = self.config.get("default_engine", ENGINE_ASK)
        if default == ENGINE_WPS:
            self.engine_type = ENGINE_WPS
            print("--> [auto] engine: WPS Office")
            return
        if default == ENGINE_MS:
            self.engine_type = ENGINE_MS
            print("--> [auto] engine: Microsoft Office")
            return

        self.print_step_title("Select Office Engine")
        print("  [1] WPS Office")
        print("  [2] Microsoft Office")
        print("-" * 60)
        while True:
            choice = input("Choose (1/2, default 1): ").strip()
            if choice in ("", "1"):
                self.engine_type = ENGINE_WPS
                break
            if choice == "2":
                self.engine_type = ENGINE_MS
                break
            print("Invalid input. Please enter 1 or 2.")
        print(f"--> Selected: {self.get_readable_engine_type()} ({self.engine_type})\n")

    def setup_logging(self):
        log_dir = self.config.get("log_folder", "./logs")
        if not os.path.isabs(log_dir):
            log_dir = os.path.join(get_app_path(), log_dir)
        os.makedirs(log_dir, exist_ok=True)

        self.log_path = os.path.join(
            log_dir, f"conversion_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        )

        logging.basicConfig(
            filename=self.log_path,
            level=logging.INFO,
            format="%(message)s",
            encoding="utf-8",
            force=True,
        )
        console = logging.StreamHandler()
        console.setLevel(logging.INFO)
        console.setFormatter(logging.Formatter("%(message)s"))
        logging.getLogger("").addHandler(console)

        engine_label = self.engine_type.upper() if self.engine_type else "N/A"
        with open(self.log_path, "a", encoding="utf-8") as f:
            f.write(f"{datetime.now()} === Task Start (v{__version__}) ===\n")
            f.write(f"run_mode: {self.run_mode} ({self.get_readable_run_mode()})\n")
            if self.run_mode in (MODE_CONVERT_ONLY, MODE_CONVERT_THEN_MERGE):
                f.write(
                    f"content_strategy: {self.content_strategy} ({self.get_readable_content_strategy()})\n"
                )
            if self.run_mode in (MODE_CONVERT_THEN_MERGE, MODE_MERGE_ONLY):
                f.write(
                    f"merge_mode: {self.merge_mode} ({self.get_readable_merge_mode()})\n"
                )
            f.write(f"engine: {engine_label}\n")
            f.write(f"source_folder: {self.config['source_folder']}\n")
            f.write(f"target_folder: {self.config['target_folder']}\n")
            f.write(f"office_reuse_app: {self._should_reuse_office_app()}\n")
            f.write(
                f"office_restart_every_n_files: {self._get_office_restart_every()}\n"
            )
            if self.config.get("enable_sandbox", True):
                f.write(f"sandbox: enabled | temp: {self.temp_sandbox}\n")
            else:
                f.write(
                    f"sandbox: disabled | temp: {self.temp_sandbox} (PDF temp only)\n"
                )
            if self.run_mode in (MODE_CONVERT_THEN_MERGE, MODE_MERGE_ONLY):
                f.write(f"merge_output_dir: {self.merge_output_dir}\n")
            f.write("=" * 60 + "\n")

    def _kill_current_app(self, app_type, force=False):
        if self.reuse_process and not force:
            return
        name_map = {
            ENGINE_WPS: {"word": "wps", "excel": "et", "ppt": "wpp"},
            ENGINE_MS: {"word": "winword", "excel": "excel", "ppt": "powerpnt"},
        }
        if self.engine_type not in name_map:
            return
        app_name = name_map[self.engine_type].get(app_type, "")
        self._kill_process_by_name(app_name)

    def _get_local_app(self, app_type):
        if not HAS_WIN32:
            raise RuntimeError(
                "Current system does not support Windows COM; Office conversion is unavailable."
            )

        pythoncom.CoInitialize()
        if self.engine_type == ENGINE_WPS:
            prog_id = {
                "word": "Kwps.Application",
                "excel": "Ket.Application",
                "ppt": "Kwpp.Application",
            }.get(app_type)
        else:
            prog_id = {
                "word": "Word.Application",
                "excel": "Excel.Application",
                "ppt": "PowerPoint.Application",
            }.get(app_type)
        app = None
        try:
            app = win32com.client.Dispatch(prog_id)
        except Exception:
            app = win32com.client.DispatchEx(prog_id)

        try:
            app.Visible = False
            if app_type != "ppt":
                app.DisplayAlerts = False
        except Exception:
            pass

        if self.engine_type == ENGINE_MS and app_type == "excel":
            try:
                app.AskToUpdateLinks = False
            except Exception:
                pass

        return app

    def close_office_apps(self):
        if not self.reuse_process and self.run_mode not in (
            MODE_MERGE_ONLY,
            MODE_COLLECT_ONLY,
        ):
            self.cleanup_all_processes()

    # =============== Path and conflict handling ===============

    def get_target_path(self, source_file_path, ext, prefix_override=None):
        filename = os.path.basename(source_file_path)
        base_name = os.path.splitext(filename)[0]
        ext_lower = ext.lower()

        if prefix_override:
            prefix = prefix_override
        else:
            prefix = ""
            word_exts = self.config["allowed_extensions"].get("word", [])
            excel_exts = self.config["allowed_extensions"].get("excel", [])
            ppt_exts = self.config["allowed_extensions"].get("powerpoint", [])
            pdf_exts = self.config["allowed_extensions"].get("pdf", [])

            if ext_lower in word_exts:
                prefix = "Word_"
            elif ext_lower in excel_exts:
                prefix = "Excel_"
            elif ext_lower in ppt_exts:
                prefix = "PPT_"
            elif ext_lower in pdf_exts:
                prefix = "PDF_"

        new_filename = f"{prefix}{base_name}.pdf"
        return os.path.join(self.config["target_folder"], new_filename)

    def handle_file_conflict(self, temp_pdf_path, target_pdf_path):
        if not os.path.exists(target_pdf_path):
            os.makedirs(os.path.dirname(target_pdf_path), exist_ok=True)
            shutil.move(temp_pdf_path, target_pdf_path)
            return "success", target_pdf_path

        if os.path.getsize(temp_pdf_path) == os.path.getsize(target_pdf_path):
            try:
                os.remove(target_pdf_path)
                shutil.move(temp_pdf_path, target_pdf_path)
                return "overwrite", target_pdf_path
            except Exception:
                return "overwrite_failed", target_pdf_path
        else:
            conflict_dir = os.path.join(os.path.dirname(target_pdf_path), "conflicts")
            os.makedirs(conflict_dir, exist_ok=True)
            fname = os.path.splitext(os.path.basename(target_pdf_path))[0]
            ts = datetime.now().strftime("%Y%m%d%H%M%S")
            new_path = os.path.join(conflict_dir, f"{fname}_{ts}.pdf")
            shutil.move(temp_pdf_path, new_path)
            return "conflict_saved", new_path

    # =============== Content scanning ===============

    def scan_pdf_content(self, pdf_path):
        if not HAS_PYPDF:
            return False
        try:
            reader = PdfReader(pdf_path)
            max_pages = min(len(reader.pages), 5)
            for i in range(max_pages):
                text = reader.pages[i].extract_text()
                if text:
                    for kw in self.price_keywords:
                        if kw in text:
                            return True
        except Exception:
            pass
        return False

    def scan_excel_content_in_thread(self, workbook):
        try:
            for sheet in workbook.Worksheets:
                try:
                    data = sheet.UsedRange.Value
                    if not data:
                        continue
                    if not isinstance(data, tuple):
                        data = ((data,),)
                    for row in data:
                        if not row:
                            continue
                        for cell in row:
                            if cell and isinstance(cell, str):
                                for kw in self.price_keywords:
                                    if kw in cell:
                                        logging.info(
                                            f"Excel matched keyword [{kw}] in sheet: {sheet.Name}"
                                        )
                                        return True
                except Exception:
                    continue
        except Exception as e:
            logging.warning(f"scan Excel content failed: {e}")
        return False

    # =============== COM safe execution helpers ===============

    def _safe_exec(self, func, *args, retries=3, **kwargs):
        for attempt in range(retries + 1):
            if not self.is_running:
                raise Exception("program stopped")
            try:
                return func(*args, **kwargs)
            except pywintypes.com_error as e:
                error_code = e.hresult
                if error_code == ERR_RPC_SERVER_BUSY:
                    time.sleep(random.randint(2, 5))
                    continue
                if attempt < retries:
                    time.sleep(1)
                    continue
                raise Exception(f"COM error ({error_code}): {e}")
            except Exception:
                if attempt < retries:
                    time.sleep(1)
                    continue
                raise

    def _unblock_file(self, file_path):
        try:
            zone_path = file_path + ":Zone.Identifier"
            try:
                os.remove(zone_path)
            except Exception:
                pass
        except Exception:
            pass

    def _setup_excel_pages(self, workbook):
        try:
            for sheet in workbook.Worksheets:
                try:
                    _ = sheet.UsedRange
                    try:
                        sheet.ResetAllPageBreaks()
                    except Exception:
                        pass
                    ps = sheet.PageSetup
                    try:
                        ps.PrintArea = ""
                    except Exception:
                        pass
                    ps.Zoom = False
                    ps.Orientation = 2
                    ps.FitToPagesWide = 1
                    ps.FitToPagesTall = False
                    ps.CenterHorizontally = True
                    try:
                        ps.LeftMargin = 20
                        ps.RightMargin = 20
                        ps.TopMargin = 20
                        ps.BottomMargin = 20
                    except Exception:
                        pass
                except Exception:
                    pass
        except Exception:
            pass

    # =============== MacOS Automation Support (Stub/Future) ===============

    def _convert_on_mac(self, file_source, sandbox_target_pdf, ext):
        """"""
        if not is_mac():
            return False

        # Simplistic implementation structure - user may need to expand this
        # or rely on LibreOffice CLI if preferred.
        # For now, we try to use 'soffice' (LibreOffice) if available, or just log error.

        # Check for soffice first (LibreOffice)
        soffice = shutil.which("soffice")
        if soffice:
            cmd = [
                soffice,
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                os.path.dirname(sandbox_target_pdf),
                file_source,
            ]
            try:
                subprocess.run(
                    cmd,
                    check=True,
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                )
                # LibreOffice output name might differ, need to rename to sandbox_target_pdf
                # But outdir is sandbox dir.
                # Assuming standard naming: source.docx -> source.pdf
                base = os.path.splitext(os.path.basename(file_source))[0]
                possible_output = os.path.join(
                    os.path.dirname(sandbox_target_pdf), base + ".pdf"
                )
                if (
                    os.path.exists(possible_output)
                    and possible_output != sandbox_target_pdf
                ):
                    shutil.move(possible_output, sandbox_target_pdf)
                return True
            except Exception as e:
                logging.error(f"LibreOffice conversion failed: {e}")

        # If no soffice, try basic AppleScript for Word/Excel?
        # (This is complex to implement fully robustly in one step without testing environment)
        logging.warning(
            "macOS Office Automation not fully implemented. Install LibreOffice for best results."
        )
        return False

    # =============== Core conversion ===============

    def convert_logic_in_thread(
        self, file_source, sandbox_target_pdf, ext, result_context
    ):
        app = None
        doc = None
        try:
            if is_mac():
                # macOS specific path
                if self._convert_on_mac(file_source, sandbox_target_pdf, ext):
                    return
                # If failed, fall through logic (which might fail if no win32com)
                if not HAS_WIN32:
                    raise RuntimeError(
                        "macOS conversion failed (LibreOffice not found?) and win32com not available."
                    )

            if ext in self.config["allowed_extensions"]["word"]:
                app = self._get_local_app("word")
                try:
                    if self.engine_type == ENGINE_WPS:
                        try:
                            doc = self._safe_exec(
                                app.Documents.Open, file_source, ReadOnly=True
                            )
                        except Exception:
                            doc = self._safe_exec(app.Documents.Open, file_source)
                    else:
                        doc = self._safe_exec(
                            app.Documents.Open,
                            file_source,
                            ReadOnly=True,
                            Visible=False,
                            OpenAndRepair=True,
                        )
                    self._safe_exec(
                        doc.ExportAsFixedFormat, sandbox_target_pdf, wdFormatPDF
                    )
                finally:
                    if doc:
                        try:
                            doc.Close(SaveChanges=False)
                        except Exception:
                            pass

            elif ext in self.config["allowed_extensions"]["excel"]:
                app = self._get_local_app("excel")
                try:
                    if self.engine_type == ENGINE_WPS:
                        try:
                            doc = self._safe_exec(
                                app.Workbooks.Open, file_source, ReadOnly=True
                            )
                        except Exception:
                            doc = self._safe_exec(app.Workbooks.Open, file_source)
                        if (
                            not result_context.get("skip_scan", False)
                            and self.content_strategy != STRATEGY_STANDARD
                        ):
                            has_kw = self.scan_excel_content_in_thread(doc)
                            if has_kw:
                                result_context["is_price"] = True
                            elif self.content_strategy == STRATEGY_PRICE_ONLY:
                                result_context["scan_aborted"] = True
                                return
                        self._setup_excel_pages(doc)
                        try:
                            self._safe_exec(
                                doc.ExportAsFixedFormat, xlTypePDF, sandbox_target_pdf
                            )
                        except Exception:
                            if os.path.exists(sandbox_target_pdf):
                                os.remove(sandbox_target_pdf)
                            self._safe_exec(
                                doc.SaveAs, sandbox_target_pdf, FileFormat=xlPDF_SaveAs
                            )
                    else:
                        doc = self._safe_exec(
                            app.Workbooks.Open,
                            file_source,
                            UpdateLinks=0,
                            ReadOnly=True,
                            IgnoreReadOnlyRecommended=True,
                            CorruptLoad=xlRepairFile,
                        )
                        if (
                            not result_context.get("skip_scan", False)
                            and self.content_strategy != STRATEGY_STANDARD
                        ):
                            has_kw = self.scan_excel_content_in_thread(doc)
                            if has_kw:
                                result_context["is_price"] = True
                            elif self.content_strategy == STRATEGY_PRICE_ONLY:
                                result_context["scan_aborted"] = True
                                return
                        self._setup_excel_pages(doc)
                        self._safe_exec(
                            doc.ExportAsFixedFormat, xlTypePDF, sandbox_target_pdf
                        )
                finally:
                    if doc:
                        try:
                            doc.Close(SaveChanges=False)
                        except Exception:
                            pass

            elif ext in self.config["allowed_extensions"]["powerpoint"]:
                app = self._get_local_app("ppt")
                try:
                    if self.engine_type == ENGINE_WPS:
                        try:
                            doc = self._safe_exec(
                                app.Presentations.Open, file_source, WithWindow=False
                            )
                        except Exception:
                            doc = self._safe_exec(app.Presentations.Open, file_source)
                        self._safe_exec(doc.SaveCopyAs, sandbox_target_pdf, ppSaveAsPDF)
                    else:
                        doc = self._safe_exec(
                            app.Presentations.Open,
                            file_source,
                            WithWindow=False,
                            ReadOnly=True,
                        )
                        try:
                            self._safe_exec(
                                doc.ExportAsFixedFormat,
                                sandbox_target_pdf,
                                ppFixedFormatTypePDF,
                            )
                        except Exception:
                            if os.path.exists(sandbox_target_pdf):
                                os.remove(sandbox_target_pdf)
                            self._safe_exec(
                                doc.SaveCopyAs, sandbox_target_pdf, ppSaveAsPDF
                            )
                finally:
                    if doc:
                        try:
                            doc.Close()
                        except Exception:
                            pass
        finally:
            if app:
                try:
                    if not self._should_reuse_office_app():
                        app.Quit()
                except Exception:
                    pass
            pythoncom.CoUninitialize()

    def copy_pdf_direct(self, source, temp_target):
        try:
            shutil.copy2(source, temp_target)
        except Exception as e:
            raise Exception(f"[PDF copy failed] {e}")

    def quarantine_failed_file(self, source_path, should_copy=True):
        if not should_copy:
            return
        try:
            fname = os.path.basename(source_path)
            target = os.path.join(self.failed_dir, fname)
            if os.path.exists(target):
                name, ext = os.path.splitext(fname)
                target = os.path.join(
                    self.failed_dir, f"{name}_{datetime.now().strftime('%H%M%S')}{ext}"
                )
            shutil.copy2(source_path, target)
        except Exception:
            pass

    def record_detailed_error(self, source_path, exception, context=None):
        """
        记录详细的错误信息，包括错误分类和处理建议。

        Args:
            source_path: 失败文件的路径
            exception: 异常对象
            context: 额外上下文信息（如运行模式、转换引擎等）

        Returns:
            dict: 错误详情字典
        """
        error_info = classify_conversion_error(exception, str(source_path))

        record = {
            "source_path": os.path.abspath(source_path),
            "file_name": os.path.basename(source_path),
            "error_type": error_info["error_type"],
            "error_category": error_info["error_category"],
            "message": error_info["message"],
            "suggestion": error_info["suggestion"],
            "is_retryable": error_info["is_retryable"],
            "requires_manual_action": error_info["requires_manual_action"],
            "raw_error": str(exception)[:500] if exception else "",
            "timestamp": datetime.now().isoformat(),
            "context": context or {},
        }

        self.detailed_error_records.append(record)

        # 更新分类统计
        error_type = error_info["error_type"]
        if error_type in self.stats:
            self.stats[error_type] += 1

        return record

    def export_failed_files_report(self, output_dir=None):
        """
        导出失败文件详细报告，包括 JSON 和 可读的 TXT 格式。

        Args:
            output_dir: 输出目录，默认为 target_folder

        Returns:
            dict: {"json_path": ..., "txt_path": ..., "summary": ...}
        """
        if not self.detailed_error_records:
            return {"json_path": None, "txt_path": None, "summary": "无失败记录"}

        output_dir = output_dir or self.config.get("target_folder", ".")
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

        # 1. 导出 JSON 报告
        json_path = os.path.join(output_dir, f"failed_files_report_{timestamp}.json")
        report_data = {
            "generated_at": datetime.now().isoformat(),
            "run_mode": self.get_readable_run_mode(),
            "total_failed": len(self.detailed_error_records),
            "statistics": {
                "by_error_type": {},
                "by_category": {"retryable": 0, "needs_manual": 0, "unrecoverable": 0},
                "retryable_count": 0,
                "manual_action_count": 0,
            },
            "records": self.detailed_error_records,
        }

        # 统计分类
        for record in self.detailed_error_records:
            et = record["error_type"]
            report_data["statistics"]["by_error_type"][et] = (
                report_data["statistics"]["by_error_type"].get(et, 0) + 1
            )

            cat = record["error_category"]
            if cat in report_data["statistics"]["by_category"]:
                report_data["statistics"]["by_category"][cat] += 1

            if record["is_retryable"]:
                report_data["statistics"]["retryable_count"] += 1
            if record["requires_manual_action"]:
                report_data["statistics"]["manual_action_count"] += 1

        try:
            with open(json_path, "w", encoding="utf-8") as f:
                json.dump(report_data, f, indent=2, ensure_ascii=False)
        except Exception as e:
            logging.error(f"Failed to write JSON report: {e}")
            json_path = None

        # 2. 导出可读的 TXT 报告
        txt_path = os.path.join(output_dir, f"failed_files_report_{timestamp}.txt")
        try:
            lines = []
            lines.append("=" * 70)
            lines.append("知喂 (ZhiWei) - 转换失败文件报告")
            lines.append(f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            lines.append(f"运行模式: {self.get_readable_run_mode()}")
            lines.append("=" * 70)
            lines.append("")

            # 统计摘要
            lines.append("## 统计摘要")
            lines.append("-" * 70)
            lines.append(f"总失败数: {len(self.detailed_error_records)}")
            lines.append(f"可重试: {report_data['statistics']['retryable_count']}")
            lines.append(
                f"需人工处理: {report_data['statistics']['manual_action_count']}"
            )
            lines.append("")

            lines.append("### 按错误类型分布:")
            for et, count in sorted(
                report_data["statistics"]["by_error_type"].items(), key=lambda x: -x[1]
            ):
                lines.append(f"  - {et}: {count}")
            lines.append("")

            # 详细列表
            lines.append("## 失败文件详情")
            lines.append("-" * 70)

            for i, record in enumerate(self.detailed_error_records, 1):
                lines.append(f"\n[{i}] {record['file_name']}")
                lines.append(f"    路径: {record['source_path']}")
                lines.append(f"    错误类型: {record['error_type']}")
                lines.append(f"    错误信息: {record['message']}")
                lines.append(f"    可重试: {'是' if record['is_retryable'] else '否'}")
                if record["requires_manual_action"]:
                    lines.append(f"    处理建议:")
                    for line in record["suggestion"].split("\n"):
                        lines.append(f"        {line}")

            lines.append("")
            lines.append("=" * 70)
            lines.append("报告结束")
            lines.append("")
            lines.append("提示:")
            lines.append("- 可重试的文件可点击「重试失败文件」按钮重新处理")
            lines.append("- 需人工处理的文件请按照建议手动修复后重新运行")
            lines.append("- JSON 格式详细报告: " + (json_path or "生成失败"))
            lines.append("=" * 70)

            with open(txt_path, "w", encoding="utf-8") as f:
                f.write("\n".join(lines))
        except Exception as e:
            logging.error(f"Failed to write TXT report: {e}")
            txt_path = None

        self.failed_report_path = txt_path

        return {
            "json_path": json_path,
            "txt_path": txt_path,
            "summary": f"共 {len(self.detailed_error_records)} 个失败文件，"
            f"其中 {report_data['statistics']['retryable_count']} 个可重试",
        }

    def get_error_summary_for_display(self):
        """
        获取用于 GUI 显示的错误摘要。

        Returns:
            dict: 按错误类型分组的文件列表和处理建议
        """
        summary = {}

        for record in self.detailed_error_records:
            et = record["error_type"]
            if et not in summary:
                summary[et] = {
                    "message": record["message"],
                    "suggestion": record["suggestion"],
                    "is_retryable": record["is_retryable"],
                    "requires_manual_action": record["requires_manual_action"],
                    "files": [],
                }
            summary[et]["files"].append(record["file_name"])

        return summary

    @staticmethod
    def _find_files_recursive(root_dir, exts):
        results = []
        if not root_dir or not os.path.isdir(root_dir):
            return results
        ext_set = tuple(str(e).lower() for e in (exts or []))
        for current_root, _, files in os.walk(root_dir):
            for name in files:
                if str(name).lower().endswith(ext_set):
                    results.append(os.path.join(current_root, name))
        return results

    def _extract_cab_with_fallback(self, cab_path, extract_dir):
        cab_abs = os.path.abspath(cab_path)
        extract_abs = os.path.abspath(extract_dir)
        os.makedirs(extract_abs, exist_ok=True)

        expand_ok = False
        if is_win():
            try:
                cmd_expand = ["expand", cab_abs, "-F:*", extract_abs]
                subprocess.run(
                    cmd_expand,
                    capture_output=True,
                    text=True,
                    encoding="gbk",
                    errors="ignore",
                    check=True,
                )
                expand_ok = True
            except Exception:
                expand_ok = False

        if expand_ok and self._find_files_recursive(
            extract_abs, (".mshc", ".htm", ".html")
        ):
            return

        seven_zip = str(self.config.get("cab_7z_path", "") or "").strip()
        if seven_zip:
            if not os.path.isabs(seven_zip):
                seven_zip = os.path.abspath(os.path.join(get_app_path(), seven_zip))
            if not os.path.isfile(seven_zip):
                seven_zip = ""
        if not seven_zip:
            seven_zip = shutil.which("7z") or shutil.which("7za") or ""
        if not seven_zip:
            raise RuntimeError(
                "CAB extraction fallback requires 7z. Please install 7-Zip or set cab_7z_path."
            )

        cmd_7z = [seven_zip, "x", cab_abs, f"-o{extract_abs}", "-y"]
        subprocess.run(
            cmd_7z,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="ignore",
            check=True,
        )

        if not self._find_files_recursive(extract_abs, (".mshc", ".htm", ".html")):
            raise RuntimeError(
                f"CAB extraction produced no MSHC/HTML payload: {cab_path}"
            )

    @staticmethod
    def _extract_mshc_payload(mshc_path, content_dir):
        os.makedirs(content_dir, exist_ok=True)
        with zipfile.ZipFile(mshc_path, "r") as zf:
            zf.extractall(content_dir)

    @staticmethod
    def _meta_content_by_names(soup, names):
        if not soup:
            return ""
        name_set = {str(n).strip().lower() for n in names}
        for meta in soup.find_all("meta"):
            meta_name = str(meta.get("name", "")).strip().lower()
            if meta_name in name_set:
                return str(meta.get("content", "") or "").strip()
        return ""

    def _parse_mshelp_topics(self, html_root):
        html_files = self._find_files_recursive(html_root, (".htm", ".html"))
        if not html_files:
            return []

        topics = {}
        for fpath in html_files:
            rel_path = os.path.relpath(fpath, html_root).replace("\\", "/")
            topic_id = rel_path
            parent_id = ""
            title = os.path.splitext(os.path.basename(fpath))[0]

            if HAS_BS4:
                try:
                    with open(fpath, "rb") as f:
                        raw = f.read()
                    soup = BeautifulSoup(raw, "html.parser")
                    meta_id = self._meta_content_by_names(soup, ["Microsoft.Help.Id"])
                    meta_parent = self._meta_content_by_names(
                        soup, ["Microsoft.Help.TocParent"]
                    )
                    meta_title = self._meta_content_by_names(soup, ["Title"])
                    title_tag = soup.find("title")
                    if meta_id:
                        topic_id = meta_id
                    if meta_parent:
                        parent_id = meta_parent
                    if meta_title:
                        title = meta_title
                    elif title_tag and title_tag.get_text(strip=True):
                        title = title_tag.get_text(strip=True)
                except Exception:
                    pass

            topics[topic_id] = {
                "id": topic_id,
                "parent": parent_id,
                "title": title or os.path.basename(fpath),
                "file": fpath,
                "children": [],
            }

        if not topics:
            return []

        roots = []
        for tid, topic in topics.items():
            pid = str(topic.get("parent", "") or "").strip()
            if not pid or pid in ("-1", tid) or pid not in topics:
                roots.append(tid)
            else:
                topics[pid]["children"].append(tid)

        for topic in topics.values():
            topic["children"].sort(
                key=lambda cid: (
                    topics[cid].get("title", ""),
                    topics[cid].get("file", ""),
                )
            )
        roots.sort(
            key=lambda rid: (topics[rid].get("title", ""), topics[rid].get("file", ""))
        )

        ordered = []
        visited = set()

        def walk(topic_id):
            if topic_id in visited or topic_id not in topics:
                return
            visited.add(topic_id)
            ordered.append(topics[topic_id])
            for child_id in topics[topic_id].get("children", []):
                walk(child_id)

        for rid in roots:
            walk(rid)
        for tid in sorted(topics.keys()):
            walk(tid)

        return ordered

    @staticmethod
    def _normalize_md_line(text):
        return re.sub(r"\s+", " ", str(text or "")).strip()

    def _table_to_markdown_lines(self, table_tag):
        rows = []
        for tr in table_tag.find_all("tr"):
            cells = tr.find_all(["th", "td"])
            if not cells:
                continue
            row = [self._normalize_md_line(c.get_text(" ", strip=True)) for c in cells]
            rows.append(row)
        if not rows:
            return []

        width = max(len(r) for r in rows)
        norm_rows = []
        for r in rows:
            norm_rows.append(r + [""] * (width - len(r)))

        header = [c.replace("|", "\\|") for c in norm_rows[0]]
        lines = [
            "| " + " | ".join(header) + " |",
            "| " + " | ".join(["---"] * width) + " |",
        ]
        for r in norm_rows[1:]:
            escaped = [c.replace("|", "\\|") for c in r]
            lines.append("| " + " | ".join(escaped) + " |")
        return lines

    def _render_html_to_markdown(self, html_path):
        if not HAS_BS4:
            with open(html_path, "r", encoding="utf-8", errors="ignore") as f:
                raw = f.read()
            text = re.sub(r"<[^>]+>", " ", raw)
            text = re.sub(r"\s+", " ", text).strip()
            return text

        with open(html_path, "rb") as f:
            soup = BeautifulSoup(f.read(), "html.parser")

        for tag in soup(["script", "style", "noscript", "svg"]):
            tag.decompose()

        body = soup.body if soup.body else soup
        lines = []

        def append_para(text):
            t = self._normalize_md_line(text)
            if t:
                lines.append(t)
                lines.append("")

        def render(node):
            name = getattr(node, "name", None)
            if not name:
                return
            name = str(name).lower()
            if name in ("script", "style", "noscript", "svg"):
                return

            if name in ("h1", "h2", "h3", "h4", "h5", "h6"):
                level = int(name[1])
                title = self._normalize_md_line(node.get_text(" ", strip=True))
                if title:
                    lines.append("#" * level + " " + title)
                    lines.append("")
                return

            if name == "pre":
                code = node.get_text("\n", strip=False).strip("\n")
                if code:
                    lines.append("```text")
                    lines.append(code)
                    lines.append("```")
                    lines.append("")
                return

            if name in ("ul", "ol"):
                lis = node.find_all("li", recursive=False)
                if lis:
                    for idx, li in enumerate(lis, 1):
                        item_text = self._normalize_md_line(
                            li.get_text(" ", strip=True)
                        )
                        if not item_text:
                            continue
                        prefix = f"{idx}. " if name == "ol" else "- "
                        lines.append(prefix + item_text)
                    lines.append("")
                return

            if name == "table":
                table_lines = self._table_to_markdown_lines(node)
                if table_lines:
                    lines.extend(table_lines)
                    lines.append("")
                return

            if name in ("p", "blockquote"):
                append_para(node.get_text(" ", strip=True))
                return

            if name in (
                "article",
                "section",
                "main",
                "body",
                "div",
                "header",
                "footer",
                "aside",
                "nav",
            ):
                for child in node.children:
                    if getattr(child, "name", None):
                        render(child)
                return

            child_tags = [c for c in node.children if getattr(c, "name", None)]
            if child_tags:
                for child in child_tags:
                    render(child)
                return
            append_para(node.get_text(" ", strip=True))

        for child in body.children:
            if getattr(child, "name", None):
                render(child)

        compact = []
        blank = False
        for line in lines:
            if str(line).strip():
                compact.append(line.rstrip())
                blank = False
            else:
                if not blank:
                    compact.append("")
                blank = True
        return "\n".join(compact).strip()

    def _convert_cab_to_markdown(self, cab_path, source_path_for_output):
        if not os.path.exists(cab_path):
            raise RuntimeError(f"CAB file not found: {cab_path}")

        if not HAS_BS4:
            logging.warning(
                "beautifulsoup4 is not installed; CAB markdown quality may be limited."
            )

        temp_root = os.path.join(self.temp_sandbox, f"cab_{uuid.uuid4().hex}")
        extract_dir = os.path.join(temp_root, "cab_extract")
        content_dir = os.path.join(temp_root, "mshc_content")

        try:
            self._extract_cab_with_fallback(cab_path, extract_dir)

            mshc_files = self._find_files_recursive(extract_dir, (".mshc",))
            html_root = extract_dir
            if mshc_files:
                self._extract_mshc_payload(mshc_files[0], content_dir)
                html_root = content_dir

            topics = self._parse_mshelp_topics(html_root)
            if not topics:
                raise RuntimeError(f"no parseable help topics in CAB: {cab_path}")

            md_path = self._build_ai_output_path_from_source(
                source_path_for_output, "Markdown", ".md"
            )
            if not md_path:
                raise RuntimeError(f"failed to build markdown path for CAB: {cab_path}")

            lines = [
                f"# {os.path.basename(source_path_for_output)}",
                "",
                f"- source_cab: {os.path.abspath(source_path_for_output)}",
                f"- topic_count: {len(topics)}",
                f"- generated_at: {datetime.now().isoformat(timespec='seconds')}",
                "",
                "## 目录",
                "",
            ]
            for idx, topic in enumerate(topics, 1):
                title = self._normalize_md_line(
                    topic.get("title", "") or topic.get("id", "")
                )
                lines.append(f"{idx}. {title or 'Untitled'}")
            lines.append("")

            rendered_count = 0
            for idx, topic in enumerate(topics, 1):
                title = self._normalize_md_line(
                    topic.get("title", "") or topic.get("id", "")
                )
                html_file = topic.get("file", "")
                if not html_file or not os.path.exists(html_file):
                    continue
                body_md = self._render_html_to_markdown(html_file)
                if not body_md:
                    continue
                lines.extend([f"## {idx}. {title or 'Untitled'}", "", body_md, ""])
                rendered_count += 1

            if rendered_count <= 0:
                raise RuntimeError(
                    f"CAB topics parsed but no readable content rendered: {cab_path}"
                )

            with open(md_path, "w", encoding="utf-8") as f:
                f.write("\n".join(lines).rstrip() + "\n")

            self.generated_markdown_outputs.append(md_path)
            self._append_mshelp_record(source_path_for_output, md_path, rendered_count)
            return md_path, rendered_count
        finally:
            try:
                if os.path.exists(temp_root):
                    shutil.rmtree(temp_root, ignore_errors=True)
            except Exception:
                pass

    def _append_mshelp_record(self, source_cab_path, markdown_path, topic_count):
        src_abs = os.path.abspath(source_cab_path)
        md_abs = os.path.abspath(markdown_path)
        folder_name = str(self.config.get("mshelpviewer_folder_name", "MSHelpViewer"))
        folder_name = folder_name.strip() or "MSHelpViewer"
        folder_name_lower = folder_name.lower()

        mshelp_dir = ""
        try:
            p = Path(src_abs)
            for parent in [p.parent, *p.parents]:
                if parent.name.lower() == folder_name_lower:
                    mshelp_dir = str(parent)
                    break
        except Exception:
            mshelp_dir = ""

        try:
            src_rel = os.path.relpath(src_abs, self._get_source_root_for_path(src_abs))
        except Exception:
            src_rel = src_abs

        self.mshelp_records.append(
            {
                "source_cab": src_abs,
                "source_cab_relpath": src_rel,
                "mshelpviewer_dir": mshelp_dir,
                "markdown_path": md_abs,
                "topic_count": int(topic_count or 0),
                "status": "success",
            }
        )

    def _find_mshelpviewer_dirs(self, root_dir):
        result = []
        if not root_dir or not os.path.isdir(root_dir):
            return result
        folder_name = str(self.config.get("mshelpviewer_folder_name", "MSHelpViewer"))
        folder_name = folder_name.strip() or "MSHelpViewer"
        folder_name_lower = folder_name.lower()

        for current_root, dirs, _ in os.walk(root_dir):
            if os.path.basename(current_root).lower() == folder_name_lower:
                result.append(current_root)
                dirs[:] = []
                continue
            matches = [d for d in dirs if d.lower() == folder_name_lower]
            for d in matches:
                result.append(os.path.join(current_root, d))
        # de-dup keep stable order
        seen = set()
        unique = []
        for d in result:
            ad = os.path.abspath(d)
            key = ad.lower() if is_win() else ad
            if key in seen:
                continue
            seen.add(key)
            unique.append(ad)
        return unique

    def _scan_mshelp_cab_candidates(self):
        source_roots = self._get_source_roots()
        dirs = []
        for source_root in source_roots:
            dirs.extend(self._find_mshelpviewer_dirs(source_root))
        cab_exts = tuple(
            e.lower()
            for e in self.config.get("allowed_extensions", {}).get("cab", [".cab"])
        )
        if not cab_exts:
            cab_exts = (".cab",)

        files = []
        seen = set()
        for d in dirs:
            for cab_path in self._find_files_recursive(d, cab_exts):
                abs_cab = os.path.abspath(cab_path)
                key = abs_cab.lower() if is_win() else abs_cab
                if key in seen:
                    continue
                seen.add(key)
                files.append(abs_cab)
        files.sort()
        return dirs, files

    def _write_mshelp_index_files(self):
        if not self.mshelp_records:
            return []
        target_root = self.config.get("target_folder", "")
        if not target_root:
            return []

        out_dir = os.path.join(target_root, "_AI", "MSHelp")
        os.makedirs(out_dir, exist_ok=True)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        json_path = os.path.join(out_dir, f"MSHelp_Index_{ts}.json")
        csv_path = os.path.join(out_dir, f"MSHelp_Index_{ts}.csv")

        records = sorted(
            self.mshelp_records,
            key=lambda x: (x.get("mshelpviewer_dir", ""), x.get("source_cab", "")),
        )
        payload = {
            "version": 1,
            "generated_at": datetime.now().isoformat(timespec="seconds"),
            "record_count": len(records),
            "records": records,
        }
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)

        fields = [
            "source_cab",
            "source_cab_relpath",
            "mshelpviewer_dir",
            "markdown_path",
            "topic_count",
            "status",
        ]
        with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=fields)
            writer.writeheader()
            writer.writerows(records)

        outputs = [json_path, csv_path]
        self.generated_mshelp_outputs.extend(outputs)
        logging.info(f"MSHelp index generated: {json_path}")
        logging.info(f"MSHelp index generated: {csv_path}")
        return outputs

    @staticmethod
    def _wrap_plain_text_for_pdf(text, width=100):
        words = str(text or "").split()
        if not words:
            return [""]
        lines = []
        current = []
        current_len = 0
        for w in words:
            if current_len + len(w) + (1 if current else 0) > width:
                lines.append(" ".join(current))
                current = [w]
                current_len = len(w)
            else:
                current.append(w)
                current_len += len(w) + (1 if current_len > 0 else 0)
        if current:
            lines.append(" ".join(current))
        return lines or [""]

    def _export_markdown_to_docx(self, md_path, out_docx):
        if not HAS_PYDOCX:
            raise RuntimeError("python-docx is not installed.")
        with open(md_path, "r", encoding="utf-8", errors="ignore") as f:
            lines = f.read().splitlines()

        doc = Document()
        in_code = False
        for raw in lines:
            line = str(raw or "")
            if line.strip().startswith("```"):
                in_code = not in_code
                continue
            if in_code:
                doc.add_paragraph(line)
                continue
            s = line.strip()
            if not s:
                doc.add_paragraph("")
                continue
            if s.startswith("### "):
                doc.add_heading(s[4:], level=3)
            elif s.startswith("## "):
                doc.add_heading(s[3:], level=2)
            elif s.startswith("# "):
                doc.add_heading(s[2:], level=1)
            elif s.startswith("- "):
                doc.add_paragraph(s[2:], style="List Bullet")
            elif re.match(r"^\d+\.\s+", s):
                text = re.sub(r"^\d+\.\s+", "", s)
                doc.add_paragraph(text, style="List Number")
            else:
                doc.add_paragraph(s)
        doc.save(out_docx)

    def _export_markdown_to_pdf(self, md_path, out_pdf):
        if not HAS_REPORTLAB:
            raise RuntimeError("reportlab is not installed.")
        with open(md_path, "r", encoding="utf-8", errors="ignore") as f:
            lines = f.read().splitlines()

        c = canvas.Canvas(out_pdf, pagesize=A4)
        page_w, page_h = A4
        x = 36
        y = page_h - 36
        line_h = 12

        for raw in lines:
            text = str(raw or "")
            wrapped = self._wrap_plain_text_for_pdf(text, width=100)
            for w in wrapped:
                if y <= 36:
                    c.showPage()
                    y = page_h - 36
                c.drawString(x, y, w)
                y -= line_h
        c.save()

    def _merge_mshelp_markdowns(self):
        if not self.mshelp_records:
            return []
        if not bool(self.config.get("enable_mshelp_merge_output", True)):
            return []

        merge_start = time.perf_counter()
        target_root = self.config.get("target_folder", "")
        if not target_root:
            return []
        out_dir = os.path.join(target_root, "_AI", "MSHelp", "Merged")
        os.makedirs(out_dir, exist_ok=True)

        # Reuse unified max_merge_size_mb (same parameter as regular PDF merge)
        try:
            max_size_mb = int(self.config.get("max_merge_size_mb", 80) or 80)
        except Exception:
            max_size_mb = 80
        max_size_bytes = max(1, max_size_mb) * 1024 * 1024

        valid = []
        for rec in self.mshelp_records:
            mdp = rec.get("markdown_path", "")
            if not mdp or not os.path.exists(mdp):
                continue
            try:
                with open(mdp, "r", encoding="utf-8", errors="ignore") as f:
                    content = f.read()
            except Exception:
                continue
            item = dict(rec)
            item["_content"] = content
            item["_bytes"] = len(content.encode("utf-8"))
            valid.append(item)

        if not valid:
            return []

        chunks = []
        current = []
        current_bytes = 0
        for rec in valid:
            rec_bytes = int(rec.get("_bytes", 0) or 0)
            if current and (current_bytes + rec_bytes > max_size_bytes):
                chunks.append(current)
                current = []
                current_bytes = 0
            current.append(rec)
            current_bytes += rec_bytes
        if current:
            chunks.append(current)

        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        outputs = []
        export_docx = bool(self.config.get("enable_mshelp_output_docx", False))
        export_pdf = bool(self.config.get("enable_mshelp_output_pdf", False))

        for idx, chunk in enumerate(chunks, 1):
            md_path = os.path.join(out_dir, f"MSHelp_API_Merged_{ts}_{idx:03d}.md")
            lines = [
                f"# MSHelp API Merged Package {idx}/{len(chunks)}",
                "",
                f"- generated_at: {datetime.now().isoformat(timespec='seconds')}",
                f"- source_root: {self.config.get('source_folder', '')}",
                f"- document_count: {len(chunk)}",
                "",
                "## Source Map",
                "",
                "| No. | CAB Source | MSHelpViewer Dir | Markdown Path | Topic Count |",
                "| --- | --- | --- | --- | ---: |",
            ]
            for j, rec in enumerate(chunk, 1):
                lines.append(
                    "| {0} | {1} | {2} | {3} | {4} |".format(
                        j,
                        str(rec.get("source_cab", "")).replace("|", "\\|"),
                        str(rec.get("mshelpviewer_dir", "")).replace("|", "\\|"),
                        str(rec.get("markdown_path", "")).replace("|", "\\|"),
                        int(rec.get("topic_count", 0) or 0),
                    )
                )
            lines.append("")
            lines.append("## Documents")
            lines.append("")

            for j, rec in enumerate(chunk, 1):
                title = os.path.basename(str(rec.get("source_cab", "") or f"doc_{j}"))
                lines.extend(
                    [
                        f"### [{j}] {title}",
                        "",
                        f"- source_cab: {rec.get('source_cab', '')}",
                        f"- source_markdown: {rec.get('markdown_path', '')}",
                        "",
                        "---",
                        "",
                        rec.get("_content", ""),
                        "",
                    ]
                )

            with open(md_path, "w", encoding="utf-8") as f:
                f.write("\n".join(lines).rstrip() + "\n")
            outputs.append(md_path)
            self.generated_mshelp_outputs.append(md_path)
            logging.info(f"MSHelp merged markdown generated: {md_path}")

            if export_docx:
                docx_path = os.path.splitext(md_path)[0] + ".docx"
                try:
                    self._export_markdown_to_docx(md_path, docx_path)
                    outputs.append(docx_path)
                    self.generated_mshelp_outputs.append(docx_path)
                    logging.info(f"MSHelp merged DOCX generated: {docx_path}")
                except Exception as e:
                    logging.warning(f"MSHelp merged DOCX skipped: {e}")

            if export_pdf:
                pdf_path = os.path.splitext(md_path)[0] + ".pdf"
                try:
                    self._export_markdown_to_pdf(md_path, pdf_path)
                    outputs.append(pdf_path)
                    self.generated_mshelp_outputs.append(pdf_path)
                    logging.info(f"MSHelp merged PDF generated: {pdf_path}")
                except Exception as e:
                    logging.warning(f"MSHelp merged PDF skipped: {e}")

        self._add_perf_seconds(
            "mshelp_merge_seconds", time.perf_counter() - merge_start
        )
        return outputs

    def process_single_file(
        self, file_path, target_path_initial, ext, progress_str, is_retry=False
    ):
        if os.path.getsize(file_path) == 0:
            self.stats["skipped"] += 1
            logging.warning(f"skip empty file: {file_path}")
            return "skip_empty", target_path_initial

        is_word = ext in self.config["allowed_extensions"].get("word", [])
        is_excel = ext in self.config["allowed_extensions"].get("excel", [])
        is_ppt = ext in self.config["allowed_extensions"].get("powerpoint", [])
        is_pdf = ext == ".pdf"
        is_cab = ext in self.config["allowed_extensions"].get("cab", [])
        is_office = is_word or is_excel or is_ppt

        filename = os.path.basename(file_path)

        # Step 1: filename keyword pre-hit
        is_filename_match = False
        if self.content_strategy != STRATEGY_STANDARD:
            for kw in self.price_keywords:
                if kw in filename:
                    is_filename_match = True
                    break

        if self.content_strategy == STRATEGY_PRICE_ONLY and not is_filename_match:
            if is_word or is_ppt:
                self.stats["skipped"] += 1
                return "skip_strategy", target_path_initial

        sandbox_pdf = os.path.join(self.temp_sandbox, f"{uuid.uuid4()}.pdf")

        use_sandbox = self.config.get("enable_sandbox", True)
        working_src = file_path
        sandbox_src_path = None

        final_target_path = target_path_initial

        result_context = {
            "is_price": is_filename_match,
            "scan_aborted": False,
            "skip_scan": is_filename_match,
        }
        output_plan = self.compute_convert_output_plan(self.run_mode, self.config)

        if is_filename_match:
            final_target_path = self.get_target_path(
                file_path, ext, prefix_override="Price_"
            )

        base_timeout = self.config.get("timeout_seconds", 60)
        ppt_timeout = self.config.get("ppt_timeout_seconds", base_timeout)
        current_timeout = ppt_timeout if is_ppt else base_timeout

        base_wait = self.config.get("pdf_wait_seconds", 15)
        ppt_wait = self.config.get("ppt_pdf_wait_seconds", base_wait)
        current_pdf_wait = ppt_wait if is_ppt else base_wait

        try:
            convert_core_start = time.perf_counter()
            if use_sandbox:
                sandbox_src_path = os.path.join(
                    self.temp_sandbox, f"{uuid.uuid4()}{ext}"
                )
                shutil.copy2(file_path, sandbox_src_path)
                self._unblock_file(sandbox_src_path)
                working_src = sandbox_src_path

            if is_cab:
                md_path, rendered_count = self._convert_cab_to_markdown(
                    working_src, file_path
                )
                self._add_perf_seconds(
                    "convert_core_seconds", time.perf_counter() - convert_core_start
                )
                return f"success_cab_md[{rendered_count}]", md_path

            if is_pdf:
                if not is_filename_match and self.content_strategy != STRATEGY_STANDARD:
                    has_kw = self.scan_pdf_content(working_src)
                    if has_kw:
                        result_context["is_price"] = True
                    elif self.content_strategy == STRATEGY_PRICE_ONLY:
                        self.stats["skipped"] += 1
                        return "skip_content", target_path_initial

                self.copy_pdf_direct(working_src, sandbox_pdf)
                self._add_perf_seconds(
                    "convert_core_seconds", time.perf_counter() - convert_core_start
                )

            else:
                convert_thread = threading.Thread(
                    target=self.convert_logic_in_thread,
                    args=(working_src, sandbox_pdf, ext, result_context),
                    daemon=True,
                )
                convert_thread.start()

                wait_start = time.time()
                while convert_thread.is_alive():
                    elapsed = time.time() - wait_start
                    if elapsed > current_timeout:
                        break
                    print(
                        f"{progress_str} converting: {filename} ({elapsed:.1f}s)    ",
                        end="",
                        flush=True,
                    )
                    time.sleep(0.1)

                convert_thread.join(timeout=0.1)

                if convert_thread.is_alive():
                    self._add_perf_seconds(
                        "convert_core_seconds",
                        time.perf_counter() - convert_core_start,
                    )
                    self.stats["timeout"] += 1
                    logging.error(f"timeout skip (>{current_timeout}s)")
                    if is_word:
                        self._kill_current_app("word", force=True)
                    elif is_excel:
                        self._kill_current_app("excel", force=True)
                    elif is_ppt:
                        self._kill_current_app("ppt", force=True)
                    raise Exception("timeout")
                self._add_perf_seconds(
                    "convert_core_seconds", time.perf_counter() - convert_core_start
                )

            if result_context["scan_aborted"]:
                self.stats["skipped"] += 1
                return "skip_content", target_path_initial

            if result_context["is_price"]:
                final_target_path = self.get_target_path(
                    file_path, ext, prefix_override="Price_"
                )

            wait_pdf_start = time.perf_counter()
            while time.perf_counter() - wait_pdf_start < current_pdf_wait:
                if os.path.exists(sandbox_pdf):
                    time.sleep(0.5)
                    self._add_perf_seconds(
                        "pdf_wait_seconds", time.perf_counter() - wait_pdf_start
                    )
                    final_path_res = ""
                    md_path_res = ""
                    if output_plan.get("need_final_pdf"):
                        result_status, final_path_res = self.handle_file_conflict(
                            sandbox_pdf, final_target_path
                        )
                        self.generated_pdfs.append(final_path_res)
                    else:
                        result_status = (
                            "success_no_output"
                            if not output_plan.get("need_markdown")
                            else "success_md_only"
                        )

                    if output_plan.get("need_markdown"):
                        markdown_start = time.perf_counter()
                        if final_path_res and os.path.exists(final_path_res):
                            md_path_res = (
                                self._export_pdf_markdown(final_path_res) or ""
                            )
                        else:
                            md_path_res = (
                                self._export_pdf_markdown(
                                    sandbox_pdf, source_path_hint=file_path
                                )
                                or ""
                            )
                        self._add_perf_seconds(
                            "markdown_seconds",
                            time.perf_counter() - markdown_start,
                        )

                    tag_info = ""
                    if is_filename_match:
                        tag_info = " [name_hit]"
                    elif result_context["is_price"]:
                        tag_info = " [content_hit]"
                    final_output_path = final_path_res or md_path_res
                    return f"{result_status}{tag_info}", final_output_path
                time.sleep(0.5)

            self._add_perf_seconds(
                "pdf_wait_seconds", time.perf_counter() - wait_pdf_start
            )
            raise Exception(
                f"conversion command sent but PDF not generated ({current_pdf_wait}s)"
            )

        finally:
            if is_office:
                self._on_office_file_processed(ext)
            try:
                if sandbox_src_path and os.path.exists(sandbox_src_path):
                    os.remove(sandbox_src_path)
                if os.path.exists(sandbox_pdf):
                    os.remove(sandbox_pdf)
            except Exception:
                pass

    # =============== Batch processing / retry ===============

    def get_progress_prefix(self, current, total):
        width = len(str(total)) if total > 0 else 1
        percent = current / total if total else 0
        bar_len = 20
        filled = int(bar_len * percent)
        bar = "#" * filled + "-" * (bar_len - filled)
        return f"[{int(percent * 100):>3}%]{bar} [{str(current).rjust(width)}/{total}]"

    def _emit_file_plan(self, file_list):
        cb = getattr(self, "file_plan_callback", None)
        if not callable(cb):
            return
        try:
            cb(list(file_list or []))
        except Exception as e:
            logging.warning(f"file_plan_callback failed: {e}")

    def _emit_file_done(self, record):
        cb = getattr(self, "file_done_callback", None)
        if not callable(cb):
            return
        if not isinstance(record, dict):
            return
        try:
            cb(dict(record))
        except Exception as e:
            logging.warning(f"file_done_callback failed: {e}")

    def run_batch(self, file_list, is_retry=False, source_alias_map=None):
        total = len(file_list)
        results = []
        source_alias_map = source_alias_map or {}

        for i, fpath in enumerate(file_list, 1):
            if not self.is_running:
                break

            logical_source = source_alias_map.get(fpath, fpath)
            fname = os.path.basename(fpath)
            ext = os.path.splitext(fpath)[1].lower()
            target_path_initial = self.get_target_path(logical_source, ext)

            progress_prefix = self.get_progress_prefix(i, total)
            if self.progress_callback:
                self.progress_callback(i, total)

            label = "[重试]" if is_retry else "正在处理"
            print(
                f"\r{progress_prefix} {label}: {fname}" + " " * 20, end="", flush=True
            )

            start = time.time()
            try:
                status, final_path = self.process_single_file(
                    fpath, target_path_initial, ext, progress_prefix, is_retry
                )
                elapsed = time.time() - start

                if status.startswith("skip"):
                    print(
                        f"\r{progress_prefix} {status}: {fname} (耗时: {elapsed:.2f}s)    "
                    )
                    logging.info(f"{status}: {logical_source}")
                    record = {
                        "source_path": os.path.abspath(logical_source),
                        "status": "skipped",
                        "detail": status,
                        "final_path": final_path,
                        "elapsed": elapsed,
                    }
                    results.append(record)
                    self._emit_file_done(record)
                else:
                    self.stats["success"] += 1
                    print(
                        f"\r{progress_prefix} {status}: {fname} (耗时: {elapsed:.2f}s)    "
                    )
                    logging.info(f"{status}: {logical_source} -> {final_path}")
                    is_pdf_output = str(final_path).lower().endswith(".pdf")
                    result_status = "success" if is_pdf_output else "success_non_pdf"
                    if is_pdf_output:
                        self._append_conversion_index_record(
                            logical_source, final_path, status
                        )
                    record = {
                        "source_path": os.path.abspath(logical_source),
                        "status": result_status,
                        "detail": status,
                        "final_path": final_path,
                        "elapsed": elapsed,
                    }
                    results.append(record)
                    self._emit_file_done(record)

            except Exception as e:
                elapsed = time.time() - start
                err_msg = str(e)

                # 使用新的错误分类和记录机制
                error_detail = self.record_detailed_error(
                    logical_source,
                    e,
                    context={
                        "run_mode": self.get_readable_run_mode(),
                        "engine": self.get_readable_engine_type(),
                        "elapsed": elapsed,
                    },
                )

                if (
                    error_detail["error_type"] == ConversionErrorType.TIMEOUT
                    or "超时" in err_msg
                ):
                    self.stats["timeout"] += 1
                    print(
                        f"\r{progress_prefix} 超时: {fname} (耗时: {elapsed:.2f}s)    "
                    )
                    record = {
                        "source_path": os.path.abspath(logical_source),
                        "status": "timeout",
                        "detail": err_msg,
                        "final_path": "",
                        "elapsed": elapsed,
                        "error": err_msg,
                        "error_type": error_detail["error_type"],
                        "error_category": error_detail["error_category"],
                        "suggestion": error_detail["suggestion"],
                    }
                    results.append(record)
                    self._emit_file_done(record)
                else:
                    self.stats["failed"] += 1
                    # 根据错误类型显示不同提示
                    error_type_display = error_detail["error_type"]
                    if error_detail["requires_manual_action"]:
                        error_type_display += " [需人工处理]"
                    elif error_detail["is_retryable"]:
                        error_type_display += " [可重试]"

                    print(
                        f"\r{progress_prefix} 失败({error_detail['error_type']}): {fname}    "
                    )
                    record = {
                        "source_path": os.path.abspath(logical_source),
                        "status": "failed",
                        "detail": err_msg,
                        "final_path": "",
                        "elapsed": elapsed,
                        "error": err_msg,
                        "error_type": error_detail["error_type"],
                        "error_category": error_detail["error_category"],
                        "suggestion": error_detail["suggestion"],
                        "is_retryable": error_detail["is_retryable"],
                        "requires_manual_action": error_detail[
                            "requires_manual_action"
                        ],
                    }
                    results.append(record)
                    self._emit_file_done(record)

                logging.error(
                    f"failed: {logical_source} | reason: {e} | type: {error_detail['error_type']}"
                )

                if not is_retry:
                    self.quarantine_failed_file(fpath)
                    self.error_records.append(logical_source)

        return results

    def ask_retry_failed_files(self, failed_count, timeout=20):
        print("\n" + "=" * 60)
        print(f"[WARN] {failed_count} files failed (including timeout).")
        if self.error_records:
            print("Failed examples (up to 10):")
            for p in self.error_records[:10]:
                print("  -", p)
        print("-" * 60)
        print("Retry failed files?")
        print("  Enter Y + Enter -> retry")
        print("  Enter N + Enter -> no retry")
        print(f"  If no input in {timeout}s, default is no retry.")
        print("=" * 60)

        if not HAS_MSVCRT:
            ans = input("Input [Y/N] and press Enter: ").strip().lower()
            return ans == "y"

        buf = ""
        start = time.time()
        last_shown = None

        while True:
            elapsed = time.time() - start
            remain = int(timeout - elapsed)
            if remain < 0:
                print("\n[INFO] timeout reached, default no retry.")
                return False

            if last_shown != remain:
                print(
                    f"\rInput [Y/N] within {remain:2d}s: {buf}",
                    end="",
                    flush=True,
                )
                last_shown = remain

            if msvcrt.kbhit():
                ch = msvcrt.getwch()
                if ch in ("\r", "\n"):
                    ans = buf.strip().lower()
                    print()
                    if ans == "y":
                        print("[SELECT] retry failed files.\n")
                        return True
                    print("[SELECT] do not retry failed files.\n")
                    return False
                if ch == "\b":
                    buf = buf[:-1]
                else:
                    buf += ch
            time.sleep(0.1)

    def _create_index_doc_and_convert(self, word_app, file_list, title):
        """"""
        try:
            doc = word_app.Documents.Add()
            word_app.Visible = False  # Ensure invisible

            # Page setup: A4 (595.35 x 841.995 pt), margins 72pt (1 inch)
            try:
                doc.PageSetup.PaperSize = 7  # wdPaperA4
                doc.PageSetup.TopMargin = 72
                doc.PageSetup.BottomMargin = 72
                doc.PageSetup.LeftMargin = 72
                doc.PageSetup.RightMargin = 72
            except Exception:
                pass

            selection = word_app.Selection

            lines_per_page = 32

            def write_header():
                selection.ParagraphFormat.Alignment = 1
                selection.Font.Name = "Microsoft YaHei"
                selection.Font.Size = 16
                selection.Font.Bold = True
                selection.ParagraphFormat.LineSpacingRule = 0
                selection.TypeText(title + "\n")

                selection.ParagraphFormat.Alignment = 0
                selection.Font.Size = 10.5
                selection.Font.Bold = False
                selection.TypeText("\n")

                selection.ParagraphFormat.LineSpacingRule = 4
                selection.ParagraphFormat.LineSpacing = 20

            write_header()

            # Coordinate estimation parameters for index link placement.
            # A4高度 ~842pt. TopMargin 72.
            # Title line (16pt) + spacing; this is an approximation.
            # In Word, 16pt lines are typically around 20-22pt high.
            # Simplified model:
            # TopMargin: 72
            # Title line: ~30 (includes spacing)
            # Empty Line: ~15
            # List Start Y (from bottom): PageHeight - (72 + 30 + 15) = 842 - 117 = 725
            # Keep this conservative and stable across fonts/renderers.
            # Assume first line baseline starts around Y=700.
            # Fine-tune during merge if needed.

            for i, fname in enumerate(file_list, 1):
                if i > 1 and (i - 1) % lines_per_page == 0:
                    selection.InsertBreak(7)
                    write_header()
                # Truncate long filenames to reduce overflow in A4 index layout.
                # Conservative cutoff at 45 characters.
                if len(fname) > 45:
                    fname = fname[:42] + "..."
                selection.TypeText(f"{i}. {fname}\n")

            temp_pdf = os.path.join(self.temp_sandbox, f"index_{uuid.uuid4().hex}.pdf")

            doc.ExportAsFixedFormat(
                OutputFileName=temp_pdf,
                ExportFormat=17,  # wdExportFormatPDF
                OpenAfterExport=False,
                OptimizeFor=0,
                CreateBookmarks=1,
                DocStructureTags=True,
            )
            doc.Close(SaveChanges=0)
            return temp_pdf
        except Exception as e:
            logging.error(f"failed to generate merge index page: {e}")
            return None

    @staticmethod
    def _format_merge_filename(pattern, category="All", idx=1, now=None):
        """
        Format merge output filename from pattern.
        Placeholders: {category}, {timestamp}, {date}, {time}, {idx}
        """
        if now is None:
            now = datetime.now()
        timestamp = now.strftime("%Y%m%d_%H%M%S")
        date_part = now.strftime("%Y%m%d")
        time_part = now.strftime("%H%M%S")
        name = (
            pattern.replace("{category}", str(category))
            .replace("{timestamp}", timestamp)
            .replace("{date}", date_part)
            .replace("{time}", time_part)
            .replace("{idx}", str(idx))
        )
        # Sanitize: only allow safe filename chars (alphanumeric, dash, underscore, dot)
        safe = re.sub(r"[^\w\-.]", "_", name)
        safe = re.sub(r"_+", "_", safe).strip("_")
        if not safe:
            safe = f"Merged_{category}_{timestamp}_{idx}"
        if not safe.lower().endswith(".pdf"):
            safe = f"{safe}.pdf"
        return safe

    def _get_merge_tasks(self):
        """Build merge task groups from scanned PDFs."""
        scan_source_type = "target"
        if self.run_mode == MODE_MERGE_ONLY:
            scan_source_type = self.config.get("merge_source", "source")

        if scan_source_type == "source":
            scan_roots = self._get_source_roots()
            print(f"  [merge_only/source] scanning {len(scan_roots)} source folder(s)")
        else:
            scan_roots = [self.config["target_folder"]]
            print(f"  [merge scan] scanning target: {scan_roots[0]}")

        all_pdfs = []
        exclude_abs_paths = set(
            map(os.path.abspath, [self.failed_dir, self.merge_output_dir])
        )
        if scan_source_type == "source":
            exclude_abs_paths.add(os.path.abspath(self.config["target_folder"]))

        for scan_folder in scan_roots:
            if not scan_folder or not os.path.isdir(scan_folder):
                continue
            for root, dirs, files in os.walk(scan_folder):
                dirs[:] = [
                    d
                    for d in dirs
                    if os.path.abspath(os.path.join(root, d)) not in exclude_abs_paths
                ]

                if os.path.abspath(root) in exclude_abs_paths:
                    continue
                for f in files:
                    if f.lower().endswith(".pdf"):
                        all_pdfs.append(os.path.join(root, f))

        if not all_pdfs:
            print("[INFO] no PDF files found for merge.")
            return []

        all_pdfs.sort()

        merge_tasks = []  # [(output_name, [pdf_paths])]
        now = datetime.now()
        pattern = (
            self.config.get("merge_filename_pattern")
            or "Merged_{category}_{timestamp}_{idx}"
        ).strip()
        if not pattern:
            pattern = "Merged_{category}_{timestamp}_{idx}"

        if self.merge_mode == MERGE_MODE_ALL_IN_ONE:
            output_name = self._format_merge_filename(
                pattern, category="All", idx=1, now=now
            )
            merge_tasks.append((output_name, all_pdfs))
        else:
            categories = {
                "Price Documents": "Price_",
                "Word Documents": "Word_",
                "Excel Sheets": "Excel_",
                "PPT Slides": "PPT_",
                "Original PDF": "PDF_",
            }
            max_size_bytes = self.config.get("max_merge_size_mb", 80) * 1024 * 1024

            for cat_name, prefix in categories.items():
                current_cat_files = [
                    p for p in all_pdfs if os.path.basename(p).startswith(prefix)
                ]
                if not current_cat_files:
                    continue
                current_cat_files.sort()

                groups = []
                current_group = []
                current_size = 0
                for pdf_path in current_cat_files:
                    try:
                        f_size = os.path.getsize(pdf_path)
                    except Exception:
                        continue

                    if f_size > max_size_bytes:
                        if current_group:
                            groups.append(current_group)
                            current_group = []
                            current_size = 0
                        groups.append([pdf_path])
                        continue

                    if current_size + f_size > max_size_bytes:
                        groups.append(current_group)
                        current_group = [pdf_path]
                        current_size = f_size
                    else:
                        current_group.append(pdf_path)
                        current_size += f_size
                if current_group:
                    groups.append(current_group)

                cat_label = prefix.rstrip("_")
                for idx, group in enumerate(groups, 1):
                    output_filename = self._format_merge_filename(
                        pattern, category=cat_label, idx=idx, now=now
                    )
                    merge_tasks.append((output_filename, group))

        return merge_tasks

    @staticmethod
    def _compute_md5(path, block_size=1024 * 1024):
        h = hashlib.md5()
        with open(path, "rb") as f:
            while True:
                chunk = f.read(block_size)
                if not chunk:
                    break
                h.update(chunk)
        return h.hexdigest()

    @staticmethod
    def _mask_md5(md5_value):
        if not md5_value or len(md5_value) < 12:
            return md5_value
        return f"{md5_value[:8]}...{md5_value[-4:]}"

    def _build_short_id(self, md5_value, taken_ids):
        length = DEFAULT_SHORT_ID_LEN
        while length <= len(md5_value):
            candidate = md5_value[:length].upper()
            if candidate not in taken_ids:
                taken_ids.add(candidate)
                return candidate
            length += 2
        candidate = md5_value.upper()
        taken_ids.add(candidate)
        return candidate

    def _write_merge_map(self, output_path, records):
        if not records:
            return None, None

        base_no_ext, _ = os.path.splitext(output_path)
        csv_path = f"{base_no_ext}.map.csv"
        json_path = f"{base_no_ext}.map.json"
        fields = [
            "merge_batch_id",
            "merged_pdf_name",
            "merged_pdf_path",
            "source_index",
            "source_filename",
            "source_abspath",
            "source_relpath",
            "source_md5",
            "source_short_id",
            "start_page_1based",
            "end_page_1based",
            "page_count",
            "bookmark_title",
        ]

        with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=fields)
            writer.writeheader()
            writer.writerows(records)

        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(
                {
                    "version": 1,
                    "record_count": len(records),
                    "records": records,
                },
                f,
                ensure_ascii=False,
                indent=2,
            )
        return csv_path, json_path

    def merge_pdfs(self):
        if not self.config.get("enable_merge", True):
            return []
        if not HAS_PYPDF:
            print("\n[INFO] pypdf not found. Skip merge step. Run: pip install pypdf")
            logging.warning("pypdf not found. Skip merge step.")
            return []

        # Try importing pypdf generic classes used for manual link annotations.
        try:
            from pypdf.generic import (
                DictionaryObject,
                NumberObject,
                FloatObject,
                NameObject,
                TextStringObject,
                ArrayObject,
                RectangleObject,
            )

            HAS_PYPDF_GENERIC = True
        except ImportError:
            HAS_PYPDF_GENERIC = False

        print("\n" + "=" * 60)
        print("  Start PDF merging ...")
        print(f"  Merge mode: {self.get_readable_merge_mode()} ({self.merge_mode})")
        print(f"  Merge output dir: {self.merge_output_dir}")
        if self.enable_merge_index:
            print("  [Option] Enable index page generation (with clickable links)")
        if self.enable_merge_excel:
            print("  [Option] Enable Excel list output (one row per source file)")
        print("=" * 60)

        # Prepare Excel merge list output.
        wb_merge = None
        ws_merge = None
        merge_excel_path = None
        if self.enable_merge_excel:
            if not HAS_OPENPYXL:
                print("  [WARN] openpyxl not found. Excel merge list is disabled.")
            else:
                timestamp_excel = datetime.now().strftime("%Y%m%d_%H%M%S")
                merge_excel_path = os.path.join(
                    self.merge_output_dir, f"Merge_List_{timestamp_excel}.xlsx"
                )
                wb_merge = Workbook()
                ws_merge = wb_merge.active
                ws_merge.title = "MergeList"
                ws_merge.append(["Merged File", "Source Files"])
                # Set width
                ws_merge.column_dimensions["A"].width = 40
                ws_merge.column_dimensions["B"].width = 60

        # Build merge task groups:
        # [(output_filename, [pdf_path1, pdf_path2, ...]), ...]
        merge_tasks = self._get_merge_tasks()

        total_tasks = len(merge_tasks)
        print(f"  Total merge tasks: {total_tasks}")

        # Prepare Word app for index-page generation when enabled.
        word_app = None
        if self.enable_merge_index and total_tasks > 0:
            try:
                pythoncom.CoInitialize()
                word_app = win32com.client.Dispatch("Word.Application")
                word_app.Visible = False
                word_app.DisplayAlerts = 0
            except Exception as e:
                logging.error(f"Failed to start Word. Cannot generate index page: {e}")
                word_app = None

        generated_outputs = []
        generated_map_outputs = []
        merge_index_records = []
        merge_batch_id = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.merge_excel_path = None

        for idx, (output_filename, group) in enumerate(merge_tasks, 1):
            print(f"  Processing [{idx}/{total_tasks}]: {output_filename}")

            # Excel records (one row per merged source file).
            if ws_merge:
                for sub_file in group:
                    ws_merge.append([output_filename, os.path.basename(sub_file)])

            output_path = os.path.join(self.merge_output_dir, output_filename)

            try:
                merger = PdfWriter()
                current_page_index = 0
                map_records = []
                short_id_taken = set()

                # 1) Generate and append index pages.
                index_pdf_path = None
                file_basenames = [os.path.basename(p) for p in group]

                # Track each file's start page in merged output.
                # Initial offset equals index page count.
                file_start_pages = []

                if self.enable_merge_index and word_app:
                    index_pdf_path = self._create_index_doc_and_convert(
                        word_app, file_basenames, "File Index"
                    )

                if index_pdf_path and os.path.exists(index_pdf_path):
                    idx_reader = PdfReader(index_pdf_path)
                    index_page_count = len(idx_reader.pages)

                    # Append index pages first.
                    for p in idx_reader.pages:
                        merger.add_page(p)

                    current_page_index += index_page_count
                else:
                    index_page_count = 0

                # 2) Append content files and record page ranges.
                for source_idx, pdf_file in enumerate(group, 1):
                    fname = os.path.basename(pdf_file)

                    # Record source start page (absolute page number in merged PDF).
                    file_start_pages.append(current_page_index)

                    try:
                        reader = PdfReader(pdf_file)
                        source_page_count = len(reader.pages)
                        source_md5 = self._compute_md5(pdf_file)
                        source_short_id = self._build_short_id(
                            source_md5, short_id_taken
                        )
                        bookmark_title = fname
                        if self.config.get("bookmark_with_short_id", True):
                            bookmark_title = f"[ID:{source_short_id}] {fname}"

                        # Add bookmark.
                        merger.add_outline_item(bookmark_title, current_page_index)

                        start_page_1based = current_page_index + 1
                        end_page_1based = current_page_index + source_page_count

                        for page in reader.pages:
                            merger.add_page(page)
                        current_page_index += source_page_count

                        source_abs_path = os.path.abspath(pdf_file)
                        source_rel_path = source_abs_path
                        try:
                            source_rel_path = os.path.relpath(
                                source_abs_path,
                                self._get_source_root_for_path(source_abs_path),
                            )
                        except Exception:
                            pass

                        map_records.append(
                            {
                                "merge_batch_id": merge_batch_id,
                                "merged_pdf_name": os.path.basename(output_path),
                                "merged_pdf_path": output_path,
                                "source_index": source_idx,
                                "source_filename": fname,
                                "source_abspath": source_abs_path,
                                "source_relpath": source_rel_path,
                                "source_md5": source_md5,
                                "source_short_id": source_short_id,
                                "start_page_1based": start_page_1based,
                                "end_page_1based": end_page_1based,
                                "page_count": source_page_count,
                                "bookmark_title": bookmark_title,
                            }
                        )

                        if self.config.get("privacy", {}).get("mask_md5_in_logs", True):
                            md5_log = self._mask_md5(source_md5)
                        else:
                            md5_log = source_md5
                        logging.info(
                            f"merge map record: {fname} | pages {start_page_1based}-{end_page_1based} | ID={source_short_id} | MD5={md5_log}"
                        )
                    except Exception as e:
                        logging.error(f"merge read failed {pdf_file}: {e}")

                # 3) Add clickable link annotations on index pages.
                if index_page_count > 0 and HAS_PYPDF_GENERIC:
                    # Parameters aligned with _create_index_doc_and_convert.
                    # A4 Height = 842 pt
                    # TopMargin = 72 pt
                    # Title + space ~= 60 pt (estimated).
                    # List start Y ~= 842 - 72 - 50 = 720.

                    # Coordinate conversion notes:
                    # Word origin is top-left; PDF origin is bottom-left.
                    # Word TopMargin 72 -> PDF Y = 842 - 72 = 770.
                    # Title 行占用约 30pt -> Y = 740.
                    # Empty line ~20pt -> Y ~= 720.
                    # First text line baseline is around 700-710.
                    # 行距 20pt.

                    start_y = 715
                    line_height = 20

                    lines_per_page = 32

                    for i, target_page_num in enumerate(file_start_pages):
                        page_idx = i // lines_per_page
                        row_idx = i % lines_per_page
                        if page_idx >= index_page_count:
                            break

                        idx_page = merger.pages[page_idx]

                        rect_top = start_y - (row_idx * line_height)
                        rect_bottom = rect_top - line_height
                        rect = [72, rect_bottom, 520, rect_top]  # [x1, y1, x2, y2]

                        # Create Link annotation.
                        # Target action: GoTo target_page_num.

                        # Resolve target page reference (IndirectObject) with compatibility handling.
                        # pypdf version differences: 3.x+ uses indirect_ref; older may use indirect_reference.
                        target_page_obj = merger.pages[target_page_num]
                        target_page_ref = getattr(target_page_obj, "indirect_ref", None)
                        if target_page_ref is None:
                            target_page_ref = getattr(
                                target_page_obj, "indirect_reference", None
                            )

                        if target_page_ref is None:
                            # If target reference cannot be resolved, skip this link to avoid breaking merge.
                            logging.warning(
                                f"cannot resolve target page reference {target_page_num} (indirect_ref/reference). skip index link."
                            )
                            continue

                        # pypdf 2.x/3.x annotation APIs vary, so use manual annotation dict.
                        link_annotation = DictionaryObject()
                        link_annotation.update(
                            {
                                NameObject("/Type"): NameObject("/Annot"),
                                NameObject("/Subtype"): NameObject("/Link"),
                                NameObject("/Rect"): ArrayObject(
                                    [FloatObject(c) for c in rect]
                                ),
                                NameObject("/Border"): ArrayObject(
                                    [NumberObject(0), NumberObject(0), NumberObject(0)]
                                ),
                                NameObject("/Dest"): ArrayObject(
                                    [target_page_ref, NameObject("/Fit")]
                                ),
                            }
                        )

                        if "/Annots" not in idx_page:
                            idx_page[NameObject("/Annots")] = ArrayObject()

                        idx_page["/Annots"].append(link_annotation)

                # 4. Write merged output
                merger.write(output_path)
                merger.close()

                if not os.path.exists(output_path):
                    raise RuntimeError(
                        f"merged output file not generated: {output_path}"
                    )
                try:
                    if os.path.getsize(output_path) <= 0:
                        raise RuntimeError(
                            f"merged output file size invalid (0 bytes): {output_path}"
                        )
                except OSError:
                    pass

                generated_outputs.append(output_path)
                if map_records:
                    merged_pdf_md5 = ""
                    try:
                        merged_pdf_md5 = self._compute_md5(output_path)
                    except Exception:
                        pass
                    for rec in map_records:
                        rec["merged_pdf_md5"] = merged_pdf_md5
                    merge_index_records.extend(map_records)

                if self.config.get("enable_merge_map", True):
                    try:
                        csv_path, json_path = self._write_merge_map(
                            output_path, map_records
                        )
                        if csv_path and json_path:
                            generated_map_outputs.extend([csv_path, json_path])
                            logging.info(f"map file generated: {csv_path}")
                            logging.info(f"map file generated: {json_path}")
                    except Exception as e:
                        logging.error(f"failed to write map files {output_path}: {e}")

                if index_pdf_path and os.path.exists(index_pdf_path):
                    os.remove(index_pdf_path)

            except Exception as e:
                print(f" [FAILED] {e}")
                logging.error(f"merge task failed {output_filename}: {e}")
                traceback.print_exc()

        if word_app:
            try:
                word_app.Quit()
            except Exception:
                pass
            pythoncom.CoUninitialize()

        self.merge_index_records = merge_index_records
        self.generated_merge_outputs = list(generated_outputs)
        self.generated_map_outputs = list(generated_map_outputs)

        if wb_merge:
            try:
                if merge_index_records:
                    ws_merge_index = wb_merge.create_sheet("MergeIndex")
                    self._write_merge_index_sheet(ws_merge_index, merge_index_records)
                if self.conversion_index_records:
                    ws_conv_index = wb_merge.create_sheet("ConvertedPDFs")
                    self._write_conversion_index_sheet(
                        ws_conv_index, self.conversion_index_records
                    )
                if ws_merge:
                    self._style_header_row(ws_merge)
                    self._auto_fit_sheet(ws_merge)
                wb_merge.save(merge_excel_path)
                self.merge_excel_path = merge_excel_path
                print(f"\n  Excel index saved: {merge_excel_path}")
            except Exception as e:
                logging.error(f"failed to save Excel merge list: {e}")

        if total_tasks <= 0:
            print(
                "\n  [INFO] No merge tasks generated. Ensure PDF files exist in scan results."
            )
        elif len(generated_outputs) <= 0:
            print(
                "\n  [INFO] Merge tasks executed, but no output was generated. Check logs and output permissions."
            )
        else:
            print("\n  Merged output files:")
            for p in generated_outputs:
                print(f"  - {p}")
            if generated_map_outputs:
                print("\n  Map files:")
                for p in generated_map_outputs:
                    print(f"  - {p}")

        return generated_outputs

    def _scan_merge_candidates_by_ext(self, ext):
        ext = str(ext or "").lower()
        if not ext.startswith("."):
            ext = "." + ext

        scan_source_type = "target"
        if self.run_mode == MODE_MERGE_ONLY:
            scan_source_type = self.config.get("merge_source", "source")

        if scan_source_type == "source":
            scan_roots = self._get_source_roots()
        else:
            scan_roots = [self.config["target_folder"]]

        files = []
        exclude_abs_paths = set(
            map(os.path.abspath, [self.failed_dir, self.merge_output_dir])
        )
        if scan_source_type == "source":
            exclude_abs_paths.add(os.path.abspath(self.config["target_folder"]))

        for scan_folder in scan_roots:
            if not scan_folder or not os.path.isdir(scan_folder):
                continue
            for root, dirs, names in os.walk(scan_folder):
                dirs[:] = [
                    d
                    for d in dirs
                    if os.path.abspath(os.path.join(root, d)) not in exclude_abs_paths
                ]
                if os.path.abspath(root) in exclude_abs_paths:
                    continue
                for name in names:
                    if name.lower().endswith(ext):
                        files.append(os.path.join(root, name))

        files.sort()
        return files

    def _build_markdown_merge_tasks(self, md_files):
        if not md_files:
            return []
        tasks = []
        if self.merge_mode == MERGE_MODE_ALL_IN_ONE:
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            tasks.append((f"Merged_All_{ts}.md", md_files))
            return tasks

        categories = {
            "Price": "Price_",
            "Word": "Word_",
            "Excel": "Excel_",
            "PPT": "PPT_",
            "PDF": "PDF_",
        }
        matched = set()
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        for cat_label, prefix in categories.items():
            group = [p for p in md_files if os.path.basename(p).startswith(prefix)]
            if not group:
                continue
            matched.update(group)
            tasks.append((f"Merged_{cat_label}_{ts}_001.md", group))

        others = [p for p in md_files if p not in matched]
        if others:
            tasks.append((f"Merged_Markdown_{ts}_001.md", others))
        return tasks

    def merge_markdowns(self, candidates=None):
        md_files = list(candidates or [])
        if not md_files:
            md_files = self._scan_merge_candidates_by_ext(".md")
        if not md_files:
            return []

        tasks = self._build_markdown_merge_tasks(md_files)
        if not tasks:
            return []

        generated = []
        for output_name, group in tasks:
            lines = [
                f"# {output_name}",
                "",
                f"- generated_at: {datetime.now().isoformat(timespec='seconds')}",
                f"- source_count: {len(group)}",
                "",
                "## Source Map",
                "",
            ]
            for idx, path in enumerate(group, 1):
                lines.append(f"{idx}. {path}")
            lines.extend(["", "## Documents", ""])

            for idx, path in enumerate(group, 1):
                try:
                    with open(path, "r", encoding="utf-8", errors="ignore") as f:
                        content = f.read().strip()
                except Exception as e:
                    logging.warning(f"skip markdown in merge {path}: {e}")
                    continue
                lines.extend(
                    [
                        f"### [{idx}] {os.path.basename(path)}",
                        "",
                        f"- source_markdown: {os.path.abspath(path)}",
                        "",
                        "---",
                        "",
                        content,
                        "",
                    ]
                )

            out_path = os.path.join(self.merge_output_dir, output_name)
            with open(out_path, "w", encoding="utf-8") as f:
                f.write("\n".join(lines).rstrip() + "\n")
            generated.append(out_path)
            logging.info(f"merged markdown generated: {out_path}")

        self.generated_merge_markdown_outputs = list(generated)
        return generated

    def _confirm_continue_missing_md_merge(self):
        msg = (
            "\n[WARN] Markdown merge requested but no .md files were found. "
            "Continue with remaining tasks? [Y/n]: "
        )
        if self.interactive:
            try:
                ans = input(msg).strip().lower()
                return ans in ("", "y", "yes")
            except Exception:
                return False
        # GUI/non-interactive: continue by default; GUI precheck can override.
        logging.warning(
            "Markdown merge requested but no .md files found. Continue by default in non-interactive mode."
        )
        return True

    def _run_merge_mode_pipeline(self, batch_results):
        pref = self._get_output_pref()
        merged_outputs = []
        submode = self._get_merge_convert_submode()

        if submode == MERGE_CONVERT_SUBMODE_PDF_TO_MD:
            pdf_files = self._scan_merge_candidates_by_ext(".pdf")
            need_markdown = pref["md"] and (pref["independent"] or pref["merged"])
            if need_markdown:
                for p in pdf_files:
                    md_path = self._export_pdf_markdown(p, source_path_hint=p)
                    batch_results.append(
                        {
                            "source_path": os.path.abspath(p),
                            "status": "success_non_pdf" if md_path else "failed",
                            "detail": "pdf_to_md" if md_path else "pdf_to_md_failed",
                            "final_path": md_path or "",
                            "elapsed": 0.0,
                        }
                    )

            if pref["merged"]:
                if pref["pdf"] and self.config.get("enable_merge", True):
                    merged_outputs.extend(self.merge_pdfs() or [])
                if pref["md"]:
                    md_candidates = [
                        p for p in self.generated_markdown_outputs if os.path.exists(p)
                    ] or self._scan_merge_candidates_by_ext(".md")
                    if not md_candidates:
                        if not self._confirm_continue_missing_md_merge():
                            raise RuntimeError("Markdown merge canceled by user.")
                    else:
                        merged_outputs.extend(self.merge_markdowns(md_candidates) or [])
            return merged_outputs

        # default: merge existing artifacts
        if pref["merged"]:
            if pref["pdf"] and self.config.get("enable_merge", True):
                merged_outputs.extend(self.merge_pdfs() or [])
            if pref["md"]:
                md_candidates = self._scan_merge_candidates_by_ext(".md")
                if not md_candidates:
                    if not self._confirm_continue_missing_md_merge():
                        raise RuntimeError("Markdown merge canceled by user.")
                    logging.info("Markdown merge skipped: no .md files found.")
                else:
                    merged_outputs.extend(self.merge_markdowns(md_candidates) or [])
        return merged_outputs

    # =============== File indexing and dedup ==================

    @staticmethod
    def _compute_file_hash(path, block_size=1024 * 1024):
        h = hashlib.sha256()
        with open(path, "rb") as f:
            while True:
                chunk = f.read(block_size)
                if not chunk:
                    break
                h.update(chunk)
        return h.hexdigest()

    @staticmethod
    def _make_file_hyperlink(path: str) -> str:
        path = os.path.abspath(path)
        return "file:///" + path.replace("\\", "/")

    def _append_conversion_index_record(self, source_path, pdf_path, status=""):
        if not source_path or not pdf_path:
            return
        if not os.path.exists(pdf_path):
            return

        src_abs = os.path.abspath(source_path)
        pdf_abs = os.path.abspath(pdf_path)

        try:
            src_md5 = self._compute_md5(src_abs)
        except Exception:
            src_md5 = ""
        try:
            pdf_md5 = self._compute_md5(pdf_abs)
        except Exception:
            pdf_md5 = ""

        try:
            src_rel = os.path.relpath(src_abs, self._get_source_root_for_path(src_abs))
        except Exception:
            src_rel = src_abs
        try:
            pdf_rel = os.path.relpath(pdf_abs, self.config["target_folder"])
        except Exception:
            pdf_rel = pdf_abs

        self.conversion_index_records.append(
            {
                "source_filename": os.path.basename(src_abs),
                "source_abspath": src_abs,
                "source_relpath": src_rel,
                "source_md5": src_md5,
                "pdf_filename": os.path.basename(pdf_abs),
                "pdf_abspath": pdf_abs,
                "pdf_relpath": pdf_rel,
                "pdf_md5": pdf_md5,
                "status": status or "",
            }
        )

    @staticmethod
    def _style_header_row(ws):
        if not HAS_OPENPYXL:
            return
        for cell in ws[1]:
            cell.font = Font(bold=True)

    @staticmethod
    def _auto_fit_sheet(ws, max_width=90):
        for col in ws.columns:
            max_length = 0
            col_letter = col[0].column_letter
            for cell in col:
                try:
                    v = str(cell.value) if cell.value is not None else ""
                    max_length = max(max_length, len(v))
                except Exception:
                    pass
            ws.column_dimensions[col_letter].width = min(max_length + 2, max_width)

    def _write_conversion_index_sheet(self, ws, records):
        headers = [
            "No.",
            "Source File",
            "Source Path",
            "Source MD5",
            "Output PDF",
            "Output PDF Path",
            "Output PDF MD5",
            "Status",
        ]
        ws.append(headers)
        self._style_header_row(ws)

        for idx, rec in enumerate(records, 1):
            ws.append(
                [
                    idx,
                    rec.get("source_filename", ""),
                    rec.get("source_abspath", ""),
                    rec.get("source_md5", ""),
                    rec.get("pdf_filename", ""),
                    rec.get("pdf_abspath", ""),
                    rec.get("pdf_md5", ""),
                    rec.get("status", ""),
                ]
            )
            src_cell = ws.cell(row=idx + 1, column=3)
            src_path = rec.get("source_abspath", "")
            if src_path:
                src_cell.hyperlink = self._make_file_hyperlink(src_path)
                src_cell.style = "Hyperlink"

            pdf_cell = ws.cell(row=idx + 1, column=6)
            pdf_path = rec.get("pdf_abspath", "")
            if pdf_path:
                pdf_cell.hyperlink = self._make_file_hyperlink(pdf_path)
                pdf_cell.style = "Hyperlink"

        self._auto_fit_sheet(ws)

    def _write_merge_index_sheet(self, ws, records):
        headers = [
            "No.",
            "Merged PDF",
            "Merged PDF Path",
            "Merged PDF MD5",
            "Source Order",
            "Source PDF",
            "Source PDF Path",
            "Source PDF MD5",
            "Short ID",
            "Page Range",
            "Start Page",
            "End Page",
            "Page Count",
        ]
        ws.append(headers)
        self._style_header_row(ws)

        for idx, rec in enumerate(records, 1):
            start_page = rec.get("start_page_1based", "")
            end_page = rec.get("end_page_1based", "")
            position = f"{start_page}-{end_page}" if start_page and end_page else ""

            ws.append(
                [
                    idx,
                    rec.get("merged_pdf_name", ""),
                    rec.get("merged_pdf_path", ""),
                    rec.get("merged_pdf_md5", ""),
                    rec.get("source_index", ""),
                    rec.get("source_filename", ""),
                    rec.get("source_abspath", ""),
                    rec.get("source_md5", ""),
                    rec.get("source_short_id", ""),
                    position,
                    start_page,
                    end_page,
                    rec.get("page_count", ""),
                ]
            )

            merged_cell = ws.cell(row=idx + 1, column=3)
            merged_path = rec.get("merged_pdf_path", "")
            if merged_path:
                merged_cell.hyperlink = self._make_file_hyperlink(merged_path)
                merged_cell.style = "Hyperlink"

            source_cell = ws.cell(row=idx + 1, column=7)
            source_path = rec.get("source_abspath", "")
            if source_path:
                source_cell.hyperlink = self._make_file_hyperlink(source_path)
                source_cell.style = "Hyperlink"

        self._auto_fit_sheet(ws)

    def _write_conversion_index_workbook(self):
        if not self.conversion_index_records:
            return None
        if not HAS_OPENPYXL:
            print(
                "\n[INFO] openpyxl not found. Conversion index Excel will not be generated."
            )
            logging.warning("openpyxl missing; conversion index Excel skipped.")
            return None

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        index_path = os.path.join(
            self.config["target_folder"], f"Convert_Index_{timestamp}.xlsx"
        )
        wb = Workbook()
        ws = wb.active
        ws.title = "ConvertedPDFs"
        self._write_conversion_index_sheet(ws, self.conversion_index_records)
        wb.save(index_path)
        self.convert_index_path = index_path
        print(f"\nConversion index saved: {index_path}")
        logging.info(f"Conversion index saved: {index_path}")
        return index_path

    def _build_ai_output_path(self, source_path, sub_dir, ext):
        target_root = self.config.get("target_folder", "")
        if not target_root:
            return None

        source_abs = os.path.abspath(source_path)
        try:
            rel = os.path.relpath(source_abs, target_root)
        except Exception:
            rel = os.path.basename(source_abs)
        if rel.startswith(".."):
            rel = os.path.basename(source_abs)
        rel_no_ext = os.path.splitext(rel)[0]

        output_path = os.path.join(target_root, "_AI", sub_dir, rel_no_ext + ext)
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        return output_path

    def _build_ai_output_path_from_source(self, source_path, sub_dir, ext):
        target_root = self.config.get("target_folder", "")
        if not target_root:
            return None

        source_abs = os.path.abspath(source_path)
        rel = os.path.basename(source_abs)
        try:
            src_root = self._get_source_root_for_path(source_abs)
            rel_try = os.path.relpath(source_abs, src_root)
            if not rel_try.startswith(".."):
                rel = rel_try
        except Exception:
            pass
        rel_no_ext = os.path.splitext(rel)[0]

        output_path = os.path.join(target_root, "_AI", sub_dir, rel_no_ext + ext)
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        return output_path

    @staticmethod
    def _normalize_extracted_text(text):
        text = (text or "").replace("\r\n", "\n").replace("\r", "\n")
        lines = [ln.rstrip() for ln in text.split("\n")]
        cleaned = []
        last_blank = False
        for ln in lines:
            if ln.strip():
                cleaned.append(ln)
                last_blank = False
            else:
                if not last_blank:
                    cleaned.append("")
                last_blank = True
        return "\n".join(cleaned).strip()

    @staticmethod
    def _normalize_margin_line(line):
        s = str(line or "").strip().lower()
        s = re.sub(r"\s+", " ", s)
        # Ignore punctuation-only differences for repeated margin detection.
        s = re.sub(r"[\W_]+", "", s, flags=re.UNICODE)
        return s

    @staticmethod
    def _looks_like_page_number_line(line):
        s = str(line or "").strip().lower()
        if not s:
            return False
        if re.fullmatch(r"[#\-\s]*\d+[#\-\s]*", s):
            return True
        if re.fullmatch(r"page\s+\d+(\s*/\s*\d+)?", s):
            return True
        if re.fullmatch(r"page\s+\d+\s+of\s+\d+", s):
            return True
        if re.fullmatch(r"\d+\s*/\s*\d+", s):
            return True
        if re.fullmatch(r"第\s*\d+\s*页(\s*/\s*共?\s*\d+\s*页)?", s):
            return True
        return False

    def _collect_margin_candidates(self, page_raw_texts):
        """Collect repeated top/bottom lines across pages as header/footer candidates."""
        header_counts = {}
        footer_counts = {}
        page_count = len(page_raw_texts)
        threshold = max(2, (page_count + 1) // 2)
        if page_count < 2:
            return set(), set()

        for raw in page_raw_texts:
            lines = self._normalize_extracted_text(raw).splitlines()
            non_empty = [ln.strip() for ln in lines if ln.strip()]
            if not non_empty:
                continue
            top = non_empty[0]
            bottom = non_empty[-1]
            top_key = self._normalize_margin_line(top)
            bottom_key = self._normalize_margin_line(bottom)
            if top_key and not self._looks_like_page_number_line(top):
                header_counts[top_key] = header_counts.get(top_key, 0) + 1
            if bottom_key and not self._looks_like_page_number_line(bottom):
                footer_counts[bottom_key] = footer_counts.get(bottom_key, 0) + 1

        header_keys = {k for k, v in header_counts.items() if v >= threshold}
        footer_keys = {k for k, v in footer_counts.items() if v >= threshold}
        return header_keys, footer_keys

    def _clean_markdown_page_lines(self, raw_text, header_keys, footer_keys):
        lines = self._normalize_extracted_text(raw_text).splitlines()
        kept = [ln.strip() for ln in lines]

        removed_header = 0
        removed_footer = 0
        removed_page_no = 0

        while kept:
            top = kept[0]
            top_key = self._normalize_margin_line(top)
            if self._looks_like_page_number_line(top):
                kept.pop(0)
                removed_page_no += 1
                continue
            if top_key and top_key in header_keys:
                kept.pop(0)
                removed_header += 1
                continue
            break

        while kept:
            bottom = kept[-1]
            bottom_key = self._normalize_margin_line(bottom)
            if self._looks_like_page_number_line(bottom):
                kept.pop()
                removed_page_no += 1
                continue
            if bottom_key and bottom_key in footer_keys:
                kept.pop()
                removed_footer += 1
                continue
            break

        return kept, {
            "removed_header_lines": removed_header,
            "removed_footer_lines": removed_footer,
            "removed_page_number_lines": removed_page_no,
            "remaining_lines": len(kept),
        }

    @staticmethod
    def _looks_like_heading_line(line):
        s = str(line or "").strip()
        if not s:
            return False
        if len(s) > 90:
            return False
        if re.match(r"^(\d+(\.\d+){0,3}|[一二三四五六七八九十]+)[\.\、\)]\s*\S+", s):
            return True
        if re.match(r"^(chapter|section)\s+\d+", s, flags=re.IGNORECASE):
            return True
        if re.match(r"^[A-Z0-9][A-Z0-9 \-_/]{3,}$", s):
            return True
        if s.endswith(":") and len(s) <= 40:
            return True
        return False

    def _render_markdown_blocks(self, lines, structured_headings=True):
        blocks = []
        buf = []
        heading_count = 0

        def flush_para():
            if not buf:
                return
            merged = ""
            for ln in buf:
                if not merged:
                    merged = ln
                    continue
                if merged.endswith("-") and ln:
                    merged = merged[:-1] + ln
                elif merged.endswith(
                    ("。", "！", "？", ".", "!", "?", "；", ";", "：", ":")
                ):
                    merged += "\n" + ln
                else:
                    merged += " " + ln
            blocks.append(merged.strip())
            buf.clear()

        for raw in lines:
            ln = str(raw or "").strip()
            if not ln:
                flush_para()
                continue
            if structured_headings and self._looks_like_heading_line(ln):
                flush_para()
                blocks.append(f"### {ln}")
                heading_count += 1
            else:
                buf.append(ln)
        flush_para()
        return blocks, heading_count

    @staticmethod
    def _json_safe_value(value):
        if value is None:
            return None
        if isinstance(value, (str, int, float, bool)):
            return value
        if isinstance(value, (datetime, dt_date, dt_time)):
            try:
                return value.isoformat()
            except Exception:
                return str(value)
        return str(value)

    @staticmethod
    def _is_empty_json_cell(value):
        if value is None:
            return True
        return isinstance(value, str) and value.strip() == ""

    def _is_effectively_empty_row(self, row_values):
        if not row_values:
            return True
        return all(self._is_empty_json_cell(v) for v in row_values)

    @staticmethod
    def _looks_like_header_row(row_values):
        if not row_values:
            return False
        non_empty = [
            v
            for v in row_values
            if not (v is None or (isinstance(v, str) and v.strip() == ""))
        ]
        if not non_empty:
            return False
        string_like = [v for v in non_empty if isinstance(v, str)]
        threshold = 1 if len(non_empty) == 1 else max(2, (len(non_empty) + 1) // 2)
        return len(string_like) >= threshold

    @staticmethod
    def _normalize_header_row(header_raw, width):
        names = []
        seen = {}
        for idx in range(width):
            base = ""
            if idx < len(header_raw):
                hv = header_raw[idx]
                if hv is not None:
                    base = str(hv).strip()
            if not base:
                base = f"col_{idx + 1}"

            if base not in seen:
                seen[base] = 1
                names.append(base)
            else:
                seen[base] += 1
                names.append(f"{base}_{seen[base]}")
        return names

    @staticmethod
    def _detect_json_value_type(value):
        if value is None:
            return "null"
        if isinstance(value, bool):
            return "boolean"
        if isinstance(value, int) and not isinstance(value, bool):
            return "integer"
        if isinstance(value, float):
            return "number"
        if isinstance(value, datetime):
            return "datetime"
        if isinstance(value, dt_date):
            return "date"
        if isinstance(value, dt_time):
            return "time"
        if isinstance(value, str):
            return "string"
        return "other"

    def _build_column_profiles(self, header, raw_rows, sample_limit):
        if not header:
            return []
        n = len(header)
        profiles = []
        for idx, name in enumerate(header):
            non_null = 0
            type_counts = {}
            sample_values = []
            for row in raw_rows[:sample_limit]:
                v = row[idx] if idx < len(row) else None
                if v is None:
                    continue
                if isinstance(v, str) and v.strip() == "":
                    continue
                non_null += 1
                t = self._detect_json_value_type(v)
                type_counts[t] = type_counts.get(t, 0) + 1
                if len(sample_values) < 3:
                    sample_values.append(self._json_safe_value(v))
            inferred_type = "null"
            if type_counts:
                inferred_type = sorted(
                    type_counts.items(), key=lambda kv: (-kv[1], kv[0])
                )[0][0]
            profiles.append(
                {
                    "index_1based": idx + 1,
                    "name": name,
                    "non_null_count": non_null,
                    "inferred_type": inferred_type,
                    "type_counts": type_counts,
                    "sample_values": sample_values,
                }
            )
        return profiles

    @staticmethod
    def _col_index_to_label(col_index_1based):
        n = int(col_index_1based)
        if n <= 0:
            return "A"
        chars = []
        while n > 0:
            n, rem = divmod(n - 1, 26)
            chars.append(chr(ord("A") + rem))
        return "".join(reversed(chars))

    @staticmethod
    def _extract_formula_sheet_refs(formula_text, current_sheet_name):
        if not formula_text or not isinstance(formula_text, str):
            return set()

        refs = set()
        # Matches quoted and unquoted sheet references like:
        # 'Sheet A'!A1, Sheet1!B2, [Book.xlsx]Sheet1!A1
        for m in re.finditer(
            r"(?:'((?:[^']|'')+)'|([A-Za-z0-9_.\[\]]+))!", formula_text
        ):
            cand = m.group(1) if m.group(1) is not None else m.group(2)
            cand = (cand or "").replace("''", "'").strip()
            if not cand:
                continue
            if "]" in cand:
                cand = cand.split("]", 1)[1].strip()
            if cand and cand != current_sheet_name:
                refs.add(cand)
        return refs

    @staticmethod
    def _extract_chart_title_text(chart):
        title_obj = getattr(chart, "title", None)
        if title_obj is None:
            return ""
        if isinstance(title_obj, str):
            return title_obj.strip()

        # openpyxl chart title is often rich text; fall back to str(title_obj).
        try:
            tx = getattr(title_obj, "tx", None)
            rich = getattr(tx, "rich", None) if tx is not None else None
            if rich is not None:
                parts = []
                for para in getattr(rich, "p", []) or []:
                    for run in getattr(para, "r", []) or []:
                        txt = getattr(run, "t", None)
                        if txt:
                            parts.append(str(txt))
                if parts:
                    return "".join(parts).strip()
        except Exception:
            pass

        try:
            return str(title_obj).strip()
        except Exception:
            return ""

    @staticmethod
    def _stringify_chart_anchor(anchor):
        if anchor is None:
            return ""
        try:
            marker = getattr(anchor, "_from", None)
            if marker is not None:
                col = int(getattr(marker, "col", 0)) + 1
                row = int(getattr(marker, "row", 0)) + 1
                return f"{OfficeConverter._col_index_to_label(col)}{row}"
        except Exception:
            pass
        try:
            return str(anchor)
        except Exception:
            return ""

    def _extract_sheet_charts(self, ws_formula, series_ref_limit=50):
        charts = []
        if ws_formula is None:
            return charts
        for idx, chart in enumerate(getattr(ws_formula, "_charts", []) or [], 1):
            refs = []
            for series in getattr(chart, "series", []) or []:
                for attr in ("val", "xVal", "yVal", "cat", "bubbleSize"):
                    ref_obj = getattr(series, attr, None)
                    if ref_obj is None:
                        continue
                    num_ref = getattr(ref_obj, "numRef", None)
                    str_ref = getattr(ref_obj, "strRef", None)
                    for ref in (num_ref, str_ref):
                        formula_ref = (
                            getattr(ref, "f", None) if ref is not None else None
                        )
                        if formula_ref:
                            refs.append(str(formula_ref))
            dedup_refs = sorted(set(refs))
            charts.append(
                {
                    "index_1based": idx,
                    "chart_type": chart.__class__.__name__,
                    "title": self._extract_chart_title_text(chart),
                    "anchor": self._stringify_chart_anchor(
                        getattr(chart, "anchor", None)
                    ),
                    "series_ref_count": len(dedup_refs),
                    "series_refs_truncated": len(dedup_refs) > series_ref_limit,
                    "series_refs": dedup_refs[:series_ref_limit],
                }
            )
        return charts

    @staticmethod
    def _extract_sheet_pivot_tables(ws_formula):
        pivots = []
        if ws_formula is None:
            return pivots
        for idx, pivot in enumerate(getattr(ws_formula, "_pivots", []) or [], 1):
            loc_ref = ""
            try:
                location = getattr(pivot, "location", None)
                loc_ref = str(getattr(location, "ref", "") or "")
            except Exception:
                loc_ref = ""
            pivots.append(
                {
                    "index_1based": idx,
                    "name": str(getattr(pivot, "name", "") or ""),
                    "cache_id": getattr(pivot, "cacheId", None),
                    "location_ref": loc_ref,
                }
            )
        return pivots

    @staticmethod
    def _extract_workbook_defined_names(wb_formula):
        names = []
        if wb_formula is None:
            return names

        dn_container = getattr(wb_formula, "defined_names", None)
        if dn_container is None:
            return names

        dn_objects = []
        raw_list = getattr(dn_container, "definedName", None)
        if raw_list:
            dn_objects.extend(list(raw_list))
        try:
            for _, dn in dn_container.items():
                if isinstance(dn, (list, tuple, set)):
                    dn_objects.extend(list(dn))
                else:
                    dn_objects.append(dn)
        except Exception:
            pass

        seen = set()
        for dn in dn_objects:
            if dn is None:
                continue
            name = str(getattr(dn, "name", "") or "")
            local_sheet_id = getattr(dn, "localSheetId", None)
            hidden = bool(getattr(dn, "hidden", False))
            attr_text = str(getattr(dn, "attr_text", "") or "")
            comment = str(getattr(dn, "comment", "") or "")

            destinations = []
            try:
                for sheet_name, ref in dn.destinations:
                    destinations.append({"sheet": str(sheet_name), "ref": str(ref)})
            except Exception:
                pass

            dedup_key = (
                name,
                str(local_sheet_id),
                hidden,
                attr_text,
                json.dumps(destinations, ensure_ascii=False, sort_keys=True),
            )
            if dedup_key in seen:
                continue
            seen.add(dedup_key)

            names.append(
                {
                    "name": name,
                    "local_sheet_id": local_sheet_id,
                    "hidden": hidden,
                    "attr_text": attr_text,
                    "is_formula": bool(attr_text.startswith("=")),
                    "comment": comment,
                    "destinations": destinations,
                }
            )

        names.sort(key=lambda x: (x.get("name", ""), x.get("local_sheet_id") or -1))
        return names

    def _export_pdf_markdown(self, pdf_path, source_path_hint=None):
        if not self.config.get(
            "output_enable_md", self.config.get("enable_markdown", True)
        ):
            return None
        if not HAS_PYPDF:
            return None
        if not pdf_path or not os.path.exists(pdf_path):
            return None

        if source_path_hint:
            md_path = self._build_ai_output_path_from_source(
                source_path_hint, "Markdown", ".md"
            )
        else:
            md_path = self._build_ai_output_path(pdf_path, "Markdown", ".md")
        if not md_path:
            return None

        try:
            reader = PdfReader(pdf_path)
            page_count = len(reader.pages)
            raw_page_texts = []
            for page in reader.pages:
                text = ""
                try:
                    text = page.extract_text() or ""
                except Exception:
                    text = ""
                raw_page_texts.append(text)

            strip_header_footer = bool(
                self.config.get("markdown_strip_header_footer", True)
            )
            structured_headings = bool(
                self.config.get("markdown_structured_headings", True)
            )
            header_keys, footer_keys = (
                self._collect_margin_candidates(raw_page_texts)
                if strip_header_footer
                else (set(), set())
            )

            lines = [
                f"# {os.path.basename(pdf_path)}",
                "",
                f"- source_pdf: {os.path.abspath(pdf_path)}",
                f"- page_count: {page_count}",
                f"- generated_at: {datetime.now().isoformat(timespec='seconds')}",
                "",
            ]

            removed_header_total = 0
            removed_footer_total = 0
            removed_page_no_total = 0
            heading_total = 0
            non_empty_pages = 0

            for idx, raw_text in enumerate(raw_page_texts, 1):
                page_lines, page_stats = self._clean_markdown_page_lines(
                    raw_text, header_keys, footer_keys
                )
                blocks, heading_count = self._render_markdown_blocks(
                    page_lines, structured_headings=structured_headings
                )
                page_body = "\n\n".join(blocks).strip()
                if page_body:
                    non_empty_pages += 1
                else:
                    page_body = "(empty)"

                removed_header_total += page_stats.get("removed_header_lines", 0)
                removed_footer_total += page_stats.get("removed_footer_lines", 0)
                removed_page_no_total += page_stats.get("removed_page_number_lines", 0)
                heading_total += heading_count

                lines.extend([f"## Page {idx}", "", page_body, ""])

            with open(md_path, "w", encoding="utf-8") as f:
                f.write("\n".join(lines).rstrip() + "\n")

            self.generated_markdown_outputs.append(md_path)
            source_pdf_for_meta = (
                os.path.abspath(source_path_hint)
                if source_path_hint
                else os.path.abspath(pdf_path)
            )
            self.markdown_quality_records.append(
                {
                    "source_pdf": source_pdf_for_meta,
                    "markdown_path": os.path.abspath(md_path),
                    "page_count": page_count,
                    "non_empty_page_count": non_empty_pages,
                    "removed_header_lines": removed_header_total,
                    "removed_footer_lines": removed_footer_total,
                    "removed_page_number_lines": removed_page_no_total,
                    "heading_count": heading_total,
                    "header_candidate_count": len(header_keys),
                    "footer_candidate_count": len(footer_keys),
                    "strip_header_footer": strip_header_footer,
                    "structured_headings": structured_headings,
                }
            )
            logging.info(f"Markdown export success: {md_path}")
            return md_path
        except Exception as e:
            logging.error(f"Markdown export failed {pdf_path}: {e}")
            return None

    def _write_markdown_quality_report(self):
        if not self.config.get("enable_markdown_quality_report", True):
            return None
        if not self.markdown_quality_records:
            return None

        target_root = self.config.get("target_folder", "")
        if not target_root:
            return None

        sample_limit = max(
            1, int(self.config.get("markdown_quality_sample_limit", 20) or 20)
        )
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_dir = os.path.join(target_root, "_AI", "MarkdownQuality")
        os.makedirs(out_dir, exist_ok=True)
        out_path = os.path.join(out_dir, f"Markdown_Quality_Report_{ts}.json")

        total_pages = 0
        non_empty_pages = 0
        header_removed = 0
        footer_removed = 0
        page_no_removed = 0
        heading_count = 0
        samples = []

        for rec in self.markdown_quality_records:
            total_pages += int(rec.get("page_count", 0) or 0)
            non_empty_pages += int(rec.get("non_empty_page_count", 0) or 0)
            header_removed += int(rec.get("removed_header_lines", 0) or 0)
            footer_removed += int(rec.get("removed_footer_lines", 0) or 0)
            page_no_removed += int(rec.get("removed_page_number_lines", 0) or 0)
            heading_count += int(rec.get("heading_count", 0) or 0)
            if len(samples) < sample_limit:
                samples.append(
                    {
                        "source_pdf": rec.get("source_pdf", ""),
                        "markdown_path": rec.get("markdown_path", ""),
                        "page_count": rec.get("page_count", 0),
                        "non_empty_page_count": rec.get("non_empty_page_count", 0),
                        "removed_header_lines": rec.get("removed_header_lines", 0),
                        "removed_footer_lines": rec.get("removed_footer_lines", 0),
                        "removed_page_number_lines": rec.get(
                            "removed_page_number_lines", 0
                        ),
                        "heading_count": rec.get("heading_count", 0),
                    }
                )

        payload = {
            "version": 1,
            "generated_at": datetime.now().isoformat(timespec="seconds"),
            "record_count": len(self.markdown_quality_records),
            "summary": {
                "total_pages": total_pages,
                "non_empty_pages": non_empty_pages,
                "removed_header_lines": header_removed,
                "removed_footer_lines": footer_removed,
                "removed_page_number_lines": page_no_removed,
                "heading_count": heading_count,
            },
            "sample_limit": sample_limit,
            "samples": samples,
            "records": self.markdown_quality_records,
        }
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)

        self.markdown_quality_report_path = out_path
        self.generated_markdown_quality_outputs.append(out_path)
        logging.info(f"Markdown quality report generated: {out_path}")
        return out_path

    def _write_records_json_exports(self):
        if not self.config.get("enable_excel_json", False):
            return []

        target_root = self.config.get("target_folder", "")
        if not target_root:
            return []

        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_dir = os.path.join(target_root, "_AI", "Records")
        os.makedirs(out_dir, exist_ok=True)

        outputs = []

        if self.conversion_index_records:
            convert_path = os.path.join(out_dir, f"Convert_Records_{ts}.json")
            with open(convert_path, "w", encoding="utf-8") as f:
                json.dump(
                    {
                        "version": 1,
                        "record_type": "convert_index",
                        "record_count": len(self.conversion_index_records),
                        "records": self.conversion_index_records,
                    },
                    f,
                    ensure_ascii=False,
                    indent=2,
                )
            outputs.append(convert_path)

        if self.merge_index_records:
            merge_path = os.path.join(out_dir, f"Merge_Records_{ts}.json")
            with open(merge_path, "w", encoding="utf-8") as f:
                json.dump(
                    {
                        "version": 1,
                        "record_type": "merge_index",
                        "record_count": len(self.merge_index_records),
                        "records": self.merge_index_records,
                    },
                    f,
                    ensure_ascii=False,
                    indent=2,
                )
            outputs.append(merge_path)

        self.generated_records_json_outputs = outputs
        for p in outputs:
            logging.info(f"Records JSON generated: {p}")
        return outputs

    @staticmethod
    def _sanitize_chromadb_collection_name(name):
        s = re.sub(r"[^A-Za-z0-9._-]+", "_", str(name or "").strip())
        s = s.strip("._-")
        if len(s) < 3:
            s = f"corpus_{s}" if s else "office_corpus"
        if len(s) > 63:
            s = s[:63].rstrip("._-")
        if not s:
            s = "office_corpus"
        return s

    def _resolve_chromadb_persist_dir(self):
        configured = str(self.config.get("chromadb_persist_dir", "") or "").strip()
        if configured:
            if os.path.isabs(configured):
                return configured
            return os.path.abspath(
                os.path.join(self.config.get("target_folder", ""), configured)
            )
        return os.path.join(
            self.config.get("target_folder", ""), "_AI", "ChromaDB", "db"
        )

    @staticmethod
    def _chunk_text_for_vector(text, max_chars=1800, overlap=200):
        content = str(text or "").strip()
        if not content:
            return []
        max_chars = max(200, int(max_chars or 1800))
        overlap = max(0, int(overlap or 0))
        if overlap >= max_chars:
            overlap = max(0, max_chars // 5)

        chunks = []
        start = 0
        n = len(content)
        while start < n:
            end = min(n, start + max_chars)
            piece = content[start:end].strip()
            if piece:
                chunks.append(piece)
            if end >= n:
                break
            start = max(0, end - overlap)
        return chunks

    def _collect_chromadb_documents(self):
        docs = []
        md_paths = []
        seen = set()
        for p in self.generated_markdown_outputs or []:
            abs_p = os.path.abspath(str(p))
            if abs_p in seen or not os.path.exists(abs_p):
                continue
            seen.add(abs_p)
            md_paths.append(abs_p)

        md_to_pdf = {}
        for rec in self.markdown_quality_records or []:
            mdp = os.path.abspath(str(rec.get("markdown_path", "") or ""))
            if mdp:
                md_to_pdf[mdp] = str(rec.get("source_pdf", "") or "")

        max_chars = int(self.config.get("chromadb_max_chars_per_chunk", 1800) or 1800)
        overlap = int(self.config.get("chromadb_chunk_overlap", 200) or 200)

        for md_path in md_paths:
            try:
                with open(md_path, "r", encoding="utf-8") as f:
                    raw = f.read()
            except Exception:
                continue
            chunks = self._chunk_text_for_vector(
                raw, max_chars=max_chars, overlap=overlap
            )
            if not chunks:
                continue
            source_pdf = md_to_pdf.get(md_path, "")
            path_hash = hashlib.sha1(
                md_path.encode("utf-8", errors="ignore")
            ).hexdigest()[:16]
            for idx, chunk in enumerate(chunks, 1):
                doc_id = f"md_{path_hash}_{idx:05d}"
                docs.append(
                    {
                        "id": doc_id,
                        "document": chunk,
                        "metadata": {
                            "kind": "markdown",
                            "source_markdown_path": md_path,
                            "source_pdf_path": source_pdf,
                            "chunk_index": idx,
                            "chunk_count": len(chunks),
                            "char_count": len(chunk),
                        },
                    }
                )
        return docs

    def _write_chromadb_export(self):
        if not self.config.get("enable_chromadb_export", False):
            return None

        target_root = self.config.get("target_folder", "")
        if not target_root:
            return None
        os.makedirs(target_root, exist_ok=True)

        out_dir = os.path.join(target_root, "_AI", "ChromaDB")
        os.makedirs(out_dir, exist_ok=True)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        manifest_path = os.path.join(out_dir, f"chroma_export_{ts}.json")
        jsonl_path = os.path.join(out_dir, f"chroma_docs_{ts}.jsonl")

        docs = self._collect_chromadb_documents()
        if not docs:
            self.generated_chromadb_outputs = []
            self.chromadb_export_manifest_path = None
            logging.info("ChromaDB export skipped: no Markdown chunks available.")
            return None
        write_jsonl = bool(self.config.get("chromadb_write_jsonl_fallback", True))
        if write_jsonl:
            with open(jsonl_path, "w", encoding="utf-8") as f:
                for item in docs:
                    f.write(json.dumps(item, ensure_ascii=False) + "\n")

        collection_name = self._sanitize_chromadb_collection_name(
            self.config.get("chromadb_collection_name", "office_corpus")
        )
        persist_dir = self._resolve_chromadb_persist_dir()
        status = "empty"
        error = ""
        collection_count = 0

        if docs and HAS_CHROMADB:
            try:
                os.makedirs(persist_dir, exist_ok=True)
                client = chromadb.PersistentClient(path=persist_dir)
                collection = client.get_or_create_collection(name=collection_name)

                batch_size = 200
                for i in range(0, len(docs), batch_size):
                    batch = docs[i : i + batch_size]
                    ids = [x["id"] for x in batch]
                    documents = [x["document"] for x in batch]
                    metadatas = []
                    for x in batch:
                        md = {}
                        for k, v in (x.get("metadata", {}) or {}).items():
                            if isinstance(v, (str, int, float, bool)):
                                md[k] = v
                            elif v is None:
                                md[k] = ""
                            else:
                                md[k] = str(v)
                        metadatas.append(md)
                    collection.upsert(ids=ids, documents=documents, metadatas=metadatas)
                collection_count = int(collection.count() or 0)
                status = "ok"
            except Exception as e:
                status = "failed"
                error = str(e)
        elif docs and not HAS_CHROMADB:
            status = "chromadb_missing"
            error = "chromadb not installed"

        payload = {
            "version": 1,
            "generated_at": datetime.now().isoformat(timespec="seconds"),
            "status": status,
            "error": error,
            "record_count": len(docs),
            "chromadb_available": HAS_CHROMADB,
            "persist_dir": os.path.abspath(persist_dir),
            "collection_name": collection_name,
            "collection_count": collection_count,
            "jsonl_path": os.path.abspath(jsonl_path) if write_jsonl else "",
        }
        with open(manifest_path, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)

        outputs = [manifest_path]
        if write_jsonl:
            outputs.append(jsonl_path)
        self.generated_chromadb_outputs = outputs
        self.chromadb_export_manifest_path = manifest_path
        logging.info(
            f"ChromaDB export status={status}, records={len(docs)}: {manifest_path}"
        )
        return manifest_path

    def _export_single_excel_json(self, source_excel_path):
        out_path = self._build_ai_output_path_from_source(
            source_excel_path, "ExcelJSON", ".json"
        )
        if not out_path:
            return None

        max_rows = max(1, int(self.config.get("excel_json_max_rows", 2000) or 2000))
        max_cols = max(1, int(self.config.get("excel_json_max_cols", 80) or 80))
        records_preview_limit = max(
            1, int(self.config.get("excel_json_records_preview", 200) or 200)
        )
        profile_rows_limit = max(
            1, int(self.config.get("excel_json_profile_rows", 500) or 500)
        )
        include_formulas = bool(self.config.get("excel_json_include_formulas", True))
        extract_sheet_links = bool(
            self.config.get("excel_json_extract_sheet_links", True)
        )
        include_merged_ranges = bool(
            self.config.get("excel_json_include_merged_ranges", True)
        )
        formula_sample_limit = max(
            1, int(self.config.get("excel_json_formula_sample_limit", 200) or 200)
        )
        merged_range_limit = max(
            1, int(self.config.get("excel_json_merged_range_limit", 500) or 500)
        )

        payload = {
            "version": 1,
            "generated_at": datetime.now().isoformat(timespec="seconds"),
            "source_excel": os.path.abspath(source_excel_path),
            "source_relpath": "",
            "parse_status": "ok",
            "error": "",
            "limits": {
                "max_rows": max_rows,
                "max_cols": max_cols,
                "records_preview_limit": records_preview_limit,
                "profile_rows_limit": profile_rows_limit,
                "include_formulas": include_formulas,
                "extract_sheet_links": extract_sheet_links,
                "include_merged_ranges": include_merged_ranges,
                "formula_sample_limit": formula_sample_limit,
                "merged_range_limit": merged_range_limit,
            },
            "sheets": [],
            "workbook_links": [],
            "workbook_defined_name_count": 0,
            "workbook_defined_names": [],
            "chart_count_total": 0,
            "pivot_table_count_total": 0,
        }
        try:
            payload["source_relpath"] = os.path.relpath(
                os.path.abspath(source_excel_path),
                self._get_source_root_for_path(source_excel_path),
            )
        except Exception:
            payload["source_relpath"] = os.path.basename(source_excel_path)

        ext = os.path.splitext(source_excel_path)[1].lower()
        if ext == ".xls":
            payload["parse_status"] = "unsupported_format_xls"
            payload["error"] = "xls not supported by openpyxl"
        elif not HAS_OPENPYXL:
            payload["parse_status"] = "openpyxl_missing"
            payload["error"] = "openpyxl not installed"
        else:
            wb_values = None
            wb_formula = None
            try:
                wb_values = load_workbook(
                    source_excel_path, data_only=True, read_only=True
                )
                if include_formulas or extract_sheet_links or include_merged_ranges:
                    # Open a non-readonly workbook for formula and merged-range metadata.
                    wb_formula = load_workbook(
                        source_excel_path, data_only=False, read_only=False
                    )

                workbook_link_counts = {}
                if wb_formula is not None:
                    payload["workbook_defined_names"] = (
                        self._extract_workbook_defined_names(wb_formula)
                    )
                    payload["workbook_defined_name_count"] = len(
                        payload["workbook_defined_names"]
                    )

                for ws in wb_values.worksheets:
                    ws_formula = None
                    if wb_formula is not None and ws.title in wb_formula.sheetnames:
                        ws_formula = wb_formula[ws.title]

                    rows_json = []
                    rows_raw = []
                    rows_meta = []
                    row_count_scanned = 0
                    row_count_empty_skipped = 0
                    truncated = False
                    formula_cells_count = 0
                    formula_samples = []
                    sheet_link_counts = {}

                    for row in ws.iter_rows(values_only=True):
                        row_count_scanned += 1
                        if row_count_scanned > max_rows:
                            truncated = True
                            break

                        raw_vals = list(row)[:max_cols]
                        vals = [self._json_safe_value(v) for v in raw_vals]
                        formula_map = {}

                        if ws_formula is not None and (
                            include_formulas or extract_sheet_links
                        ):
                            for ci in range(1, max_cols + 1):
                                fval = ws_formula.cell(
                                    row=row_count_scanned, column=ci
                                ).value
                                if isinstance(fval, str) and fval.startswith("="):
                                    formula_map[ci - 1] = fval
                                    formula_cells_count += 1
                                    if len(formula_samples) < formula_sample_limit:
                                        addr = f"{self._col_index_to_label(ci)}{row_count_scanned}"
                                        formula_samples.append(
                                            {
                                                "cell": addr,
                                                "formula": fval,
                                                "value": vals[ci - 1]
                                                if ci - 1 < len(vals)
                                                else None,
                                            }
                                        )
                                    if extract_sheet_links:
                                        refs = self._extract_formula_sheet_refs(
                                            fval, ws.title
                                        )
                                        for ref_sheet in refs:
                                            sheet_link_counts[ref_sheet] = (
                                                sheet_link_counts.get(ref_sheet, 0) + 1
                                            )

                        if len(row) > max_cols:
                            truncated = True

                        # Trim trailing empty cells, but keep trailing formula cells.
                        last_keep_idx = -1
                        for idx in range(len(vals)):
                            if (not self._is_empty_json_cell(vals[idx])) or (
                                idx in formula_map
                            ):
                                last_keep_idx = idx
                        if last_keep_idx >= 0:
                            vals = vals[: last_keep_idx + 1]
                            raw_vals = raw_vals[: last_keep_idx + 1]
                        else:
                            vals = []
                            raw_vals = []
                            formula_map = {}

                        if self._is_effectively_empty_row(vals) and not formula_map:
                            row_count_empty_skipped += 1
                            continue

                        rows_json.append(vals)
                        rows_raw.append(raw_vals)
                        rows_meta.append(
                            {
                                "source_row_index_1based": row_count_scanned,
                                "formulas_by_col_index0": formula_map,
                            }
                        )

                    width = 0
                    for r in rows_json:
                        if len(r) > width:
                            width = len(r)

                    header_detected = False
                    header_row_index = None
                    header_raw = []
                    data_rows_json = rows_json
                    data_rows_raw = rows_raw
                    data_rows_meta = rows_meta
                    if rows_json and self._looks_like_header_row(rows_json[0]):
                        header_detected = True
                        header_row_index = 0
                        header_raw = rows_json[0]
                        data_rows_json = rows_json[1:]
                        data_rows_raw = rows_raw[1:]
                        data_rows_meta = rows_meta[1:]
                        if len(header_raw) > width:
                            width = len(header_raw)

                    header = self._normalize_header_row(header_raw, width)

                    records_preview = []
                    for row_vals, row_meta in zip(
                        data_rows_json[:records_preview_limit],
                        data_rows_meta[:records_preview_limit],
                    ):
                        rec = {}
                        formula_named = {}
                        for idx, col_name in enumerate(header):
                            v = row_vals[idx] if idx < len(row_vals) else None
                            rec[col_name] = v
                            ftxt = row_meta.get("formulas_by_col_index0", {}).get(idx)
                            if ftxt:
                                formula_named[col_name] = ftxt
                        rec["_source_row_index_1based"] = row_meta.get(
                            "source_row_index_1based"
                        )
                        if formula_named:
                            rec["__formulas"] = formula_named
                        records_preview.append(rec)

                    column_profiles = self._build_column_profiles(
                        header, data_rows_raw, profile_rows_limit
                    )

                    merged_ranges = []
                    merged_ranges_truncated = False
                    if include_merged_ranges and ws_formula is not None:
                        all_ranges = list(ws_formula.merged_cells.ranges)
                        if len(all_ranges) > merged_range_limit:
                            merged_ranges_truncated = True
                        for rg in all_ranges[:merged_range_limit]:
                            merged_ranges.append(
                                {
                                    "range": str(rg),
                                    "top_left_row_1based": int(rg.min_row),
                                    "top_left_col_1based": int(rg.min_col),
                                    "top_left_value": self._json_safe_value(
                                        ws_formula.cell(
                                            row=rg.min_row, column=rg.min_col
                                        ).value
                                    ),
                                }
                            )

                    linked_sheets = sorted(sheet_link_counts.keys())
                    for to_sheet, ref_count in sheet_link_counts.items():
                        edge_key = (ws.title, to_sheet)
                        workbook_link_counts[edge_key] = (
                            workbook_link_counts.get(edge_key, 0) + ref_count
                        )
                    charts = self._extract_sheet_charts(ws_formula)
                    pivots = self._extract_sheet_pivot_tables(ws_formula)
                    payload["chart_count_total"] += len(charts)
                    payload["pivot_table_count_total"] += len(pivots)

                    payload["sheets"].append(
                        {
                            "name": ws.title,
                            "row_count_scanned": row_count_scanned,
                            "row_count_exported": len(rows_json),
                            "row_count_empty_skipped": row_count_empty_skipped,
                            "max_cols_exported": max_cols,
                            "truncated": truncated,
                            "header_detected": header_detected,
                            "header_row_index_exported": header_row_index,
                            "header": header,
                            "data_row_count": len(data_rows_json),
                            "rows": rows_json,
                            "records_preview": records_preview,
                            "column_profiles": column_profiles,
                            "formula_stats": {
                                "formula_cells_count": formula_cells_count,
                                "formula_sample_count": len(formula_samples),
                                "formula_samples": formula_samples,
                                "linked_sheets": linked_sheets,
                                "linked_sheet_ref_counts": sheet_link_counts,
                            },
                            "merged_ranges_count": len(merged_ranges),
                            "merged_ranges_truncated": merged_ranges_truncated,
                            "merged_ranges": merged_ranges,
                            "charts_count": len(charts),
                            "charts": charts,
                            "pivot_tables_count": len(pivots),
                            "pivot_tables": pivots,
                        }
                    )

                payload["workbook_links"] = [
                    {
                        "from_sheet": k[0],
                        "to_sheet": k[1],
                        "ref_count": v,
                    }
                    for k, v in sorted(
                        workbook_link_counts.items(),
                        key=lambda kv: (kv[0][0], kv[0][1]),
                    )
                ]
            except Exception as e:
                payload["parse_status"] = "parse_failed"
                payload["error"] = str(e)
            finally:
                if wb_values is not None:
                    try:
                        wb_values.close()
                    except Exception:
                        pass
                if wb_formula is not None:
                    try:
                        wb_formula.close()
                    except Exception:
                        pass

        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
        return out_path

    def _write_excel_structured_json_exports(self):
        if not self.config.get("enable_excel_json", False):
            return []

        excel_exts = set(
            e.lower()
            for e in self.config.get("allowed_extensions", {}).get("excel", [])
        )
        source_paths = []
        seen = set()
        for rec in self.conversion_index_records:
            src = os.path.abspath(rec.get("source_abspath", "") or "")
            if not src or not os.path.exists(src):
                continue
            if os.path.splitext(src)[1].lower() not in excel_exts:
                continue
            if src in seen:
                continue
            seen.add(src)
            source_paths.append(src)

        outputs = []
        for src in source_paths:
            try:
                out_path = self._export_single_excel_json(src)
                if out_path and os.path.exists(out_path):
                    outputs.append(out_path)
                    logging.info(f"Excel JSON generated: {out_path}")
            except Exception as e:
                logging.error(f"Excel JSON export failed {src}: {e}")

        self.generated_excel_json_outputs = outputs
        return outputs

    def _safe_file_meta(self, path):
        if not path:
            return None
        abs_path = os.path.abspath(path)
        if not os.path.exists(abs_path):
            return None
        try:
            stat = os.stat(abs_path)
        except OSError:
            return None

        try:
            rel_path = os.path.relpath(abs_path, self.config["target_folder"])
        except Exception:
            rel_path = abs_path

        md5_value = ""
        sha256_value = ""
        try:
            md5_value = self._compute_md5(abs_path)
        except Exception:
            pass
        try:
            sha256_value = self._compute_file_hash(abs_path)
        except Exception:
            pass

        return {
            "path_abs": abs_path,
            "path_rel_to_target": rel_path,
            "size_bytes": int(stat.st_size),
            "mtime": datetime.fromtimestamp(stat.st_mtime).isoformat(
                timespec="seconds"
            ),
            "md5": md5_value,
            "sha256": sha256_value,
        }

    def _add_artifact(self, artifacts, kind, path):
        meta = self._safe_file_meta(path)
        if not meta:
            return
        item = {"kind": kind}
        item.update(meta)
        artifacts.append(item)

    def _maybe_build_llm_delivery_hub(self, target_folder, artifacts):
        if not self.config.get("enable_llm_delivery_hub", True):
            return None

        if not artifacts:
            return None

        llm_root = self.config.get("llm_delivery_root") or os.path.join(
            target_folder, "_LLM_UPLOAD"
        )
        try:
            os.makedirs(llm_root, exist_ok=True)
        except Exception as e:
            logging.error(f"failed to create LLM hub root {llm_root}: {e}")
            return None

        include_pdf = self.config.get("llm_delivery_include_pdf", False)
        flatten = self.config.get("llm_delivery_flatten", False)

        hub_items = []
        counts = {
            "markdown": 0,
            "json": 0,
            "pdf": 0,
            "other": 0,
        }

        # Kind whitelist: only include content files useful for LLM ingestion
        _LLM_CONTENT_KINDS = {
            "markdown_export",  # Converted Markdown from Office/PDF
            "merged_markdown",  # General merged markdown packages
            "mshelp_merged_markdown",  # MSHelp merged documentation
            "excel_structured_json",  # Structured Excel data (JSON)
        }
        _LLM_PDF_KINDS = {
            "merged_pdf",  # Merged PDF volumes
            "converted_pdf",  # Individual converted PDFs
            "mshelp_merged_pdf",  # MSHelp merged PDF
        }

        # ---- Merge dedup: if merged docs exist, skip individual sources ----
        dedup_enabled = self.config.get("upload_dedup_merged", True)
        has_merged_pdf = dedup_enabled and any(
            a.get("kind") == "merged_pdf" for a in artifacts
        )
        has_merged_md = dedup_enabled and any(
            a.get("kind") == "merged_markdown" for a in artifacts
        )
        has_mshelp_merged = dedup_enabled and any(
            a.get("kind") == "mshelp_merged_markdown" for a in artifacts
        )

        # Collect paths of individual MSHelp markdowns (they are in _AI/MSHelp/ but NOT in Merged/)
        _mshelp_source_paths = set()
        if has_mshelp_merged:
            for rec in getattr(self, "mshelp_records", []) or []:
                mdp = rec.get("markdown_path", "")
                if mdp:
                    _mshelp_source_paths.add(os.path.normcase(os.path.abspath(mdp)))

        for art in artifacts:
            kind = art.get("kind", "")
            rel = art.get("path_rel_to_target") or ""
            abs_path = art.get("path_abs") or ""
            if not rel or not abs_path:
                continue

            is_content = kind in _LLM_CONTENT_KINDS
            is_pdf = kind in _LLM_PDF_KINDS

            if not (is_content or (include_pdf and is_pdf)):
                continue

            # Skip individual converted PDFs when merged PDFs exist
            if has_merged_pdf and kind == "converted_pdf":
                continue

            # Skip individual markdown exports when merged markdown packages exist
            if has_merged_md and kind == "markdown_export":
                continue

            # Skip individual MSHelp markdown sources when merged packages exist
            if has_mshelp_merged and kind == "markdown_export":
                norm = os.path.normcase(os.path.abspath(abs_path))
                if norm in _mshelp_source_paths:
                    continue

            ext = os.path.splitext(rel.lower())[1]
            if ext == ".md":
                counts["markdown"] += 1
                category = "Markdown"
            elif ext in (".json", ".jsonl"):
                counts["json"] += 1
                category = "JSON"
            elif ext == ".pdf":
                counts["pdf"] += 1
                category = "PDF"
            else:
                counts["other"] += 1
                category = "Files"

            if flatten:
                # Flat: use clean original basename, collision-safe
                base_name = os.path.basename(rel)
                hub_rel = base_name
                hub_rel_base, hub_rel_ext = os.path.splitext(hub_rel)
                candidate = os.path.join(llm_root, hub_rel)
                collision_idx = 1
                while os.path.exists(candidate):
                    hub_rel = f"{hub_rel_base}_{collision_idx}{hub_rel_ext}"
                    candidate = os.path.join(llm_root, hub_rel)
                    collision_idx += 1
            else:
                hub_rel = os.path.join(category, rel)

            hub_abs = os.path.join(llm_root, hub_rel)
            hub_dir = os.path.dirname(hub_abs)
            try:
                os.makedirs(hub_dir, exist_ok=True)
            except Exception as e:
                logging.error(f"failed to create LLM hub subdir {hub_dir}: {e}")
                continue

            try:
                shutil.copy2(abs_path, hub_abs)
            except Exception as e:
                logging.error(f"failed to copy to LLM hub {hub_abs}: {e}")
                continue

            try:
                size_bytes = os.path.getsize(hub_abs)
            except Exception:
                size_bytes = 0

            hub_items.append(
                {
                    "kind": kind,
                    "source_abs_path": abs_path,
                    "delivery_rel_path": hub_rel.replace("\\", "/"),
                    "size_bytes": int(size_bytes),
                    "md5": art.get("md5", ""),
                    "sha256": art.get("sha256", ""),
                }
            )

        if not hub_items:
            return None

        manifest = {
            "version": 1,
            "generated_at": datetime.now().isoformat(timespec="seconds"),
            "run_mode": self.run_mode,
            "source_folder": self.config.get("source_folder", ""),
            "target_folder": target_folder,
            "hub_root": llm_root,
            "items": hub_items,
            "summary": counts,
        }

        manifest_path = None
        if self.config.get("enable_upload_json_manifest", True):
            manifest_path = os.path.join(llm_root, "llm_upload_manifest.json")
            try:
                with open(manifest_path, "w", encoding="utf-8") as f:
                    json.dump(manifest, f, ensure_ascii=False, indent=2)
            except Exception as e:
                logging.error(f"failed to write LLM hub manifest {manifest_path}: {e}")

        # ---- Generate readable text manifest (清单, gated by config) ----
        if self.config.get("enable_upload_readme", True):
            readme_path = os.path.join(llm_root, "README_UPLOAD_LIST.txt")
            try:
                total_size = sum(it["size_bytes"] for it in hub_items)
                readme_lines = [
                    "=== LLM Upload File List / 上传文件清单 ===",
                    f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
                    f"Total files: {len(hub_items)}",
                    f"Total size: {total_size / 1024 / 1024:.1f} MB",
                    f"  Markdown: {counts['markdown']}  |  JSON: {counts['json']}  |  PDF: {counts['pdf']}",
                    "",
                    "--- File List ---",
                ]
                for idx, it in enumerate(hub_items, 1):
                    size_mb = it["size_bytes"] / 1024 / 1024
                    readme_lines.append(
                        f"  {idx:3d}. [{it['kind']}] {it['delivery_rel_path']}  ({size_mb:.2f} MB)"
                    )
                readme_lines.append("")
                readme_lines.append("--- Notes ---")
                if has_merged_pdf:
                    readme_lines.append(
                        "* Individual converted PDFs are excluded (already in merged volumes)."
                    )
                if has_merged_md:
                    readme_lines.append(
                        "* Individual markdown files are excluded (already in merged markdown packages)."
                    )
                if has_mshelp_merged:
                    readme_lines.append(
                        "* Individual MSHelp markdowns are excluded (already in merged packages)."
                    )
                readme_lines.append(
                    "* Metadata files (manifests, quality reports, index) are excluded."
                )
                readme_lines.append("")
                with open(readme_path, "w", encoding="utf-8") as f:
                    f.write("\n".join(readme_lines))
            except Exception as e:
                logging.warning(f"failed to write upload readme: {e}")

        logging.info(
            "LLM hub built at %s | files: %s (md=%s, json=%s, pdf=%s)",
            llm_root,
            len(hub_items),
            counts["markdown"],
            counts["json"],
            counts["pdf"],
        )

        # expose for GUI summary
        self.llm_hub_root = llm_root

        return {
            "kind": "llm_delivery_hub",
            "path_abs": llm_root,
            "path_rel_to_target": os.path.relpath(llm_root, target_folder).replace(
                "\\", "/"
            ),
            "size_bytes": 0,
            "mtime": datetime.now().isoformat(timespec="seconds"),
            "md5": "",
            "sha256": "",
            "manifest_path": manifest_path,
        }

    def _write_corpus_manifest(self, merge_outputs=None):
        if not self.config.get("enable_corpus_manifest", True):
            return None

        target_folder = self.config.get("target_folder", "")
        if not target_folder:
            return None
        os.makedirs(target_folder, exist_ok=True)

        artifacts = []
        seen = set()

        def _append(kind, path):
            if not path:
                return
            abs_path = os.path.abspath(path)
            key = (kind, abs_path)
            if key in seen:
                return
            self._add_artifact(artifacts, kind, abs_path)
            seen.add(key)

        for p in self.generated_pdfs:
            _append("converted_pdf", p)
        for p in merge_outputs or self.generated_merge_outputs or []:
            p_low = str(p).lower()
            if p_low.endswith(".md"):
                _append("merged_markdown", p)
            else:
                _append("merged_pdf", p)
        for p in self.generated_merge_markdown_outputs:
            _append("merged_markdown", p)
        for p in self.generated_map_outputs:
            if str(p).lower().endswith(".map.csv"):
                _append("merge_map_csv", p)
            elif str(p).lower().endswith(".map.json"):
                _append("merge_map_json", p)
            else:
                _append("merge_map_file", p)
        for p in self.generated_markdown_outputs:
            _append("markdown_export", p)
        for p in self.generated_markdown_quality_outputs:
            _append("markdown_quality_report", p)
        for p in self.generated_excel_json_outputs:
            _append("excel_structured_json", p)
        for p in self.generated_records_json_outputs:
            _append("records_json", p)
        for p in self.generated_chromadb_outputs:
            p_low = str(p).lower()
            if p_low.endswith(".jsonl"):
                _append("chromadb_docs_jsonl", p)
            elif p_low.endswith(".json"):
                _append("chromadb_export_manifest", p)
            else:
                _append("chromadb_export_file", p)
        for p in self.generated_update_package_outputs:
            p_low = str(p).lower()
            if p_low.endswith("incremental_manifest.json"):
                _append("update_package_manifest", p)
            elif p_low.endswith("incremental_index.xlsx"):
                _append("update_package_index_xlsx", p)
            elif p_low.endswith("incremental_index.csv"):
                _append("update_package_index_csv", p)
            elif p_low.endswith("incremental_index.json"):
                _append("update_package_index_json", p)
            else:
                _append("update_package_file", p)
        for p in self.generated_mshelp_outputs:
            p_low = str(p).lower()
            if p_low.endswith(".json") and "mshelp_index_" in p_low:
                _append("mshelp_index_json", p)
            elif p_low.endswith(".csv") and "mshelp_index_" in p_low:
                _append("mshelp_index_csv", p)
            elif p_low.endswith(".docx"):
                _append("mshelp_merged_docx", p)
            elif p_low.endswith(".pdf"):
                _append("mshelp_merged_pdf", p)
            elif p_low.endswith(".md"):
                _append("mshelp_merged_markdown", p)
            else:
                _append("mshelp_output_file", p)

        _append("convert_index_excel", self.convert_index_path)
        _append("collect_index_excel", self.collect_index_path)
        _append("merge_list_excel", self.merge_excel_path)

        manifest = {
            "version": 1,
            "generated_at": datetime.now().isoformat(timespec="seconds"),
            "run_mode": self.run_mode,
            "collect_mode": self.collect_mode,
            "merge_mode": self.merge_mode,
            "content_strategy": self.content_strategy,
            "source_folder": self.config.get("source_folder", ""),
            "target_folder": target_folder,
            "artifacts": artifacts,
            "conversion_records": self.conversion_index_records,
            "merge_records": self.merge_index_records,
            "summary": {
                "converted_pdf_count": len(self.generated_pdfs),
                "merged_pdf_count": len(
                    [
                        p
                        for p in (merge_outputs or self.generated_merge_outputs or [])
                        if not str(p).lower().endswith(".md")
                    ]
                ),
                "merged_markdown_count": len(self.generated_merge_markdown_outputs),
                "merge_map_count": len(self.generated_map_outputs),
                "markdown_count": len(self.generated_markdown_outputs),
                "markdown_quality_report_count": len(
                    self.generated_markdown_quality_outputs
                ),
                "excel_structured_json_count": len(self.generated_excel_json_outputs),
                "records_json_count": len(self.generated_records_json_outputs),
                "chromadb_export_file_count": len(self.generated_chromadb_outputs),
                "update_package_file_count": len(self.generated_update_package_outputs),
                "mshelp_output_file_count": len(self.generated_mshelp_outputs),
                "conversion_record_count": len(self.conversion_index_records),
                "merge_record_count": len(self.merge_index_records),
                "artifact_count": len(artifacts),
            },
        }

        # Optional: build LLM delivery hub on top of collected artifacts
        try:
            llm_hub_meta = self._maybe_build_llm_delivery_hub(target_folder, artifacts)
            if llm_hub_meta:
                artifacts.append(llm_hub_meta)
                manifest["summary"]["artifact_count"] = len(artifacts)
        except Exception as e:
            logging.error(f"failed to build LLM delivery hub: {e}")

        manifest_path = os.path.join(target_folder, "corpus.json")
        with open(manifest_path, "w", encoding="utf-8") as f:
            json.dump(manifest, f, ensure_ascii=False, indent=2)

        self.corpus_manifest_path = manifest_path
        print(f"\nCorpus manifest generated: {manifest_path}")
        logging.info(f"Corpus manifest generated: {manifest_path}")
        return manifest_path

    def collect_office_files_and_build_excel(self):
        if not HAS_OPENPYXL:
            print("\n[ERROR] openpyxl not found. Cannot generate Excel report.")
            print("Run pip install openpyxl and retry.")
            logging.error("openpyxl missing; collect_only mode cannot continue.")
            return

        source_roots = self._get_source_roots()
        if not source_roots:
            print("[WARN] No source folder(s) to scan. collect_only skipped.")
            return
        target_root = self.config["target_folder"]
        os.makedirs(target_root, exist_ok=True)

        exts_word = self.config["allowed_extensions"].get("word", [])
        exts_excel = self.config["allowed_extensions"].get("excel", [])
        exts_ppt = self.config["allowed_extensions"].get("powerpoint", [])
        office_exts = set(exts_word + exts_excel + exts_ppt)

        excl_config = self.config.get("excluded_folders", [])
        excl_names = {
            x.lower()
            for x in excl_config
            if not os.path.isabs(x) and os.sep not in x and "/" not in x
        }
        excl_paths = {
            os.path.abspath(x).lower()
            for x in excl_config
            if os.path.isabs(x) or os.sep in x or "/" in x
        }

        print("\n" + "=" * 60)
        print(" File collection & dedup mode")
        print("=" * 60)
        print(f" Source dir(s) : {len(source_roots)} folder(s)")
        print(f" Target dir : {target_root}")
        print(f" Sub mode   : {self.get_readable_collect_mode()} ({self.collect_mode})")
        print(f" Filter ext : {office_exts}")
        print("=" * 60)

        all_files = []
        for source_root in source_roots:
            if not os.path.isdir(source_root):
                continue
            for root, dirs, files in os.walk(source_root):
                dirs[:] = [
                    d
                    for d in dirs
                    if d.lower() not in excl_names
                    and os.path.abspath(os.path.join(root, d)).lower() not in excl_paths
                ]
                for name in files:
                    if name.startswith("~$"):
                        continue
                    ext = os.path.splitext(name)[1].lower()
                    if ext in office_exts:
                        full_path = os.path.join(root, name)
                        try:
                            size = os.path.getsize(full_path)
                        except OSError:
                            continue
                        all_files.append((full_path, size, ext))

        total = len(all_files)
        print(f"Scanned Office files: {total}")
        logging.info(f"[collect_only] scanned Office files: {total}")

        if total == 0:
            print("[INFO] No Office files found. collect_only finished.")
            return

        size_groups = {}
        for path, size, ext in all_files:
            size_groups.setdefault(size, []).append((path, ext))

        unique_records = []
        duplicate_records = []
        group_id_counter = 1

        for size, files in size_groups.items():
            if not self.is_running:
                break
            if len(files) == 1:
                src_path, ext = files[0]
                rel = os.path.relpath(
                    src_path, self._get_source_root_for_path(src_path)
                )
                dst_path = os.path.join(target_root, rel)
                unique_records.append(
                    {
                        "group_id": None,
                        "src": src_path,
                        "dst": dst_path,
                        "size": size,
                        "ext": ext,
                    }
                )
                continue

            hash_groups = {}
            for src_path, ext in files:
                file_hash = self._compute_file_hash(src_path)
                hash_groups.setdefault(file_hash, []).append((src_path, ext))

            for file_hash, same_hash_files in hash_groups.items():
                if len(same_hash_files) == 1:
                    src_path, ext = same_hash_files[0]
                    rel = os.path.relpath(
                        src_path, self._get_source_root_for_path(src_path)
                    )
                    dst_path = os.path.join(target_root, rel)
                    unique_records.append(
                        {
                            "group_id": None,
                            "src": src_path,
                            "dst": dst_path,
                            "size": size,
                            "ext": ext,
                        }
                    )
                else:
                    group_id = f"G{group_id_counter}"
                    group_id_counter += 1

                    keep_src, keep_ext = same_hash_files[0]
                    keep_rel = os.path.relpath(
                        keep_src, self._get_source_root_for_path(keep_src)
                    )
                    keep_dst = os.path.join(target_root, keep_rel)

                    unique_records.append(
                        {
                            "group_id": group_id,
                            "src": keep_src,
                            "dst": keep_dst,
                            "size": size,
                            "ext": keep_ext,
                        }
                    )

                    for dup_src, dup_ext in same_hash_files[1:]:
                        duplicate_records.append(
                            {
                                "group_id": group_id,
                                "src": dup_src,
                                "size": size,
                                "ext": dup_ext,
                                "keep_src": keep_src,
                                "keep_dst": keep_dst,
                            }
                        )

        print(f"\nDedup completed:")
        print(f"  Unique files    : {len(unique_records)}")
        print(f"  Duplicate files : {len(duplicate_records)}")
        logging.info(
            f"[collect_only] unique={len(unique_records)}, duplicate={len(duplicate_records)}"
        )

        copied_count = 0
        if self.collect_mode == COLLECT_MODE_COPY_AND_INDEX:
            print("\nCopying unique files to target directory...")
            for idx, rec in enumerate(unique_records, 1):
                if not self.is_running:
                    break
                src = rec["src"]
                dst = rec["dst"]
                dst_dir = os.path.dirname(dst)
                os.makedirs(dst_dir, exist_ok=True)
                try:
                    if not os.path.exists(dst):
                        shutil.copy2(src, dst)
                    rec["copied"] = True
                    copied_count += 1
                except Exception as e:
                    logging.error(f"[collect_only] copy failed: {src} -> {dst} | {e}")
                    rec["copied"] = False

                if idx % 20 == 0 or idx == len(unique_records):
                    print(
                        f"\rProcessed {idx}/{len(unique_records)} unique files...",
                        end="",
                        flush=True,
                    )
            print(f"\rCopy finished, copied {copied_count} files.         ")
        else:
            print("\nCurrent mode is [index_only]; skip file copy.")
            for rec in unique_records:
                rec["copied"] = False

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        excel_path = os.path.join(target_root, f"office_index_{timestamp}.xlsx")

        wb = Workbook()
        ws_unique = wb.active
        ws_unique.title = "UniqueFiles"
        ws_dup = wb.create_sheet("Duplicates")

        if self.collect_mode == COLLECT_MODE_COPY_AND_INDEX:
            headers_unique = [
                "No.",
                "GroupID",
                "FileName",
                "Ext",
                "Size(KB)",
                "SourcePath",
                "TargetPath",
            ]
        else:
            headers_unique = [
                "No.",
                "GroupID",
                "FileName",
                "Ext",
                "Size(KB)",
                "SourcePath",
            ]

        ws_unique.append(headers_unique)
        for cell in ws_unique[1]:
            cell.font = Font(bold=True)

        for idx, rec in enumerate(unique_records, 1):
            src = rec["src"]
            dst = rec["dst"]
            size_kb = round(rec["size"] / 1024, 2)
            group_id = rec["group_id"] or ""
            file_name = os.path.basename(src)
            ext = rec["ext"]

            if self.collect_mode == COLLECT_MODE_COPY_AND_INDEX:
                row = [idx, group_id, file_name, ext, size_kb, src, dst]
                ws_unique.append(row)
                dst_cell = ws_unique.cell(row=idx + 1, column=7)
                dst_cell.hyperlink = self._make_file_hyperlink(dst)
                dst_cell.style = "Hyperlink"
            else:
                row = [idx, group_id, file_name, ext, size_kb, src]
                ws_unique.append(row)
                src_cell = ws_unique.cell(row=idx + 1, column=6)
                src_cell.hyperlink = self._make_file_hyperlink(src)
                src_cell.style = "Hyperlink"

        if self.collect_mode == COLLECT_MODE_COPY_AND_INDEX:
            headers_dup = [
                "No.",
                "GroupID",
                "FileName",
                "Ext",
                "Size(KB)",
                "SourcePath",
                "KeptTargetPath",
            ]
        else:
            headers_dup = [
                "No.",
                "GroupID",
                "FileName",
                "Ext",
                "Size(KB)",
                "SourcePath",
                "KeptSourcePath",
            ]

        ws_dup.append(headers_dup)
        for cell in ws_dup[1]:
            cell.font = Font(bold=True)

        for idx, rec in enumerate(duplicate_records, 1):
            src = rec["src"]
            keep_src = rec["keep_src"]
            keep_dst = rec["keep_dst"]
            size_kb = round(rec["size"] / 1024, 2)
            group_id = rec["group_id"]
            file_name = os.path.basename(src)
            ext = rec["ext"]

            if self.collect_mode == COLLECT_MODE_COPY_AND_INDEX:
                row = [idx, group_id, file_name, ext, size_kb, src, keep_dst]
                ws_dup.append(row)
                src_cell = ws_dup.cell(row=idx + 1, column=6)
                src_cell.hyperlink = self._make_file_hyperlink(src)
                src_cell.style = "Hyperlink"
            else:
                row = [idx, group_id, file_name, ext, size_kb, src, keep_src]
                ws_dup.append(row)
                src_cell = ws_dup.cell(row=idx + 1, column=6)
                src_cell.hyperlink = self._make_file_hyperlink(src)
                src_cell.style = "Hyperlink"
                keep_cell = ws_dup.cell(row=idx + 1, column=7)
                keep_cell.hyperlink = self._make_file_hyperlink(keep_src)
                keep_cell.style = "Hyperlink"

        for ws in (ws_unique, ws_dup):
            for col in ws.columns:
                max_length = 0
                col_letter = col[0].column_letter
                for cell in col:
                    try:
                        v = str(cell.value) if cell.value is not None else ""
                        max_length = max(max_length, len(v))
                    except Exception:
                        pass
                ws.column_dimensions[col_letter].width = min(max_length + 2, 80)

        wb.save(excel_path)
        self.collect_index_path = excel_path
        print(f"\nExcel index generated: {excel_path}")
        logging.info(f"[collect_only] Excel index generated: {excel_path}")

        print("\n=== collect_only summary ===")
        print(f"scanned          : {total}")
        print(
            f"unique files     : {len(unique_records)} (copied {copied_count}; valid in copy mode)"
        )
        print(f"duplicate files  : {len(duplicate_records)}")
        print(f"index file       : {excel_path}")
        print("========================\n")

    # =============== Main flow ===============

    def _get_source_roots(self):
        """Return list of source folder paths to scan (multi-folder or single)."""
        roots = self.config.get("source_folders")
        if isinstance(roots, list) and roots:
            roots = [os.path.abspath(str(p).strip()) for p in roots if str(p).strip()]
            roots = [r for r in roots if os.path.isdir(r)]
            if roots:
                return roots
        single = self.config.get("source_folder", "")
        if single and os.path.isdir(single):
            return [os.path.abspath(single)]
        return []

    def _get_source_root_for_path(self, abs_path):
        """Return the source root that contains abs_path (for relpath). Prefer longest match."""
        abs_path = os.path.abspath(abs_path)
        roots = self._get_source_roots()
        if not roots:
            return self.config.get("source_folder", "") or ""
        match = ""
        for r in roots:
            r = os.path.abspath(r)
            if (
                abs_path == r
                or abs_path.startswith(r + os.sep)
                or abs_path.startswith(r + "/")
            ):
                if len(r) > len(match):
                    match = r
        return match or (roots[0] if roots else self.config.get("source_folder", ""))

    def _scan_convert_candidates(self):
        files = []
        source_roots = self._get_source_roots()
        if not source_roots:
            single = self.config.get("source_folder", "")
            print(f"\n[WARN] Source directory does not exist or empty: {single}")
            logging.error("source directory does not exist or source_folders empty")
            return files

        excl_config = self.config.get("excluded_folders", [])
        excl_names = {
            x.lower()
            for x in excl_config
            if not os.path.isabs(x) and os.sep not in x and "/" not in x
        }
        excl_paths = {
            os.path.abspath(x).lower()
            for x in excl_config
            if os.path.isabs(x) or os.sep in x or "/" in x
        }

        valid_exts = set()
        for sub in self.config.get("allowed_extensions", {}).values():
            if isinstance(sub, list):
                for e in sub:
                    if isinstance(e, str) and e:
                        valid_exts.add(e.lower())

        for source_folder in source_roots:
            if not os.path.isdir(source_folder):
                continue
            for root, dirs, filenames in os.walk(source_folder):
                dirs[:] = [
                    d
                    for d in dirs
                    if d.lower() not in excl_names
                    and os.path.abspath(os.path.join(root, d)).lower() not in excl_paths
                ]
                for fname in filenames:
                    if fname.startswith("~$"):
                        continue
                    ext = os.path.splitext(fname)[1].lower()
                    if ext not in valid_exts:
                        continue

                    full_path = os.path.join(root, fname)
                    if self.filter_date:
                        try:
                            ctime = os.path.getctime(full_path)
                            file_date = datetime.fromtimestamp(ctime).date()
                            filter_d = self.filter_date.date()
                            if self.filter_mode == "after" and file_date < filter_d:
                                continue
                            if self.filter_mode == "before" and file_date > filter_d:
                                continue
                        except Exception:
                            pass

                    files.append(full_path)

        return files

    def _apply_source_priority_filter(self, files):
        if not self.config.get("source_priority_skip_same_name_pdf", False):
            return files, []

        office_exts = set()
        office_exts.update(
            e.lower() for e in self.config.get("allowed_extensions", {}).get("word", [])
        )
        office_exts.update(
            e.lower()
            for e in self.config.get("allowed_extensions", {}).get("excel", [])
        )
        office_exts.update(
            e.lower()
            for e in self.config.get("allowed_extensions", {}).get("powerpoint", [])
        )

        office_keys = {}
        for p in files:
            ext = os.path.splitext(p)[1].lower()
            if ext not in office_exts:
                continue
            parent = os.path.abspath(os.path.dirname(p))
            if is_win():
                parent = parent.lower()
            stem = os.path.splitext(os.path.basename(p))[0].lower()
            office_keys[(parent, stem)] = p

        kept = []
        skipped = []
        for p in files:
            ext = os.path.splitext(p)[1].lower()
            if ext != ".pdf":
                kept.append(p)
                continue

            parent = os.path.abspath(os.path.dirname(p))
            if is_win():
                parent = parent.lower()
            stem = os.path.splitext(os.path.basename(p))[0].lower()
            office_path = office_keys.get((parent, stem))
            if office_path:
                skipped.append(
                    {
                        "source_path": os.path.abspath(p),
                        "status": "source_priority_skipped",
                        "detail": "same_dir_same_stem_office_exists",
                        "final_path": "",
                        "preferred_source": os.path.abspath(office_path),
                    }
                )
                continue

            kept.append(p)

        if skipped:
            logging.info(
                f"[source_priority] 跳过同目录同名 PDF: {len(skipped)}，保留 Office 版本"
            )
        return kept, skipped

    def _ext_bucket(self, path):
        ext = os.path.splitext(path)[1].lower()
        if ext in [
            e.lower() for e in self.config.get("allowed_extensions", {}).get("word", [])
        ]:
            return "word"
        if ext in [
            e.lower()
            for e in self.config.get("allowed_extensions", {}).get("excel", [])
        ]:
            return "excel"
        if ext in [
            e.lower()
            for e in self.config.get("allowed_extensions", {}).get("powerpoint", [])
        ]:
            return "powerpoint"
        if ext == ".pdf":
            return "pdf"
        return ext or "unknown"

    def _apply_global_md5_dedup(self, files):
        if not self.config.get("global_md5_dedup", False):
            return files, []

        seen = {}
        kept = []
        skipped = []

        for path in files:
            bucket = self._ext_bucket(path)
            try:
                md5_value = self._compute_md5(path)
            except Exception as e:
                logging.warning(
                    f"[global_md5] failed to compute MD5, keep file: {path} | {e}"
                )
                kept.append(path)
                continue

            key = (bucket, md5_value)
            if key in seen:
                skipped.append(
                    {
                        "source_path": os.path.abspath(path),
                        "status": "dedup_skipped",
                        "detail": "same_type_same_md5",
                        "final_path": "",
                        "md5": md5_value,
                        "duplicate_of": os.path.abspath(seen[key]),
                    }
                )
                continue

            seen[key] = path
            kept.append(path)

        if skipped:
            logging.info(f"[global_md5] skipped duplicate source files: {len(skipped)}")
        return kept, skipped

    def _resolve_incremental_registry_path(self):
        configured = str(self.config.get("incremental_registry_path", "") or "").strip()
        if configured:
            if not os.path.isabs(configured):
                return os.path.abspath(
                    os.path.join(self.config.get("target_folder", ""), configured)
                )
            return configured
        return os.path.join(
            self.config.get("target_folder", ""),
            "_AI",
            "registry",
            "incremental_registry.json",
        )

    def _build_source_meta(self, path, include_hash=False):
        abs_path = os.path.abspath(path)
        try:
            stat = os.stat(abs_path)
        except OSError:
            return None

        source_hash = ""
        if include_hash:
            try:
                source_hash = self._compute_file_hash(abs_path)
            except Exception as e:
                logging.warning(
                    f"[incremental] failed to compute source hash: {abs_path} | {e}"
                )

        return {
            "source_path": abs_path,
            "ext": os.path.splitext(abs_path)[1].lower(),
            "source_size": int(stat.st_size),
            "source_mtime_ns": int(
                getattr(stat, "st_mtime_ns", int(stat.st_mtime * 1_000_000_000))
            ),
            "source_mtime": datetime.fromtimestamp(stat.st_mtime).isoformat(
                timespec="seconds"
            ),
            "source_hash_sha256": source_hash,
        }

    def _apply_incremental_filter(self, files):
        context = {
            "enabled": False,
            "registry": None,
            "registry_path": "",
            "scan_meta": {},
            "scanned_count": len(files),
            "added_count": 0,
            "modified_count": 0,
            "renamed_count": 0,
            "unchanged_count": 0,
            "deleted_count": 0,
            "deleted_paths": [],
            "renamed_pairs": [],
            "reprocess_renamed": False,
        }

        if not self.config.get("enable_incremental_mode", False):
            self._incremental_context = None
            self.incremental_registry_path = ""
            return files, context

        verify_hash = bool(self.config.get("incremental_verify_hash", False))
        reprocess_renamed = bool(
            self.config.get("incremental_reprocess_renamed", False)
        )
        registry_path = self._resolve_incremental_registry_path()
        registry = FileRegistry(
            registry_path, base_root=self.config.get("source_folder", "")
        )
        registry.load()

        process_files = []
        scan_meta = {}
        added = 0
        modified = 0
        renamed = 0
        unchanged = 0

        for path in files:
            meta = self._build_source_meta(path, include_hash=verify_hash)
            if not meta:
                continue

            prev = registry.get(path)
            state = "added"
            if isinstance(prev, dict):
                prev_size = int(prev.get("source_size", -1))
                prev_mtime_ns = int(prev.get("source_mtime_ns", -1))
                same_size = prev_size == meta["source_size"]
                same_mtime = prev_mtime_ns == meta["source_mtime_ns"]
                if verify_hash:
                    prev_hash = str(prev.get("source_hash_sha256", "") or "")
                    curr_hash = meta.get("source_hash_sha256", "")
                    same_hash = bool(prev_hash and curr_hash and prev_hash == curr_hash)
                else:
                    same_hash = True

                if same_size and same_mtime and same_hash:
                    state = "unchanged"
                else:
                    state = "modified"

            meta["change_state"] = state
            scan_meta[path] = meta

            if state == "unchanged":
                unchanged += 1
            elif state == "added":
                added += 1
                process_files.append(path)
            else:
                modified += 1
                process_files.append(path)

        current_keys = {registry.normalize_path(p) for p in scan_meta.keys()}
        old_keys = set(registry.keys())
        deleted_keys = sorted(old_keys - current_keys)
        deleted_set = set(deleted_keys)

        # Rename detection: match current "added" files to previous deleted entries.
        renamed_pairs = []
        deleted_entry_map = {}
        if isinstance(registry.entries, dict):
            for key in deleted_keys:
                entry = registry.entries.get(key)
                if isinstance(entry, dict):
                    deleted_entry_map[key] = entry

        added_paths = [
            p for p, m in scan_meta.items() if str(m.get("change_state", "")) == "added"
        ]

        for src_path in sorted(added_paths):
            meta = scan_meta.get(src_path, {})
            ext = str(meta.get("ext", "") or "")
            size = int(meta.get("source_size", -1))
            mtime_ns = int(meta.get("source_mtime_ns", -1))

            candidates = []
            for old_key in list(deleted_set):
                old_entry = deleted_entry_map.get(old_key) or {}
                if str(old_entry.get("ext", "")) != ext:
                    continue
                if int(old_entry.get("source_size", -2)) != size:
                    continue
                candidates.append((old_key, old_entry))

            if not candidates:
                continue

            curr_hash = str(meta.get("source_hash_sha256", "") or "")
            if not curr_hash:
                try:
                    curr_hash = self._compute_file_hash(src_path)
                    meta["source_hash_sha256"] = curr_hash
                except Exception:
                    curr_hash = ""

            matched = None
            if curr_hash:
                for old_key, old_entry in candidates:
                    old_hash = str(old_entry.get("source_hash_sha256", "") or "")
                    if old_hash and old_hash == curr_hash:
                        matched = (old_key, old_entry, "hash")
                        break

            if matched is None:
                for old_key, old_entry in candidates:
                    old_mtime_ns = int(old_entry.get("source_mtime_ns", -1))
                    if old_mtime_ns == mtime_ns and old_mtime_ns >= 0:
                        matched = (old_key, old_entry, "mtime")
                        break

            if matched is None and len(candidates) == 1:
                old_key, old_entry = candidates[0]
                matched = (old_key, old_entry, "ext_size_unique")

            if matched is None:
                continue

            old_key, old_entry, match_type = matched
            old_path = str(old_entry.get("source_path", "") or old_key)

            meta["change_state"] = "renamed"
            meta["renamed_from"] = old_path
            meta["renamed_from_key"] = old_key
            meta["rename_match_type"] = match_type

            deleted_set.discard(old_key)
            deleted_entry_map.pop(old_key, None)

            renamed += 1
            added = max(0, added - 1)

            renamed_pairs.append(
                {
                    "from_path": old_path,
                    "to_path": os.path.abspath(src_path),
                    "match_type": match_type,
                }
            )

            if not reprocess_renamed and src_path in process_files:
                process_files.remove(src_path)

        deleted_keys = sorted(deleted_set)

        context = {
            "enabled": True,
            "registry": registry,
            "registry_path": registry_path,
            "scan_meta": scan_meta,
            "scanned_count": len(scan_meta),
            "added_count": added,
            "modified_count": modified,
            "renamed_count": renamed,
            "unchanged_count": unchanged,
            "deleted_count": len(deleted_keys),
            "deleted_paths": deleted_keys,
            "renamed_pairs": renamed_pairs,
            "reprocess_renamed": reprocess_renamed,
        }

        self._incremental_context = context
        self.incremental_registry_path = registry_path
        logging.info(
            "[incremental] 扫描完成: scanned=%s added=%s modified=%s renamed=%s unchanged=%s deleted=%s",
            context["scanned_count"],
            added,
            modified,
            renamed,
            unchanged,
            context["deleted_count"],
        )
        return process_files, context

    def _flush_incremental_registry(self, process_results):
        context = self._incremental_context or {}
        if not context.get("enabled"):
            return

        registry = context.get("registry")
        if not registry:
            return

        result_map = {}
        for item in process_results or []:
            src = item.get("source_path")
            if not src:
                continue
            result_map[registry.normalize_path(src)] = item

        now_iso = datetime.now().isoformat(timespec="seconds")
        new_entries = {}

        for src_path, meta in (context.get("scan_meta") or {}).items():
            key = registry.normalize_path(src_path)
            prev = (
                registry.entries.get(key, {})
                if isinstance(registry.entries, dict)
                else {}
            )
            rename_from_key = str(meta.get("renamed_from_key", "") or "")
            if (not prev) and rename_from_key and isinstance(registry.entries, dict):
                prev = registry.entries.get(rename_from_key, {})
            entry = dict(prev) if isinstance(prev, dict) else {}

            entry.update(
                {
                    "source_path": os.path.abspath(src_path),
                    "ext": meta.get("ext", ""),
                    "source_size": meta.get("source_size", 0),
                    "source_mtime": meta.get("source_mtime", ""),
                    "source_mtime_ns": meta.get("source_mtime_ns", 0),
                    "source_hash_sha256": meta.get("source_hash_sha256", ""),
                    "change_state": meta.get("change_state", ""),
                    "renamed_from": meta.get("renamed_from", ""),
                    "rename_match_type": meta.get("rename_match_type", ""),
                    "last_seen_at": now_iso,
                    "last_run_mode": self.run_mode,
                }
            )

            result = result_map.get(key)
            if result:
                entry["last_status"] = result.get("status", "")
                entry["last_error"] = result.get("error", "")
                entry["last_processed_at"] = now_iso
                final_path = result.get("final_path", "")
                if final_path and os.path.exists(final_path):
                    entry["last_output_pdf"] = os.path.abspath(final_path)
                    try:
                        entry["last_output_pdf_md5"] = self._compute_md5(final_path)
                    except Exception:
                        pass
            elif meta.get("change_state") == "unchanged":
                entry["last_status"] = entry.get("last_status") or "unchanged"
            elif meta.get("change_state") == "renamed":
                entry["last_status"] = "renamed_detected"
                entry["last_processed_at"] = now_iso

            new_entries[key] = entry

        registry.entries = new_entries
        run_summary = {
            "scanned_count": context.get("scanned_count", 0),
            "added_count": context.get("added_count", 0),
            "modified_count": context.get("modified_count", 0),
            "renamed_count": context.get("renamed_count", 0),
            "unchanged_count": context.get("unchanged_count", 0),
            "deleted_count": context.get("deleted_count", 0),
            "processed_result_count": len(result_map),
        }
        registry.save(run_summary=run_summary)
        logging.info(
            f"[incremental] registry updated: {context.get('registry_path', '')}"
        )

    def _resolve_update_package_root(self):
        configured = str(self.config.get("update_package_root", "") or "").strip()
        if configured:
            if not os.path.isabs(configured):
                return os.path.abspath(
                    os.path.join(self.config.get("target_folder", ""), configured)
                )
            return configured
        return os.path.join(
            self.config.get("target_folder", ""), "_AI", "Update_Package"
        )

    def _write_update_package_index_xlsx(self, xlsx_path, records):
        if not HAS_OPENPYXL:
            return None
        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "IncrementalIndex"
            headers = [
                "seq",
                "change_state",
                "process_status",
                "source_file",
                "source_path",
                "source_md5",
                "source_sha256",
                "renamed_from",
                "rename_match_type",
                "packaged_pdf",
                "packaged_pdf_path",
                "packaged_pdf_md5",
                "note",
            ]
            ws.append(headers)
            for cell in ws[1]:
                cell.font = Font(bold=True)

            for rec in records:
                ws.append(
                    [
                        rec.get("seq", 0),
                        rec.get("change_state", ""),
                        rec.get("process_status", ""),
                        rec.get("source_file", ""),
                        rec.get("source_path", ""),
                        rec.get("source_md5", ""),
                        rec.get("source_sha256", ""),
                        rec.get("renamed_from", ""),
                        rec.get("rename_match_type", ""),
                        rec.get("packaged_pdf", ""),
                        rec.get("packaged_pdf_path", ""),
                        rec.get("packaged_pdf_md5", ""),
                        rec.get("note", ""),
                    ]
                )
                row_idx = ws.max_row
                src_cell = ws.cell(row=row_idx, column=5)
                if rec.get("source_path"):
                    src_cell.hyperlink = self._make_file_hyperlink(rec["source_path"])
                    src_cell.style = "Hyperlink"
                renamed_cell = ws.cell(row=row_idx, column=8)
                if rec.get("renamed_from"):
                    renamed_cell.hyperlink = self._make_file_hyperlink(
                        rec["renamed_from"]
                    )
                    renamed_cell.style = "Hyperlink"
                pdf_cell = ws.cell(row=row_idx, column=11)
                if rec.get("packaged_pdf_path"):
                    pdf_cell.hyperlink = self._make_file_hyperlink(
                        rec["packaged_pdf_path"]
                    )
                    pdf_cell.style = "Hyperlink"

            for col in ws.columns:
                col_letter = col[0].column_letter
                max_len = 0
                for cell in col:
                    value = "" if cell.value is None else str(cell.value)
                    if len(value) > max_len:
                        max_len = len(value)
                ws.column_dimensions[col_letter].width = min(max_len + 2, 80)

            wb.save(xlsx_path)
            return xlsx_path
        except Exception as e:
            logging.error(f"[update_package] failed to write XLSX index: {e}")
            return None

    def _generate_update_package(self, process_results):
        context = self._incremental_context or {}
        if not context.get("enabled"):
            return None
        if not self.config.get("enable_update_package", True):
            return None
        registry = context.get("registry")
        if registry is None:
            registry = FileRegistry(
                context.get("registry_path", ""),
                base_root=self.config.get("source_folder", ""),
            )

        scan_meta = context.get("scan_meta") or {}
        changed_sources = []
        for src_path, meta in scan_meta.items():
            state = str(meta.get("change_state", ""))
            if state in ("added", "modified", "renamed"):
                changed_sources.append(src_path)
        if not changed_sources:
            return None

        result_map = {}
        for item in process_results or []:
            src = item.get("source_path")
            if not src:
                continue
            result_map[registry.normalize_path(src)] = item

        package_root = self._resolve_update_package_root()
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        package_dir = os.path.join(package_root, f"Update_Package_{timestamp}")
        pdf_dir = os.path.join(package_dir, "PDF")
        os.makedirs(pdf_dir, exist_ok=True)

        records = []
        packaged_pdfs = []
        used_pdf_names = {}

        for idx, src_path in enumerate(sorted(changed_sources), 1):
            norm_key = registry.normalize_path(src_path)
            meta = scan_meta.get(src_path, {})
            result = result_map.get(norm_key, {})

            status = str(result.get("status", "pending"))
            detail = str(result.get("detail", "") or "")
            final_path = str(result.get("final_path", "") or "")
            if status == "pending" and str(meta.get("change_state", "")) == "renamed":
                status = "renamed_detected"
                if not detail:
                    detail = "rename detected; no reconvert"
            source_md5 = ""
            try:
                source_md5 = self._compute_md5(src_path)
            except Exception:
                pass

            packaged_pdf = ""
            packaged_pdf_path = ""
            packaged_pdf_md5 = ""
            if status == "success" and final_path and os.path.exists(final_path):
                base_name = os.path.basename(final_path)
                stem, ext = os.path.splitext(base_name)
                count = used_pdf_names.get(base_name, 0)
                if count > 0:
                    base_name = f"{stem}_{count}{ext}"
                used_pdf_names[os.path.basename(final_path)] = count + 1
                packaged_pdf_path = os.path.join(pdf_dir, base_name)
                try:
                    shutil.copy2(final_path, packaged_pdf_path)
                    packaged_pdf = base_name
                    packaged_pdfs.append(packaged_pdf_path)
                    packaged_pdf_md5 = self._compute_md5(packaged_pdf_path)
                except Exception as e:
                    detail = (
                        f"{detail}; copy_failed={e}" if detail else f"copy_failed={e}"
                    )

            record = {
                "seq": idx,
                "change_state": meta.get("change_state", ""),
                "process_status": status,
                "source_file": os.path.basename(src_path),
                "source_path": os.path.abspath(src_path),
                "source_md5": source_md5,
                "source_sha256": meta.get("source_hash_sha256", ""),
                "renamed_from": meta.get("renamed_from", ""),
                "rename_match_type": meta.get("rename_match_type", ""),
                "packaged_pdf": packaged_pdf,
                "packaged_pdf_path": os.path.abspath(packaged_pdf_path)
                if packaged_pdf_path
                else "",
                "packaged_pdf_md5": packaged_pdf_md5,
                "note": detail,
            }
            records.append(record)

        index_json = os.path.join(package_dir, "incremental_index.json")
        index_csv = os.path.join(package_dir, "incremental_index.csv")
        fields = [
            "seq",
            "change_state",
            "process_status",
            "source_file",
            "source_path",
            "source_md5",
            "source_sha256",
            "renamed_from",
            "rename_match_type",
            "packaged_pdf",
            "packaged_pdf_path",
            "packaged_pdf_md5",
            "note",
        ]

        with open(index_json, "w", encoding="utf-8") as f:
            json.dump(
                {"version": 1, "record_count": len(records), "records": records},
                f,
                ensure_ascii=False,
                indent=2,
            )

        with open(index_csv, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=fields)
            writer.writeheader()
            writer.writerows(records)

        index_xlsx = os.path.join(package_dir, "incremental_index.xlsx")
        xlsx_path = self._write_update_package_index_xlsx(index_xlsx, records)

        status_counts = {}
        for rec in records:
            key = rec.get("process_status", "unknown")
            status_counts[key] = status_counts.get(key, 0) + 1

        manifest = {
            "version": 1,
            "generated_at": datetime.now().isoformat(timespec="seconds"),
            "package_dir": os.path.abspath(package_dir),
            "run_mode": self.run_mode,
            "target_folder": self.config.get("target_folder", ""),
            "incremental_registry_path": context.get("registry_path", ""),
            "scan_summary": {
                "scanned_count": context.get("scanned_count", 0),
                "added_count": context.get("added_count", 0),
                "modified_count": context.get("modified_count", 0),
                "renamed_count": context.get("renamed_count", 0),
                "unchanged_count": context.get("unchanged_count", 0),
                "deleted_count": context.get("deleted_count", 0),
            },
            "deleted_paths": context.get("deleted_paths", []),
            "renamed_pairs": context.get("renamed_pairs", []),
            "record_count": len(records),
            "packaged_pdf_count": len(packaged_pdfs),
            "status_counts": status_counts,
            "records": records,
        }

        package_manifest = os.path.join(package_dir, "incremental_manifest.json")
        with open(package_manifest, "w", encoding="utf-8") as f:
            json.dump(manifest, f, ensure_ascii=False, indent=2)

        outputs = [package_manifest, index_json, index_csv]
        if xlsx_path:
            outputs.append(xlsx_path)
        outputs.extend(packaged_pdfs)

        self.generated_update_package_outputs = outputs
        self.update_package_manifest_path = package_manifest
        logging.info(f"[update_package] update package generated: {package_dir}")
        return package_manifest

    def _build_perf_summary(self):
        m = self.perf_metrics or {}
        lines = [
            "",
            "=== 性能统计 ===",
            f"扫描耗时: {m.get('scan_seconds', 0.0):.2f}s",
            f"转换主流程耗时: {m.get('batch_seconds', 0.0):.2f}s",
            f"  - Office/PDF核心耗时: {m.get('convert_core_seconds', 0.0):.2f}s",
            f"  - 等待PDF落盘耗时: {m.get('pdf_wait_seconds', 0.0):.2f}s",
            f"  - Markdown导出耗时: {m.get('markdown_seconds', 0.0):.2f}s",
            f"  - MSHelp合并耗时: {m.get('mshelp_merge_seconds', 0.0):.2f}s",
            f"合并耗时: {m.get('merge_seconds', 0.0):.2f}s",
            f"后处理耗时: {m.get('postprocess_seconds', 0.0):.2f}s",
            f"总耗时: {m.get('total_seconds', 0.0):.2f}s",
        ]
        success_count = self.stats.get("success", 0) or 0
        if success_count > 0:
            avg = m.get("total_seconds", 0.0) / success_count
            lines.append(f"平均每成功文件耗时: {avg:.2f}s")
            if m.get("total_seconds", 0.0) > 0:
                lines.append(
                    f"吞吐率(成功文件/分钟): {success_count / m.get('total_seconds', 1.0) * 60:.2f}"
                )
        return "\n".join(lines) + "\n"

    def _run_mshelp_only(self):
        logging.info("scanning MSHelpViewer folders...")
        scan_start = time.perf_counter()
        mshelp_dirs, cab_files = self._scan_mshelp_cab_candidates()
        self._add_perf_seconds("scan_seconds", time.perf_counter() - scan_start)

        logging.info("MSHelpViewer folder count: %s", len(mshelp_dirs))
        logging.info("MSHelp CAB candidate count: %s", len(cab_files))

        self.stats["total"] = len(cab_files)
        results = []
        if cab_files:
            logging.info("start processing MSHelp CAB files: %s", len(cab_files))
            batch_start = time.perf_counter()
            results.extend(self.run_batch(cab_files))
            self._add_perf_seconds("batch_seconds", time.perf_counter() - batch_start)
        else:
            print("\n[INFO] No MSHelp CAB files found under source folder.")

        mshelp_index_outputs = self._write_mshelp_index_files()
        mshelp_merged_outputs = self._merge_mshelp_markdowns()
        return results, mshelp_dirs, mshelp_index_outputs, mshelp_merged_outputs

    def run(self, resume_file_list=None):
        self.setup_logging()
        self.print_runtime_summary()
        self._reset_perf_metrics()
        self._office_file_counter = 0
        total_start = time.perf_counter()

        merge_outputs = []
        batch_results = []
        self.generated_pdfs = []
        self.generated_merge_outputs = []
        self.generated_merge_markdown_outputs = []
        self.generated_map_outputs = []
        self.generated_markdown_outputs = []
        self.generated_markdown_quality_outputs = []
        self.generated_excel_json_outputs = []
        self.generated_records_json_outputs = []
        self.generated_chromadb_outputs = []
        self.generated_update_package_outputs = []
        self.generated_mshelp_outputs = []
        self.markdown_quality_records = []
        self.conversion_index_records = []
        self.merge_index_records = []
        self.mshelp_records = []
        self.collect_index_path = None
        self.convert_index_path = None
        self.merge_excel_path = None
        self.corpus_manifest_path = None
        self.update_package_manifest_path = None
        self.markdown_quality_report_path = None
        self.chromadb_export_manifest_path = None
        self.llm_hub_root = ""
        self.incremental_registry_path = ""
        self._incremental_context = None

        try:
            self._check_sandbox_free_space_or_raise()
        except Exception as e:
            logging.error(f"sandbox free space precheck failed: {e}")
            raise

        if self.run_mode == MODE_COLLECT_ONLY:
            self.collect_office_files_and_build_excel()
        elif self.run_mode == MODE_MSHELP_ONLY:
            (
                batch_results,
                mshelp_dirs,
                mshelp_index_outputs,
                mshelp_merged_outputs,
            ) = self._run_mshelp_only()
            summary = (
                f"\n=== MSHelp 最终统计(v{__version__}) ===\n"
                f"MSHelpViewer目录: {len(mshelp_dirs)}\n"
                f"总处理(CAB): {self.stats['total']}\n"
                f"成功: {self.stats['success']}\n"
                f"失败: {self.stats['failed']}\n"
                f"超时: {self.stats['timeout']}\n"
                f"跳过: {self.stats['skipped']}\n"
                f"索引输出: {len(mshelp_index_outputs)}\n"
                f"合并输出: {len(mshelp_merged_outputs)}\n"
            )
            logging.info(summary)
            print(summary)
        else:
            incremental_ctx = {}
            if self.run_mode in (MODE_CONVERT_ONLY, MODE_CONVERT_THEN_MERGE):
                if resume_file_list is not None:
                    files = []
                    for p in resume_file_list:
                        p = str(p or "").strip()
                        if not p:
                            continue
                        files.append(os.path.abspath(p))
                    logging.info(
                        "resume mode active: using provided file list count=%s",
                        len(files),
                    )
                else:
                    scan_start = time.perf_counter()
                    logging.info("scanning files...")
                    files = self._scan_convert_candidates()
                    logging.info("scan candidate file count: %s", len(files))

                    files, source_priority_skips = self._apply_source_priority_filter(
                        files
                    )
                    if source_priority_skips:
                        self.stats["skipped"] += len(source_priority_skips)
                        batch_results.extend(source_priority_skips)

                    files, incremental_ctx = self._apply_incremental_filter(files)
                    if incremental_ctx.get("enabled"):
                        self.stats["skipped"] += incremental_ctx.get(
                            "unchanged_count", 0
                        )
                        renamed_pairs = incremental_ctx.get("renamed_pairs", []) or []
                        if renamed_pairs and not incremental_ctx.get(
                            "reprocess_renamed", False
                        ):
                            self.stats["skipped"] += len(renamed_pairs)
                            for item in renamed_pairs:
                                batch_results.append(
                                    {
                                        "source_path": item.get("to_path", ""),
                                        "status": "renamed_detected",
                                        "detail": "rename detected; no reconvert",
                                        "final_path": "",
                                        "renamed_from": item.get("from_path", ""),
                                    }
                                )

                    files, dedup_skips = self._apply_global_md5_dedup(files)
                    if dedup_skips:
                        self.stats["skipped"] += len(dedup_skips)
                        batch_results.extend(dedup_skips)
                    self._add_perf_seconds(
                        "scan_seconds", time.perf_counter() - scan_start
                    )

                self._emit_file_plan(files)

                self.stats["total"] = len(files)
                if files:
                    logging.info("start processing %s files", len(files))
                    batch_start = time.perf_counter()
                    batch_results.extend(self.run_batch(files))
                    self._add_perf_seconds(
                        "batch_seconds", time.perf_counter() - batch_start
                    )
                else:
                    if resume_file_list is not None:
                        print("\n[INFO] Resume mode: no pending files.")
                    elif incremental_ctx.get("enabled"):
                        print(
                            "\n[INFO] Incremental mode: no added/modified files found."
                        )
                    else:
                        print(
                            "\n[INFO] No convertible Office files found in source directory."
                        )

                self.close_office_apps()

                failed_count = self.stats["failed"] + self.stats["timeout"]
                should_retry = False
                if failed_count > 0:
                    if self.config.get("auto_retry_failed", False):
                        should_retry = True
                        print(f"\n[CONFIG] Auto retry failed files ({failed_count})...")
                    elif self.interactive:
                        should_retry = self.ask_retry_failed_files(
                            failed_count, timeout=20
                        )

                if should_retry:
                    print("\n" + "=" * 60)
                    print("  Start retrying failed files...")
                    print("  Re-checking and cleaning related processes...")
                    print("=" * 60)

                    if not self.reuse_process:
                        self.cleanup_all_processes()

                    retry_files = []
                    retry_alias_map = {}
                    if os.path.exists(self.failed_dir):
                        if self.config.get(
                            "enable_sandbox", True
                        ) and not os.path.exists(self.temp_sandbox):
                            os.makedirs(self.temp_sandbox)

                        valid_exts = [
                            e
                            for sub in self.config.get(
                                "allowed_extensions", {}
                            ).values()
                            for e in sub
                        ]
                        for f in os.listdir(self.failed_dir):
                            if f.startswith("~$"):
                                continue
                            ext = os.path.splitext(f)[1].lower()
                            if ext in valid_exts:
                                retry_path = os.path.join(self.failed_dir, f)
                                retry_files.append(retry_path)

                        name_map = {}
                        for src in self.error_records:
                            name_map.setdefault(os.path.basename(src), []).append(src)
                        for retry_path in retry_files:
                            name = os.path.basename(retry_path)
                            if name in name_map and name_map[name]:
                                retry_alias_map[retry_path] = name_map[name].pop(0)

                    if retry_files:
                        retry_start = time.perf_counter()
                        batch_results.extend(
                            self.run_batch(
                                retry_files,
                                is_retry=True,
                                source_alias_map=retry_alias_map,
                            )
                        )
                        self._add_perf_seconds(
                            "batch_seconds", time.perf_counter() - retry_start
                        )
                    else:
                        print("No retryable files found in failed directory.")

                    self.close_office_apps()

                if self.run_mode == MODE_CONVERT_ONLY and self.enable_merge_excel:
                    try:
                        self._write_conversion_index_workbook()
                    except Exception as e:
                        logging.error(f"failed to write conversion index: {e}")

                try:
                    self._flush_incremental_registry(batch_results)
                except Exception as e:
                    logging.error(f"failed to write incremental registry: {e}")
                try:
                    self._generate_update_package(batch_results)
                except Exception as e:
                    logging.error(f"failed to generate update package: {e}")

            elif self.run_mode == MODE_MERGE_ONLY:
                print("Current mode is merge_and_convert. Conversion step skipped.")
                merge_start = time.perf_counter()
                merge_outputs = self._run_merge_mode_pipeline(batch_results) or []
                self._add_perf_seconds(
                    "merge_seconds", time.perf_counter() - merge_start
                )

            if (
                self.run_mode == MODE_CONVERT_THEN_MERGE
                and self.config.get("enable_merge", True)
                and bool(self.config.get("output_enable_merged", True))
            ):
                merge_start = time.perf_counter()
                merge_outputs = []
                if bool(self.config.get("output_enable_pdf", True)):
                    merge_outputs.extend(self.merge_pdfs() or [])
                if bool(self.config.get("output_enable_md", True)):
                    md_outputs = (
                        self.merge_markdowns(candidates=self.generated_markdown_outputs)
                        or []
                    )
                    merge_outputs.extend(md_outputs)
                self._add_perf_seconds(
                    "merge_seconds", time.perf_counter() - merge_start
                )

            summary = (
                f"\n=== 最终统计(v{__version__}) ===\n"
                f"总处理: {self.stats['total']}\n"
                f"成功: {self.stats['success']}\n"
                f"失败: {self.stats['failed']}\n"
                f"超时: {self.stats['timeout']}\n"
                f"跳过(含空文件/策略): {self.stats['skipped']}\n"
            )

            if self._incremental_context and self._incremental_context.get("enabled"):
                inc = self._incremental_context
                summary += (
                    f"增量扫描: {inc.get('scanned_count', 0)} | "
                    f"新增: {inc.get('added_count', 0)} | "
                    f"修改: {inc.get('modified_count', 0)} | "
                    f"重命名: {inc.get('renamed_count', 0)} | "
                    f"未变更: {inc.get('unchanged_count', 0)} | "
                    f"删除: {inc.get('deleted_count', 0)}\n"
                )
                if self.incremental_registry_path:
                    summary += f"增量账本: {self.incremental_registry_path}\n"
                if self.update_package_manifest_path:
                    summary += f"增量包清单: {self.update_package_manifest_path}\n"

            logging.info(summary)
            print(summary)

        postprocess_start = time.perf_counter()
        try:
            self._write_excel_structured_json_exports()
        except Exception as e:
            logging.error(f"failed to write Excel JSON: {e}")

        try:
            self._write_records_json_exports()
        except Exception as e:
            logging.error(f"failed to write Records JSON: {e}")

        try:
            self._write_chromadb_export()
        except Exception as e:
            logging.error(f"failed to write ChromaDB export: {e}")

        try:
            self._write_markdown_quality_report()
        except Exception as e:
            logging.error(f"failed to write Markdown quality report: {e}")

        try:
            self._write_corpus_manifest(merge_outputs=merge_outputs)
        except Exception as e:
            logging.error(f"failed to write corpus manifest: {e}")

        # 导出失败文件详细报告
        if self.detailed_error_records:
            try:
                report_result = self.export_failed_files_report()
                if report_result.get("txt_path"):
                    failed_summary = (
                        f"\n=== 失败文件报告 ===\n"
                        f"报告路径: {report_result['txt_path']}\n"
                        f"摘要: {report_result['summary']}\n"
                    )
                    print(failed_summary)
                    logging.info(failed_summary)
            except Exception as e:
                logging.error(f"failed to export failed files report: {e}")
        self._add_perf_seconds(
            "postprocess_seconds", time.perf_counter() - postprocess_start
        )
        self.perf_metrics["total_seconds"] = time.perf_counter() - total_start
        perf_summary = self._build_perf_summary()
        logging.info(perf_summary)
        print(perf_summary)

        try:
            open_dir = self.config["target_folder"]
            if (
                merge_outputs
                and self.merge_output_dir
                and os.path.isdir(self.merge_output_dir)
            ):
                open_dir = self.merge_output_dir
            os.startfile(open_dir)
        except Exception:
            pass

        if self.temp_sandbox and os.path.exists(self.temp_sandbox):
            try:
                shutil.rmtree(self.temp_sandbox, ignore_errors=True)
            except Exception:
                pass


def create_default_config(config_path):
    """Create a default config file."""
    try:
        default_config = {
            "source_folder": "C:\\Docs",
            "target_folder": "C:\\PDFs",
            "log_folder": "./logs",
            "enable_sandbox": True,
            "default_engine": "ask",
            "kill_process_mode": "ask",
            "auto_retry_failed": False,
            "office_reuse_app": True,
            "office_restart_every_n_files": 25,
            "timeout_seconds": 60,
            "pdf_wait_seconds": 15,
            "ppt_timeout_seconds": 180,
            "ppt_pdf_wait_seconds": 30,
            "enable_merge": True,
            "output_enable_pdf": True,
            "output_enable_md": True,
            "output_enable_merged": True,
            "output_enable_independent": False,
            "merge_convert_submode": MERGE_CONVERT_SUBMODE_MERGE_ONLY,
            "enable_corpus_manifest": True,
            "markdown_strip_header_footer": True,
            "markdown_structured_headings": True,
            "enable_markdown_quality_report": True,
            "markdown_quality_sample_limit": 20,
            "enable_excel_json": False,
            "enable_chromadb_export": False,
            "chromadb_persist_dir": "",
            "chromadb_collection_name": "office_corpus",
            "chromadb_max_chars_per_chunk": 1800,
            "chromadb_chunk_overlap": 200,
            "chromadb_write_jsonl_fallback": True,
            "enable_incremental_mode": False,
            "incremental_verify_hash": False,
            "incremental_reprocess_renamed": False,
            "incremental_registry_path": "",
            "source_priority_skip_same_name_pdf": False,
            "global_md5_dedup": False,
            "enable_update_package": True,
            "update_package_root": "",
            "cab_7z_path": "",
            "mshelpviewer_folder_name": "MSHelpViewer",
            "enable_mshelp_merge_output": True,
            "enable_mshelp_output_docx": False,
            "enable_mshelp_output_pdf": False,
            "excel_json_max_rows": 2000,
            "excel_json_max_cols": 80,
            "excel_json_records_preview": 200,
            "excel_json_profile_rows": 500,
            "excel_json_include_formulas": True,
            "excel_json_extract_sheet_links": True,
            "excel_json_include_merged_ranges": True,
            "excel_json_formula_sample_limit": 200,
            "excel_json_merged_range_limit": 500,
            "max_merge_size_mb": 80,
            "merge_filename_pattern": "Merged_{category}_{timestamp}_{idx}",
            "price_keywords": ["报价", "价格表", "Price", "Quotation"],
            "excluded_folders": ["temp", "backup", "archive"],
            "allowed_extensions": {
                "word": [".doc", ".docx"],
                "excel": [".xls", ".xlsx"],
                "powerpoint": [".ppt", ".pptx"],
                "pdf": [".pdf"],
                "cab": [".cab"],
            },
            "overwrite_same_size": True,
            "merge_mode": MERGE_MODE_CATEGORY,
            "run_mode": MODE_CONVERT_THEN_MERGE,
            "merge_source": "target",
            "temp_sandbox_root": "",
            "sandbox_min_free_gb": 10,
            "sandbox_low_space_policy": "block",
            "enable_llm_delivery_hub": True,
            "llm_delivery_root": "",
            "llm_delivery_flatten": True,
            "llm_delivery_include_pdf": False,
            "enable_gdrive_upload": False,
            "gdrive_client_secrets_path": "",
            "gdrive_folder_id": "",
            "gdrive_token_path": "",
            "enable_upload_readme": True,
            "enable_upload_json_manifest": True,
            "upload_dedup_merged": True,
            "enable_merge_index": False,
            "enable_merge_excel": False,
            "enable_merge_map": True,
            "bookmark_with_short_id": True,
            "everything": {
                "enabled": True,
                "es_path": "",
                "prefer_path_exact": True,
                "timeout_ms": 1500,
            },
            "listary": {"enabled": True, "copy_query_on_locate": True},
            "privacy": {"mask_md5_in_logs": True},
            "ui": {
                "tooltip_delay_ms": 500,
                "tooltip_bg": "#FFF7D6",
                "tooltip_fg": "#202124",
                "tooltip_font_family": "System",
                "tooltip_font_size": 9,
                "tooltip_auto_theme": True,
            },
        }
        with open(config_path, "w", encoding="utf-8") as f:
            json.dump(default_config, f, indent=4, ensure_ascii=False)
        print(f"Default config created: {config_path}")
        return True
    except Exception as e:
        print(f"Failed to create default config: {e}")
        return False


# =============== CLI entry ===============

if __name__ == "__main__":
    clear_console()
    script_dir = get_app_path()
    default_config_path = os.path.join(script_dir, "config.json")
    parser = argparse.ArgumentParser()
    parser.add_argument("--config", default=default_config_path)
    args = parser.parse_args()

    if not os.path.exists(args.config):
        try:
            create_default_config(args.config)
        except Exception:
            pass

    converter = OfficeConverter(args.config, interactive=True)
    converter.cli_wizard()
    converter.run()
