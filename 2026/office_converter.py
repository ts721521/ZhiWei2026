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
from datetime import datetime
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

__version__ = "5.19.1"

from converter.constants import (
    wdFormatPDF,
    xlTypePDF,
    ppSaveAsPDF,
    ppFixedFormatTypePDF,
    xlPDF_SaveAs,
    xlRepairFile,
    ENGINE_WPS,
    ENGINE_MS,
    ENGINE_ASK,
    KILL_MODE_ASK,
    KILL_MODE_AUTO,
    KILL_MODE_KEEP,
    MODE_CONVERT_ONLY,
    MODE_MERGE_ONLY,
    MODE_CONVERT_THEN_MERGE,
    MODE_COLLECT_ONLY,
    MODE_MSHELP_ONLY,
    MERGE_CONVERT_SUBMODE_MERGE_ONLY,
    MERGE_CONVERT_SUBMODE_PDF_TO_MD,
    COLLECT_MODE_COPY_AND_INDEX,
    COLLECT_MODE_INDEX_ONLY,
    MERGE_MODE_CATEGORY,
    MERGE_MODE_ALL_IN_ONE,
    STRATEGY_STANDARD,
    STRATEGY_SMART_TAG,
    STRATEGY_PRICE_ONLY,
    ERR_RPC_SERVER_BUSY,
)



from converter.errors import ConversionErrorType, classify_conversion_error
from converter.default_config import create_default_config
from converter.platform_utils import clear_console, get_app_path, is_mac, is_win
from converter.file_registry import FileRegistry
from converter.output_plan import compute_convert_output_plan
from converter.config_defaults import apply_config_defaults
from converter.readable import (
    readable_collect_mode,
    readable_content_strategy,
    readable_engine_type,
    readable_merge_mode,
    readable_run_mode,
)
from converter.runtime_prefs import get_merge_convert_submode, get_output_pref
from converter.path_config import get_path_from_config
from converter.office_cycle import (
    get_app_type_for_ext,
    get_office_restart_every,
    should_reuse_office_app,
)
from converter.process_policy import resolve_process_handling
from converter.perf_summary import build_perf_summary
from converter.process_ops import kill_process_by_name, process_names_for_engine
from converter.hash_utils import (
    build_short_id,
    compute_file_hash,
    compute_md5,
    make_file_hyperlink,
    mask_md5,
)
from converter.excel_json_utils import (
    build_column_profiles,
    col_index_to_label,
    detect_json_value_type,
    extract_formula_sheet_refs,
    is_effectively_empty_row,
    is_empty_json_cell,
    json_safe_value,
    looks_like_header_row,
    normalize_header_row,
)
from converter.artifact_meta import add_artifact, safe_file_meta
from converter.failure_stage import (
    get_failure_output_expectation,
    infer_failure_stage,
    sanitize_failure_log_stem,
)
from converter.failure_trace_utils import (
    build_failed_file_trace_payload,
    write_failed_file_trace_log,
)
from converter.chromadb_utils import (
    chunk_text_for_vector,
    resolve_chromadb_persist_dir,
    sanitize_chromadb_collection_name,
)
from converter.markdown_text_utils import (
    clean_markdown_page_lines,
    collect_margin_candidates,
    looks_like_heading_line,
    looks_like_page_number_line,
    normalize_extracted_text,
    normalize_margin_line,
    render_markdown_blocks,
)
from converter.naming_utils import ext_bucket, format_merge_filename
from converter.excel_chart_utils import extract_chart_title_text, stringify_chart_anchor
from converter.ai_paths import build_ai_output_path, build_ai_output_path_from_source
from converter.text_helpers import (
    extract_mshc_payload,
    find_files_recursive,
    meta_content_by_names,
    normalize_md_line,
    wrap_plain_text_for_pdf,
)
from converter.excel_sheet_utils import auto_fit_sheet, style_header_row
from converter.merge_candidates import (
    build_markdown_merge_tasks,
    scan_candidates_by_ext,
)
from converter.checkpoint_utils import (
    clear_checkpoint_file,
    get_checkpoint_path,
    mark_file_done_in_checkpoint,
    save_checkpoint,
)
from converter.runtime_paths import (
    resolve_incremental_registry_path,
    resolve_update_package_root,
)
from converter.batch_helpers import collect_retry_candidates, get_progress_prefix
from converter.callback_utils import emit_file_done, emit_file_plan
from converter.batch_parallel import run_batch_parallel as run_batch_parallel_impl
from converter.batch_sequential import run_batch as run_batch_impl
from converter.process_single import process_single_file as process_single_file_impl
from converter.run_workflow import run as run_workflow_impl
from converter.corpus_manifest import (
    maybe_build_llm_delivery_hub as maybe_build_llm_delivery_hub_impl,
)
from converter.corpus_manifest import write_corpus_manifest as write_corpus_manifest_impl
from converter.collect_index import (
    collect_office_files_and_build_excel as collect_office_files_and_build_excel_impl,
)
from converter.merge_pdfs import merge_pdfs as merge_pdfs_impl
from converter.source_roots import (
    get_configured_source_roots,
    get_source_root_for_path,
    get_source_roots,
    probe_source_root_access,
)
from converter.safe_exec import safe_exec
from converter.file_ops import (
    copy_pdf_direct,
    handle_file_conflict,
    quarantine_failed_file,
    unblock_file,
)
from converter.interactive_prompts import confirm_continue_missing_md_merge
from converter.incremental_filters import (
    apply_global_md5_dedup,
    apply_source_priority_filter,
)
from converter.incremental_registry_ops import (
    build_source_meta,
    flush_incremental_registry,
)
from converter.incremental_scan import apply_incremental_filter
from converter.scan_convert_candidates import scan_convert_candidates
from converter.mshelp_scan import find_mshelpviewer_dirs, scan_mshelp_cab_candidates
from converter.mshelp_topics import parse_mshelp_topics
from converter.update_package_index import write_update_package_index_xlsx
from converter.mshelp_records import build_mshelp_record, write_mshelp_index_files
from converter.failure_report import (
    export_failed_files_report as export_failed_files_report_impl,
)
from converter.chromadb_docs import collect_chromadb_documents
from converter.markdown_docx_export import (
    export_markdown_to_docx as export_markdown_to_docx_impl,
)
from converter.markdown_pdf_export import (
    export_markdown_to_pdf as export_markdown_to_pdf_impl,
)
from converter.chromadb_export import write_chromadb_export as write_chromadb_export_impl
from converter.mshelp_merge import merge_mshelp_markdowns as merge_mshelp_markdowns_impl
from converter.excel_json_export import (
    export_single_excel_json as export_single_excel_json_impl,
)
from converter.update_package_export import (
    generate_update_package as generate_update_package_impl,
)


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
            "permission_denied": 0,  # 鏉冮檺閿欒璁℃暟
            "file_locked": 0,  # 鏂囦欢閿佸畾璁℃暟
            "file_corrupted": 0,  # 鏂囦欢鎹熷潖璁℃暟
            "com_error": 0,  # COM閿欒璁℃暟
        }
        self.error_records = []  # 绠€鍗曡矾寰勫垪琛紙鍏煎鏃ч€昏緫锛?
        self.detailed_error_records = []  # 缁撴瀯鍖栭敊璇褰?
        self.failed_report_path = None  # 澶辫触鎶ュ憡璺緞

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
        print(f" 鐭ュ杺 ZhiWei 路 鐭ヨ瘑鎶曞杺宸ュ叿  v{__version__}")
        print(" Supports WPS / Microsoft Office, CLI / GUI dual mode")
        print("=" * 60)
        print(f"Config file: {self.config_path}\n")

    def print_step_title(self, text):
        print("\n" + "-" * 60)
        print(text)
        print("-" * 60)

    def get_readable_run_mode(self):
        return readable_run_mode(self.run_mode)

    def get_readable_collect_mode(self):
        return readable_collect_mode(self.collect_mode)

    def get_readable_content_strategy(self):
        return readable_content_strategy(self.content_strategy)

    def get_readable_engine_type(self):
        return readable_engine_type(self.engine_type)

    def get_readable_merge_mode(self):
        return readable_merge_mode(self.merge_mode)

    compute_convert_output_plan = staticmethod(compute_convert_output_plan)

    def _get_output_pref(self):
        return get_output_pref(self.config)

    def _get_merge_convert_submode(self):
        return get_merge_convert_submode(self.config)

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
        for app in process_names_for_engine(self.engine_type):
            self._kill_process_by_name(app)

    def _kill_process_by_name(self, app_name):
        try:
            kill_process_by_name(app_name, has_win32=HAS_WIN32, run_cmd=subprocess.run)
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
        self.config["obsidian_root"] = self._get_path_from_config("obsidian_root")
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
        runtime = apply_config_defaults(
            cfg,
            run_mode_default=self.run_mode,
            collect_mode_default=self.collect_mode,
            content_strategy_default=self.content_strategy,
            enable_merge_index_default=self.enable_merge_index,
            enable_merge_excel_default=self.enable_merge_excel,
        )
        self.price_keywords = runtime["price_keywords"]
        self.excluded_folders = runtime["excluded_folders"]
        self.merge_mode = runtime["merge_mode"]
        self.run_mode = runtime["run_mode"]
        self.collect_mode = runtime["collect_mode"]
        self.content_strategy = runtime["content_strategy"]
        self.enable_merge_index = runtime["enable_merge_index"]
        self.enable_merge_excel = runtime["enable_merge_excel"]

    def _should_reuse_office_app(self):
        return should_reuse_office_app(
            self.config, has_win32=HAS_WIN32, is_mac_platform=is_mac()
        )

    def _get_office_restart_every(self):
        return get_office_restart_every(self.config)

    def _get_app_type_for_ext(self, ext):
        return get_app_type_for_ext(self.config, ext)

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
        return get_path_from_config(
            self.config,
            key_base,
            prefer_win=is_win(),
            prefer_mac=is_mac(),
        )

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
        decision = resolve_process_handling(
            run_mode=self.run_mode,
            kill_process_mode=self.config.get("kill_process_mode", KILL_MODE_ASK),
            interactive=self.interactive,
        )
        if decision["skip"]:
            return
        if decision["cleanup_all"]:
            self.cleanup_all_processes()
        self.reuse_process = decision["reuse_process"]

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
        return handle_file_conflict(temp_pdf_path, target_pdf_path)

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
        return safe_exec(
            func,
            *args,
            retries=retries,
            is_running_getter=lambda: bool(self.is_running),
            sleep_fn=time.sleep,
            randint_fn=random.randint,
            com_error_cls=pywintypes.com_error,
            rpc_server_busy_code=ERR_RPC_SERVER_BUSY,
            **kwargs,
        )

    def _unblock_file(self, file_path):
        unblock_file(file_path)

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
        copy_pdf_direct(source, temp_target)

    def quarantine_failed_file(self, source_path, should_copy=True):
        return quarantine_failed_file(source_path, self.failed_dir, should_copy=should_copy)

    _sanitize_failure_log_stem = staticmethod(sanitize_failure_log_stem)

    def _get_failure_output_expectation(self):
        return get_failure_output_expectation(
            self.run_mode,
            self.config,
            self.compute_convert_output_plan,
        )

    def _infer_failure_stage(self, source_path, raw_error="", context=None):
        return infer_failure_stage(
            source_path,
            raw_error=raw_error,
            context=context,
            cab_extensions=self.config.get("allowed_extensions", {}).get("cab", []),
            expected_outputs_getter=self._get_failure_output_expectation,
        )

    def _build_failed_file_trace_payload(
        self,
        *,
        source_path,
        error_detail,
        status,
        elapsed,
        is_retry,
        failed_copy_path=None,
        extra_context=None,
    ):
        return build_failed_file_trace_payload(
            source_path=source_path,
            error_detail=error_detail,
            status=status,
            elapsed=elapsed,
            is_retry=is_retry,
            failed_copy_path=failed_copy_path,
            extra_context=extra_context,
            get_failure_output_expectation_fn=self._get_failure_output_expectation,
            get_readable_run_mode_fn=self.get_readable_run_mode,
            get_readable_engine_type_fn=self.get_readable_engine_type,
            infer_failure_stage_fn=self._infer_failure_stage,
        )

    def _write_failed_file_trace_log(self, payload, failed_copy_path=None):
        return write_failed_file_trace_log(
            payload,
            failed_copy_path=failed_copy_path,
            enable_trace_log=self.config.get("enable_failed_file_trace_log", True),
            failed_dir=self.failed_dir,
            target_folder=self.config.get("target_folder", ""),
            sanitize_stem_fn=self._sanitize_failure_log_stem,
            log_error=logging.error,
        )

    def record_detailed_error(self, source_path, exception, context=None):
        """
        璁板綍璇︾粏鐨勯敊璇俊鎭紝鍖呮嫭閿欒鍒嗙被鍜屽鐞嗗缓璁€?

        Args:
            source_path: 澶辫触鏂囦欢鐨勮矾寰?
            exception: 寮傚父瀵硅薄
            context: 棰濆涓婁笅鏂囦俊鎭紙濡傝繍琛屾ā寮忋€佽浆鎹㈠紩鎿庣瓑锛?

        Returns:
            dict: 閿欒璇︽儏瀛楀吀
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
        record["failure_stage"] = self._infer_failure_stage(
            record["source_path"], raw_error=record["raw_error"], context=record["context"]
        )
        record["expected_outputs"] = self._get_failure_output_expectation()

        self.detailed_error_records.append(record)

        # 鏇存柊鍒嗙被缁熻
        error_type = error_info["error_type"]
        if error_type in self.stats:
            self.stats[error_type] += 1

        return record

    def export_failed_files_report(self, output_dir=None):
        output_dir = output_dir or self.config.get("target_folder", ".")
        result = export_failed_files_report_impl(
            self.detailed_error_records,
            output_dir,
            run_mode=self.get_readable_run_mode(),
            now_fn=datetime.now,
            log_error=logging.error,
        )
        self.failed_report_path = result.get("txt_path")
        return result

    def get_error_summary_for_display(self):
        """
        鑾峰彇鐢ㄤ簬 GUI 鏄剧ず鐨勯敊璇憳瑕併€?

        Returns:
            dict: 鎸夐敊璇被鍨嬪垎缁勭殑鏂囦欢鍒楄〃鍜屽鐞嗗缓璁?
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

    _find_files_recursive = staticmethod(find_files_recursive)

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

    _extract_mshc_payload = staticmethod(extract_mshc_payload)
    _meta_content_by_names = staticmethod(meta_content_by_names)

    def _parse_mshelp_topics(self, html_root):
        return parse_mshelp_topics(
            html_root,
            find_files_recursive_fn=self._find_files_recursive,
            has_bs4=HAS_BS4,
            beautifulsoup_cls=BeautifulSoup if HAS_BS4 else None,
            meta_content_by_names_fn=self._meta_content_by_names,
        )

    _normalize_md_line = staticmethod(normalize_md_line)

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
                "## 鐩綍",
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
        self.mshelp_records.append(
            build_mshelp_record(
                source_cab_path,
                markdown_path,
                topic_count,
                folder_name=self.config.get("mshelpviewer_folder_name", "MSHelpViewer"),
                get_source_root_for_path_fn=self._get_source_root_for_path,
            )
        )

    def _find_mshelpviewer_dirs(self, root_dir):
        return find_mshelpviewer_dirs(
            root_dir,
            folder_name=self.config.get("mshelpviewer_folder_name", "MSHelpViewer"),
            is_win_fn=is_win,
        )

    def _scan_mshelp_cab_candidates(self):
        return scan_mshelp_cab_candidates(
            self.config,
            self._get_source_roots(),
            find_mshelpviewer_dirs_fn=self._find_mshelpviewer_dirs,
            find_files_recursive_fn=self._find_files_recursive,
            is_win_fn=is_win,
        )

    def _write_mshelp_index_files(self):
        return write_mshelp_index_files(
            self.mshelp_records,
            self.config.get("target_folder", ""),
            generated_mshelp_outputs=self.generated_mshelp_outputs,
            log_info=logging.info,
        )

    _wrap_plain_text_for_pdf = staticmethod(wrap_plain_text_for_pdf)

    def _export_markdown_to_docx(self, md_path, out_docx):
        return export_markdown_to_docx_impl(
            md_path,
            out_docx,
            has_pydocx=HAS_PYDOCX,
            document_cls=Document if HAS_PYDOCX else None,
            re_module=re,
        )

    def _export_markdown_to_pdf(self, md_path, out_pdf):
        return export_markdown_to_pdf_impl(
            md_path,
            out_pdf,
            has_reportlab=HAS_REPORTLAB,
            canvas_cls=canvas.Canvas if HAS_REPORTLAB else None,
            page_size=A4 if HAS_REPORTLAB else None,
            wrap_plain_text_for_pdf_fn=self._wrap_plain_text_for_pdf,
        )

    def _merge_mshelp_markdowns(self):
        merge_start = time.perf_counter()
        outputs = merge_mshelp_markdowns_impl(
            self.mshelp_records,
            self.config,
            generated_outputs=self.generated_mshelp_outputs,
            export_markdown_to_docx_fn=self._export_markdown_to_docx,
            export_markdown_to_pdf_fn=self._export_markdown_to_pdf,
            now_fn=datetime.now,
            log_info=logging.info,
            log_warning=logging.warning,
        )
        if outputs:
            self._add_perf_seconds(
                "mshelp_merge_seconds", time.perf_counter() - merge_start
            )
        return outputs

    def process_single_file(
        self, file_path, target_path_initial, ext, progress_str, is_retry=False
    ):
        return process_single_file_impl(
            self,
            file_path,
            target_path_initial,
            ext,
            progress_str,
            is_retry=is_retry,
        )

    # =============== Batch processing / retry ===============

    def get_progress_prefix(self, current, total):
        return get_progress_prefix(current, total)

    def _emit_file_plan(self, file_list):
        emit_file_plan(
            getattr(self, "file_plan_callback", None),
            file_list,
            warn_func=logging.warning,
        )

    def _emit_file_done(self, record):
        emit_file_done(
            getattr(self, "file_done_callback", None),
            record,
            warn_func=logging.warning,
        )

    def run_batch(self, file_list, is_retry=False, source_alias_map=None):
        return run_batch_impl(
            self,
            file_list,
            is_retry=is_retry,
            source_alias_map=source_alias_map,
        )

    # ========== 鏂偣缁紶鍜屽苟鍙戝鐞嗗姛鑳?==========

    def _get_checkpoint_path(self):
        """Return checkpoint path for current task."""
        return get_checkpoint_path(self.config)

    def _init_checkpoint(self, file_list):
        """Initialize checkpoint and return (checkpoint_data, pending_files)."""
        if not self.config.get("enable_checkpoint", True):
            return None, file_list

        checkpoint_path = self._get_checkpoint_path()
        if not checkpoint_path:
            return None, file_list

        # 妫€鏌ユ槸鍚﹀瓨鍦ㄦ湭瀹屾垚鐨勬柇鐐?
        if os.path.exists(checkpoint_path):
            try:
                with open(checkpoint_path, "r", encoding="utf-8") as f:
                    checkpoint = json.load(f)

                if checkpoint.get("status") == "running":
                    completed = set(checkpoint.get("completed_files", []))
                    pending = [f for f in file_list if f not in completed]

                    if pending and self.config.get("checkpoint_auto_resume", True):
                        completed_count = len(completed)
                        total_count = len(file_list)
                        print(
                            f"\n[checkpoint] detected unfinished task: {completed_count}/{total_count} completed"
                        )
                        print(f"[checkpoint] pending files: {len(pending)}")

                        if hasattr(self, "checkpoint_resume_callback") and callable(
                            self.checkpoint_resume_callback
                        ):
                            should_resume = self.checkpoint_resume_callback(
                                completed_count, total_count
                            )
                            if not should_resume:
                                print("[checkpoint] user declined resume; restart from scratch")
                                os.remove(checkpoint_path)
                                return None, file_list

                        print("[checkpoint] resuming from checkpoint...")
                        return checkpoint, pending
            except Exception as e:
                logging.warning(f"Failed to load checkpoint: {e}")

        # 鍒涘缓鏂版柇鐐?
        checkpoint = {
            "version": 1,
            "created_at": datetime.now().isoformat(timespec="seconds"),
            "updated_at": datetime.now().isoformat(timespec="seconds"),
            "planned_files": list(file_list),
            "completed_files": [],
            "status": "running",
        }
        self._save_checkpoint(checkpoint)
        return checkpoint, file_list

    def _save_checkpoint(self, checkpoint):
        """Persist checkpoint to file."""
        save_checkpoint(checkpoint, self._get_checkpoint_path())

    def _mark_file_done_in_checkpoint(self, checkpoint, file_path):
        """Mark a file as completed in checkpoint data."""
        return mark_file_done_in_checkpoint(checkpoint, file_path)

    def _clear_checkpoint(self):
        """娓呴櫎鏂偣鏂囦欢"""
        clear_checkpoint_file(self._get_checkpoint_path())

    def _convert_single_file_threadsafe(
        self, fpath, target_path_initial, ext, progress_prefix, is_retry=False
    ):
        """
        绾跨▼瀹夊叏鐨勫崟鏂囦欢杞崲鏂规硶銆?
        姣忎釜绾跨▼鐙珛鍒濆鍖?COM锛岄伩鍏嶈法绾跨▼闂銆?
        """
        # 鍒濆鍖?COM锛堟瘡涓嚎绋嬪繀椤荤嫭绔嬭皟鐢級
        if HAS_WIN32:
            pythoncom.CoInitialize()

        try:
            # 璋冪敤鐜版湁鐨勫崟鏂囦欢澶勭悊閫昏緫
            return self.process_single_file(
                fpath, target_path_initial, ext, progress_prefix, is_retry
            )
        finally:
            # 娓呯悊 COM
            if HAS_WIN32:
                pythoncom.CoUninitialize()

    def run_batch_parallel(self, file_list, is_retry=False, source_alias_map=None):
        return run_batch_parallel_impl(
            self,
            file_list,
            is_retry=is_retry,
            source_alias_map=source_alias_map,
        )

    def _collect_retry_candidates(self):
        return collect_retry_candidates(
            self.failed_dir,
            self.config.get("allowed_extensions", {}),
            self.error_records,
            self.detailed_error_records,
        )

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
            # A4楂樺害 ~842pt. TopMargin 72.
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

    _format_merge_filename = staticmethod(format_merge_filename)

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

    _compute_md5 = staticmethod(compute_md5)
    _mask_md5 = staticmethod(mask_md5)
    _build_short_id = staticmethod(build_short_id)

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
        return merge_pdfs_impl(
            self,
            has_pypdf=HAS_PYPDF,
            has_openpyxl=HAS_OPENPYXL,
            pdf_writer_cls=PdfWriter if HAS_PYPDF else None,
            pdf_reader_cls=PdfReader if HAS_PYPDF else None,
            workbook_cls=Workbook if HAS_OPENPYXL else None,
            pythoncom_mod=pythoncom,
            win32_client=win32com.client,
        )

    def _scan_merge_candidates_by_ext(self, ext):
        scan_source_type = "target"
        if self.run_mode == MODE_MERGE_ONLY:
            scan_source_type = self.config.get("merge_source", "source")

        if scan_source_type == "source":
            scan_roots = self._get_source_roots()
        else:
            scan_roots = [self.config["target_folder"]]

        exclude_abs_paths = [self.failed_dir, self.merge_output_dir]
        if scan_source_type == "source":
            exclude_abs_paths.append(self.config["target_folder"])
        return scan_candidates_by_ext(ext, scan_roots, exclude_abs_paths)

    def _build_markdown_merge_tasks(self, md_files):
        return build_markdown_merge_tasks(md_files, self.merge_mode)

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
        return confirm_continue_missing_md_merge(
            self.interactive,
            input_fn=input,
            warn_func=logging.warning,
        )

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

    _compute_file_hash = staticmethod(compute_file_hash)
    _make_file_hyperlink = staticmethod(make_file_hyperlink)

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
        style_header_row(ws, has_openpyxl=HAS_OPENPYXL, font_cls=Font if HAS_OPENPYXL else None)

    @staticmethod
    def _auto_fit_sheet(ws, max_width=90):
        auto_fit_sheet(ws, max_width=max_width)

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
        return build_ai_output_path(
            source_path,
            sub_dir,
            ext,
            self.config.get("target_folder", ""),
        )

    def _build_ai_output_path_from_source(self, source_path, sub_dir, ext):
        return build_ai_output_path_from_source(
            source_path,
            sub_dir,
            ext,
            self.config.get("target_folder", ""),
            source_root_resolver=self._get_source_root_for_path,
        )

    _normalize_extracted_text = staticmethod(normalize_extracted_text)
    _normalize_margin_line = staticmethod(normalize_margin_line)
    _looks_like_page_number_line = staticmethod(looks_like_page_number_line)
    _collect_margin_candidates = staticmethod(collect_margin_candidates)
    _clean_markdown_page_lines = staticmethod(clean_markdown_page_lines)
    _looks_like_heading_line = staticmethod(looks_like_heading_line)
    _render_markdown_blocks = staticmethod(render_markdown_blocks)

    _json_safe_value = staticmethod(json_safe_value)
    _is_empty_json_cell = staticmethod(is_empty_json_cell)
    _is_effectively_empty_row = staticmethod(is_effectively_empty_row)
    _looks_like_header_row = staticmethod(looks_like_header_row)
    _normalize_header_row = staticmethod(normalize_header_row)
    _detect_json_value_type = staticmethod(detect_json_value_type)
    _build_column_profiles = staticmethod(build_column_profiles)
    _col_index_to_label = staticmethod(col_index_to_label)
    _extract_formula_sheet_refs = staticmethod(extract_formula_sheet_refs)

    _extract_chart_title_text = staticmethod(extract_chart_title_text)
    _stringify_chart_anchor = staticmethod(stringify_chart_anchor)

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

    _sanitize_chromadb_collection_name = staticmethod(
        sanitize_chromadb_collection_name
    )

    def _resolve_chromadb_persist_dir(self):
        return resolve_chromadb_persist_dir(self.config)

    _chunk_text_for_vector = staticmethod(chunk_text_for_vector)

    def _collect_chromadb_documents(self):
        return collect_chromadb_documents(
            generated_markdown_outputs=self.generated_markdown_outputs,
            markdown_quality_records=self.markdown_quality_records,
            config=self.config,
            chunk_text_for_vector_fn=self._chunk_text_for_vector,
        )

    def _write_chromadb_export(self):
        if not self.config.get("enable_chromadb_export", False):
            return None

        target_root = self.config.get("target_folder", "")
        if not target_root:
            return None

        docs = self._collect_chromadb_documents()
        if not docs:
            self.generated_chromadb_outputs = []
            self.chromadb_export_manifest_path = None
            logging.info("ChromaDB export skipped: no Markdown chunks available.")
            return None
        manifest_path, outputs = write_chromadb_export_impl(
            docs,
            config=self.config,
            target_root=target_root,
            has_chromadb=HAS_CHROMADB,
            chromadb_module=chromadb if HAS_CHROMADB else None,
            sanitize_collection_name_fn=self._sanitize_chromadb_collection_name,
            resolve_persist_dir_fn=self._resolve_chromadb_persist_dir,
            now_fn=datetime.now,
            log_info=logging.info,
        )
        self.generated_chromadb_outputs = outputs
        self.chromadb_export_manifest_path = manifest_path
        return manifest_path

    def _export_single_excel_json(self, source_excel_path):
        return export_single_excel_json_impl(
            source_excel_path,
            config=self.config,
            build_ai_output_path_from_source_fn=self._build_ai_output_path_from_source,
            source_root_resolver=self._get_source_root_for_path,
            has_openpyxl=HAS_OPENPYXL,
            load_workbook_fn=load_workbook if HAS_OPENPYXL else None,
            extract_workbook_defined_names_fn=self._extract_workbook_defined_names,
            extract_sheet_charts_fn=self._extract_sheet_charts,
            extract_sheet_pivot_tables_fn=self._extract_sheet_pivot_tables,
        )

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
        return safe_file_meta(path, self.config.get("target_folder", ""))

    def _add_artifact(self, artifacts, kind, path):
        add_artifact(artifacts, kind, path, self.config.get("target_folder", ""))

    def _maybe_build_llm_delivery_hub(self, target_folder, artifacts):
        return maybe_build_llm_delivery_hub_impl(self, target_folder, artifacts)

    def _write_corpus_manifest(self, merge_outputs=None):
        return write_corpus_manifest_impl(self, merge_outputs=merge_outputs)

    def collect_office_files_and_build_excel(self):
        return collect_office_files_and_build_excel_impl(
            self,
            has_openpyxl=HAS_OPENPYXL,
            workbook_cls=Workbook if HAS_OPENPYXL else None,
            font_cls=Font if HAS_OPENPYXL else None,
        )

    # =============== Main flow ===============

    def _get_configured_source_roots(self):
        return get_configured_source_roots(self.config)

    def _get_source_roots(self):
        """Return list of accessible source folder paths to scan."""
        return get_source_roots(self.config, is_dir_func=os.path.isdir)

    def _record_scan_access_skip(
        self, path, exception, context=None, seen_keys=None, silent=False
    ):
        abs_path = os.path.abspath(path) if path else ""
        key = abs_path.lower() if is_win() else abs_path
        if seen_keys is not None:
            if key in seen_keys:
                return None
            seen_keys.add(key)

        detail = self.record_detailed_error(
            abs_path or "<unknown>",
            exception,
            context=dict(context or {}, phase="scan", skip_only=True),
        )
        self.stats["skipped"] += 1

        msg = (
            f"[scan_skip] inaccessible path skipped: {abs_path or '<unknown>'} | "
            f"type={detail.get('error_type')} | err={exception}"
        )
        trace_payload = self._build_failed_file_trace_payload(
            source_path=abs_path or "<unknown>",
            error_detail=detail,
            status="skipped_scan",
            elapsed=(detail.get("context") or {}).get("elapsed", 0.0),
            is_retry=False,
            failed_copy_path=None,
            extra_context={"scan_only": True},
        )
        self._write_failed_file_trace_log(trace_payload, failed_copy_path=None)
        if not silent:
            print(f"[WARN] {msg}")
        logging.warning(msg)
        return detail

    def _probe_source_root_access(self, source_root, context=None, seen_keys=None):
        return probe_source_root_access(
            source_root,
            self._record_scan_access_skip,
            context=context,
            seen_keys=seen_keys,
            is_dir_func=os.path.isdir,
            listdir_fn=os.listdir,
        )

    def _get_source_root_for_path(self, abs_path):
        """Return the source root that contains abs_path (for relpath). Prefer longest match."""
        return get_source_root_for_path(
            abs_path,
            self._get_source_roots(),
            fallback=self.config.get("source_folder", "") or "",
        )

    def _scan_convert_candidates(self):
        return scan_convert_candidates(
            self.config,
            self._get_configured_source_roots(),
            probe_source_root_access_fn=self._probe_source_root_access,
            record_scan_access_skip_fn=self._record_scan_access_skip,
            filter_date=self.filter_date,
            filter_mode=self.filter_mode,
            print_warn_fn=print,
            log_error_fn=logging.error,
        )

    def _apply_source_priority_filter(self, files):
        return apply_source_priority_filter(
            files,
            self.config,
            is_win_fn=is_win,
            log_info=logging.info,
        )

    def _ext_bucket(self, path):
        return ext_bucket(path, self.config.get("allowed_extensions", {}))

    def _apply_global_md5_dedup(self, files):
        return apply_global_md5_dedup(
            files,
            self.config.get("global_md5_dedup", False),
            self._ext_bucket,
            self._compute_md5,
            log_warning=logging.warning,
            log_info=logging.info,
        )

    def _resolve_incremental_registry_path(self):
        return resolve_incremental_registry_path(self.config)

    def _build_source_meta(self, path, include_hash=False):
        return build_source_meta(
            path,
            include_hash=include_hash,
            compute_file_hash_fn=self._compute_file_hash,
            log_warning=logging.warning,
        )

    def _apply_incremental_filter(self, files):
        process_files, context = apply_incremental_filter(
            files,
            self.config,
            resolve_registry_path_fn=self._resolve_incremental_registry_path,
            build_source_meta_fn=self._build_source_meta,
            compute_file_hash_fn=self._compute_file_hash,
            log_info=logging.info,
        )
        if not context.get("enabled"):
            self._incremental_context = None
            self.incremental_registry_path = ""
            return process_files, context
        self._incremental_context = context
        self.incremental_registry_path = context.get("registry_path", "")
        return process_files, context

    def _flush_incremental_registry(self, process_results):
        return flush_incremental_registry(
            self._incremental_context,
            process_results,
            run_mode=self.run_mode,
            compute_md5_fn=self._compute_md5,
            log_info=logging.info,
        )

    def _resolve_update_package_root(self):
        return resolve_update_package_root(self.config)

    def _write_update_package_index_xlsx(self, xlsx_path, records):
        return write_update_package_index_xlsx(
            xlsx_path,
            records,
            has_openpyxl=HAS_OPENPYXL,
            workbook_cls=Workbook if HAS_OPENPYXL else None,
            font_cls=Font if HAS_OPENPYXL else None,
            make_file_hyperlink_fn=self._make_file_hyperlink,
            log_error=logging.error,
        )

    def _generate_update_package(self, process_results):
        package_manifest, outputs = generate_update_package_impl(
            process_results,
            incremental_context=self._incremental_context,
            config=self.config,
            run_mode=self.run_mode,
            resolve_update_package_root_fn=self._resolve_update_package_root,
            compute_md5_fn=self._compute_md5,
            write_update_package_index_xlsx_fn=self._write_update_package_index_xlsx,
            logger_info=logging.info,
        )
        if package_manifest:
            self.generated_update_package_outputs = outputs
            self.update_package_manifest_path = package_manifest
        return package_manifest

    def _build_perf_summary(self):
        return build_perf_summary(self.perf_metrics, self.stats)

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
        return run_workflow_impl(self, resume_file_list=resume_file_list, app_version=__version__)


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


