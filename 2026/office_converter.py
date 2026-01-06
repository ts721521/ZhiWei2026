# -*- coding: utf-8 -*-
"""
Office 文档批量转换 & 文件梳理工具 - 核心逻辑

说明：
- 可单独命令行运行，也可被 GUI 调用（office_gui.py）。
- GUI 模式通过 interactive=False 禁用所有 CLI input() 交互。
"""

import os
import sys
import time
import json
import shutil
import logging
import argparse
import uuid
import tempfile
import subprocess
import threading
import signal
import random
import hashlib
from datetime import datetime
from pathlib import Path

import win32com.client
import pythoncom
import pywintypes

# Windows 下用于带超时的键盘输入（仅 CLI 重试选择用）
try:
    import msvcrt

    HAS_MSVCRT = True
except ImportError:
    HAS_MSVCRT = False

# pypdf 用于 PDF 合并 & 内容扫描
try:
    from pypdf import PdfWriter, PdfReader

    HAS_PYPDF = True
except ImportError:
    HAS_PYPDF = False

# openpyxl 用于 collect_only 模式下生成 Excel 索引
try:
    from openpyxl import Workbook
    from openpyxl.styles import Font

    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

__version__ = "5.15.1"

# Office 常量
wdFormatPDF = 17
xlTypePDF = 0
ppSaveAsPDF = 32
ppFixedFormatTypePDF = 2
xlPDF_SaveAs = 57
xlRepairFile = 1

# 引擎类型
ENGINE_WPS = "wps"
ENGINE_MS = "ms"
ENGINE_ASK = "ask"

# 进程处理策略
KILL_MODE_ASK = "ask"
KILL_MODE_AUTO = "auto"
KILL_MODE_KEEP = "keep"

# 主运行模式
MODE_CONVERT_ONLY = "convert_only"
MODE_MERGE_ONLY = "merge_only"
MODE_CONVERT_THEN_MERGE = "convert_then_merge"
MODE_COLLECT_ONLY = "collect_only"  # 文件梳理去重模式

# collect_only 子模式
COLLECT_MODE_COPY_AND_INDEX = "copy_and_index"  # 去重 + 拷贝 + Excel
COLLECT_MODE_INDEX_ONLY = "index_only"  # 仅 Excel（不拷贝）

# 合并模式
MERGE_MODE_CATEGORY = "category_split"  # 按 Price_/Word_/Excel_... 分类拆卷（原有行为）
MERGE_MODE_ALL_IN_ONE = "all_in_one"  # 所有 PDF 合成一个总文件（忽略分类）

# 内容处理策略（仅转换模式）
STRATEGY_STANDARD = "standard"  # 仅按扩展名分类
STRATEGY_SMART_TAG = "smart_tag"  # 文件名/内容命中报价关键字则 Price_
STRATEGY_PRICE_ONLY = "price_only"  # 仅处理包含关键字的文件

ERR_RPC_SERVER_BUSY = -2147417846


def get_app_path():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))


def clear_console():
    try:
        if sys.stdout.isatty():
            os.system("cls" if os.name == "nt" else "clear")
    except Exception:
        pass


class OfficeConverter:
    def __init__(self, config_path: str, interactive: bool = True):
        """
        interactive:
          - True : CLI 模式（允许 input()）
          - False: GUI 模式（不得调用任何 input()）
        """
        self.config_path = config_path
        self.interactive = interactive

        self.temp_sandbox = None
        self.temp_sandbox_root = None
        self.failed_dir = None
        self.merge_output_dir = None

        self.engine_type = None
        self.is_running = True
        self.reuse_process = False

        # 运行模式 & 策略默认值（可被配置 / CLI / GUI 覆盖）
        self.run_mode = MODE_CONVERT_THEN_MERGE
        self.collect_mode = COLLECT_MODE_COPY_AND_INDEX
        self.merge_mode = MERGE_MODE_CATEGORY
        self.content_strategy = STRATEGY_STANDARD

        self.price_keywords = []
        self.excluded_folders = []

        self.progress_callback = None  # GUI 回调钩子: func(current, total)
        self.generated_pdfs = []

        # 仅在主线程下注册 signal（GUI 后台线程不会注册）
        if threading.current_thread() is threading.main_thread():
            try:
                signal.signal(signal.SIGINT, self.signal_handler)
                signal.signal(signal.SIGTERM, self.signal_handler)
            except Exception:
                pass

        # 1. 读取配置（只负责给 self.config 填默认值）
        self.load_config(config_path)

        # 2. 初始化目录（基于当前 config）
        self._init_paths_from_config()

        # 3. 初始化统计
        self.stats = {
            "total": 0,
            "success": 0,
            "failed": 0,
            "timeout": 0,
            "skipped": 0,
        }
        self.error_records = []

    # =============== 基础初始化 ===============

    def _init_paths_from_config(self):
        """根据当前 config 初始化临时目录 / 失败目录 / 合并目录"""
        # 临时转换目录根
        temp_root = self.config.get("temp_sandbox_root", "").strip()
        if temp_root:
            if not os.path.isabs(temp_root):
                temp_root = os.path.abspath(os.path.join(get_app_path(), temp_root))
        else:
            temp_root = tempfile.gettempdir()

        self.temp_sandbox_root = temp_root
        self.temp_sandbox = os.path.join(temp_root, "OfficeToPDF_Sandbox")
        os.makedirs(self.temp_sandbox, exist_ok=True)

        # 失败文件目录
        self.failed_dir = os.path.join(self.config["target_folder"], "_FAILED_FILES")
        os.makedirs(self.failed_dir, exist_ok=True)

        # 合并输出目录
        self.merge_output_dir = os.path.join(self.config["target_folder"], "_MERGED")
        os.makedirs(self.merge_output_dir, exist_ok=True)

    # =============== 通用显示 ===============

    def print_welcome(self):
        print("=" * 60)
        print(f" Office 文档批量转换 & 文件梳理工具  v{__version__}")
        print(" 支持 WPS / Microsoft Office · CLI / GUI 双模式")
        print("=" * 60)
        print(f"配置文件: {self.config_path}\n")

    def print_step_title(self, text):
        print("\n" + "-" * 60)
        print(text)
        print("-" * 60)

    def get_readable_run_mode(self):
        m = {
            MODE_CONVERT_ONLY: "仅转换",
            MODE_MERGE_ONLY: "仅合并",
            MODE_CONVERT_THEN_MERGE: "先转换再合并",
            MODE_COLLECT_ONLY: "文件梳理去重",
        }
        return m.get(self.run_mode, self.run_mode)

    def get_readable_collect_mode(self):
        m = {
            COLLECT_MODE_COPY_AND_INDEX: "去重 + 拷贝 + Excel 索引",
            COLLECT_MODE_INDEX_ONLY: "仅生成 Excel 索引（不拷贝）",
        }
        return m.get(self.collect_mode, self.collect_mode)

    def get_readable_content_strategy(self):
        m = {
            STRATEGY_STANDARD: "标准分类",
            STRATEGY_SMART_TAG: "智能标记（文件名/内容识别报价）",
            STRATEGY_PRICE_ONLY: "报价猎手（仅处理报价相关）",
        }
        return m.get(self.content_strategy, self.content_strategy)

    def get_readable_engine_type(self):
        m = {
            ENGINE_WPS: "WPS Office",
            ENGINE_MS: "Microsoft Office",
            None: "未使用（仅合并/梳理模式）",
        }
        return m.get(self.engine_type, str(self.engine_type))

    def get_readable_merge_mode(self):
        m = {
            MERGE_MODE_CATEGORY: "按类别拆卷合并",
            MERGE_MODE_ALL_IN_ONE: "全部合并成一个文件",
        }
        return m.get(self.merge_mode, self.merge_mode)

    def print_runtime_summary(self):
        """只打印信息，不做任何交互（GUI 可以直接用这个输出到日志）"""
        print("\n" + "=" * 60)
        print(" 运行参数总览")
        print("=" * 60)
        print(f"  源目录      : {self.config['source_folder']}")
        print(f"  目标目录    : {self.config['target_folder']}")
        print(f"  运行模式    : {self.get_readable_run_mode()} ({self.run_mode})")

        if self.run_mode == MODE_COLLECT_ONLY:
            print(
                f"  梳理子模式  : {self.get_readable_collect_mode()} ({self.collect_mode})"
            )

        if self.run_mode in (MODE_CONVERT_ONLY, MODE_CONVERT_THEN_MERGE):
            print(f"  内容策略    : {self.get_readable_content_strategy()}")
            print(f"  转换引擎    : {self.get_readable_engine_type()}")
            print(
                f"  进程处理策略: {self.config.get('kill_process_mode', KILL_MODE_ASK)}"
            )

        print(f"  排除文件夹  : {self.excluded_folders}")

        # 沙箱信息
        if self.config.get("enable_sandbox", True):
            print(f"  沙箱模式    : 启用")
            print(f"  临时转换目录: {self.temp_sandbox}")
        else:
            print(f"  沙箱模式    : 关闭（直接对源文件操作）")
            print(
                f"  临时转换目录: {self.temp_sandbox}  [仅用于中间 PDF，不拷贝源文件]"
            )

        # 合并信息
        if self.run_mode in (MODE_CONVERT_THEN_MERGE, MODE_MERGE_ONLY):
            if self.config.get("enable_merge", True):
                print(f"  启用合并    : 是")
                print(
                    f"  合并大小上限: {self.config.get('max_merge_size_mb', 80)} MB/卷"
                )
                print(
                    f"  合并模式    : {self.get_readable_merge_mode()} ({self.merge_mode})"
                )
            else:
                print(f"  启用合并    : 否")

        print("=" * 60)
        print("如有误，可修改 config.json 或在 GUI 中调整配置后再运行。\n")

    # =============== 信号 ===============

    def signal_handler(self, signum, frame):
        print("\n\n" + "!" * 60)
        print("  接收到终止信号 (Ctrl+C)！正在紧急停止...")
        self.is_running = False
        if not self.reuse_process and self.run_mode not in (
            MODE_MERGE_ONLY,
            MODE_COLLECT_ONLY,
        ):
            print("  正在清理后台 Office 进程...")
            self.cleanup_all_processes()

        if self.temp_sandbox and os.path.exists(self.temp_sandbox):
            try:
                shutil.rmtree(self.temp_sandbox, ignore_errors=True)
            except Exception:
                pass

        print("--> 程序已退出。")
        sys.exit(0)

    # =============== 进程管理 ===============

    def get_process_list(self, process_names):
        running_list = []
        try:
            cmd = "tasklist /FO CSV /NH"
            output = subprocess.check_output(cmd, shell=True).decode(
                "gbk", errors="ignore"
            )
            for line in output.splitlines():
                parts = line.split(",")
                if len(parts) > 1:
                    p_name = parts[0].strip('"').lower()
                    p_pid = parts[1].strip('"')
                    base_name = p_name.replace(".exe", "")
                    if base_name in process_names:
                        running_list.append(f"{p_name} (PID: {p_pid})")
        except Exception:
            pass
        return running_list

    def check_and_handle_running_processes(self):
        """注意：KILL_MODE_ASK 仅适用于 CLI；GUI 不要设置为 ask"""
        if self.run_mode in (MODE_MERGE_ONLY, MODE_COLLECT_ONLY):
            return

        mode = self.config.get("kill_process_mode", KILL_MODE_ASK)
        if mode == KILL_MODE_KEEP:
            self.reuse_process = True
            print("--> 根据配置，已启用[进程复用]模式（不主动杀进程）。")
            return

        target_apps = (
            ["wps", "et", "wpp", "wpscenter"]
            if self.engine_type == ENGINE_WPS
            else ["winword", "excel", "powerpnt"]
        )
        app_label = (
            "WPS Office" if self.engine_type == ENGINE_WPS else "Microsoft Office"
        )

        print("\n正在扫描系统中的", app_label, "进程...")
        running = self.get_process_list(target_apps)

        if not running:
            print("--> 未发现相关残留进程，环境干净。")
            return

        if mode == KILL_MODE_AUTO:
            print(f"--> [自动模式] 检测到 {len(running)} 个进程，正在清理...")
            self.cleanup_all_processes()
            self.reuse_process = False
            return

        # 只有 CLI 会走到这里（GUI 不要设置 kill_process_mode=ask）
        print("\n" + "!" * 60)
        print(f" [警告] 检测到以下 {app_label} 程序正在运行：")
        for p in running:
            print(f"  - {p}")
        print("!" * 60)
        print("请选择操作：")
        print("  [1] 强力清理：自动杀掉上述进程（推荐）")
        print("  [2] 尝试复用：不杀进程（如果有弹窗，可能导致失败）")
        print("  [3] 退出程序")
        print("-" * 60)

        while True:
            choice = input("请输入序号 (1-3): ").strip()
            if choice == "1":
                self.cleanup_all_processes()
                self.reuse_process = False
                break
            elif choice == "2":
                self.reuse_process = True
                break
            elif choice == "3":
                sys.exit(0)
            else:
                print("输入无效，请重新输入 1 / 2 / 3。")
        print()

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
        if not app_name:
            return
        try:
            cmd = f"taskkill /F /IM {app_name}.exe"
            subprocess.run(
                cmd, shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
            )
        except Exception:
            pass

    # =============== 配置 ===============

    def load_config(self, path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                content = f.read().replace("\\", "/")
                self.config = json.loads(content)
        except Exception as e:
            print(f"[错误] 无法加载配置文件: {e}")
            sys.exit(1)

        self.config["source_folder"] = os.path.abspath(self.config["source_folder"])
        self.config["target_folder"] = os.path.abspath(self.config["target_folder"])

        # 通用默认值
        if "timeout_seconds" not in self.config:
            self.config["timeout_seconds"] = 60
        if "enable_sandbox" not in self.config:
            self.config["enable_sandbox"] = True
        if "default_engine" not in self.config:
            self.config["default_engine"] = ENGINE_ASK
        if "kill_process_mode" not in self.config:
            self.config["kill_process_mode"] = KILL_MODE_ASK
        if "auto_retry_failed" not in self.config:
            self.config["auto_retry_failed"] = False
        if "pdf_wait_seconds" not in self.config:
            self.config["pdf_wait_seconds"] = 15
        if "ppt_timeout_seconds" not in self.config:
            self.config["ppt_timeout_seconds"] = self.config["timeout_seconds"]
        if "ppt_pdf_wait_seconds" not in self.config:
            self.config["ppt_pdf_wait_seconds"] = self.config["pdf_wait_seconds"]
        if "enable_merge" not in self.config:
            self.config["enable_merge"] = True
        if "max_merge_size_mb" not in self.config:
            self.config["max_merge_size_mb"] = 80
        if "temp_sandbox_root" not in self.config:
            self.config["temp_sandbox_root"] = ""

        if "merge_mode" not in self.config:
            self.config["merge_mode"] = MERGE_MODE_CATEGORY

        # 关键字 & 排除目录
        if "price_keywords" not in self.config:
            self.config["price_keywords"] = ["报价", "价格表", "Price", "Quotation"]
        self.price_keywords = self.config["price_keywords"]

        if "excluded_folders" not in self.config:
            self.config["excluded_folders"] = ["temp", "backup", "archive"]
        self.excluded_folders = self.config["excluded_folders"]

        # 允许扩展名
        exts = self.config.setdefault("allowed_extensions", {})
        exts.setdefault("word", [])
        exts.setdefault("excel", [])
        exts.setdefault("powerpoint", [])
        if "pdf" not in exts:
            exts["pdf"] = [".pdf"]

        # 从配置中恢复一些模式默认值（GUI 存进去后可用）
        self.merge_mode = self.config.get("merge_mode", MERGE_MODE_CATEGORY)
        self.run_mode = self.config.get("run_mode", self.run_mode)
        self.collect_mode = self.config.get("collect_mode", self.collect_mode)
        self.content_strategy = self.config.get(
            "content_strategy", self.content_strategy
        )

    def save_config(self):
        try:
            with open(self.config_path, "w", encoding="utf-8") as f:
                json.dump(self.config, f, indent=4, ensure_ascii=False)
            print("--> 配置文件已更新。")
        except Exception:
            pass

    # =============== CLI 交互向导（GUI 不会用） ===============

    def cli_wizard(self):
        """仅 CLI 模式调用：做一轮交互式向导，然后准备好所有参数"""
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

        # 引擎相关
        if self.run_mode == MODE_MERGE_ONLY:
            print("\n[模式] 当前为【仅合并】模式，不需要启动 Office 引擎。")
        elif self.run_mode == MODE_COLLECT_ONLY:
            print("\n[模式] 当前为【文件梳理去重】模式，不涉及 Office 转换。")
        else:
            self.select_engine_mode()
            self.check_and_handle_running_processes()

        # 目录可能因 ask_for_subfolder 改变，重新初始化
        self._init_paths_from_config()

    def confirm_config_in_terminal(self):
        self.print_step_title("步骤 1 / 4 ：确认源 / 目标目录")
        print("当前配置路径：")
        print(f"  源目录 : {self.config['source_folder']}")
        print(f"  目标目录 : {self.config['target_folder']}")
        print("-" * 60)

        while True:
            choice = (
                input("是否需要修改上述路径？[Y/n] (回车默认为 n): ").strip().lower()
            )
            if choice in ("", "n"):
                break
            elif choice == "y":
                print("\n=== 配置修改模式 ===")
                print(f"当前源目录: {self.config['source_folder']}")
                new_s = (
                    input("请输入新源目录 [回车保持不变]: ")
                    .strip()
                    .replace('"', "")
                    .replace("'", "")
                )
                if new_s:
                    self.config["source_folder"] = os.path.abspath(new_s)

                print(f"\n当前目标目录: {self.config['target_folder']}")
                new_t = (
                    input("请输入新目标目录 [回车保持不变]: ")
                    .strip()
                    .replace('"', "")
                    .replace("'", "")
                )
                if new_t:
                    self.config["target_folder"] = os.path.abspath(new_t)

                self.save_config()
                print("配置已保存。")
                print("-" * 60)
                print("更新后的配置路径：")
                print(f"  源目录 : {self.config['source_folder']}")
                print(f"  目标目录 : {self.config['target_folder']}")
                print("-" * 60)
            else:
                print("输入无效，请输入 Y 或 N。")
        print("--> 路径确认完毕。\n")

    def ask_for_subfolder(self):
        self.print_step_title("步骤 2 / 4 ：设置本次输出子文件夹（可选）")
        print("可以为本次任务创建一个子文件夹，方便与历史任务区分。")
        print("-" * 60)
        sub = input("请输入本次输出子文件夹名称 (直接回车表示不创建): ").strip()
        if sub:
            for char in '<>:"/\\|?*':
                sub = sub.replace(char, "")
            self.config["target_folder"] = os.path.abspath(
                os.path.join(self.config["target_folder"], sub)
            )
            print(f"--> 已为本次任务指定子目录：{self.config['target_folder']}")
        else:
            print("--> 使用配置中的目标目录，不创建额外子文件夹。")

    def select_run_mode(self):
        self.print_step_title("步骤 3 / 4 ：选择运行模式")
        print("  [1] 仅转换")
        print("  [2] 仅合并")
        print("  [3] 先转换再合并（推荐）")
        print("  [4] 文件梳理去重")
        print("-" * 60)
        choice = input("请输入序号 (1/2/3/4，回车默认为 3): ").strip()
        if choice == "1":
            self.run_mode = MODE_CONVERT_ONLY
        elif choice == "2":
            self.run_mode = MODE_MERGE_ONLY
        elif choice == "4":
            self.run_mode = MODE_COLLECT_ONLY
        else:
            self.run_mode = MODE_CONVERT_THEN_MERGE
        print(f"--> 运行模式: {self.get_readable_run_mode()} ({self.run_mode})")

    def select_collect_mode(self):
        self.print_step_title("文件梳理子模式选择")
        print("  [1] 去重 + 拷贝 + Excel 索引")
        print("  [2] 仅生成 Excel 索引（不拷贝）")
        print("-" * 60)
        choice = input("请输入序号 (1/2，回车默认为 1): ").strip()
        if choice == "2":
            self.collect_mode = COLLECT_MODE_INDEX_ONLY
        else:
            self.collect_mode = COLLECT_MODE_COPY_AND_INDEX
        print(
            f"--> 梳理子模式: {self.get_readable_collect_mode()} ({self.collect_mode})"
        )

    def select_merge_mode(self):
        if not self.config.get("enable_merge", True):
            self.merge_mode = MERGE_MODE_CATEGORY
            return

        cfg_mode = self.config.get("merge_mode", MERGE_MODE_CATEGORY)
        if cfg_mode in (MERGE_MODE_ALL_IN_ONE, MERGE_MODE_CATEGORY):
            self.merge_mode = cfg_mode
            print(
                f"--> 按配置使用合并模式: {self.get_readable_merge_mode()} ({self.merge_mode})"
            )
            return

        self.print_step_title("合并模式选择")
        print("  [1] 分类拆卷（按 Price_/Word_/Excel_/PPT_/PDF_ 分类合并）")
        print("  [2] 全部合并成一个 PDF 文件（忽略分类）")
        print("-" * 60)
        choice = input("请输入序号 (1/2，回车默认为 1): ").strip()
        if choice == "2":
            self.merge_mode = MERGE_MODE_ALL_IN_ONE
        else:
            self.merge_mode = MERGE_MODE_CATEGORY
        print(f"--> 本次合并模式: {self.get_readable_merge_mode()} ({self.merge_mode})")

    def select_content_strategy(self):
        self.print_step_title("步骤 4 / 4 ：选择内容处理策略（只在涉及转换时生效）")
        print("  [1] 标准分类")
        print("  [2] 智能标记（识别报价）")
        print("  [3] 报价猎手（仅报价相关）")
        print("-" * 60)
        print(f"当前关键字: {self.price_keywords}")
        choice = input("请输入序号 (1/2/3，回车默认为 1): ").strip()
        if choice == "2":
            self.content_strategy = STRATEGY_SMART_TAG
        elif choice == "3":
            self.content_strategy = STRATEGY_PRICE_ONLY
        else:
            self.content_strategy = STRATEGY_STANDARD
        print(
            f"--> 策略: {self.get_readable_content_strategy()} ({self.content_strategy})\n"
        )

    def select_engine_mode(self):
        default = self.config.get("default_engine", ENGINE_ASK)
        if default == ENGINE_WPS:
            self.engine_type = ENGINE_WPS
            print("--> [自动选择] 使用引擎: WPS Office")
            return
        elif default == ENGINE_MS:
            self.engine_type = ENGINE_MS
            print("--> [自动选择] 使用引擎: Microsoft Office")
            return

        self.print_step_title("Office 转换引擎选择")
        print("  [1] WPS Office")
        print("  [2] Microsoft Office")
        print("-" * 60)
        while True:
            choice = input("请输入序号 (1/2，回车默认为 1-WPS): ").strip()
            if choice in ("", "1"):
                self.engine_type = ENGINE_WPS
                break
            elif choice == "2":
                self.engine_type = ENGINE_MS
                break
            else:
                print("输入无效，请输入 1 或 2。")
        print(f"--> 已选择: {self.get_readable_engine_type()} ({self.engine_type})\n")

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
            f.write(f"{datetime.now()} === 任务开始 (v{__version__}) ===\n")
            f.write(f"运行模式: {self.run_mode} ({self.get_readable_run_mode()})\n")
            if self.run_mode in (MODE_CONVERT_ONLY, MODE_CONVERT_THEN_MERGE):
                f.write(
                    f"内容策略: {self.content_strategy} ({self.get_readable_content_strategy()})\n"
                )
                f.write(f"使用引擎: {engine_label}\n")
            if self.run_mode == MODE_COLLECT_ONLY:
                f.write(
                    f"梳理子模式: {self.collect_mode} ({self.get_readable_collect_mode()})\n"
                )
            if self.run_mode in (
                MODE_CONVERT_THEN_MERGE,
                MODE_MERGE_ONLY,
            ) and self.config.get("enable_merge", True):
                f.write(
                    f"合并模式: {self.merge_mode} ({self.get_readable_merge_mode()})\n"
                )
            f.write(f"源目录: {self.config['source_folder']}\n")
            f.write(f"目标目录: {self.config['target_folder']}\n")
            f.write(f"排除文件夹: {self.excluded_folders}\n")

            if self.config.get("enable_sandbox", True):
                f.write(f"沙箱模式: 启用 | 临时转换目录: {self.temp_sandbox}\n")
            else:
                f.write(
                    f"沙箱模式: 关闭 | 临时转换目录: {self.temp_sandbox}（仅中间PDF）\n"
                )

            if self.run_mode != MODE_COLLECT_ONLY:
                f.write(f"合并输出目录: {self.merge_output_dir}\n")

    # =============== Office 应用管理 ===============

    def _kill_current_app(self, app_type):
        if self.reuse_process:
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

    # =============== 路径与前缀逻辑 ===============

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
            return "成功", target_pdf_path

        if os.path.getsize(temp_pdf_path) == os.path.getsize(target_pdf_path):
            try:
                os.remove(target_pdf_path)
                shutil.move(temp_pdf_path, target_pdf_path)
                return "覆盖", target_pdf_path
            except Exception:
                return "覆盖失败", target_pdf_path
        else:
            conflict_dir = os.path.join(os.path.dirname(target_pdf_path), "conflicts")
            os.makedirs(conflict_dir, exist_ok=True)
            fname = os.path.splitext(os.path.basename(target_pdf_path))[0]
            ts = datetime.now().strftime("%Y%m%d%H%M%S")
            new_path = os.path.join(conflict_dir, f"{fname}_{ts}.pdf")
            shutil.move(temp_pdf_path, new_path)
            return "冲突备份", new_path

    # =============== 内容扫描 ===============

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
                                            f"Excel匹配到关键词 [{kw}] in Sheet: {sheet.Name}"
                                        )
                                        return True
                except Exception:
                    continue
        except Exception as e:
            logging.warning(f"扫描Excel内容失败: {e}")
        return False

    # =============== COM 安全调用辅助 ===============

    def _safe_exec(self, func, *args, retries=3, **kwargs):
        for attempt in range(retries + 1):
            if not self.is_running:
                raise Exception("程序已终止")
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
                raise Exception(f"COM错误 ({error_code}): {e}")
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

    # =============== 核心转换 ===============

    def convert_logic_in_thread(
        self, file_source, sandbox_target_pdf, ext, result_context
    ):
        app = None
        doc = None
        try:
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
                    app.Quit()
                except Exception:
                    pass
            pythoncom.CoUninitialize()

    def copy_pdf_direct(self, source, temp_target):
        try:
            shutil.copy2(source, temp_target)
        except Exception as e:
            raise Exception(f"[PDF复制失败] {e}")

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

    def process_single_file(
        self, file_path, target_path_initial, ext, progress_str, is_retry=False
    ):
        if os.path.getsize(file_path) == 0:
            self.stats["skipped"] += 1
            logging.warning(f"跳过空文件: {file_path}")
            return "跳过(0KB)", target_path_initial

        is_word = ext in self.config["allowed_extensions"].get("word", [])
        is_excel = ext in self.config["allowed_extensions"].get("excel", [])
        is_ppt = ext in self.config["allowed_extensions"].get("powerpoint", [])
        is_pdf = ext == ".pdf"

        filename = os.path.basename(file_path)

        # Step 1: 文件名优先匹配
        is_filename_match = False
        if self.content_strategy != STRATEGY_STANDARD:
            for kw in self.price_keywords:
                if kw in filename:
                    is_filename_match = True
                    break

        if self.content_strategy == STRATEGY_PRICE_ONLY and not is_filename_match:
            if is_word or is_ppt:
                self.stats["skipped"] += 1
                return "跳过(类型)", target_path_initial

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
            if use_sandbox:
                sandbox_src_path = os.path.join(
                    self.temp_sandbox, f"{uuid.uuid4()}{ext}"
                )
                shutil.copy2(file_path, sandbox_src_path)
                self._unblock_file(sandbox_src_path)
                working_src = sandbox_src_path

            if is_pdf:
                if not is_filename_match and self.content_strategy != STRATEGY_STANDARD:
                    has_kw = self.scan_pdf_content(working_src)
                    if has_kw:
                        result_context["is_price"] = True
                    elif self.content_strategy == STRATEGY_PRICE_ONLY:
                        self.stats["skipped"] += 1
                        return "跳过(无关键字)", target_path_initial

                self.copy_pdf_direct(working_src, sandbox_pdf)

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
                    # CLI 进度输出（GUI 下日志也会看到）
                    print(
                        f"\r{progress_str} 正在处理: {filename} ({elapsed:.1f}s)    ",
                        end="",
                        flush=True,
                    )
                    time.sleep(0.1)

                convert_thread.join(timeout=0.1)

                if convert_thread.is_alive():
                    self.stats["timeout"] += 1
                    logging.error(f"超时跳过 (>{current_timeout}s)")
                    if not self.reuse_process:
                        if is_word:
                            self._kill_current_app("word")
                        elif is_excel:
                            self._kill_current_app("excel")
                        elif is_ppt:
                            self._kill_current_app("ppt")
                    raise Exception("超时")

            if result_context["scan_aborted"]:
                self.stats["skipped"] += 1
                return "跳过(无关键字)", target_path_initial

            if result_context["is_price"]:
                final_target_path = self.get_target_path(
                    file_path, ext, prefix_override="Price_"
                )

            wait_pdf_start = time.time()
            while time.time() - wait_pdf_start < current_pdf_wait:
                if os.path.exists(sandbox_pdf):
                    time.sleep(0.5)
                    result_status, final_path_res = self.handle_file_conflict(
                        sandbox_pdf, final_target_path
                    )
                    self.generated_pdfs.append(final_path_res)

                    tag_info = ""
                    if is_filename_match:
                        tag_info = " [文件名命中]"
                    elif result_context["is_price"]:
                        tag_info = " [内容命中]"

                    return f"{result_status}{tag_info}", final_path_res
                time.sleep(0.5)

            raise Exception(
                f"转换指令已发送但未生成PDF ({current_pdf_wait}s内未检测到文件)"
            )

        finally:
            try:
                if sandbox_src_path and os.path.exists(sandbox_src_path):
                    os.remove(sandbox_src_path)
                if os.path.exists(sandbox_pdf):
                    os.remove(sandbox_pdf)
            except Exception:
                pass

    # =============== 批处理 / 重试 ===============

    def get_progress_prefix(self, current, total):
        width = len(str(total)) if total > 0 else 1
        percent = current / total if total else 0
        bar_len = 20
        filled = int(bar_len * percent)
        bar = "█" * filled + "░" * (bar_len - filled)
        return f"[{int(percent * 100):>3}%]{bar} [{str(current).rjust(width)}/{total}]"

    def run_batch(self, file_list, is_retry=False):
        total = len(file_list)
        for i, fpath in enumerate(file_list, 1):
            if not self.is_running:
                break

            fname = os.path.basename(fpath)
            ext = os.path.splitext(fpath)[1].lower()
            target_path_initial = self.get_target_path(fpath, ext)

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

                if "跳过" in status:
                    print(
                        f"\r{progress_prefix} {status}: {fname} (耗时: {elapsed:.2f}s)    "
                    )
                    logging.info(f"{status}: {fpath}")
                else:
                    self.stats["success"] += 1
                    print(
                        f"\r{progress_prefix} {status}: {fname} (耗时: {elapsed:.2f}s)    "
                    )
                    logging.info(f"{status}: {fpath} -> {final_path}")

            except Exception as e:
                elapsed = time.time() - start
                err_msg = str(e)
                if "超时" in err_msg:
                    print(
                        f"\r{progress_prefix} 超时: {fname} (耗时: {elapsed:.2f}s)    "
                    )
                else:
                    self.stats["failed"] += 1
                    print(
                        f"\r{progress_prefix} 失败: {fname} (耗时: {elapsed:.2f}s)    "
                    )

                logging.error(f"失败: {fpath} | 原因: {e}")

                if not is_retry:
                    self.quarantine_failed_file(fpath)
                    self.error_records.append(fpath)

    def ask_retry_failed_files(self, failed_count, timeout=20):
        print("\n" + "=" * 60)
        print(f"[警告] 本次共有 {failed_count} 个文件处理失败（含超时）。")
        if self.error_records:
            print("失败样例（最多 10 条）：")
            for p in self.error_records[:10]:
                print("  -", p)
        print("-" * 60)
        print("是否尝试重新处理这些失败的文件？")
        print("  输入 Y 然后回车  -> 进行重试")
        print("  输入 N 然后回车  -> 不重试")
        print(f"  如果 {timeout} 秒内未输入，则默认【不重试】")
        print("=" * 60)

        if not HAS_MSVCRT:
            ans = input("请输入 [Y/N] 并回车确认（无超时限制）: ").strip().lower()
            return ans == "y"

        buf = ""
        start = time.time()
        last_shown = None

        while True:
            elapsed = time.time() - start
            remain = int(timeout - elapsed)
            if remain < 0:
                print("\n[提示] 已超过倒计时，默认不重试失败文件。")
                return False

            if last_shown != remain:
                print(
                    f"\r请在 {remain:2d} 秒内输入 [Y/N] 并回车确认: {buf}",
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
                        print("[选择] 将尝试重试失败文件。\n")
                        return True
                    else:
                        print("[选择] 不重试失败文件。\n")
                        return False
                elif ch == "\b":
                    buf = buf[:-1]
                else:
                    buf += ch
            time.sleep(0.1)

    # =============== 合并功能 ===============

    def merge_pdfs(self):
        if not self.config.get("enable_merge", True):
            return
        if not HAS_PYPDF:
            print(
                "\n[提示] 未检测到 pypdf 库，跳过合并步骤。请运行 pip install pypdf 安装。"
            )
            logging.warning("未检测到 pypdf 库，跳过合并。")
            return

        print("\n" + "=" * 60)
        print("  开始 PDF 合并 ...")
        print(f"  合并模式: {self.get_readable_merge_mode()} ({self.merge_mode})")
        print(f"  合并输出目录: {self.merge_output_dir}")
        print("=" * 60)

        scan_folder = self.config["target_folder"]
        
        # 确定扫描来源：默认目标目录
        # 如果是“仅合并”模式，根据配置决定是源目录还是目标目录
        scan_source_type = "target"
        if self.run_mode == MODE_MERGE_ONLY:
            scan_source_type = self.config.get("merge_source", "source")
        
        if scan_source_type == "source":
            scan_folder = self.config["source_folder"]
            print(f"  [仅合并模式] 正在扫描源目录: {scan_folder}")
        else:
            # convert_then_merge 或 merge_only(选了target)
            print(f"  [合并扫描] 正在扫描目标目录: {scan_folder}")

        all_pdfs = []
        # 如果目标目录在源目录里（或者是子目录），要防止扫描到输出目录
        exclude_abs_paths = set(map(os.path.abspath, [self.failed_dir, self.merge_output_dir]))
        
        # 如果扫描的是源目录，且目标目录也在里面，则排除目标目录
        # 如果扫描的是目标目录，本身就在里面，自然会扫描到（除了 excluded 的 failed/merged）
        if scan_source_type == "source":
             exclude_abs_paths.add(os.path.abspath(self.config["target_folder"]))

        for root, dirs, files in os.walk(scan_folder):
            if os.path.abspath(root) in exclude_abs_paths:
                continue
            for f in files:
                if f.lower().endswith(".pdf"):
                    all_pdfs.append(os.path.join(root, f))

        if not all_pdfs:
            print("[提示] 目标目录中没有 PDF 文件，无法合并。")
            return

        all_pdfs.sort()

        if self.merge_mode == MERGE_MODE_ALL_IN_ONE:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_name = f"Merged_All_{timestamp}.pdf"
            output_path = os.path.join(self.merge_output_dir, output_name)
            print(f">> 全量合并，共 {len(all_pdfs)} 个 PDF -> {output_name}")
            try:
                merger = PdfWriter()
                for p in all_pdfs:
                    merger.append(p)
                merger.write(output_path)
                merger.close()
                logging.info(f"全量合并成功: {output_path}")
                print("--> 全部合并完成。")
            except Exception as e:
                logging.error(f"全量合并失败: {e}")
                print(f"[错误] 全部合并失败: {e}")
            return

        # 分类拆卷模式（原有行为）
        categories = {
            "报价单文件": "Price_",
            "Word文档": "Word_",
            "Excel表格": "Excel_",
            "PPT幻灯片": "PPT_",
            "原PDF文件": "PDF_",
        }

        max_size_bytes = self.config.get("max_merge_size_mb", 80) * 1024 * 1024
        total_merged_count = 0

        for cat_name, prefix in categories.items():
            current_cat_files = [
                p for p in all_pdfs if os.path.basename(p).startswith(prefix)
            ]
            if not current_cat_files:
                continue
            current_cat_files.sort()
            print(f"\n>> 正在处理类别: {cat_name} (共 {len(current_cat_files)} 个文件)")

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

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            cat_label = prefix.rstrip("_")

            for idx, group in enumerate(groups, 1):
                output_filename = f"Merged_{cat_label}_{timestamp}_{idx}.pdf"
                output_path = os.path.join(self.merge_output_dir, output_filename)

                print(f"   生成第 {idx} 卷 ({len(group)} 个文件)...", end="")
                try:
                    merger = PdfWriter()
                    for pdf in group:
                        merger.append(pdf)
                    merger.write(output_path)
                    merger.close()
                    print(" [完成]")
                    logging.info(f"分类合并成功 [{cat_name}]: {output_path}")
                    total_merged_count += 1
                except Exception as e:
                    print(f" [失败] {e}")
                    logging.error(f"分类合并失败 [{cat_name}]: {output_path} | {e}")

        print(f"\n--> 分类拆卷合并完成，共生成 {total_merged_count} 个汇总文件。")

    # =============== 文件梳理去重 ==================

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

    def collect_office_files_and_build_excel(self):
        if not HAS_OPENPYXL:
            print("\n[错误] 未检测到 openpyxl 库，无法生成 Excel 报表。")
            print("请运行: pip install openpyxl  然后重试。")
            logging.error("缺少 openpyxl 库，collect_only 模式无法执行。")
            return

        source_root = self.config["source_folder"]
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
        print(" 文件梳理去重模式")
        print("=" * 60)
        print(f" 源目录   : {source_root}")
        print(f" 目标目录 : {target_root}")
        print(f" 子模式   : {self.get_readable_collect_mode()} ({self.collect_mode})")
        print(f" 筛选类型 : {office_exts}")
        print("=" * 60)

        all_files = []
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
        print(f"共扫描到 Office 文件: {total} 个。")
        logging.info(f"[collect_only] 扫描到 Office 文件: {total} 个。")

        if total == 0:
            print("[提示] 没有找到 Office 文件，梳理结束。")
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
                rel = os.path.relpath(src_path, source_root)
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
                    rel = os.path.relpath(src_path, source_root)
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
                    keep_rel = os.path.relpath(keep_src, source_root)
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

        print(f"\n去重完成：")
        print(f"  唯一文件 : {len(unique_records)} 个")
        print(f"  重复文件 : {len(duplicate_records)} 个")
        logging.info(
            f"[collect_only] 唯一文件: {len(unique_records)}，重复文件: {len(duplicate_records)}"
        )

        copied_count = 0
        if self.collect_mode == COLLECT_MODE_COPY_AND_INDEX:
            print("\n正在拷贝唯一文件到目标目录...")
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
                    logging.error(f"[collect_only] 拷贝失败: {src} -> {dst} | {e}")
                    rec["copied"] = False

                if idx % 20 == 0 or idx == len(unique_records):
                    print(
                        f"\r已处理 {idx}/{len(unique_records)} 个唯一文件...",
                        end="",
                        flush=True,
                    )
            print(f"\r拷贝完成，成功拷贝 {copied_count} 个文件。          ")
        else:
            print("\n当前为【仅生成 Excel 索引】模式，不执行文件拷贝。")
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
                "序号",
                "组ID",
                "文件名",
                "扩展名",
                "大小(KB)",
                "源文件路径",
                "目标文件路径",
            ]
        else:
            headers_unique = [
                "序号",
                "组ID",
                "文件名",
                "扩展名",
                "大小(KB)",
                "源文件路径",
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
                "序号",
                "组ID",
                "文件名",
                "扩展名",
                "大小(KB)",
                "源文件路径",
                "保留文件目标路径",
            ]
        else:
            headers_dup = [
                "序号",
                "组ID",
                "文件名",
                "扩展名",
                "大小(KB)",
                "源文件路径",
                "保留文件源路径",
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
        print(f"\nExcel 索引已生成：{excel_path}")
        logging.info(f"[collect_only] Excel 索引生成: {excel_path}")

        print("\n=== 文件梳理去重结果 ===")
        print(f"扫描总数   : {total}")
        print(
            f"唯一文件   : {len(unique_records)} (拷贝成功 {copied_count}，仅拷贝模式有效)"
        )
        print(f"重复文件   : {len(duplicate_records)}")
        print(f"索引文件   : {excel_path}")
        print("========================\n")

    # =============== 主流程 ===============

    def run(self):
        # 日志初始化 & 参数总览
        self.setup_logging()
        self.print_runtime_summary()

        if self.run_mode == MODE_COLLECT_ONLY:
            self.collect_office_files_and_build_excel()
        else:
            # 转换阶段
            if self.run_mode in (MODE_CONVERT_ONLY, MODE_CONVERT_THEN_MERGE):
                files = []
                logging.info("正在扫描文件...")

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

                if not os.path.exists(self.config["source_folder"]):
                    print(f"\n[警告] 源目录不存在: {self.config['source_folder']}")
                    logging.error(f"源目录不存在: {self.config['source_folder']}")
                else:
                    valid_exts = [
                        e
                        for sub in self.config["allowed_extensions"].values()
                        for e in sub
                    ]
                    for root, dirs, filenames in os.walk(self.config["source_folder"]):
                        dirs[:] = [
                            d
                            for d in dirs
                            if d.lower() not in excl_names
                            and os.path.abspath(os.path.join(root, d)).lower()
                            not in excl_paths
                        ]
                        for f in filenames:
                            if not f.startswith("~$"):
                                ext = os.path.splitext(f)[1].lower()
                                if ext in valid_exts:
                                    files.append(os.path.join(root, f))

                    logging.info(f"开始处理 {len(files)} 个文件...")
                    self.stats["total"] = len(files)

                    if len(files) > 0:
                        self.run_batch(files)
                    else:
                        print("\n[提示] 源目录中没有发现可转换的 Office 文件。")

                self.close_office_apps()

                failed_count = self.stats["failed"] + self.stats["timeout"]
                should_retry = False
                if failed_count > 0:
                    if self.config.get("auto_retry_failed", False):
                        should_retry = True
                        print(f"\n[配置] 自动重试失败文件 ({failed_count}个)...")
                    elif self.interactive:
                        should_retry = self.ask_retry_failed_files(
                            failed_count, timeout=20
                        )

                if should_retry:
                    print("\n" + "=" * 60)
                    print("  开始重试失败文件...")
                    print("  正在重新检查并清理进程...")
                    print("=" * 60)

                    if not self.reuse_process:
                        self.cleanup_all_processes()

                    retry_files = []
                    if os.path.exists(self.failed_dir):
                        if self.config.get(
                            "enable_sandbox", True
                        ) and not os.path.exists(self.temp_sandbox):
                            os.makedirs(self.temp_sandbox)

                        valid_exts = [
                            e
                            for sub in self.config["allowed_extensions"].values()
                            for e in sub
                        ]
                        for f in os.listdir(self.failed_dir):
                            if not f.startswith("~$"):
                                ext = os.path.splitext(f)[1].lower()
                                if ext in valid_exts:
                                    retry_files.append(os.path.join(self.failed_dir, f))

                    if retry_files:
                        self.run_batch(retry_files, is_retry=True)
                    else:
                        print("未在失败目录找到可重试的文件。")

                    self.close_office_apps()

            elif self.run_mode == MODE_MERGE_ONLY:
                print("当前模式为仅合并，跳过转换步骤。")

            # 合并阶段
            if self.run_mode in (
                MODE_CONVERT_THEN_MERGE,
                MODE_MERGE_ONLY,
            ) and self.config.get("enable_merge", True):
                self.merge_pdfs()

            summary = (
                f"\n=== 最终统计 (v{__version__}) ===\n"
                f"总处理: {self.stats['total']}\n"
                f"成功: {self.stats['success']}\n"
                f"失败: {self.stats['failed']}\n"
                f"超时: {self.stats['timeout']}\n"
                f"跳过(含空文件/策略): {self.stats['skipped']}\n"
            )
            logging.info(summary)
            print(summary)

        # 打开目标目录
        try:
            os.startfile(self.config["target_folder"])
        except Exception:
            pass

        # 清理沙箱
        if self.temp_sandbox and os.path.exists(self.temp_sandbox):
            try:
                shutil.rmtree(self.temp_sandbox, ignore_errors=True)
            except Exception:
                pass


def create_default_config(config_path):
    """生成默认配置文件"""
    try:
        default_config = {
            "source_folder": "C:\\Docs",
            "target_folder": "C:\\PDFs",
            "log_folder": "./logs",
            "enable_sandbox": True,
            "default_engine": "ask",
            "kill_process_mode": "ask",
            "auto_retry_failed": False,
            "timeout_seconds": 60,
            "pdf_wait_seconds": 15,
            "ppt_timeout_seconds": 180,
            "ppt_pdf_wait_seconds": 30,
            "enable_merge": True,
            "max_merge_size_mb": 80,
            "price_keywords": ["报价", "价格表", "Price", "Quotation"],
            "excluded_folders": ["temp", "backup", "archive"],
            "allowed_extensions": {
                "word": [".doc", ".docx"],
                "excel": [".xls", ".xlsx"],
                "powerpoint": [".ppt", ".pptx"],
                "pdf": [".pdf"],
            },
            "overwrite_same_size": True,
            "merge_mode": MERGE_MODE_CATEGORY,
            "merge_source": "source",
            "temp_sandbox_root": "",
        }
        with open(config_path, "w", encoding="utf-8") as f:
            json.dump(default_config, f, indent=4, ensure_ascii=False)
        print(f"已生成默认配置文件: {config_path}")
        return True
    except Exception as e:
        print(f"生成默认配置失败: {e}")
        return False


# =============== 命令行入口 ===============

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
