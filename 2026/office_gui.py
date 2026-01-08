# -*- coding: utf-8 -*-
"""
office_gui.py - Office 文档批量转换 & 梳理工具 GUI 版

说明：
- 依赖 office_converter.py 中的 OfficeConverter（你已经更新到 v5.15.1）
- GUI 中：
    * “运行参数”页：选择源/目标目录、运行模式、内容策略、合并模式、沙箱等
    * “配置管理”页：直接编辑 config.json 的部分配置（日志目录、排除目录、关键字、超时参数等）
- “保存配置”按钮：写入 config.json
- “开始运行”按钮：用当前界面参数启动转换/梳理（不会自动改 config.json）
- “停止”按钮：设置 converter.is_running=False，优雅停止
"""

import os
import sys
import json
import threading
import queue
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import tempfile
import traceback
from datetime import datetime

from office_converter import (
    OfficeConverter,
    MODE_CONVERT_ONLY,
    MODE_MERGE_ONLY,
    MODE_CONVERT_THEN_MERGE,
    MODE_COLLECT_ONLY,
    COLLECT_MODE_COPY_AND_INDEX,
    COLLECT_MODE_INDEX_ONLY,
    MERGE_MODE_CATEGORY,
    MERGE_MODE_ALL_IN_ONE,
    ENGINE_WPS,
    ENGINE_MS,
    KILL_MODE_AUTO,
    KILL_MODE_KEEP,
    __version__,
    get_app_path,
    create_default_config,
)

LOG_QUEUE = queue.Queue()


class TkLogHandler:
    """简单 handler：把 stdout/stderr 文本丢到队列，由 GUI 定时刷到 Text 里"""

    def write(self, msg: str):
        msg = msg.rstrip("\n")
        if msg:
            LOG_QUEUE.put(msg)

    def flush(self):
        pass


# ========= GUI 专用 Converter 子类：屏蔽 CLI 输入 =========


class GUIOfficeConverter(OfficeConverter):
    """
    覆盖所有会在 __init__ 中调用 input() 的方法，使其在 GUI 环境下不弹出 CLI 交互。
    run_mode / collect_mode / engine_type 等由 GUI 在实例化之后再覆盖。
    """

    def print_welcome(self):
        # GUI 环境下不需要在控制台打印欢迎界面（日志里会有头部）
        pass

    def confirm_config_in_terminal(self):
        # 不在 GUI 环境下再次询问源/目标目录
        pass

    def ask_for_subfolder(self):
        # GUI 中不做“子目录”询问；如需子目录可直接在目标路径里体现
        pass

    def select_run_mode(self):
        # 运行模式由 GUI 设置
        pass

    def select_collect_mode(self):
        # 梳理子模式由 GUI 设置
        pass

    def select_merge_mode(self):
        # 合并模式由 GUI 设置（配置或界面 Radio）
        pass

    def select_content_strategy(self):
        # 内容策略由 GUI 设置
        pass

    def select_engine_mode(self):
        # 引擎类型由 GUI 设置
        pass

    def print_runtime_summary(self):
        # 简化：GUI 环境下不打印 CLI 风格总览，日志由 GUI 捕获 print 即可
        pass


# ========= 主窗口 =========


class OfficeGUI(tk.Tk):
    def __init__(self, config_path=None):
        super().__init__()
        self.title(f"Office 批量转换 & 梳理工具 v{__version__}")
        self.geometry("980x680")
        self.minsize(900, 620)

        self.script_dir = get_app_path()
        self.config_path = config_path or os.path.join(self.script_dir, "config.json")

        # 当前任务线程 & 转换器实例
        self.worker_thread = None
        self.current_converter = None
        self.stop_requested = False

        # 把 stdout/stderr 重定向到 GUI 日志窗口
        sys.stdout = TkLogHandler()
        sys.stderr = TkLogHandler()

        if not os.path.exists(self.config_path):
            success = create_default_config(self.config_path)
            if success:
                messagebox.showinfo(
                    "提示",
                    "未找到 config.json，已为您自动生成默认配置。\n",
                )

        self._build_ui()
        self._load_config_to_ui()

        # 定时刷新日志
        self.after(200, self._poll_log_queue)

    # ===================== UI 构建 =====================

    def _build_ui(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)

        # 顶部 Notebook：运行参数 / 配置管理
        notebook = ttk.Notebook(self)
        notebook.grid(row=0, column=0, sticky="nsew", padx=8, pady=6)
        self.notebook = notebook

        tab_run = ttk.Frame(notebook)
        tab_cfg = ttk.Frame(notebook)

        notebook.add(tab_run, text="运行参数")
        notebook.add(tab_cfg, text="配置管理")

        # ===== 运行参数页 =====
        self._build_run_tab(tab_run)

        # ===== 配置管理页 =====
        self._build_config_tab(tab_cfg)

        # 状态 + 进度条
        status_frame = ttk.Frame(self)
        status_frame.grid(row=1, column=0, sticky="ew", padx=8)
        status_frame.columnconfigure(0, weight=1)

        self.var_status = tk.StringVar(value="就绪")
        self.status_label = ttk.Label(
            status_frame, textvariable=self.var_status, anchor="w"
        )
        self.status_label.grid(row=0, column=0, sticky="ew")

        self.progress = ttk.Progressbar(
            status_frame, orient="horizontal", mode="determinate", length=100
        )
        self.progress.grid(row=1, column=0, sticky="ew", pady=(2, 4))

        # 日志区
        log_frame = ttk.LabelFrame(self, text="运行日志")
        log_frame.grid(row=2, column=0, sticky="nsew", padx=8, pady=(0, 8))
        log_frame.rowconfigure(0, weight=1)
        log_frame.columnconfigure(0, weight=1)

        self.txt_log = tk.Text(
            log_frame,
            height=16,
            wrap="none",
            font=("Consolas", 9),
        )
        self.txt_log.grid(row=0, column=0, sticky="nsew")
        y_scroll = ttk.Scrollbar(log_frame, command=self.txt_log.yview)
        y_scroll.grid(row=0, column=1, sticky="ns")
        self.txt_log.config(yscrollcommand=y_scroll.set)

        # 初始模式联动
        self._on_run_mode_change()
        self._on_toggle_sandbox()

    def _build_run_tab(self, tab_run: ttk.Frame):
        tab_run.columnconfigure(1, weight=1)
        row = 0

        # === 分组1：路径设置 ===
        group_path = ttk.LabelFrame(tab_run, text="1. 路径设置", padding=(10, 5))
        group_path.grid(row=row, column=0, columnspan=4, sticky="ew", padx=4, pady=4)
        group_path.columnconfigure(1, weight=1)

        # 配置文件
        r_path = 0
        ttk.Label(group_path, text="配置文件:").grid(row=r_path, column=0, sticky="w")
        self.var_config_path = tk.StringVar(value=self.config_path)
        ent_cfg = ttk.Entry(group_path, textvariable=self.var_config_path)
        ent_cfg.grid(row=r_path, column=1, sticky="ew", padx=4)
        ttk.Button(
            group_path, text="打开所在目录", command=self.open_config_folder
        ).grid(row=r_path, column=2, padx=4)
        r_path += 1

        # 源目录
        ttk.Label(group_path, text="源目录:").grid(row=r_path, column=0, sticky="w")
        self.var_source_folder = tk.StringVar()
        ent_src = ttk.Entry(group_path, textvariable=self.var_source_folder)
        ent_src.grid(row=r_path, column=1, sticky="ew", padx=4)
        ttk.Button(group_path, text="浏览...", command=self.browse_source).grid(
            row=r_path, column=2, padx=2
        )
        ttk.Button(group_path, text="打开", command=self.open_source_folder).grid(
            row=r_path, column=3, padx=2
        )
        r_path += 1

        # 目标目录
        ttk.Label(group_path, text="目标目录:").grid(row=r_path, column=0, sticky="w")
        self.var_target_folder = tk.StringVar()
        ent_tgt = ttk.Entry(group_path, textvariable=self.var_target_folder)
        ent_tgt.grid(row=r_path, column=1, sticky="ew", padx=4)
        ttk.Button(group_path, text="浏览...", command=self.browse_target).grid(
            row=r_path, column=2, padx=2
        )
        ttk.Button(group_path, text="打开", command=self.open_target_folder).grid(
            row=r_path, column=3, padx=2
        )

        row += 1

        # === 分组2：任务模式 ===
        group_mode = ttk.LabelFrame(tab_run, text="2. 任务模式与策略", padding=(10, 5))
        group_mode.grid(row=row, column=0, columnspan=4, sticky="ew", padx=4, pady=4)
        group_mode.columnconfigure(1, weight=1)
        r_mode = 0

        # 运行模式
        ttk.Label(group_mode, text="运行模式:").grid(row=r_mode, column=0, sticky="w")
        self.var_run_mode = tk.StringVar(value=MODE_CONVERT_THEN_MERGE)
        frm_run = ttk.Frame(group_mode)
        frm_run.grid(row=r_mode, column=1, columnspan=3, sticky="w")
        ttk.Radiobutton(
            frm_run,
            text="仅转换",
            value=MODE_CONVERT_ONLY,
            variable=self.var_run_mode,
            command=self._on_run_mode_change,
        ).pack(side="left", padx=4)
        ttk.Radiobutton(
            frm_run,
            text="仅合并",
            value=MODE_MERGE_ONLY,
            variable=self.var_run_mode,
            command=self._on_run_mode_change,
        ).pack(side="left", padx=4)
        ttk.Radiobutton(
            frm_run,
            text="先转换再合并",
            value=MODE_CONVERT_THEN_MERGE,
            variable=self.var_run_mode,
            command=self._on_run_mode_change,
        ).pack(side="left", padx=4)
        ttk.Radiobutton(
            frm_run,
            text="文件梳理去重",
            value=MODE_COLLECT_ONLY,
            variable=self.var_run_mode,
            command=self._on_run_mode_change,
        ).pack(side="left", padx=4)
        r_mode += 1

        ttk.Separator(group_mode, orient="horizontal").grid(
            row=r_mode, column=0, columnspan=4, sticky="ew", pady=5
        )
        r_mode += 1

        # 梳理子模式
        self.lbl_collect_mode = ttk.Label(
            group_mode, text="梳理子模式:", foreground="#555"
        )
        self.lbl_collect_mode.grid(row=r_mode, column=0, sticky="w")
        self.var_collect_mode = tk.StringVar(value=COLLECT_MODE_COPY_AND_INDEX)
        frm_collect = ttk.Frame(group_mode)
        frm_collect.grid(row=r_mode, column=1, columnspan=3, sticky="w")
        self.rb_collect_copy = ttk.Radiobutton(
            frm_collect,
            text="去重 + 拷贝 + Excel",
            value=COLLECT_MODE_COPY_AND_INDEX,
            variable=self.var_collect_mode,
        )
        self.rb_collect_copy.pack(side="left", padx=4)
        self.rb_collect_index = ttk.Radiobutton(
            frm_collect,
            text="仅 Excel 索引",
            value=COLLECT_MODE_INDEX_ONLY,
            variable=self.var_collect_mode,
        )
        self.rb_collect_index.pack(side="left", padx=4)
        r_mode += 1

        # 内容策略
        self.lbl_strategy = ttk.Label(group_mode, text="内容处理策略:")
        self.lbl_strategy.grid(row=r_mode, column=0, sticky="w")
        self.var_strategy = tk.StringVar(value="standard")
        frm_strategy = ttk.Frame(group_mode)
        frm_strategy.grid(row=r_mode, column=1, columnspan=3, sticky="w")
        ttk.Radiobutton(
            frm_strategy,
            text="标准分类",
            value="standard",
            variable=self.var_strategy,
        ).pack(side="left", padx=4)
        ttk.Radiobutton(
            frm_strategy,
            text="智能标记",
            value="smart_tag",
            variable=self.var_strategy,
        ).pack(side="left", padx=4)
        ttk.Radiobutton(
            frm_strategy,
            text="报价猎手",
            value="price_only",
            variable=self.var_strategy,
        ).pack(side="left", padx=4)
        r_mode += 1

        # 合并选项
        self.lbl_merge = ttk.Label(group_mode, text="合并选项:")
        self.lbl_merge.grid(row=r_mode, column=0, sticky="w")
        frm_merge = ttk.Frame(group_mode)
        frm_merge.grid(row=r_mode, column=1, columnspan=3, sticky="w")
        self.var_enable_merge = tk.IntVar(value=1)
        self.chk_enable_merge = ttk.Checkbutton(
            frm_merge,
            text="启用合并",
            variable=self.var_enable_merge,
        )
        self.chk_enable_merge.pack(side="left", padx=4)

        self.lbl_merge_mode = ttk.Label(frm_merge, text="模式:")
        self.lbl_merge_mode.pack(side="left", padx=(10, 2))
        self.var_merge_mode = tk.StringVar(value=MERGE_MODE_CATEGORY)
        self.rb_merge_cat = ttk.Radiobutton(
            frm_merge,
            text="分类拆卷",
            value=MERGE_MODE_CATEGORY,
            variable=self.var_merge_mode,
        )
        self.rb_merge_cat.pack(side="left", padx=4)
        self.rb_merge_all = ttk.Radiobutton(
            frm_merge,
            text="全部合一",
            value=MERGE_MODE_ALL_IN_ONE,
            variable=self.var_merge_mode,
        )
        self.rb_merge_all.pack(side="left", padx=4)

        # 合并来源（仅合并模式下显示/有效）
        self.lbl_merge_source = ttk.Label(frm_merge, text="来源:")
        self.lbl_merge_source.pack(side="left", padx=(10, 2))
        self.var_merge_source = tk.StringVar(value="source")
        self.rb_merge_src_source = ttk.Radiobutton(
            frm_merge,
            text="源目录",
            value="source",
            variable=self.var_merge_source,
        )
        self.rb_merge_src_source.pack(side="left", padx=4)
        self.rb_merge_src_target = ttk.Radiobutton(
            frm_merge,
            text="目标目录",
            value="target",
            variable=self.var_merge_source,
        )
        self.rb_merge_src_target.pack(side="left", padx=4)

        r_mode += 1

        # 日期过滤
        self.lbl_date_filter = ttk.Label(group_mode, text="日期过滤:")
        self.lbl_date_filter.grid(row=r_mode, column=0, sticky="w")
        
        frm_date = ttk.Frame(group_mode)
        frm_date.grid(row=r_mode, column=1, columnspan=3, sticky="w")
        
        self.var_enable_date_filter = tk.IntVar(value=0)
        self.chk_date_filter = ttk.Checkbutton(
            frm_date, 
            text="启用", 
            variable=self.var_enable_date_filter,
            command=self._on_toggle_date_filter
        )
        self.chk_date_filter.pack(side="left", padx=4)
        
        from datetime import datetime
        today_str = datetime.now().strftime("%Y-%m-%d")
        self.var_date_str = tk.StringVar(value=today_str)  # YYYY-MM-DD
        self.ent_date = ttk.Entry(frm_date, textvariable=self.var_date_str, width=12, state="disabled")
        self.ent_date.pack(side="left", padx=4)
        ttk.Label(frm_date, text="(YYYY-MM-DD)").pack(side="left")
        
        self.var_filter_mode = tk.StringVar(value="after")
        self.rb_filter_after = ttk.Radiobutton(
            frm_date,
            text="之后",
            value="after",
            variable=self.var_filter_mode,
            state="disabled"
        )
        self.rb_filter_after.pack(side="left", padx=(10, 4))
        
        self.rb_filter_before = ttk.Radiobutton(
            frm_date,
            text="之前",
            value="before",
            variable=self.var_filter_mode,
            state="disabled"
        )
        self.rb_filter_before.pack(side="left", padx=4)

        row += 1

        # === 分组3：执行配置 (引擎/沙箱) ===
        self.group_exec = ttk.LabelFrame(tab_run, text="3. 执行配置", padding=(10, 5))
        self.group_exec.grid(
            row=row, column=0, columnspan=4, sticky="ew", padx=4, pady=4
        )
        self.group_exec.columnconfigure(1, weight=1)
        r_exec = 0

        # 引擎 & 进程策略
        ttk.Label(self.group_exec, text="转换引擎:").grid(
            row=r_exec, column=0, sticky="w"
        )
        frm_engine = ttk.Frame(self.group_exec)
        frm_engine.grid(row=r_exec, column=1, columnspan=3, sticky="w")
        self.var_engine = tk.StringVar(value=ENGINE_WPS)
        ttk.Radiobutton(
            frm_engine,
            text="WPS Office",
            value=ENGINE_WPS,
            variable=self.var_engine,
        ).pack(side="left", padx=4)
        ttk.Radiobutton(
            frm_engine,
            text="Microsoft Office",
            value=ENGINE_MS,
            variable=self.var_engine,
        ).pack(side="left", padx=4)

        ttk.Label(frm_engine, text="进程策略:").pack(side="left", padx=(10, 2))
        self.var_kill_mode = tk.StringVar(value=KILL_MODE_AUTO)
        ttk.Radiobutton(
            frm_engine,
            text="自动清理",
            value=KILL_MODE_AUTO,
            variable=self.var_kill_mode,
        ).pack(side="left", padx=4)
        ttk.Radiobutton(
            frm_engine,
            text="复用进程",
            value=KILL_MODE_KEEP,
            variable=self.var_kill_mode,
        ).pack(side="left", padx=4)
        r_exec += 1

        # 沙箱 & 临时目录
        self.lbl_sandbox = ttk.Label(self.group_exec, text="沙箱 / 临时目录:")
        self.lbl_sandbox.grid(row=r_exec, column=0, sticky="nw")
        frm_sandbox = ttk.Frame(self.group_exec)
        frm_sandbox.grid(row=r_exec, column=1, columnspan=3, sticky="ew")
        frm_sandbox.columnconfigure(1, weight=1)

        self.var_enable_sandbox = tk.IntVar(value=1)
        self.chk_enable_sandbox = ttk.Checkbutton(
            frm_sandbox,
            text="启用沙箱 (推荐，避免污染源目录)",
            variable=self.var_enable_sandbox,
            command=self._on_toggle_sandbox,
        )
        self.chk_enable_sandbox.grid(row=0, column=0, columnspan=3, sticky="w", pady=2)

        ttk.Label(frm_sandbox, text="临时转换目录:").grid(
            row=1, column=0, sticky="w", pady=2
        )
        self.var_temp_sandbox_root = tk.StringVar()
        self.entry_temp_sandbox_root = ttk.Entry(
            frm_sandbox,
            textvariable=self.var_temp_sandbox_root,
            width=50,
        )
        self.entry_temp_sandbox_root.grid(row=1, column=1, sticky="ew", padx=4)
        self.btn_temp_sandbox_root = ttk.Button(
            frm_sandbox, text="浏览...", command=self.browse_temp_sandbox_root
        )
        self.btn_temp_sandbox_root.grid(row=1, column=2, padx=4)

        row += 1

        # 操作按钮
        btn_frame = ttk.Frame(tab_run)
        btn_frame.grid(row=row, column=0, columnspan=4, sticky="ew", pady=(6, 0))
        btn_frame.columnconfigure(0, weight=1)
        btn_frame.columnconfigure(1, weight=1)
        btn_frame.columnconfigure(2, weight=1)

        self.btn_run = ttk.Button(btn_frame, text="开始运行", command=self.start_task)
        self.btn_run.grid(row=0, column=0, padx=4, sticky="ew")

        self.btn_stop = ttk.Button(
            btn_frame,
            text="停止",
            command=self.stop_task,
            state="disabled",
        )
        self.btn_stop.grid(row=0, column=1, padx=4, sticky="ew")

        self.btn_save_cfg = ttk.Button(
            btn_frame,
            text="保存配置",
            command=self.save_config_from_ui,
        )
        self.btn_save_cfg.grid(row=0, column=2, padx=4, sticky="ew")

    def _build_config_tab(self, tab_cfg: ttk.Frame):
        """配置管理界面：直接编辑 config.json 的部分字段"""
        tab_cfg.columnconfigure(1, weight=1)
        row = 0

        # 日志目录
        ttk.Label(tab_cfg, text="日志目录 log_folder:").grid(
            row=row, column=0, sticky="nw", pady=4
        )
        self.var_log_folder = tk.StringVar()
        ent_log = ttk.Entry(tab_cfg, textvariable=self.var_log_folder)
        ent_log.grid(row=row, column=1, sticky="ew", padx=4)
        ttk.Button(tab_cfg, text="浏览...", command=self.browse_log_folder).grid(
            row=row, column=2, padx=4
        )
        row += 1

        # 排除文件夹
        ttk.Label(tab_cfg, text="排除文件夹 excluded_folders:\n（一行一个）").grid(
            row=row, column=0, sticky="nw", pady=4
        )
        self.txt_excluded_folders = tk.Text(tab_cfg, height=5, width=40)
        self.txt_excluded_folders.grid(row=row, column=1, sticky="ew", padx=4)
        row += 1

        # 报价关键字
        ttk.Label(tab_cfg, text="报价关键字 price_keywords:\n（一行一个）").grid(
            row=row, column=0, sticky="nw", pady=4
        )
        self.txt_price_keywords = tk.Text(tab_cfg, height=5, width=40)
        self.txt_price_keywords.grid(row=row, column=1, sticky="ew", padx=4)
        row += 1

        # 超时 & 合并大小等
        frame_adv = ttk.LabelFrame(tab_cfg, text="高级参数（数值）")
        frame_adv.grid(row=row, column=0, columnspan=3, sticky="ew", padx=2, pady=8)
        frame_adv.columnconfigure(1, weight=1)

        r2 = 0
        ttk.Label(frame_adv, text="转换超时（秒）timeout_seconds:").grid(
            row=r2, column=0, sticky="w", pady=2
        )
        self.var_timeout_seconds = tk.StringVar()
        ttk.Entry(frame_adv, textvariable=self.var_timeout_seconds, width=10).grid(
            row=r2, column=1, sticky="w", padx=4
        )
        r2 += 1

        ttk.Label(frame_adv, text="PDF 生成等待（秒）pdf_wait_seconds:").grid(
            row=r2, column=0, sticky="w", pady=2
        )
        self.var_pdf_wait_seconds = tk.StringVar()
        ttk.Entry(frame_adv, textvariable=self.var_pdf_wait_seconds, width=10).grid(
            row=r2, column=1, sticky="w", padx=4
        )
        r2 += 1

        ttk.Label(frame_adv, text="PPT 超时（秒）ppt_timeout_seconds:").grid(
            row=r2, column=0, sticky="w", pady=2
        )
        self.var_ppt_timeout_seconds = tk.StringVar()
        ttk.Entry(frame_adv, textvariable=self.var_ppt_timeout_seconds, width=10).grid(
            row=r2, column=1, sticky="w", padx=4
        )
        r2 += 1

        ttk.Label(frame_adv, text="PPT PDF 等待（秒）ppt_pdf_wait_seconds:").grid(
            row=r2, column=0, sticky="w", pady=2
        )
        self.var_ppt_pdf_wait_seconds = tk.StringVar()
        ttk.Entry(frame_adv, textvariable=self.var_ppt_pdf_wait_seconds, width=10).grid(
            row=r2, column=1, sticky="w", padx=4
        )
        r2 += 1

        ttk.Label(frame_adv, text="合并大小上限（MB）max_merge_size_mb:").grid(
            row=r2, column=0, sticky="w", pady=2
        )
        self.var_max_merge_size_mb = tk.StringVar()
        ttk.Entry(frame_adv, textvariable=self.var_max_merge_size_mb, width=10).grid(
            row=r2, column=1, sticky="w", padx=4
        )

    # ===================== UI 联动 =====================

    def _on_run_mode_change(self):
        mode = self.var_run_mode.get()
        is_collect = mode == MODE_COLLECT_ONLY
        is_convert = mode in (MODE_CONVERT_ONLY, MODE_CONVERT_THEN_MERGE)
        is_merge_related = mode in (MODE_CONVERT_THEN_MERGE, MODE_MERGE_ONLY)

        # 引擎与沙箱组（只有涉及转换时才需要）
        if is_collect or mode == MODE_MERGE_ONLY:
            # 禁用整个执行配置组中的控件
            for child in self.group_exec.winfo_children():
                try:
                    child.configure(state="disabled")
                except Exception:
                    # Frame 可能没有 state 属性，忽略
                    pass
                # 递归禁用 Frame 内部的子控件
                for grandchild in child.winfo_children():
                    try:
                        grandchild.configure(state="disabled")
                    except Exception:
                        pass
        else:
            # 启用
            for child in self.group_exec.winfo_children():
                try:
                    child.configure(state="normal")
                except Exception:
                    pass
                for grandchild in child.winfo_children():
                    try:
                        grandchild.configure(state="normal")
                    except Exception:
                        pass
            # 重新触发一次 sandbox toggle 逻辑以确保输入框状态正确
            self._on_toggle_sandbox()

        # 梳理子模式
        state_collect = "normal" if is_collect else "disabled"
        self.lbl_collect_mode.configure(state=state_collect)
        self.rb_collect_copy.configure(state=state_collect)
        self.rb_collect_index.configure(state=state_collect)

        # 内容策略仅转换模式用
        state_strategy = "normal" if is_convert else "disabled"
        self.lbl_strategy.configure(state=state_strategy)

        # 合并选项
        state_merge = "normal" if is_merge_related else "disabled"
        self.lbl_merge.configure(state=state_merge)
        self.chk_enable_merge.configure(state=state_merge)
        self.lbl_merge_mode.configure(state=state_merge)
        self.rb_merge_cat.configure(state=state_merge)
        self.rb_merge_all.configure(state=state_merge)

        # 合并来源：只有在 MODE_MERGE_ONLY 时才允许选择；如果是“先转换再合并”，强制为 target
        if mode == MODE_MERGE_ONLY and is_merge_related:
             state_merge_source = "normal"
        else:
             state_merge_source = "disabled"
             # 如果不是仅合并模式，且是转换再合并，则强制设为 target 以免误解
             if mode == MODE_CONVERT_THEN_MERGE:
                 self.var_merge_source.set("target")

        self.lbl_merge_source.configure(state=state_merge_source)
        self.rb_merge_src_source.configure(state=state_merge_source)
        self.rb_merge_src_target.configure(state=state_merge_source)

    def _on_toggle_date_filter(self):
        enabled = bool(self.var_enable_date_filter.get())
        state = "normal" if enabled else "disabled"
        self.ent_date.configure(state=state)
        self.rb_filter_after.configure(state=state)
        self.rb_filter_before.configure(state=state)

    def _on_toggle_sandbox(self):
        # 如果当前模式本身就禁用了引擎组，这里就不应该启用
        mode = self.var_run_mode.get()
        is_disabled_globally = mode == MODE_COLLECT_ONLY or mode == MODE_MERGE_ONLY

        enabled = bool(self.var_enable_sandbox.get()) and not is_disabled_globally
        state = "normal" if enabled else "disabled"
        self.entry_temp_sandbox_root.configure(state=state)
        self.btn_temp_sandbox_root.configure(state=state)

    # ===================== 目录/按钮动作 =====================

    def browse_source(self):
        path = filedialog.askdirectory(title="选择源目录")
        if path:
            self.var_source_folder.set(path)

    def browse_target(self):
        path = filedialog.askdirectory(title="选择目标目录")
        if path:
            self.var_target_folder.set(path)

    def open_source_folder(self):
        path = self.var_source_folder.get().strip()
        if path and os.path.isdir(path):
            os.startfile(path)

    def open_target_folder(self):
        path = self.var_target_folder.get().strip()
        if path and os.path.isdir(path):
            os.startfile(path)

    def open_config_folder(self):
        folder = os.path.dirname(self.config_path)
        if folder and os.path.isdir(folder):
            os.startfile(folder)

    def browse_temp_sandbox_root(self):
        path = filedialog.askdirectory(title="选择临时转换根目录")
        if path:
            self.var_temp_sandbox_root.set(path)

    def browse_log_folder(self):
        path = filedialog.askdirectory(title="选择日志目录")
        if path:
            self.var_log_folder.set(path)

    # ===================== 配置读写 =====================

    def _load_config_to_ui(self):
        """启动时从 config.json 读取一次作为默认值"""
        if not os.path.exists(self.config_path):
            return
        try:
            with open(self.config_path, "r", encoding="utf-8") as f:
                cfg = json.load(f)
        except Exception:
            return

        # 运行参数
        self.var_source_folder.set(cfg.get("source_folder", ""))
        self.var_target_folder.set(cfg.get("target_folder", ""))
        self.var_enable_sandbox.set(1 if cfg.get("enable_sandbox", True) else 0)
        self.var_temp_sandbox_root.set(cfg.get("temp_sandbox_root", ""))

        self.var_enable_merge.set(1 if cfg.get("enable_merge", True) else 0)
        self.var_merge_mode.set(cfg.get("merge_mode", MERGE_MODE_CATEGORY))
        self.var_merge_source.set(cfg.get("merge_source", "source"))

        # 运行模式 / 子模式 / 策略（作为默认）
        self.var_run_mode.set(cfg.get("run_mode", MODE_CONVERT_THEN_MERGE))
        self.var_collect_mode.set(cfg.get("collect_mode", COLLECT_MODE_COPY_AND_INDEX))
        self.var_strategy.set(cfg.get("content_strategy", "standard"))

        # 引擎 & 进程策略
        default_engine = cfg.get("default_engine", ENGINE_WPS)
        if default_engine not in (ENGINE_WPS, ENGINE_MS):
            default_engine = ENGINE_WPS
        self.var_engine.set(default_engine)

        kill_mode = cfg.get("kill_process_mode", KILL_MODE_AUTO)
        if kill_mode not in (KILL_MODE_AUTO, KILL_MODE_KEEP):
            kill_mode = KILL_MODE_AUTO
        self.var_kill_mode.set(kill_mode)

        # 配置管理页
        self.var_log_folder.set(cfg.get("log_folder", "./logs"))

        excluded = cfg.get("excluded_folders", [])
        self.txt_excluded_folders.delete("1.0", "end")
        if isinstance(excluded, list):
            self.txt_excluded_folders.insert("end", "\n".join(excluded))

        price_kws = cfg.get("price_keywords", [])
        self.txt_price_keywords.delete("1.0", "end")
        if isinstance(price_kws, list):
            self.txt_price_keywords.insert("end", "\n".join(price_kws))

        self.var_timeout_seconds.set(str(cfg.get("timeout_seconds", 60)))
        self.var_pdf_wait_seconds.set(str(cfg.get("pdf_wait_seconds", 15)))
        self.var_ppt_timeout_seconds.set(str(cfg.get("ppt_timeout_seconds", 180)))
        self.var_ppt_pdf_wait_seconds.set(str(cfg.get("ppt_pdf_wait_seconds", 30)))
        self.var_max_merge_size_mb.set(str(cfg.get("max_merge_size_mb", 80)))

        # 联动刷新
        self._on_run_mode_change()
        self._on_toggle_sandbox()

    def save_config_from_ui(self):
        """保存当前 UI 参数到 config.json（作为默认值）"""
        cfg = {}
        if os.path.exists(self.config_path):
            try:
                with open(self.config_path, "r", encoding="utf-8") as f:
                    cfg = json.load(f)
            except Exception:
                cfg = {}

        # 运行参数
        cfg["source_folder"] = self.var_source_folder.get().strip()
        cfg["target_folder"] = self.var_target_folder.get().strip()
        cfg["enable_sandbox"] = bool(self.var_enable_sandbox.get())
        cfg["temp_sandbox_root"] = self.var_temp_sandbox_root.get().strip()

        cfg["enable_merge"] = bool(self.var_enable_merge.get())
        cfg["merge_mode"] = self.var_merge_mode.get()
        cfg["merge_source"] = self.var_merge_source.get()

        cfg["run_mode"] = self.var_run_mode.get()
        cfg["collect_mode"] = self.var_collect_mode.get()
        cfg["content_strategy"] = self.var_strategy.get()

        cfg["default_engine"] = self.var_engine.get()
        cfg["kill_process_mode"] = self.var_kill_mode.get()

        # 配置管理页
        cfg["log_folder"] = self.var_log_folder.get().strip() or "./logs"

        excluded_text = self.txt_excluded_folders.get("1.0", "end").strip()
        excluded_list = [
            line.strip() for line in excluded_text.splitlines() if line.strip()
        ]
        cfg["excluded_folders"] = excluded_list

        kw_text = self.txt_price_keywords.get("1.0", "end").strip()
        kw_list = [line.strip() for line in kw_text.splitlines() if line.strip()]
        cfg["price_keywords"] = kw_list

        def _to_int(var, default):
            try:
                v = int(var.get().strip())
                return v if v > 0 else default
            except Exception:
                return default

        cfg["timeout_seconds"] = _to_int(self.var_timeout_seconds, 60)
        cfg["pdf_wait_seconds"] = _to_int(self.var_pdf_wait_seconds, 15)
        cfg["ppt_timeout_seconds"] = _to_int(self.var_ppt_timeout_seconds, 180)
        cfg["ppt_pdf_wait_seconds"] = _to_int(self.var_ppt_pdf_wait_seconds, 30)
        cfg["max_merge_size_mb"] = _to_int(self.var_max_merge_size_mb, 80)

        try:
            with open(self.config_path, "w", encoding="utf-8") as f:
                json.dump(cfg, f, indent=4, ensure_ascii=False)
            messagebox.showinfo("保存配置", "配置已保存到 config.json。")
        except Exception as e:
            messagebox.showerror("保存失败", f"写入配置文件失败：\n{e}")

    # ===================== 日志 & 状态 =====================

    def _poll_log_queue(self):
        try:
            while True:
                msg = LOG_QUEUE.get_nowait()
                self.txt_log.insert("end", msg + "\n")
                self.txt_log.see("end")
        except queue.Empty:
            pass
        self.after(200, self._poll_log_queue)

    def _set_running_ui_state(self, running: bool):
        if running:
            self.btn_run.configure(state="disabled")
            self.btn_stop.configure(state="normal")
            self.btn_save_cfg.configure(state="disabled")
            self.progress["mode"] = "determinate"
            self.progress["value"] = 0
            self.var_status.set("正在初始化...")
        else:
            self.btn_run.configure(state="normal")
            self.btn_stop.configure(state="disabled")
            self.btn_save_cfg.configure(state="normal")
            self.progress.stop()
            self.progress["value"] = 100
            self.var_status.set("就绪")

    def on_progress_update(self, current, total):
        """Core 转换器在后台线程调用的回调"""

        def _update():
            if total > 0:
                pct = (current / total) * 100
                self.progress["value"] = pct
                self.var_status.set(f"正在处理: {current}/{total} ({pct:.1f}%)")
            else:
                self.progress["mode"] = "indeterminate"
                self.progress.start(20)
                self.var_status.set(f"正在处理: {current}/...")

        # 线程安全调用
        self.after(0, _update)

    # ===================== 任务控制 =====================

    def start_task(self):
        if self.worker_thread and self.worker_thread.is_alive():
            messagebox.showinfo("提示", "已有任务在运行，请先停止或等待完成。")
            return

        # 自动去除路径可能包含的引号
        source = self.var_source_folder.get().strip().strip('"').strip("'")
        target = self.var_target_folder.get().strip().strip('"').strip("'")

        # 回写去引号后的值到 UI
        self.var_source_folder.set(source)
        self.var_target_folder.set(target)

        if not source or not os.path.isdir(source):
            messagebox.showerror("错误", "请先设置有效的【源目录】。")
            return
        if not target:
            messagebox.showerror("错误", "请先设置有效的【目标目录】。")
            return

        self.stop_requested = False
        self.txt_log.insert("end", "\n========== GUI 后台任务启动 ==========\n")
        self.txt_log.see("end")

        def worker():
            try:
                print(f"[GUI] 使用配置文件: {self.config_path}")
                converter = GUIOfficeConverter(self.config_path)
                # 注入进度回调
                converter.progress_callback = self.on_progress_update
                self.current_converter = converter

                # 用当前界面参数覆盖 runtime（不写回 config.json）
                cfg = converter.config
                cfg["source_folder"] = source
                cfg["target_folder"] = target
                cfg["enable_sandbox"] = bool(self.var_enable_sandbox.get())
                cfg["temp_sandbox_root"] = self.var_temp_sandbox_root.get().strip()
                cfg["enable_merge"] = bool(self.var_enable_merge.get())
                cfg["merge_mode"] = self.var_merge_mode.get()
                cfg["merge_source"] = self.var_merge_source.get()
                cfg["kill_process_mode"] = self.var_kill_mode.get()
                cfg["default_engine"] = self.var_engine.get()

                converter.run_mode = self.var_run_mode.get()
                converter.collect_mode = self.var_collect_mode.get()
                converter.content_strategy = self.var_strategy.get()
                converter.merge_mode = self.var_merge_mode.get()
                converter.engine_type = self.var_engine.get()

                # 设置日期过滤
                if self.var_enable_date_filter.get():
                    date_str = self.var_date_str.get().strip()
                    try:
                        converter.filter_date = datetime.strptime(date_str, "%Y-%m-%d")
                        converter.filter_mode = self.var_filter_mode.get()
                        print(f"[GUI] 已启用日期过滤: {converter.filter_mode} {date_str}")
                    except ValueError:
                        print(f"[GUI] [警告] 日期格式无效: {date_str}，将忽略日期过滤。")

                # 重新根据覆盖后的 target / temp_sandbox_root 计算路径（简单重算一遍）
                # 临时目录
                temp_root = cfg.get("temp_sandbox_root", "").strip()
                if temp_root:
                    if not os.path.isabs(temp_root):
                        temp_root = os.path.abspath(
                            os.path.join(get_app_path(), temp_root)
                        )
                else:
                    temp_root = tempfile.gettempdir()
                converter.temp_sandbox_root = temp_root
                converter.temp_sandbox = os.path.join(temp_root, "OfficeToPDF_Sandbox")
                os.makedirs(converter.temp_sandbox, exist_ok=True)
                print(f"[GUI] 本次临时转换目录: {converter.temp_sandbox}")

                # 失败目录 & 合并目录
                converter.failed_dir = os.path.join(
                    cfg["target_folder"], "_FAILED_FILES"
                )
                os.makedirs(converter.failed_dir, exist_ok=True)
                converter.merge_output_dir = os.path.join(
                    cfg["target_folder"], "_MERGED"
                )
                os.makedirs(converter.merge_output_dir, exist_ok=True)

                # 运行
                converter.run()
                print("[GUI] 任务已完成。")

            except Exception as e:
                print(f"[GUI] 运行出错: {e}")
                print(traceback.format_exc())
                messagebox.showerror("运行错误", f"任务运行出错：\n{e}")
            finally:
                self.current_converter = None
                self.stop_requested = False
                self.after(0, lambda: self._set_running_ui_state(False))

        self.worker_thread = threading.Thread(target=worker, daemon=True)
        self.worker_thread.start()
        self._set_running_ui_state(True)

    def stop_task(self):
        if self.current_converter is None:
            return
        if messagebox.askyesno("停止任务", "确定要请求停止当前任务吗？"):
            self.stop_requested = True
            self.current_converter.is_running = False
            print("[GUI] 已请求停止任务，正在等待当前步骤结束...")
            self.var_status.set("已请求停止，等待当前文件处理结束...")

    # ===================== 程序入口 =====================


if __name__ == "__main__":
    app = OfficeGUI()
    app.mainloop()
