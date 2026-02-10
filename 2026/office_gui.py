# -*- coding: utf-8 -*-
"""
office_gui.py - Office 文档批量转换 & 梳理工具 GUI 版

说明：
- 依赖 office_converter.py 中的 OfficeConverter（你已经更新到 v5.15.6）
- GUI 中：
    * “运行参数”页：选择源/目标目录、运行模式、内容策略、合并模式、沙箱等
    * “配置管理”页：直接编辑 config.json 的部分配置（日志目录、排除目录、关键字、超时参数等）
- “保存配置”按钮：写入 config.json
- “开始运行”按钮：用当前界面参数启动转换/梳理（不会自动改 config.json）
- “停止”按钮：设置 converter.is_running=False，优雅停止
"""

import os
import sys
import subprocess
import glob

import json
import threading
import queue
import re
from types import SimpleNamespace
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, colorchooser
from tkinter.constants import *  # Explicitly import standard constants
import ttkbootstrap as tb
from ttkbootstrap.constants import *
from ttkbootstrap.widgets.scrolled import ScrolledText

import tempfile
import traceback
from datetime import datetime
from ui_translations import TRANSLATIONS
from locate_source import locate_by_page, locate_by_short_id
from search_adapter import EverythingAdapter, build_listary_query

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


class HoverTip:
    """Minimal tooltip for Tk/ttk widgets."""

    def __init__(
        self,
        widget,
        text_func,
        style_func=None,
        delay_ms=500,
        bg="#FFF7D6",
        fg="#202124",
        font_family="System",
        font_size=9,
    ):
        self.widget = widget
        self.text_func = text_func
        self.style_func = style_func
        self.delay_ms = delay_ms
        self.bg = bg
        self.fg = fg
        self.font_family = font_family
        self.font_size = font_size
        self.tipwindow = None
        self._after_id = None
        self._last_x_root = None
        self._last_y_root = None
        widget.bind("<Enter>", self._schedule, add="+")
        widget.bind("<Motion>", self._on_motion, add="+")
        widget.bind("<Leave>", self._hide, add="+")
        widget.bind("<FocusOut>", self._hide, add="+")
        widget.bind("<Destroy>", self._hide, add="+")
        widget.bind("<ButtonPress>", self._hide, add="+")

    def _schedule(self, event=None):
        self._cache_pointer(event)
        self._cancel_schedule()
        self._after_id = self.widget.after(self.delay_ms, self._show)

    def _on_motion(self, event=None):
        self._cache_pointer(event)
        if self.tipwindow is not None:
            x, y = self._resolve_position()
            try:
                self.tipwindow.wm_geometry(f"+{x}+{y}")
            except Exception:
                pass

    def _cache_pointer(self, event=None):
        if event is not None:
            try:
                self._last_x_root = int(event.x_root)
                self._last_y_root = int(event.y_root)
                return
            except Exception:
                pass
        try:
            self._last_x_root = int(self.widget.winfo_pointerx())
            self._last_y_root = int(self.widget.winfo_pointery())
        except Exception:
            self._last_x_root = None
            self._last_y_root = None

    def _resolve_position(self):
        if self._last_x_root is not None and self._last_y_root is not None:
            return self._last_x_root + 12, self._last_y_root + 18
        return self.widget.winfo_rootx() + 14, self.widget.winfo_rooty() + self.widget.winfo_height() + 8

    def _cancel_schedule(self):
        if self._after_id is not None:
            try:
                self.widget.after_cancel(self._after_id)
            except Exception:
                pass
            self._after_id = None

    def _show(self):
        text = self.text_func() if callable(self.text_func) else str(self.text_func)
        if not text:
            return
        if self.tipwindow is not None:
            return

        bg = self.bg
        fg = self.fg
        font_family = self.font_family
        font_size = self.font_size
        if callable(self.style_func):
            try:
                style = self.style_func() or {}
                bg = style.get("bg", bg)
                fg = style.get("fg", fg)
                font_family = style.get("font_family", font_family)
                font_size = style.get("font_size", font_size)
            except Exception:
                pass

        x, y = self._resolve_position()
        tw = tk.Toplevel(self.widget.winfo_toplevel())
        try:
            tw.wm_overrideredirect(True)
        except Exception:
            pass
        try:
            tw.wm_attributes("-topmost", True)
        except Exception:
            pass
        tw.wm_geometry(f"+{x}+{y}")
        label = tk.Label(
            tw,
            text=text,
            justify=LEFT,
            background=bg,
            foreground=fg,
            relief="solid",
            borderwidth=1,
            padx=8,
            pady=4,
            font=(font_family, font_size),
        )
        label.pack()
        try:
            tw.lift()
        except Exception:
            pass
        self.tipwindow = tw

    def _hide(self, _event=None):
        self._cancel_schedule()
        if self.tipwindow is not None:
            try:
                self.tipwindow.destroy()
            except Exception:
                pass
            self.tipwindow = None


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


class OfficeGUI(tb.Window):
    TOOLTIP_DEFAULTS = {
        "tooltip_delay_ms": 300,
        "tooltip_bg": "#FFF7D6",
        "tooltip_fg": "#202124",
        "tooltip_font_size": 9,
        "tooltip_auto_theme": True,
    }

    def __init__(self, config_path=None):
        super().__init__(themename="cosmo")
        self.current_lang = "zh"  # Default language
        self.title(f"{self.tr('title')} v{__version__}")
        self.geometry("1100x760")
        self.minsize(900, 620)
        
        # Ensure style is available for theme toggling
        # tb.Window automatically creates a style object, accessible via self.style if needed, 
        # but self.style is not a standard attribute of tk.Tk. 
        # ttkbootstrap.Window has a 'style' attribute.

        self.script_dir = get_app_path()
        self.config_path = config_path or os.path.join(self.script_dir, "config.json")

        # 当前任务线程 & 转换器实例
        self.worker_thread = None
        self.current_converter = None
        self.stop_requested = False
        self._tooltips = []
        self._tooltip_widget_ids = set()
        self._normalizing_inputs = False
        self.tooltip_delay_ms = self.TOOLTIP_DEFAULTS["tooltip_delay_ms"]
        self.tooltip_bg = self.TOOLTIP_DEFAULTS["tooltip_bg"]
        self.tooltip_fg = self.TOOLTIP_DEFAULTS["tooltip_fg"]
        self.tooltip_font_family = "System"
        self.tooltip_font_size = self.TOOLTIP_DEFAULTS["tooltip_font_size"]
        self.tooltip_auto_theme = self.TOOLTIP_DEFAULTS["tooltip_auto_theme"]

        # 把 stdout/stderr 重定向到 GUI 日志窗口
        # sys.stdout = TkLogHandler()
        # sys.stderr = TkLogHandler()

        if not os.path.exists(self.config_path):
            success = create_default_config(self.config_path)
            if success:
                messagebox.showinfo(
                    "提示",
                    "未找到 config.json，已为您自动生成默认配置。\n",
                )

        self._build_ui()
        self._load_config_to_ui()
        self.locator_short_id_index = {}

        # 定时刷新日志
        self.after(200, self._poll_log_queue)


    # ===================== UI 构建 =====================

    # ===================== UI 构建 (Modern Layout) =====================

    def tr(self, key):
        """Translate key to current language"""
        return TRANSLATIONS.get(self.current_lang, TRANSLATIONS["zh"]).get(key, key)

    def _attach_tooltip(self, widget, key):
        if id(widget) in self._tooltip_widget_ids:
            return
        self._tooltip_widget_ids.add(id(widget))
        self._tooltips.append(
            HoverTip(
                widget,
                lambda k=key: self.tr(k),
                style_func=self._get_tooltip_style,
                delay_ms=self.tooltip_delay_ms,
                bg=self.tooltip_bg,
                fg=self.tooltip_fg,
                font_family=self.tooltip_font_family,
                font_size=self.tooltip_font_size,
            )
        )

    def _attach_tooltip_text(self, widget, text):
        if id(widget) in self._tooltip_widget_ids:
            return
        self._tooltip_widget_ids.add(id(widget))
        self._tooltips.append(
            HoverTip(
                widget,
                lambda t=text: t,
                style_func=self._get_tooltip_style,
                delay_ms=self.tooltip_delay_ms,
                bg=self.tooltip_bg,
                fg=self.tooltip_fg,
                font_family=self.tooltip_font_family,
                font_size=self.tooltip_font_size,
            )
        )

    def _auto_attach_action_tooltips(self, root):
        specific_key_by_text = {
            self.tr("mode_convert"): "tip_mode_convert",
            self.tr("mode_merge"): "tip_mode_merge",
            self.tr("mode_convert_merge"): "tip_mode_convert_merge",
            self.tr("mode_collect"): "tip_mode_collect",
            self.tr("lbl_sandbox"): "tip_toggle_sandbox",
            self.tr("lbl_filter_date"): "tip_toggle_date_filter",
            self.tr("chk_merge_index"): "tip_toggle_merge_index",
            self.tr("chk_merge_excel"): "tip_toggle_merge_excel",
            self.tr("chk_tooltip_auto_theme"): "tip_toggle_tooltip_auto_theme",
        }
        for child in root.winfo_children():
            self._auto_attach_action_tooltips(child)
            try:
                keys = set(child.keys())
            except Exception:
                continue
            if "text" not in keys:
                continue
            text = str(child.cget("text")).strip()
            if not text:
                continue
            is_option = "variable" in keys and ("value" in keys or "onvalue" in keys)
            is_button = "command" in keys and not is_option
            if not (is_option or is_button):
                continue
            specific_key = specific_key_by_text.get(text)
            if specific_key:
                self._attach_tooltip(child, specific_key)
                continue
            if is_button:
                tip = self.tr("tip_auto_button_action").format(text)
            else:
                tip = self.tr("tip_auto_option_action").format(text)
            self._attach_tooltip_text(child, tip)

    def _get_tooltip_style(self):
        if self.tooltip_auto_theme:
            theme_name = self.var_theme.get() if hasattr(self, "var_theme") else "cosmo"
            if theme_name == "superhero":
                return {
                    "bg": "#2D3748",
                    "fg": "#F8FAFC",
                    "font_family": self.tooltip_font_family,
                    "font_size": self.tooltip_font_size,
                }
            return {
                "bg": "#FFF7D6",
                "fg": "#202124",
                "font_family": self.tooltip_font_family,
                "font_size": self.tooltip_font_size,
            }

        return {
            "bg": self.tooltip_bg,
            "fg": self.tooltip_fg,
            "font_family": self.tooltip_font_family,
            "font_size": self.tooltip_font_size,
        }

    def _toggle_language(self):
        """Switch language and rebuild UI"""
        # Save current state to config first
        try:
            self._save_settings_to_file(show_msg=False)
        except Exception:
            pass
        
        self.current_lang = "en" if self.current_lang == "zh" else "zh"
        self._tooltips = []
        self._tooltip_widget_ids = set()
        
        # Destroy all children to rebuild
        for widget in self.winfo_children():
            widget.destroy()
            
        # Rebuild
        self.title(f"{self.tr('title')} v{__version__}")
        self._build_ui()
        self._load_config_to_ui()

    # ===================== UI 构建 =====================

    def _build_ui(self):
        self._tooltips = []
        self._tooltip_widget_ids = set()
        # 1. Header
        header_frame = tb.Frame(self, bootstyle="light")
        header_frame.pack(fill=X, side=TOP)
        tb.Label(
            header_frame,
            text=self.tr("header_title"),
            font=("Helvetica", 16, "bold"),
            bootstyle="inverse-light",
        ).pack(side=LEFT, padx=20, pady=10)

        ctrl_frame = tb.Frame(header_frame, bootstyle="light")
        ctrl_frame.pack(side=RIGHT, padx=20)

        self.btn_lang_toggle = tb.Button(
            ctrl_frame,
            text=self.tr("lang_toggle"),
            command=self._toggle_language,
            bootstyle="outline-secondary",
            width=6,
        )
        self.btn_lang_toggle.pack(side=LEFT, padx=10)
        self._attach_tooltip(self.btn_lang_toggle, "tip_lang_toggle")

        self.var_theme = tk.StringVar(value="cosmo")

        def toggle_theme():
            t = self.var_theme.get()
            new_theme = "superhero" if t == "cosmo" else "cosmo"
            self.style.theme_use(new_theme)
            self.var_theme.set(new_theme)

        self.chk_theme_toggle = tb.Checkbutton(
            ctrl_frame,
            text=self.tr("theme_dark"),
            bootstyle="round-toggle",
            variable=self.var_theme,
            onvalue="superhero",
            offvalue="cosmo",
            command=toggle_theme,
        )
        self.chk_theme_toggle.pack(side=LEFT)
        self._attach_tooltip(self.chk_theme_toggle, "tip_theme_toggle")

        # 2. Footer (progress + actions)
        footer_wrap = tb.Frame(self)
        footer_wrap.pack(side=BOTTOM, fill=X)
        footer_frame = tb.Frame(footer_wrap, padding=15, bootstyle="light")
        footer_frame.pack(side=BOTTOM, fill=X)
        self.progress = tb.Progressbar(
            footer_wrap,
            bootstyle="success-striped",
            mode="determinate",
            length=100,
        )
        self.progress.pack(side=BOTTOM, fill=X)
        self._build_footer(footer_frame)

        # 3. Main body (tabs + optional log pane)
        body_frame = tb.Frame(self, padding=10)
        body_frame.pack(fill=BOTH, expand=YES)
        self.paned = tb.Panedwindow(body_frame, orient=VERTICAL)
        self.paned.pack(fill=BOTH, expand=YES)

        self.config_container = tb.Frame(self.paned)
        self.paned.add(self.config_container, weight=3)
        self.main_notebook = tb.Notebook(self.config_container)
        self.main_notebook.pack(fill=BOTH, expand=YES)

        self.tab_run = tb.Frame(self.main_notebook)
        self.tab_config = tb.Frame(self.main_notebook)
        self.main_notebook.add(self.tab_run, text=self.tr("tab_run_center"))
        self.main_notebook.add(self.tab_config, text=self.tr("tab_config_center"))

        run_content = self._create_scrollable_page(self.tab_run)
        cfg_content = self._create_scrollable_page(self.tab_config)
        self._build_run_tab_content(run_content)
        self._build_config_tab_content(cfg_content)
        self.main_notebook.select(0)

        self.log_pane = tb.Frame(self.paned)
        self.txt_log = ScrolledText(
            self.log_pane, height=10, font=("Consolas", 9), bootstyle="primary-round"
        )
        self.txt_log.pack(fill=BOTH, expand=YES)
        self.txt_log.text.tag_config("INFO", foreground="#007bff")
        self.txt_log.text.tag_config("SUCCESS", foreground="#28a745")
        self.txt_log.text.tag_config("WARNING", foreground="#ffc107")
        self.txt_log.text.tag_config("ERROR", foreground="#dc3545")
        self.txt_log.text.tag_config("DIM", foreground="#6c757d")

        self._on_run_mode_change()
        self._on_toggle_sandbox()
        self._set_running_ui_state(False)

    def _create_scrollable_page(self, parent):
        canvas = tk.Canvas(parent, highlightthickness=0)
        scrollbar = tb.Scrollbar(parent, orient="vertical", command=canvas.yview)
        content = tb.Frame(canvas)
        content.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        win_id = canvas.create_window((0, 0), window=content, anchor="nw", width=canvas.winfo_reqwidth())

        def on_canvas_configure(e):
            canvas.itemconfig(win_id, width=e.width)

        canvas.bind("<Configure>", on_canvas_configure)
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side=LEFT, fill=BOTH, expand=YES)
        scrollbar.pack(side=RIGHT, fill=Y)
        return content

    def _create_path_row(self, parent, label_key, var, cmd_browse, cmd_open):
        f = tb.Frame(parent)
        f.pack(fill=X, pady=(5, 0))
        tb.Label(f, text=self.tr(label_key), font=("System", 9, "bold")).pack(anchor="w")

        f_in = tb.Frame(f)
        f_in.pack(fill=X)
        entry = tb.Entry(f_in, textvariable=var, font=("System", 9))
        entry.pack(side=LEFT, fill=X, expand=YES)
        path_tip_key_by_label = {
            "lbl_source": "tip_input_source_folder",
            "lbl_target": "tip_input_target_folder",
            "lbl_config": "tip_input_config_path",
        }
        tip_key = path_tip_key_by_label.get(label_key)
        if tip_key:
            self._attach_tooltip(entry, tip_key)
        btn_browse = tb.Button(f_in, text="...", command=cmd_browse, bootstyle="outline", width=3)
        btn_browse.pack(side=LEFT, padx=(2, 0))
        self._attach_tooltip(btn_browse, "tip_browse_folder")
        if cmd_open:
            btn_open = tb.Button(f_in, text="↗", command=cmd_open, bootstyle="link", width=2)
            btn_open.pack(side=LEFT)
            self._attach_tooltip(btn_open, "tip_open_folder")

    def _add_section_help(self, parent, tip_key):
        row = tb.Frame(parent)
        row.pack(fill=X, pady=(0, 4))
        spacer = tb.Label(row, text="")
        spacer.pack(side=LEFT, fill=X, expand=YES)
        btn = tb.Button(row, text="?", width=2, bootstyle="info-outline")
        btn.pack(side=RIGHT)
        self._attach_tooltip(btn, tip_key)

    def _build_run_tab_content(self, parent):
        # Section 1: run mode
        lf_mode = tb.Labelframe(parent, text=self.tr("sec_mode"), padding=10)
        lf_mode.pack(fill=X, pady=5)
        self._add_section_help(lf_mode, "tip_section_run_mode")
        self.var_run_mode = tk.StringVar(value=MODE_CONVERT_THEN_MERGE)
        grid_frame = tb.Frame(lf_mode)
        grid_frame.pack(fill=X)
        tb.Radiobutton(
            grid_frame, text=self.tr("mode_convert"), variable=self.var_run_mode, value=MODE_CONVERT_ONLY,
            command=self._on_run_mode_change, bootstyle="toolbutton-outline"
        ).grid(row=0, column=0, sticky="ew", padx=2, pady=2)
        tb.Radiobutton(
            grid_frame, text=self.tr("mode_merge"), variable=self.var_run_mode, value=MODE_MERGE_ONLY,
            command=self._on_run_mode_change, bootstyle="toolbutton-outline"
        ).grid(row=0, column=1, sticky="ew", padx=2, pady=2)
        tb.Radiobutton(
            grid_frame, text=self.tr("mode_convert_merge"), variable=self.var_run_mode, value=MODE_CONVERT_THEN_MERGE,
            command=self._on_run_mode_change, bootstyle="toolbutton-outline"
        ).grid(row=1, column=0, sticky="ew", padx=2, pady=2)
        tb.Radiobutton(
            grid_frame, text=self.tr("mode_collect"), variable=self.var_run_mode, value=MODE_COLLECT_ONLY,
            command=self._on_run_mode_change, bootstyle="toolbutton-outline"
        ).grid(row=1, column=1, sticky="ew", padx=2, pady=2)
        grid_frame.columnconfigure(0, weight=1)
        grid_frame.columnconfigure(1, weight=1)

        self.frm_collect_opts = tb.Frame(lf_mode, padding=(10, 5))
        self.frm_collect_opts.pack(fill=X)
        self.var_collect_mode = tk.StringVar(value=COLLECT_MODE_COPY_AND_INDEX)
        tb.Radiobutton(
            self.frm_collect_opts, text="Copy + Index",
            variable=self.var_collect_mode, value=COLLECT_MODE_COPY_AND_INDEX
        ).pack(anchor="w")
        tb.Radiobutton(
            self.frm_collect_opts, text="Index Only",
            variable=self.var_collect_mode, value=COLLECT_MODE_INDEX_ONLY
        ).pack(anchor="w")

        # Section 2: paths (runtime only)
        lf_paths = tb.Labelframe(parent, text=self.tr("sec_paths"), padding=10)
        lf_paths.pack(fill=X, pady=5)
        self._add_section_help(lf_paths, "tip_section_run_paths")
        self.var_source_folder = tk.StringVar()
        self._create_path_row(lf_paths, "lbl_source", self.var_source_folder, self.browse_source, self.open_source_folder)
        self.var_target_folder = tk.StringVar()
        self._create_path_row(lf_paths, "lbl_target", self.var_target_folder, self.browse_target, self.open_target_folder)

        # Section 3: execution options
        lf_settings = tb.Labelframe(parent, text=self.tr("sec_advanced"), padding=10)
        lf_settings.pack(fill=X, pady=5)
        self._add_section_help(lf_settings, "tip_section_run_advanced")
        self.group_exec = tb.Frame(lf_settings)
        self.group_exec.pack(fill=X, pady=5)
        tb.Label(self.group_exec, text=self.tr("lbl_engine"), bootstyle="primary").pack(anchor="w")
        self.var_engine = tk.StringVar(value=ENGINE_WPS)
        frm_eng = tb.Frame(self.group_exec)
        frm_eng.pack(anchor="w")
        tb.Radiobutton(frm_eng, text="WPS Office", variable=self.var_engine, value=ENGINE_WPS).pack(side=LEFT, padx=5)
        tb.Radiobutton(frm_eng, text="MS Office", variable=self.var_engine, value=ENGINE_MS).pack(side=LEFT, padx=5)
        self.var_enable_sandbox = tk.IntVar(value=1)
        self.chk_enable_sandbox = tb.Checkbutton(
            self.group_exec, text=self.tr("lbl_sandbox"), variable=self.var_enable_sandbox,
            bootstyle="success-round-toggle", command=self._on_toggle_sandbox
        )
        self.chk_enable_sandbox.pack(anchor="w", pady=(10, 2))
        self.frm_sandbox_path = tb.Frame(self.group_exec)
        self.frm_sandbox_path.pack(fill=X, padx=20)
        self.var_temp_sandbox_root = tk.StringVar()
        self.entry_temp_sandbox_root = tb.Entry(self.frm_sandbox_path, textvariable=self.var_temp_sandbox_root, font=("Consolas", 8))
        self.entry_temp_sandbox_root.pack(side=LEFT, fill=X, expand=YES)
        self.btn_temp_sandbox_root = tb.Button(self.frm_sandbox_path, text=self.tr("btn_browse"), command=self.browse_temp_sandbox_root, bootstyle="outline", width=3)
        self.btn_temp_sandbox_root.pack(side=LEFT, padx=2)
        self._attach_tooltip(self.btn_temp_sandbox_root, "tip_choose_temp")

        tb.Separator(lf_settings).pack(fill=X, pady=10)
        self.lbl_merge = tb.Label(lf_settings, text=self.tr("lbl_merge_logic"), bootstyle="primary")
        self.lbl_merge.pack(anchor="w")
        self.var_enable_merge = tk.IntVar(value=1)
        self.chk_enable_merge = tb.Checkbutton(lf_settings, text="Enable Merge", variable=self.var_enable_merge, bootstyle="square-toggle")
        self.chk_enable_merge.pack(anchor="w")
        self.frm_merge_opts = tb.Frame(lf_settings, padding=(20, 0))
        self.frm_merge_opts.pack(fill=X)
        self.var_merge_mode = tk.StringVar(value=MERGE_MODE_CATEGORY)
        tb.Radiobutton(self.frm_merge_opts, text=self.tr("rad_category"), variable=self.var_merge_mode, value=MERGE_MODE_CATEGORY).pack(anchor="w")
        tb.Radiobutton(self.frm_merge_opts, text=self.tr("rad_all_in_one"), variable=self.var_merge_mode, value=MERGE_MODE_ALL_IN_ONE).pack(anchor="w")
        tb.Separator(self.frm_merge_opts).pack(fill=X, pady=5)
        self.var_enable_merge_index = tk.IntVar(value=0)
        self.chk_merge_index = tb.Checkbutton(self.frm_merge_opts, text=self.tr("chk_merge_index"), variable=self.var_enable_merge_index)
        self.chk_merge_index.pack(anchor="w")
        self.var_enable_merge_excel = tk.IntVar(value=0)
        self.chk_merge_excel = tb.Checkbutton(self.frm_merge_opts, text=self.tr("chk_merge_excel"), variable=self.var_enable_merge_excel)
        self.chk_merge_excel.pack(anchor="w")
        tb.Separator(self.frm_merge_opts).pack(fill=X, pady=5)
        self.lbl_m_src = tb.Label(self.frm_merge_opts, text=self.tr("lbl_merge_src"), font=("System", 9))
        self.lbl_m_src.pack(anchor="w")
        self.var_merge_source = tk.StringVar(value="source")
        frm_m_src = tb.Frame(self.frm_merge_opts)
        frm_m_src.pack(fill=X)
        tb.Radiobutton(frm_m_src, text=self.tr("rad_src_dir"), variable=self.var_merge_source, value="source").pack(side=LEFT)
        tb.Radiobutton(frm_m_src, text=self.tr("rad_tgt_dir"), variable=self.var_merge_source, value="target").pack(side=LEFT, padx=10)

        # Section 4: strategy + date filter
        lf_extra = tb.Labelframe(parent, text=self.tr("sec_filters"), padding=10)
        lf_extra.pack(fill=X, pady=5)
        self._add_section_help(lf_extra, "tip_section_run_filters")
        self.lbl_strategy = tb.Label(lf_extra, text=self.tr("lbl_strategy"))
        self.lbl_strategy.pack(anchor="w")
        self.var_strategy = tk.StringVar(value="standard")
        cb_strat = tb.Combobox(lf_extra, textvariable=self.var_strategy, values=["standard", "smart_tag", "price_only"], state="readonly")
        cb_strat.pack(fill=X, pady=(0, 5))
        self.var_enable_date_filter = tk.IntVar(value=0)
        self.chk_date_filter = tb.Checkbutton(lf_extra, text=self.tr("lbl_filter_date"), variable=self.var_enable_date_filter, command=self._on_toggle_date_filter)
        self.chk_date_filter.pack(anchor="w", pady=(5, 0))
        self.frm_date = tb.Frame(lf_extra)
        self.frm_date.pack(fill=X, padx=20)
        today_str = datetime.now().strftime("%Y-%m-%d")
        self.var_date_str = tk.StringVar(value=today_str)
        try:
            self.ent_date = tb.DateEntry(self.frm_date, dateformat="%Y-%m-%d", firstweekday=0, startdate=datetime.now())
            self.ent_date.entry.configure(textvariable=self.var_date_str)
            self.ent_date.pack(fill=X)
        except Exception:
            self.ent_date = tb.Entry(self.frm_date, textvariable=self.var_date_str)
            self.ent_date.pack(fill=X)
        self.var_filter_mode = tk.StringVar(value="after")
        frm_dt_mode = tb.Frame(self.frm_date)
        frm_dt_mode.pack(fill=X)
        self.rb_filter_after = tb.Radiobutton(frm_dt_mode, text=self.tr("rad_after"), variable=self.var_filter_mode, value="after")
        self.rb_filter_after.pack(side=LEFT)
        self.rb_filter_before = tb.Radiobutton(frm_dt_mode, text=self.tr("rad_before"), variable=self.var_filter_mode, value="before")
        self.rb_filter_before.pack(side=LEFT, padx=10)

        # Section 5: NotebookLM locator (runtime)
        lf_locator = tb.Labelframe(parent, text=self.tr("sec_locator"), padding=10)
        lf_locator.pack(fill=X, pady=5)
        self._add_section_help(lf_locator, "tip_section_run_locator")
        tb.Label(lf_locator, text=self.tr("lbl_locator_merged")).pack(anchor="w")
        self.var_locator_merged = tk.StringVar()
        self.cb_locator_merged = tb.Combobox(lf_locator, textvariable=self.var_locator_merged, state="readonly", values=[])
        self.cb_locator_merged.pack(fill=X, pady=(0, 4))
        self._attach_tooltip(self.cb_locator_merged, "tip_input_locator_merged")
        row_locator = tb.Frame(lf_locator)
        row_locator.pack(fill=X)
        tb.Label(row_locator, text=self.tr("lbl_locator_page")).pack(side=LEFT)
        self.var_locator_page = tk.StringVar()
        self.ent_locator_page = tb.Entry(row_locator, textvariable=self.var_locator_page, width=8)
        self.ent_locator_page.pack(side=LEFT, padx=(6, 12))
        self._attach_tooltip(self.ent_locator_page, "tip_input_locator_page")
        tb.Label(row_locator, text=self.tr("lbl_locator_id")).pack(side=LEFT)
        self.var_locator_short_id = tk.StringVar()
        self.ent_locator_short_id = tb.Entry(row_locator, textvariable=self.var_locator_short_id, width=14)
        self.ent_locator_short_id.pack(side=LEFT, padx=(6, 0))
        self._attach_tooltip(self.ent_locator_short_id, "tip_input_locator_short_id")
        row_locator_btn = tb.Labelframe(lf_locator, text=self.tr("lbl_locator_group_actions"), padding=(8, 6))
        row_locator_btn.pack(fill=X, pady=(6, 0))
        self.btn_locator_refresh = tb.Button(row_locator_btn, text=self.tr("btn_locator_refresh"), command=self.refresh_locator_maps, bootstyle="secondary-outline", width=12)
        self.btn_locator_refresh.pack(side=LEFT, padx=2)
        self._attach_tooltip(self.btn_locator_refresh, "tip_locator_refresh")
        self.btn_locator_locate = tb.Button(row_locator_btn, text=self.tr("btn_locator_locate"), command=self.run_locator_query, bootstyle="primary", width=10)
        self.btn_locator_locate.pack(side=LEFT, padx=2)
        self._attach_tooltip(self.btn_locator_locate, "tip_locator_locate")
        self.btn_locator_open_file = tb.Button(row_locator_btn, text=self.tr("btn_locator_open_file"), command=self.open_locator_file, bootstyle="success-outline", width=10)
        self.btn_locator_open_file.pack(side=LEFT, padx=2)
        self._attach_tooltip(self.btn_locator_open_file, "tip_locator_open_file")
        self.btn_locator_open_dir = tb.Button(row_locator_btn, text=self.tr("btn_locator_open_dir"), command=self.open_locator_folder, bootstyle="info-outline", width=10)
        self.btn_locator_open_dir.pack(side=LEFT, padx=2)
        self._attach_tooltip(self.btn_locator_open_dir, "tip_locator_open_dir")
        row_locator_btn2 = tb.Labelframe(lf_locator, text=self.tr("lbl_locator_group_external"), padding=(8, 6))
        row_locator_btn2.pack(fill=X, pady=(4, 0))
        self.btn_locator_everything = tb.Button(row_locator_btn2, text=self.tr("btn_locator_everything"), command=self.search_with_everything, bootstyle="warning-outline", width=14)
        self.btn_locator_everything.pack(side=LEFT, padx=2)
        self._attach_tooltip(self.btn_locator_everything, "tip_locator_everything")
        self.btn_locator_copy_listary = tb.Button(row_locator_btn2, text=self.tr("btn_locator_copy_listary"), command=self.copy_listary_query, bootstyle="dark-outline", width=16)
        self.btn_locator_copy_listary.pack(side=LEFT, padx=2)
        self._attach_tooltip(self.btn_locator_copy_listary, "tip_locator_listary")
        self.var_locator_result = tk.StringVar(value=self.tr("msg_locator_waiting"))
        tb.Label(lf_locator, textvariable=self.var_locator_result, bootstyle="secondary", wraplength=880, justify=LEFT).pack(anchor="w", pady=(6, 0))
        self.last_locate_record = None
        self._set_locator_action_state(False)
        self._auto_attach_action_tooltips(lf_mode)
        self._auto_attach_action_tooltips(lf_paths)
        self._auto_attach_action_tooltips(lf_settings)
        self._auto_attach_action_tooltips(lf_extra)
        self._auto_attach_action_tooltips(lf_locator)
        self._attach_tooltip(self.entry_temp_sandbox_root, "tip_input_sandbox_root")
        self._attach_tooltip(cb_strat, "tip_input_strategy")
        self._attach_tooltip(self.ent_date, "tip_input_date")
        self._bind_var_validation(self.var_locator_page, lambda: self._normalize_then_validate(self.var_locator_page, self._normalize_numeric_var, "locator"))
        self._bind_var_validation(self.var_locator_short_id, lambda: self._normalize_then_validate(self.var_locator_short_id, self._normalize_short_id_var, "locator"))
        self._bind_var_validation(self.var_date_str, lambda: self._normalize_then_validate(self.var_date_str, self._normalize_date_var, "run"))
        self._bind_var_validation(self.var_enable_date_filter, lambda: self.validate_runtime_inputs(silent=False, scope="run"))

    def _build_config_tab_content(self, parent):
        # Config path and log path
        lf_cfg_path = tb.Labelframe(parent, text=self.tr("sec_paths"), padding=10)
        lf_cfg_path.pack(fill=X, pady=5)
        self._add_section_help(lf_cfg_path, "tip_section_cfg_paths")
        self.var_config_path = tk.StringVar(value=self.config_path)
        self._create_path_row(lf_cfg_path, "lbl_config", self.var_config_path, self.open_config_folder, None)

        # Process & limits (persistent defaults)
        lf_proc = tb.Labelframe(parent, text=self.tr("sec_process"), padding=10)
        lf_proc.pack(fill=X, pady=5)
        self._add_section_help(lf_proc, "tip_section_cfg_process")
        tb.Label(lf_proc, text=self.tr("lbl_kill_mode"), font=("System", 9)).pack(anchor="w")
        self.var_kill_mode = tk.StringVar(value=KILL_MODE_AUTO)
        frm_kill = tb.Frame(lf_proc)
        frm_kill.pack(fill=X)
        tb.Radiobutton(frm_kill, text=self.tr("rad_auto_kill"), variable=self.var_kill_mode, value=KILL_MODE_AUTO).pack(side=LEFT)
        tb.Radiobutton(frm_kill, text=self.tr("rad_keep_running"), variable=self.var_kill_mode, value=KILL_MODE_KEEP).pack(side=LEFT, padx=10)

        tb.Label(lf_proc, text=self.tr("lbl_log_folder"), font=("System", 9)).pack(anchor="w", pady=(5, 0))
        frm_log = tb.Frame(lf_proc)
        frm_log.pack(fill=X)
        self.var_log_folder = tk.StringVar(value="./logs")
        self.ent_log_folder = tb.Entry(frm_log, textvariable=self.var_log_folder)
        self.ent_log_folder.pack(side=LEFT, fill=X, expand=YES)
        self._attach_tooltip(self.ent_log_folder, "tip_input_log_folder")
        self.btn_log_folder = tb.Button(frm_log, text=self.tr("btn_browse"), command=self.browse_log_folder, bootstyle="outline", width=3)
        self.btn_log_folder.pack(side=LEFT, padx=2)
        self._attach_tooltip(self.btn_log_folder, "tip_choose_log")

        frm_time = tb.Frame(lf_proc)
        frm_time.pack(fill=X, pady=5)
        self.var_timeout_seconds = tk.StringVar(value="60")
        tb.Label(frm_time, text=self.tr("lbl_gen_timeout")).grid(row=0, column=0, sticky="e")
        self.ent_timeout_seconds = tb.Entry(frm_time, textvariable=self.var_timeout_seconds, width=5)
        self.ent_timeout_seconds.grid(row=0, column=1, sticky="w", padx=5)
        self._attach_tooltip(self.ent_timeout_seconds, "tip_input_timeout_seconds")
        self.var_pdf_wait_seconds = tk.StringVar(value="15")
        tb.Label(frm_time, text=self.tr("lbl_pdf_wait")).grid(row=0, column=2, sticky="e")
        self.ent_pdf_wait_seconds = tb.Entry(frm_time, textvariable=self.var_pdf_wait_seconds, width=5)
        self.ent_pdf_wait_seconds.grid(row=0, column=3, sticky="w", padx=5)
        self._attach_tooltip(self.ent_pdf_wait_seconds, "tip_input_pdf_wait_seconds")
        self.var_ppt_timeout_seconds = tk.StringVar(value="180")
        tb.Label(frm_time, text=self.tr("lbl_ppt_timeout")).grid(row=1, column=0, sticky="e")
        self.ent_ppt_timeout_seconds = tb.Entry(frm_time, textvariable=self.var_ppt_timeout_seconds, width=5)
        self.ent_ppt_timeout_seconds.grid(row=1, column=1, sticky="w", padx=5)
        self._attach_tooltip(self.ent_ppt_timeout_seconds, "tip_input_ppt_timeout_seconds")
        self.var_ppt_pdf_wait_seconds = tk.StringVar(value="30")
        tb.Label(frm_time, text=self.tr("lbl_ppt_wait")).grid(row=1, column=2, sticky="e")
        self.ent_ppt_pdf_wait_seconds = tb.Entry(frm_time, textvariable=self.var_ppt_pdf_wait_seconds, width=5)
        self.ent_ppt_pdf_wait_seconds.grid(row=1, column=3, sticky="w", padx=5)
        self._attach_tooltip(self.ent_ppt_pdf_wait_seconds, "tip_input_ppt_pdf_wait_seconds")
        self.var_max_merge_size_mb = tk.StringVar(value="80")
        tb.Label(frm_time, text=self.tr("lbl_max_mb")).grid(row=2, column=0, sticky="e")
        self.ent_max_merge_size_mb = tb.Entry(frm_time, textvariable=self.var_max_merge_size_mb, width=5)
        self.ent_max_merge_size_mb.grid(row=2, column=1, sticky="w", padx=5)
        self._attach_tooltip(self.ent_max_merge_size_mb, "tip_input_max_merge_size_mb")

        tb.Separator(lf_proc).pack(fill=X, pady=6)
        tb.Label(lf_proc, text=self.tr("lbl_tooltip_cfg"), font=("System", 9, "bold")).pack(anchor="w")
        frm_tip = tb.Frame(lf_proc)
        frm_tip.pack(fill=X, pady=(4, 0))
        self.var_tooltip_auto_theme = tk.IntVar(value=1)
        self.chk_tooltip_auto_theme = tb.Checkbutton(frm_tip, text=self.tr("chk_tooltip_auto_theme"), variable=self.var_tooltip_auto_theme)
        self.chk_tooltip_auto_theme.grid(row=0, column=0, sticky="w", padx=(0, 8))
        self.var_tooltip_delay_ms = tk.StringVar(value="300")
        tb.Label(frm_tip, text=self.tr("lbl_tooltip_delay")).grid(row=0, column=1, sticky="e")
        self.ent_tooltip_delay = tb.Entry(frm_tip, textvariable=self.var_tooltip_delay_ms, width=6)
        self.ent_tooltip_delay.grid(row=0, column=2, sticky="w", padx=4)
        self.var_tooltip_font_size = tk.StringVar(value="9")
        tb.Label(frm_tip, text=self.tr("lbl_tooltip_font_size")).grid(row=0, column=3, sticky="e")
        self.ent_tooltip_font_size = tb.Entry(frm_tip, textvariable=self.var_tooltip_font_size, width=6)
        self.ent_tooltip_font_size.grid(row=0, column=4, sticky="w", padx=4)
        self.var_tooltip_bg = tk.StringVar(value="#FFF7D6")
        tb.Label(frm_tip, text=self.tr("lbl_tooltip_bg")).grid(row=1, column=1, sticky="e")
        self.ent_tooltip_bg = tb.Entry(frm_tip, textvariable=self.var_tooltip_bg, width=10)
        self.ent_tooltip_bg.grid(row=1, column=2, sticky="w", padx=4)
        self.btn_pick_tooltip_bg = tb.Button(frm_tip, text="🎨", width=3, command=lambda: self.pick_tooltip_color("bg"))
        self.btn_pick_tooltip_bg.grid(row=1, column=2, sticky="e", padx=(0, 0))
        self._attach_tooltip(self.btn_pick_tooltip_bg, "tip_pick_color")
        self.var_tooltip_fg = tk.StringVar(value="#202124")
        tb.Label(frm_tip, text=self.tr("lbl_tooltip_fg")).grid(row=1, column=3, sticky="e")
        self.ent_tooltip_fg = tb.Entry(frm_tip, textvariable=self.var_tooltip_fg, width=10)
        self.ent_tooltip_fg.grid(row=1, column=4, sticky="w", padx=4)
        self.btn_pick_tooltip_fg = tb.Button(frm_tip, text="🎨", width=3, command=lambda: self.pick_tooltip_color("fg"))
        self.btn_pick_tooltip_fg.grid(row=1, column=4, sticky="e", padx=(0, 0))
        self._attach_tooltip(self.btn_pick_tooltip_fg, "tip_pick_color")
        self.lbl_tooltip_bg_preview = tb.Label(frm_tip, text=self.tr("lbl_tooltip_preview_bg"), width=12, anchor="center")
        self.lbl_tooltip_bg_preview.grid(row=2, column=1, columnspan=2, sticky="w", pady=(4, 0))
        self.lbl_tooltip_fg_preview = tb.Label(frm_tip, text=self.tr("lbl_tooltip_preview_fg"), width=12, anchor="center")
        self.lbl_tooltip_fg_preview.grid(row=2, column=3, columnspan=2, sticky="w", pady=(4, 0))
        self.lbl_tooltip_sample_preview = tb.Label(frm_tip, text=self.tr("lbl_tooltip_preview_sample"), anchor="center", padding=(8, 4))
        self.lbl_tooltip_sample_preview.grid(row=3, column=1, columnspan=4, sticky="ew", pady=(4, 0))
        self.btn_apply_tooltip = tb.Button(frm_tip, text=self.tr("btn_apply_tooltip"), command=self.apply_tooltip_settings, bootstyle="secondary-outline")
        self.btn_apply_tooltip.grid(row=1, column=0, sticky="w", pady=(4, 0))
        self._attach_tooltip(self.btn_apply_tooltip, "tip_apply_tooltip")
        self.btn_reset_tooltip = tb.Button(frm_tip, text=self.tr("btn_reset_tooltip"), command=self.reset_tooltip_settings, bootstyle="secondary")
        self.btn_reset_tooltip.grid(row=2, column=0, sticky="w", pady=(4, 0))
        self._attach_tooltip(self.btn_reset_tooltip, "tip_reset_tooltip")
        for v in (self.var_tooltip_delay_ms, self.var_tooltip_font_size, self.var_tooltip_bg, self.var_tooltip_fg, self.var_tooltip_auto_theme):
            v.trace_add("write", lambda *_: self.validate_tooltip_inputs(silent=True))

        # Exclude/keyword defaults
        lf_lists = tb.Labelframe(parent, text=self.tr("sec_lists"), padding=10)
        lf_lists.pack(fill=X, pady=5)
        self._add_section_help(lf_lists, "tip_section_cfg_lists")
        tb.Label(lf_lists, text=self.tr("lbl_excluded")).pack(anchor="w")
        self.txt_excluded_folders = ScrolledText(lf_lists, height=4, font=("Consolas", 8), bootstyle="default")
        self.txt_excluded_folders.pack(fill=X, pady=(0, 5))
        tb.Label(lf_lists, text=self.tr("lbl_keywords")).pack(anchor="w")
        self.txt_price_keywords = ScrolledText(lf_lists, height=3, font=("Consolas", 8), bootstyle="default")
        self.txt_price_keywords.pack(fill=X)

        # Emphasized save in config tab
        cfg_actions = tb.Frame(parent)
        cfg_actions.pack(fill=X, pady=(8, 12))
        self.btn_save_cfg_tab = tb.Button(
            cfg_actions,
            text=self.tr("btn_save_cfg"),
            command=self._save_settings_to_file,
            bootstyle="success",
            width=20,
        )
        self.btn_save_cfg_tab.pack(side=LEFT)
        self._attach_tooltip(self.btn_save_cfg_tab, "tip_save_config")
        self._auto_attach_action_tooltips(lf_cfg_path)
        self._auto_attach_action_tooltips(lf_proc)
        self._auto_attach_action_tooltips(lf_lists)
        self._auto_attach_action_tooltips(cfg_actions)
        self._attach_tooltip(self.txt_excluded_folders, "tip_input_excluded_folders")
        self._attach_tooltip(self.txt_price_keywords, "tip_input_price_keywords")
        self._bind_var_validation(self.var_timeout_seconds, lambda: self._normalize_then_validate(self.var_timeout_seconds, self._normalize_numeric_var, "config"))
        self._bind_var_validation(self.var_pdf_wait_seconds, lambda: self._normalize_then_validate(self.var_pdf_wait_seconds, self._normalize_numeric_var, "config"))
        self._bind_var_validation(self.var_ppt_timeout_seconds, lambda: self._normalize_then_validate(self.var_ppt_timeout_seconds, self._normalize_numeric_var, "config"))
        self._bind_var_validation(self.var_ppt_pdf_wait_seconds, lambda: self._normalize_then_validate(self.var_ppt_pdf_wait_seconds, self._normalize_numeric_var, "config"))
        self._bind_var_validation(self.var_max_merge_size_mb, lambda: self._normalize_then_validate(self.var_max_merge_size_mb, self._normalize_numeric_var, "config"))

    def _build_footer(self, parent):
        """底部按钮 + Status"""
        parent.columnconfigure(1, weight=1) # Spacer between status and buttons
        
        # Status Label (Left)
        if not hasattr(self, "var_status"):
             self.var_status = tk.StringVar(value=self.tr("status_ready"))
        
        tb.Label(parent, textvariable=self.var_status, bootstyle="secondary").grid(row=0, column=0, padx=10, sticky="w")
        
        # Buttons
        # Show Logs
        self.btn_toggle_logs = tb.Button(parent, text=self.tr("btn_toggle_logs"), command=self._toggle_logs, bootstyle="info-outline")
        self.btn_toggle_logs.grid(row=0, column=2, padx=5)
        self._attach_tooltip(self.btn_toggle_logs, "tip_toggle_logs")

        # Save
        self.btn_save_cfg = tb.Button(parent, text=self.tr("btn_save_cfg"), command=self._save_settings_to_file, bootstyle="secondary-outline")
        self.btn_save_cfg.grid(row=0, column=3, padx=5)
        self._attach_tooltip(self.btn_save_cfg, "tip_save_config")
        
        # Start
        self.btn_start = tb.Button(parent, text=self.tr("btn_start"), command=self._on_click_start, bootstyle="success", width=20)
        self.btn_start.grid(row=0, column=4, padx=5)
        self._attach_tooltip(self.btn_start, "tip_start_task")
        
        # Stop
        self.btn_stop = tb.Button(parent, text=self.tr("btn_stop"), command=self._on_click_stop, bootstyle="danger-outline", state="disabled")
        self.btn_stop.grid(row=0, column=5, padx=5)
        self._attach_tooltip(self.btn_stop, "tip_stop_task")
        self._auto_attach_action_tooltips(parent)

    def _toggle_logs(self):
        """Toggle log pane visibility"""
        if self.log_pane in self.paned.panes():
             self.paned.forget(self.log_pane)
        else:
             self.paned.add(self.log_pane, weight=1)

    # ===================== UI 联动更新 (Adapt for new structure) =====================

    def _on_run_mode_change(self):
        mode = self.var_run_mode.get()
        is_collect = mode == MODE_COLLECT_ONLY
        is_convert = mode in (MODE_CONVERT_ONLY, MODE_CONVERT_THEN_MERGE)
        is_merge_related = mode in (MODE_CONVERT_THEN_MERGE, MODE_MERGE_ONLY)

        # Collect options
        if is_collect:
             for child in self.frm_collect_opts.winfo_children(): child.configure(state="normal")
        else:
             for child in self.frm_collect_opts.winfo_children(): child.configure(state="disabled")

        # Engine & Sandbox (Enable only if converting)
        state_exec = "normal" if not (is_collect or mode == MODE_MERGE_ONLY) else "disabled"
        for child in self.group_exec.winfo_children():
            try: child.configure(state=state_exec)
            except: pass
        
        # Trigger sandbox toggle to refresh sub-widgets
        self._on_toggle_sandbox()

        # Merge Options
        state_merge = "normal" if is_merge_related else "disabled"
        self.lbl_merge.configure(state=state_merge)
        self.chk_enable_merge.configure(state=state_merge)
        for child in self.frm_merge_opts.winfo_children():
             try: child.configure(state=state_merge)
             except: pass

    def _on_toggle_date_filter(self):
        enabled = bool(self.var_enable_date_filter.get())
        state = "normal" if enabled else "disabled"
        # DateEntry complicates state, usually just disable the internal entry key binding or similar
        # For tb.DateEntry, we can try disabling the frame or buttons
        for child in self.frm_date.winfo_children():
            try: child.configure(state=state)
            except: pass
        self.ent_date.configure(state=state)

    def _on_toggle_sandbox(self):
        mode = self.var_run_mode.get()
        is_disabled_globally = mode == MODE_COLLECT_ONLY or mode == MODE_MERGE_ONLY
        
        # If group is disabled, sandbox should look disabled
        if is_disabled_globally:
            self.chk_enable_sandbox.configure(state="disabled")
            self.entry_temp_sandbox_root.configure(state="disabled")
            self.btn_temp_sandbox_root.configure(state="disabled")
            return

        # Otherwise standard toggle logic
        self.chk_enable_sandbox.configure(state="normal")
        enabled = bool(self.var_enable_sandbox.get())
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
            self.refresh_locator_maps()

    def open_source_folder(self):
        path = self.var_source_folder.get().strip()
        self._open_path(path)

    def open_target_folder(self):
        path = self.var_target_folder.get().strip()
        self._open_path(path)

    def open_config_folder(self):
        folder = os.path.dirname(self.config_path)
        self._open_path(folder)

    def _open_path(self, path):
        if not path or not os.path.exists(path):
            return
        if sys.platform == "win32":
            os.startfile(path)
        elif sys.platform == "darwin":
            subprocess.run(["open", path])
        else:
            try:
                subprocess.run(["xdg-open", path])
            except Exception:
                pass


    def browse_temp_sandbox_root(self):
        path = filedialog.askdirectory(title="选择临时转换根目录")
        if path:
            self.var_temp_sandbox_root.set(path)

    def browse_log_folder(self):
        path = filedialog.askdirectory(title="选择日志目录")
        if path:
            self.var_log_folder.set(path)

    def get_locator_map_dir(self):
        target = self.var_target_folder.get().strip()
        if not target:
            return ""
        return os.path.join(target, "_MERGED")

    def _set_locator_action_state(self, enabled: bool):
        state = "normal" if enabled else "disabled"
        for btn_name in (
            "btn_locator_open_file",
            "btn_locator_open_dir",
            "btn_locator_everything",
            "btn_locator_copy_listary",
        ):
            btn = getattr(self, btn_name, None)
            if btn is not None:
                btn.configure(state=state)

    def refresh_locator_maps(self):
        map_dir = self.get_locator_map_dir()
        self.locator_short_id_index = {}
        self.last_locate_record = None
        self._set_locator_action_state(False)
        if not map_dir or not os.path.isdir(map_dir):
            self.cb_locator_merged.configure(values=[])
            self.var_locator_result.set(self.tr("msg_locator_no_merged_dir"))
            return

        merged_names = []
        map_files = glob.glob(os.path.join(map_dir, "*.map.json"))
        map_files.sort(key=lambda p: os.path.getmtime(p), reverse=True)
        for p in map_files:
            stem = os.path.basename(p)[:-9]  # remove .map.json
            merged_names.append(f"{stem}.pdf")
            try:
                with open(p, "r", encoding="utf-8") as f:
                    payload = json.load(f)
                for rec in payload.get("records", []):
                    sid = str(rec.get("source_short_id", "")).strip().upper()
                    if sid:
                        self.locator_short_id_index.setdefault(sid, []).append(rec)
            except Exception:
                continue

        self.cb_locator_merged.configure(values=merged_names)
        current = self.var_locator_merged.get().strip()
        if merged_names:
            if current not in merged_names:
                self.var_locator_merged.set(merged_names[0])
        else:
            self.var_locator_merged.set("")
        self.var_locator_result.set(self.tr("msg_locator_loaded_maps").format(len(merged_names)))

    def run_locator_query(self):
        if not self.validate_runtime_inputs(silent=False, scope="locator"):
            return
        map_dir = self.get_locator_map_dir()
        if not os.path.isdir(map_dir):
            self.var_locator_result.set(self.tr("msg_locator_map_missing"))
            return

        merged_name = self.var_locator_merged.get().strip()
        page_raw = self.var_locator_page.get().strip()
        short_id = self.var_locator_short_id.get().strip()
        priority_note = ""

        result = None
        if page_raw and short_id and merged_name:
            priority_note = f"{self.tr('msg_locator_priority_page')} "
        if page_raw and merged_name:
            try:
                page = int(page_raw)
            except ValueError:
                self.var_locator_result.set(self.tr("msg_locator_invalid_page"))
                return
            result = locate_by_page(merged_name, page, map_dir)
        elif short_id:
            sid = short_id.upper()
            cache_hits = self.locator_short_id_index.get(sid, [])
            if len(cache_hits) == 1:
                r = cache_hits[0]
                t = SimpleNamespace(
                    source_filename=r.get("source_filename", ""),
                    start_page_1based=int(r.get("start_page_1based", 0)),
                    end_page_1based=int(r.get("end_page_1based", 0)),
                    source_short_id=r.get("source_short_id", ""),
                    source_abspath=r.get("source_abspath", ""),
                    source_md5=r.get("source_md5", ""),
                )
                result = SimpleNamespace()
                result.record = t
                result.alternatives = []
                result.status = "ok"
            elif len(cache_hits) > 1:
                result = SimpleNamespace()
                result.record = None
                result.status = "ambiguous"
                alts = []
                for r in cache_hits[:2]:
                    a = SimpleNamespace(
                        source_filename=r.get("source_filename", ""),
                        start_page_1based=int(r.get("start_page_1based", 0)),
                        end_page_1based=int(r.get("end_page_1based", 0)),
                    )
                    alts.append(a)
                result.alternatives = alts
            else:
                result = locate_by_short_id(short_id, map_dir)
        else:
            self.var_locator_result.set(self.tr("msg_locator_need_input"))
            return

        if result.record:
            self.last_locate_record = result.record
            self.var_locator_result.set(
                priority_note + self.tr("msg_locator_hit").format(
                    result.record.source_filename,
                    result.record.start_page_1based,
                    result.record.end_page_1based,
                    result.record.source_short_id,
                )
            )
            if self._read_config_value(["listary", "copy_query_on_locate"], False):
                self.copy_listary_query(silent=True)
            self._set_locator_action_state(True)
            return

        self.last_locate_record = None
        self._set_locator_action_state(False)
        if result.alternatives:
            alt = "；".join(
                [f"{x.source_filename}({x.start_page_1based}-{x.end_page_1based})" for x in result.alternatives[:2]]
            )
            self.var_locator_result.set(priority_note + self.tr("msg_locator_miss_alt").format(alt))
        else:
            self.var_locator_result.set(priority_note + self.tr("msg_locator_status").format(result.status))

    def open_locator_file(self):
        if not self.last_locate_record:
            self.var_locator_result.set(self.tr("msg_locator_run_first"))
            return
        path = self.last_locate_record.source_abspath
        if not os.path.exists(path):
            self.var_locator_result.set(self.tr("msg_locator_file_missing").format(path))
            return
        self._open_path(path)

    def open_locator_folder(self):
        if not self.last_locate_record:
            self.var_locator_result.set(self.tr("msg_locator_run_first"))
            return
        path = self.last_locate_record.source_abspath
        folder = os.path.dirname(path)
        if not os.path.isdir(folder):
            self.var_locator_result.set(self.tr("msg_locator_dir_missing").format(folder))
            return
        if sys.platform == "win32":
            subprocess.run(["explorer", "/select,", path], check=False)
        else:
            self._open_path(folder)

    def search_with_everything(self):
        if not self.last_locate_record:
            self.var_locator_result.set(self.tr("msg_locator_run_first"))
            return

        if not self._read_config_value(["everything", "enabled"], True):
            self.var_locator_result.set(self.tr("msg_locator_everything_disabled"))
            return

        es_path = self._read_config_value(["everything", "es_path"], "")
        timeout_ms = self._read_config_value(["everything", "timeout_ms"], 1500)
        prefer_path_exact = self._read_config_value(["everything", "prefer_path_exact"], True)
        adapter = EverythingAdapter(es_path=es_path, timeout_ms=int(timeout_ms))
        if not adapter.is_available():
            self.var_locator_result.set(self.tr("msg_locator_everything_notfound"))
            return

        directory = os.path.dirname(self.last_locate_record.source_abspath) if prefer_path_exact else ""
        ret = adapter.run_query(self.last_locate_record.source_filename, directory)
        if ret.ok:
            self.var_locator_result.set(self.tr("msg_locator_everything_ok"))
        else:
            self.var_locator_result.set(self.tr("msg_locator_everything_fail").format(ret.stderr))

    def copy_listary_query(self, silent=False):
        if not self.last_locate_record:
            if not silent:
                self.var_locator_result.set(self.tr("msg_locator_run_first"))
            return
        query = build_listary_query(
            self.last_locate_record.source_short_id,
            self.last_locate_record.source_md5,
            self.last_locate_record.source_filename,
            self.last_locate_record.source_abspath,
        )
        self.clipboard_clear()
        self.clipboard_append(query)
        self.update_idletasks()
        if not silent:
            self.var_locator_result.set(self.tr("msg_locator_listary_copied"))

    def _read_config_value(self, key_path, default_value):
        try:
            with open(self.config_path, "r", encoding="utf-8") as f:
                cfg = json.load(f)
            val = cfg
            for k in key_path:
                if not isinstance(val, dict):
                    return default_value
                val = val.get(k)
            return default_value if val is None else val
        except Exception:
            return default_value

    @staticmethod
    def _is_valid_hex_color(text):
        return bool(re.fullmatch(r"#(?:[0-9A-Fa-f]{6})", str(text).strip()))

    def _set_entry_valid_state(self, entry, valid=True):
        if entry is None:
            return
        try:
            entry.configure(bootstyle="default" if valid else "danger")
        except Exception:
            pass

    def _bind_var_validation(self, var, callback):
        if var is None:
            return
        try:
            var.trace_add("write", lambda *_: callback())
        except Exception:
            pass

    def _to_halfwidth(self, text):
        if text is None:
            return ""
        out = []
        for ch in str(text):
            code = ord(ch)
            if code == 0x3000:
                out.append(" ")
            elif 0xFF01 <= code <= 0xFF5E:
                out.append(chr(code - 0xFEE0))
            else:
                out.append(ch)
        return "".join(out)

    def _normalize_numeric_var(self, var):
        raw = var.get()
        norm = self._to_halfwidth(raw).strip()
        if norm != raw:
            var.set(norm)

    def _normalize_short_id_var(self, var):
        raw = var.get()
        norm = self._to_halfwidth(raw).strip().replace(" ", "").upper()
        if norm != raw:
            var.set(norm)

    def _normalize_date_var(self, var):
        raw = var.get()
        norm = self._to_halfwidth(raw).strip()
        norm = norm.replace("/", "-").replace(".", "-")
        norm = re.sub(r"-{2,}", "-", norm)
        m = re.fullmatch(r"(\d{4})-(\d{1,2})-(\d{1,2})", norm)
        if m:
            norm = f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
        if norm != raw:
            var.set(norm)

    def _normalize_then_validate(self, var, normalizer, scope):
        if self._normalizing_inputs:
            return
        self._normalizing_inputs = True
        try:
            normalizer(var)
        finally:
            self._normalizing_inputs = False
        self.validate_runtime_inputs(silent=False, scope=scope)

    def _set_status_validation_error(self, message):
        if hasattr(self, "var_status"):
            self.var_status.set(message)
        if hasattr(self, "var_locator_result"):
            self.var_locator_result.set(message)

    @staticmethod
    def _is_valid_short_id(text):
        return bool(re.fullmatch(r"[0-9A-Za-z]{4,32}", str(text).strip()))

    def validate_runtime_inputs(self, silent=True, scope="all"):
        first_error = None

        def _mark(entry, ok, message_key, label_key):
            nonlocal first_error
            self._set_entry_valid_state(entry, ok)
            if (not ok) and first_error is None:
                first_error = self.tr(message_key).format(self.tr(label_key))

        # Locator quick check
        if scope in ("all", "locator"):
            page_raw = self.var_locator_page.get().strip() if hasattr(self, "var_locator_page") else ""
            page_ok = True
            if page_raw:
                page_ok = page_raw.isdigit() and int(page_raw) > 0
            _mark(getattr(self, "ent_locator_page", None), page_ok, "msg_validation_invalid_number", "lbl_locator_page")

            short_id = self.var_locator_short_id.get().strip() if hasattr(self, "var_locator_short_id") else ""
            sid_ok = True if not short_id else self._is_valid_short_id(short_id)
            _mark(getattr(self, "ent_locator_short_id", None), sid_ok, "msg_validation_invalid_short_id", "lbl_locator_id")

        # Runtime date filter check
        if scope in ("all", "run"):
            date_entry_widget = None
            if hasattr(self, "ent_date"):
                date_entry_widget = getattr(self.ent_date, "entry", self.ent_date)
            date_ok = True
            if hasattr(self, "var_enable_date_filter") and self.var_enable_date_filter.get():
                date_str = self.var_date_str.get().strip() if hasattr(self, "var_date_str") else ""
                try:
                    datetime.strptime(date_str, "%Y-%m-%d")
                except Exception:
                    date_ok = False
            _mark(date_entry_widget, date_ok, "msg_validation_invalid_date", "lbl_filter_date")

        # Config numeric defaults
        if scope in ("all", "config"):
            numeric_fields = [
                ("var_timeout_seconds", "ent_timeout_seconds", "lbl_gen_timeout"),
                ("var_pdf_wait_seconds", "ent_pdf_wait_seconds", "lbl_pdf_wait"),
                ("var_ppt_timeout_seconds", "ent_ppt_timeout_seconds", "lbl_ppt_timeout"),
                ("var_ppt_pdf_wait_seconds", "ent_ppt_pdf_wait_seconds", "lbl_ppt_wait"),
                ("var_max_merge_size_mb", "ent_max_merge_size_mb", "lbl_max_mb"),
            ]
            for var_name, ent_name, label_key in numeric_fields:
                raw = getattr(self, var_name).get().strip() if hasattr(self, var_name) else ""
                ok = raw.isdigit() and int(raw) > 0
                _mark(getattr(self, ent_name, None), ok, "msg_validation_invalid_number", label_key)

        if first_error and not silent:
            self._set_status_validation_error(first_error)
        elif not first_error and not silent and hasattr(self, "var_status"):
            self.var_status.set(self.tr("status_ready"))

        return first_error is None

    def _update_tooltip_color_preview(self):
        bg = self.var_tooltip_bg.get().strip() if hasattr(self, "var_tooltip_bg") else ""
        fg = self.var_tooltip_fg.get().strip() if hasattr(self, "var_tooltip_fg") else ""
        bg_valid = self._is_valid_hex_color(bg)
        fg_valid = self._is_valid_hex_color(fg)
        if hasattr(self, "lbl_tooltip_bg_preview"):
            try:
                self.lbl_tooltip_bg_preview.configure(
                    background=(bg if bg_valid else "#F8D7DA"),
                    foreground="#202124",
                )
            except Exception:
                pass
        if hasattr(self, "lbl_tooltip_fg_preview"):
            try:
                self.lbl_tooltip_fg_preview.configure(
                    background="#FFFFFF",
                    foreground=(fg if fg_valid else "#D32F2F"),
                )
            except Exception:
                pass
        if hasattr(self, "lbl_tooltip_sample_preview"):
            try:
                preview_bg = bg if bg_valid else "#F8D7DA"
                preview_fg = fg if fg_valid else "#D32F2F"
                self.lbl_tooltip_sample_preview.configure(
                    background=preview_bg,
                    foreground=preview_fg,
                    font=(self.tooltip_font_family, self.tooltip_font_size),
                )
            except Exception:
                pass

    def pick_tooltip_color(self, target):
        if target not in ("bg", "fg"):
            return
        initial = self.var_tooltip_bg.get().strip() if target == "bg" else self.var_tooltip_fg.get().strip()
        _, hex_color = colorchooser.askcolor(color=initial, title=self.tr("tip_pick_color"))
        if not hex_color:
            return
        hex_color = hex_color.upper()
        if target == "bg":
            self.var_tooltip_bg.set(hex_color)
        else:
            self.var_tooltip_fg.set(hex_color)
        self.validate_tooltip_inputs(silent=True)

    def validate_tooltip_inputs(self, silent=False):
        invalid_label = None
        if hasattr(self, "var_tooltip_delay_ms"):
            ok = str(self.var_tooltip_delay_ms.get()).strip().isdigit()
            self._set_entry_valid_state(getattr(self, "ent_tooltip_delay", None), ok)
            if not ok:
                invalid_label = self.tr("lbl_tooltip_delay")
        if hasattr(self, "var_tooltip_font_size"):
            ok = str(self.var_tooltip_font_size.get()).strip().isdigit()
            self._set_entry_valid_state(getattr(self, "ent_tooltip_font_size", None), ok)
            if not ok and invalid_label is None:
                invalid_label = self.tr("lbl_tooltip_font_size")
        if hasattr(self, "var_tooltip_bg"):
            ok = self._is_valid_hex_color(self.var_tooltip_bg.get())
            self._set_entry_valid_state(getattr(self, "ent_tooltip_bg", None), ok)
            if not ok and invalid_label is None:
                invalid_label = self.tr("lbl_tooltip_bg")
        if hasattr(self, "var_tooltip_fg"):
            ok = self._is_valid_hex_color(self.var_tooltip_fg.get())
            self._set_entry_valid_state(getattr(self, "ent_tooltip_fg", None), ok)
            if not ok and invalid_label is None:
                invalid_label = self.tr("lbl_tooltip_fg")

        self._update_tooltip_color_preview()
        if invalid_label and not silent:
            if invalid_label in (self.tr("lbl_tooltip_bg"), self.tr("lbl_tooltip_fg")):
                self.var_locator_result.set(self.tr("msg_tooltip_invalid_color").format(invalid_label))
            else:
                self.var_locator_result.set(self.tr("msg_tooltip_invalid_number").format(invalid_label))
        return invalid_label is None

    def reset_tooltip_settings(self):
        self.var_tooltip_delay_ms.set(str(self.TOOLTIP_DEFAULTS["tooltip_delay_ms"]))
        self.var_tooltip_font_size.set(str(self.TOOLTIP_DEFAULTS["tooltip_font_size"]))
        self.var_tooltip_bg.set(self.TOOLTIP_DEFAULTS["tooltip_bg"])
        self.var_tooltip_fg.set(self.TOOLTIP_DEFAULTS["tooltip_fg"])
        self.var_tooltip_auto_theme.set(1 if self.TOOLTIP_DEFAULTS["tooltip_auto_theme"] else 0)
        self.apply_tooltip_settings(silent=True)
        self.var_locator_result.set(self.tr("msg_tooltip_reset"))

    def apply_tooltip_settings(self, silent=False):
        def _to_int(v, default, min_value=1, max_value=10000):
            try:
                out = int(str(v).strip())
                if out < min_value:
                    return min_value
                if out > max_value:
                    return max_value
                return out
            except Exception:
                return default
        if not self.validate_tooltip_inputs(silent=silent):
            return

        self.tooltip_delay_ms = _to_int(
            self.var_tooltip_delay_ms.get(), self.tooltip_delay_ms, min_value=50, max_value=5000
        )
        self.var_tooltip_delay_ms.set(str(self.tooltip_delay_ms))
        self.tooltip_font_size = _to_int(
            self.var_tooltip_font_size.get(), self.tooltip_font_size, min_value=8, max_value=20
        )
        self.var_tooltip_font_size.set(str(self.tooltip_font_size))
        self.tooltip_auto_theme = bool(self.var_tooltip_auto_theme.get())
        self.tooltip_bg = self.var_tooltip_bg.get().strip()
        self.tooltip_fg = self.var_tooltip_fg.get().strip()

        for tip in getattr(self, "_tooltips", []):
            tip.delay_ms = self.tooltip_delay_ms
            tip.bg = self.tooltip_bg
            tip.fg = self.tooltip_fg
            tip.font_family = self.tooltip_font_family
            tip.font_size = self.tooltip_font_size

        if not silent:
            self.var_locator_result.set(self.tr("msg_tooltip_applied"))

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

        ui_cfg = cfg.get("ui", {}) if isinstance(cfg.get("ui", {}), dict) else {}
        self.tooltip_delay_ms = int(ui_cfg.get("tooltip_delay_ms", 300) or 300)
        self.tooltip_bg = str(ui_cfg.get("tooltip_bg", "#FFF7D6") or "#FFF7D6")
        self.tooltip_fg = str(ui_cfg.get("tooltip_fg", "#202124") or "#202124")
        self.tooltip_font_family = str(
            ui_cfg.get("tooltip_font_family", "System") or "System"
        )
        self.tooltip_font_size = int(ui_cfg.get("tooltip_font_size", 9) or 9)
        self.tooltip_auto_theme = bool(ui_cfg.get("tooltip_auto_theme", True))
        if hasattr(self, "var_tooltip_auto_theme"):
            self.var_tooltip_auto_theme.set(1 if self.tooltip_auto_theme else 0)
        if hasattr(self, "var_tooltip_delay_ms"):
            self.var_tooltip_delay_ms.set(str(self.tooltip_delay_ms))
        if hasattr(self, "var_tooltip_font_size"):
            self.var_tooltip_font_size.set(str(self.tooltip_font_size))
        if hasattr(self, "var_tooltip_bg"):
            self.var_tooltip_bg.set(self.tooltip_bg)
        if hasattr(self, "var_tooltip_fg"):
            self.var_tooltip_fg.set(self.tooltip_fg)
        self.apply_tooltip_settings(silent=True)
        self.validate_tooltip_inputs(silent=True)

        # 运行参数
        self.var_source_folder.set(cfg.get("source_folder", ""))
        self.var_target_folder.set(cfg.get("target_folder", ""))
        self.var_enable_sandbox.set(1 if cfg.get("enable_sandbox", True) else 0)
        self.var_temp_sandbox_root.set(cfg.get("temp_sandbox_root", ""))

        self.var_enable_merge.set(1 if cfg.get("enable_merge", True) else 0)
        self.var_merge_mode.set(cfg.get("merge_mode", MERGE_MODE_CATEGORY))
        self.var_merge_source.set(cfg.get("merge_source", "source"))
        self.var_enable_merge_index.set(1 if cfg.get("enable_merge_index", False) else 0)
        self.var_enable_merge_excel.set(1 if cfg.get("enable_merge_excel", False) else 0)

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
        self.refresh_locator_maps()
        self._normalize_numeric_var(self.var_timeout_seconds)
        self._normalize_numeric_var(self.var_pdf_wait_seconds)
        self._normalize_numeric_var(self.var_ppt_timeout_seconds)
        self._normalize_numeric_var(self.var_ppt_pdf_wait_seconds)
        self._normalize_numeric_var(self.var_max_merge_size_mb)
        self._normalize_short_id_var(self.var_locator_short_id)
        self._normalize_date_var(self.var_date_str)
        self.validate_runtime_inputs(silent=True, scope="all")

    def _save_settings_to_file(self, show_msg=True):
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
        cfg["enable_merge_index"] = bool(self.var_enable_merge_index.get())
        cfg["enable_merge_excel"] = bool(self.var_enable_merge_excel.get())

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
        cfg["enable_merge_map"] = cfg.get("enable_merge_map", True)
        cfg["bookmark_with_short_id"] = cfg.get("bookmark_with_short_id", True)
        cfg["everything"] = cfg.get(
            "everything",
            {
                "enabled": True,
                "es_path": "",
                "prefer_path_exact": True,
                "timeout_ms": 1500,
            },
        )
        cfg["listary"] = cfg.get(
            "listary",
            {
                "enabled": True,
                "copy_query_on_locate": True,
            },
        )
        cfg["privacy"] = cfg.get("privacy", {"mask_md5_in_logs": True})
        self.apply_tooltip_settings(silent=True)
        cfg["ui"] = {
            "tooltip_delay_ms": self.tooltip_delay_ms,
            "tooltip_bg": self.tooltip_bg,
            "tooltip_fg": self.tooltip_fg,
            "tooltip_font_family": self.tooltip_font_family,
            "tooltip_font_size": self.tooltip_font_size,
            "tooltip_auto_theme": self.tooltip_auto_theme,
        }

        try:
            with open(self.config_path, "w", encoding="utf-8") as f:
                json.dump(cfg, f, indent=4, ensure_ascii=False)
            if show_msg:
                messagebox.showinfo(self.tr("btn_save_cfg"), self.tr("msg_save_ok"))
        except Exception as e:
            if show_msg:
                messagebox.showerror(self.tr("btn_save_cfg"), self.tr("msg_save_fail").format(e))

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
            if hasattr(self, "btn_start"): self.btn_start.configure(state="disabled")
            if hasattr(self, "btn_stop"): self.btn_stop.configure(state="normal")
            if hasattr(self, "btn_save_cfg"): self.btn_save_cfg.configure(state="disabled")
            if hasattr(self, "btn_save_cfg_tab"): self.btn_save_cfg_tab.configure(state="disabled")
            self.progress["mode"] = "determinate"
            self.progress["value"] = 0
            self.var_status.set(self.tr("status_init") if hasattr(self, "tr") else "Initializing...")
        else:
            if hasattr(self, "btn_start"): self.btn_start.configure(state="normal")
            if hasattr(self, "btn_stop"): self.btn_stop.configure(state="disabled")
            if hasattr(self, "btn_save_cfg"): self.btn_save_cfg.configure(state="normal")
            if hasattr(self, "btn_save_cfg_tab"): self.btn_save_cfg_tab.configure(state="normal")
            self.progress.stop()
            self.progress["value"] = 100
            self.var_status.set(self.tr("status_ready") if hasattr(self, "tr") else "Ready")

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

    def _on_click_start(self):
        if self.worker_thread and self.worker_thread.is_alive():
            messagebox.showinfo("提示", "已有任务在运行，请先停止或等待完成。")
            return
        if not self.validate_runtime_inputs(silent=False, scope="all"):
            messagebox.showerror(self.tr("btn_start"), self.tr("msg_validation_fix_before_run"))
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
                cfg["enable_merge_index"] = bool(self.var_enable_merge_index.get())
                cfg["enable_merge_excel"] = bool(self.var_enable_merge_excel.get())
                cfg["kill_process_mode"] = self.var_kill_mode.get()
                cfg["default_engine"] = self.var_engine.get()

                converter.run_mode = self.var_run_mode.get()
                converter.collect_mode = self.var_collect_mode.get()
                converter.content_strategy = self.var_strategy.get()
                converter.merge_mode = self.var_merge_mode.get()
                converter.engine_type = self.var_engine.get()
                converter.enable_merge_index = bool(self.var_enable_merge_index.get())
                converter.enable_merge_excel = bool(self.var_enable_merge_excel.get())

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
                self.after(0, self.refresh_locator_maps)

        self.worker_thread = threading.Thread(target=worker, daemon=True)
        self.worker_thread.start()
        self._set_running_ui_state(True)

    def _on_click_stop(self):
        if self.current_converter is None:
            return
        if messagebox.askyesno("停止任务", "确定要请求停止当前任务吗？"):
            self.stop_requested = True
            self.current_converter.is_running = False
            print("[GUI] 已请求停止任务，正在等待当前步骤结束...")
            self.var_status.set("已请求停止，等待当前文件处理结束...")

    # ===================== 程序入口 =====================


if __name__ == "__main__":
    try:
        app = OfficeGUI()
        app.mainloop()
    except Exception:
        traceback.print_exc()
