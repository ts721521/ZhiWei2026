# -*- coding: utf-8 -*-
"""
office_gui.py - Office 文档批量转换 & 整理工具 GUI 版

说明：
- 依赖 office_converter.py 中的 OfficeConverter（已更新到 v5.19.1）
- GUI 中含：
    * "运行参数"页：选择源/目标目录、运行模式、内容策略、合并模式、沙箱等
    * "配置管理"页：直接编辑 config.json 的部分配置（日志目录、排除目录、关键字、超时参数等）
- "保存配置"按钮：写入 config.json
- "开始运行"按钮：用当前界面参数启动转换/整理（不会自动改 config.json）
- "停止"按钮：设置 converter.is_running=False，优雅停止
"""

import os
import sys

# Remove third-party site-packages (e.g. pipecad) that pollute PIL/Pillow imports
sys.path[:] = [p for p in sys.path if "pipecad" not in p.lower()]

import subprocess
import glob

import json
import threading
import queue
import re
from types import SimpleNamespace
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, colorchooser, simpledialog
from tkinter.constants import *  # Explicitly import standard constants

# Avoid UnicodeEncodeError on Windows consoles with legacy code pages.
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(errors="replace")
    except Exception:
        pass

try:
    import ttkbootstrap as tb
    from ttkbootstrap.constants import *
    from ttkbootstrap.widgets.scrolled import ScrolledText

    HAS_TTKBOOTSTRAP = True

    def _patch_ttkbootstrap_widget_bootstyle(widget_name):
        widget_cls = getattr(tb, widget_name, None)
        if widget_cls is None:
            return

        class _SafeWidget(widget_cls):
            def __init__(self, *args, **kwargs):
                try:
                    super().__init__(*args, **kwargs)
                except tk.TclError as exc:
                    msg = str(exc)
                    if "Layout" not in msg or "not found" not in msg:
                        raise
                    fallback_kwargs = dict(kwargs)
                    fallback_kwargs.pop("bootstyle", None)
                    super().__init__(*args, **fallback_kwargs)

            def configure(self, cnf=None, **kwargs):
                try:
                    return super().configure(cnf, **kwargs)
                except tk.TclError as exc:
                    msg = str(exc)
                    if "Layout" not in msg or "not found" not in msg:
                        raise
                    kwargs = dict(kwargs)
                    kwargs.pop("bootstyle", None)
                    if isinstance(cnf, dict):
                        cnf = dict(cnf)
                        cnf.pop("bootstyle", None)
                    return super().configure(cnf, **kwargs)

            config = configure

        setattr(tb, widget_name, _SafeWidget)

    for _widget_name in (
        "Button",
        "Checkbutton",
        "Radiobutton",
        "Progressbar",
        "Combobox",
        "Labelframe",
        "Scrollbar",
    ):
        _patch_ttkbootstrap_widget_bootstyle(_widget_name)
except ModuleNotFoundError:
    from tkinter.scrolledtext import ScrolledText as _TkScrolledText

    HAS_TTKBOOTSTRAP = False

    class _FallbackStyle:
        def __init__(self, theme_name="default"):
            self._theme_name = theme_name or "default"

        def theme_use(self, theme_name=None):
            if theme_name is None:
                return self._theme_name
            self._theme_name = theme_name
            return self._theme_name

    class _BootstyleMixin:
        def __init__(self, *args, **kwargs):
            kwargs.pop("bootstyle", None)
            super().__init__(*args, **kwargs)

        def configure(self, cnf=None, **kwargs):
            kwargs.pop("bootstyle", None)
            if isinstance(cnf, dict) and "bootstyle" in cnf:
                cnf = dict(cnf)
                cnf.pop("bootstyle", None)
            return super().configure(cnf, **kwargs)

        config = configure

    class _FallbackWindow(tk.Tk):
        def __init__(self, *args, themename=None, **kwargs):
            super().__init__(*args, **kwargs)
            self.style = _FallbackStyle(themename or "default")

    class _FallbackFrame(_BootstyleMixin, ttk.Frame):
        pass

    class _FallbackLabel(_BootstyleMixin, ttk.Label):
        pass

    class _FallbackButton(_BootstyleMixin, ttk.Button):
        pass

    class _FallbackCheckbutton(_BootstyleMixin, ttk.Checkbutton):
        pass

    class _FallbackProgressbar(_BootstyleMixin, ttk.Progressbar):
        pass

    class _FallbackNotebook(_BootstyleMixin, ttk.Notebook):
        pass

    class _FallbackScrollbar(_BootstyleMixin, ttk.Scrollbar):
        pass

    class _FallbackEntry(_BootstyleMixin, ttk.Entry):
        pass

    class _FallbackLabelframe(_BootstyleMixin, ttk.Labelframe):
        pass

    class _FallbackRadiobutton(_BootstyleMixin, ttk.Radiobutton):
        pass

    class _FallbackSeparator(_BootstyleMixin, ttk.Separator):
        pass

    class _FallbackCombobox(_BootstyleMixin, ttk.Combobox):
        pass

    class _FallbackDateEntry(_FallbackEntry):
        def __init__(self, *args, **kwargs):
            kwargs.pop("dateformat", None)
            kwargs.pop("firstweekday", None)
            kwargs.pop("startdate", None)
            super().__init__(*args, **kwargs)
            self.entry = self

    class ScrolledText(_TkScrolledText):
        def __init__(self, *args, **kwargs):
            kwargs.pop("bootstyle", None)
            super().__init__(*args, **kwargs)
            self.text = self

        def configure(self, cnf=None, **kwargs):
            kwargs.pop("bootstyle", None)
            if isinstance(cnf, dict) and "bootstyle" in cnf:
                cnf = dict(cnf)
                cnf.pop("bootstyle", None)
            return super().configure(cnf, **kwargs)

        config = configure

    tb = SimpleNamespace(
        Window=_FallbackWindow,
        Frame=_FallbackFrame,
        Label=_FallbackLabel,
        Button=_FallbackButton,
        Checkbutton=_FallbackCheckbutton,
        Progressbar=_FallbackProgressbar,
        Notebook=_FallbackNotebook,
        Scrollbar=_FallbackScrollbar,
        Entry=_FallbackEntry,
        Labelframe=_FallbackLabelframe,
        Radiobutton=_FallbackRadiobutton,
        Separator=_FallbackSeparator,
        Combobox=_FallbackCombobox,
        DateEntry=_FallbackDateEntry,
    )

import tempfile
import traceback
from datetime import datetime
from ui_translations import TRANSLATIONS

from office_converter import (
    OfficeConverter,
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
    ENGINE_WPS,
    ENGINE_MS,
    KILL_MODE_AUTO,
    KILL_MODE_KEEP,
    __version__,
    get_app_path,
    create_default_config,
)
from task_manager import TaskStore
from gui_task_workflow_mixin import TaskWorkflowMixin
from gui_run_tab_mixin import RunTabUIMixin
from gui_run_mode_state_mixin import RunModeStateMixin
from gui_source_folder_mixin import SourceFolderMixin
from gui_config_tab_mixin import ConfigTabUIMixin
from gui_config_dirty_mixin import ConfigDirtyStateMixin
from gui_config_io_mixin import ConfigIOMixin
from gui_config_compose_mixin import ConfigComposeMixin
from gui_config_save_mixin import ConfigSaveMixin
from gui_tooltip_settings_mixin import TooltipSettingsMixin
from gui_config_logic_mixin import ConfigLogicMixin
from gui_runtime_status_mixin import RuntimeStatusMixin
from gui_execution_mixin import ExecutionFlowMixin
from gui_profile_mixin import ProfileManagementMixin
from gui_ui_shell_mixin import UIShellMixin
from gui_locator_mixin import LocatorMixin
from gui_gdrive_mixin import GDriveMixin
from gui_misc_ui_mixin import MiscUIMixin
from gui_tooltip_mixin import TooltipMixin

LOG_QUEUE = queue.Queue()


class TkLogHandler:
    """Simple handler that forwards stdout/stderr lines into the GUI log queue."""

    def __init__(self, log_queue=None):
        self.log_queue = log_queue if log_queue is not None else LOG_QUEUE

    def write(self, msg: str):
        msg = msg.rstrip("\n")
        if msg:
            self.log_queue.put(msg)

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
        return (
            self.widget.winfo_rootx() + 14,
            self.widget.winfo_rooty() + self.widget.winfo_height() + 8,
        )

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


# ========= GUI 涓撶敤 Converter 瀛愮被锛氬睆钄?CLI 杈撳叆 =========


class GUIOfficeConverter(OfficeConverter):
    """GUI-only converter: disable interactive CLI prompts."""

    def print_welcome(self):
        # GUI 鐜涓嬩笉闇€瑕佸湪鎺у埗鍙版墦鍗版杩庣晫闈紙鏃ュ織閲屼細鏈夊ご閮級
        pass

    def confirm_config_in_terminal(self):
        # 涓嶅湪 GUI 鐜涓嬪啀娆¤闂簮/鐩爣鐩綍
        pass

    def ask_for_subfolder(self):
        # GUI 涓笉鍋?瀛愮洰褰?璇㈤棶锛涘闇€瀛愮洰褰曞彲鐩存帴鍦ㄧ洰鏍囪矾寰勯噷浣撶幇
        pass

    def select_run_mode(self):
        # 杩愯妯″紡鐢?GUI 璁剧疆
        pass

    def select_collect_mode(self):
        # 姊崇悊瀛愭ā寮忕敱 GUI 璁剧疆
        pass

    def select_merge_mode(self):
        # 鍚堝苟妯″紡鐢?GUI 璁剧疆锛堥厤缃垨鐣岄潰 Radio锛?
        pass

    def select_content_strategy(self):
        # 鍐呭绛栫暐鐢?GUI 璁剧疆
        pass

    def select_engine_mode(self):
        # 寮曟搸绫诲瀷鐢?GUI 璁剧疆
        pass

    def print_runtime_summary(self):
        # 绠€鍖栵細GUI 鐜涓嬩笉鎵撳嵃 CLI 椋庢牸鎬昏锛屾棩蹇楃敱 GUI 鎹曡幏 print 鍗冲彲
        pass


# ========= 涓荤獥鍙?========


class OfficeGUI(TaskWorkflowMixin, RunTabUIMixin, RunModeStateMixin, SourceFolderMixin, ConfigTabUIMixin, ConfigDirtyStateMixin, ConfigIOMixin, ConfigComposeMixin, ConfigSaveMixin, TooltipSettingsMixin, ConfigLogicMixin, RuntimeStatusMixin, ExecutionFlowMixin, ProfileManagementMixin, UIShellMixin, LocatorMixin, GDriveMixin, MiscUIMixin, TooltipMixin, tb.Window):
    TOOLTIP_DEFAULTS = {
        "tooltip_delay_ms": 300,
        "tooltip_bg": "#FFF7D6",
        "tooltip_fg": "#202124",
        "tooltip_font_size": 9,
        "tooltip_auto_theme": True,
    }

    def __init__(self, config_path=None):
        super().__init__(themename="cosmo")
        if not HAS_TTKBOOTSTRAP:
            print("[GUI] ttkbootstrap not found, using tkinter compatibility mode.")
        self.current_lang = "zh"  # 仅中文界面，保留变量供 tr 等兼容
        self.title(f"{self.tr('title')} v{__version__}")

        # 绐楀彛榛樿灏哄涓庡睆骞曡嚜閫傚簲
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        if screen_h >= 1080:
            default_w, default_h = 1360, 920
        else:
            default_w, default_h = 1280, 860
        self.geometry(f"{default_w}x{default_h}")
        self.minsize(1000, 700)
        self.update_idletasks()

        # Ensure style is available for theme toggling
        # tb.Window automatically creates a style object, accessible via self.style if needed,
        # but self.style is not a standard attribute of tk.Tk.
        # ttkbootstrap.Window has a 'style' attribute.

        self.script_dir = get_app_path()
        self.config_path = config_path or os.path.join(self.script_dir, "config.json")
        # Initialize app mode early so callbacks/tests can read it before delayed UI build.
        self.var_app_mode = tk.StringVar(value="classic")
        try:
            if os.path.isfile(self.config_path):
                with open(self.config_path, "r", encoding="utf-8") as f:
                    raw_cfg = json.load(f)
                mode = str(raw_cfg.get("app_mode", "classic")).strip().lower()
                if mode in {"classic", "task"}:
                    self.var_app_mode.set(mode)
        except Exception:
            pass
        self.var_profile_active_path = tk.StringVar(value=self.config_path)
        self.profile_manager_win = None
        self.profile_tree = None
        self._profile_tree_rows = {}
        self.save_profile_dialog = None
        self.load_profile_dialog = None
        self.load_profile_tree = None
        self._load_profile_tree_rows = {}

        # 褰撳墠浠诲姟绾跨▼ & 杞崲鍣ㄥ疄渚?
        self.worker_thread = None
        self.current_converter = None
        self._converter_cls = GUIOfficeConverter
        self.current_task_id = None
        self.current_run_context = "manual"
        self.stop_requested = False
        self._hover_tip_cls = HoverTip
        self._tooltips = []
        self._tooltip_widget_ids = set()
        self._normalizing_inputs = False
        self._suspend_cfg_dirty = False
        self.cfg_dirty = False
        self._baseline_config_snapshot = {}
        self._cfg_tab_meta = []
        self._last_section_dirty = {}
        self._ui_running = False
        self.source_folders_list = []  # List to store multiple source folders
        self.tooltip_delay_ms = self.TOOLTIP_DEFAULTS["tooltip_delay_ms"]
        self.tooltip_bg = self.TOOLTIP_DEFAULTS["tooltip_bg"]
        self.tooltip_fg = self.TOOLTIP_DEFAULTS["tooltip_fg"]
        self.tooltip_font_family = "System"
        self.tooltip_font_size = self.TOOLTIP_DEFAULTS["tooltip_font_size"]
        self.tooltip_auto_theme = self.TOOLTIP_DEFAULTS["tooltip_auto_theme"]
        self.task_store = TaskStore(self.script_dir)
        self.log_queue = queue.Queue()

        # 鎶?stdout/stderr 閲嶅畾鍚戝埌 GUI 鏃ュ織绐楀彛
        # sys.stdout = TkLogHandler(self.log_queue)
        # sys.stderr = TkLogHandler(self.log_queue)

        # 先出窗口：显示「正在加载界面…」，再在 after 回调里构建完整 UI，避免主线程长时间阻塞导致窗口不显示
        self._loading_frame = tk.Frame(self)
        tk.Label(
            self._loading_frame,
            text="正在加载界面…",
            font=("Microsoft YaHei", 14),
        ).pack(expand=True)
        self._loading_frame.pack(fill=tk.BOTH, expand=True)
        self.update_idletasks()
        self.lift()
        self.after_idle(self._finish_init)

    def tr(self, key):
        """取中文文案（仅支持中文界面）"""
        lang_map = TRANSLATIONS.get("zh", {})
        if key in lang_map:
            return lang_map[key]
        return TRANSLATIONS.get("en", {}).get(key, key)


if __name__ == "__main__":
    try:
        app = OfficeGUI()
        app.update_idletasks()
        app.lift()
        app.attributes("-topmost", True)
        app.after(150, lambda: app.attributes("-topmost", False))
        app.mainloop()
    except Exception as e:
        traceback.print_exc()
        try:
            messagebox.showerror(
                "启动失败",
                "程序启动异常，请关闭其他知喂/OfficeGUI 进程后重试。\n\n错误信息：\n"
                + str(e),
            )
        except Exception:
            pass
        raise

