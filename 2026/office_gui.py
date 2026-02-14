# -*- coding: utf-8 -*-
"""
office_gui.py - Office 鏂囨。鎵归噺杞崲 & 姊崇悊宸ュ叿 GUI 鐗?

璇存槑锛?
- 渚濊禆 office_converter.py 涓殑 OfficeConverter锛堝凡鏇存柊鍒?v5.17.0锛?
- GUI 涓惈锛?
    * "杩愯鍙傛暟"椤碉細閫夋嫨婧?鐩爣鐩綍銆佽繍琛屾ā寮忋€佸唴瀹圭瓥鐣ャ€佸悎骞舵ā寮忋€佹矙绠辩瓑
    * "閰嶇疆绠＄悊"椤碉細鐩存帴缂栬緫 config.json 鐨勯儴鍒嗛厤缃紙鏃ュ織鐩綍銆佹帓闄ょ洰褰曘€佸叧閿瓧銆佽秴鏃跺弬鏁扮瓑锛?
- "淇濆瓨閰嶇疆"鎸夐挳锛氬啓鍏?config.json
- "寮€濮嬭繍琛?鎸夐挳锛氱敤褰撳墠鐣岄潰鍙傛暟鍚姩杞崲/姊崇悊锛堜笉浼氳嚜鍔ㄦ敼 config.json锛?
- "鍋滄"鎸夐挳锛氳缃?converter.is_running=False锛屼紭闆呭仠姝?
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
from locate_source import locate_by_page, locate_by_short_id
from search_adapter import EverythingAdapter, build_listary_query

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
from task_manager import (
    TaskStore,
    build_task_runtime_config,
    create_checkpoint,
    mark_checkpoint_file_done,
    remove_task_registry_if_exists,
)

LOG_QUEUE = queue.Queue()


class TkLogHandler:
    """Simple handler that forwards stdout/stderr lines into the GUI log queue."""

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
        
        # Ensure style is available for theme toggling
        # tb.Window automatically creates a style object, accessible via self.style if needed, 
        # but self.style is not a standard attribute of tk.Tk. 
        # ttkbootstrap.Window has a 'style' attribute.

        self.script_dir = get_app_path()
        self.config_path = config_path or os.path.join(self.script_dir, "config.json")
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
        self.current_task_id = None
        self.current_run_context = "manual"
        self.stop_requested = False
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

        # 鎶?stdout/stderr 閲嶅畾鍚戝埌 GUI 鏃ュ織绐楀彛
        # sys.stdout = TkLogHandler()
        # sys.stderr = TkLogHandler()

        if not os.path.exists(self.config_path):
            success = create_default_config(self.config_path)
            if success:
                info_title = "提示"
                messagebox.showinfo(info_title, self.tr("msg_no_config"))

        self._build_ui()
        self._load_config_to_ui()
        self.locator_short_id_index = {}

        # 璁颁綇绐楀彛鍏抽棴鍓嶇殑灏哄涓庝綅缃?
        try:
            self.protocol("WM_DELETE_WINDOW", self._on_close_main_window)
        except Exception:
            pass

        # 瀹氭椂鍒锋柊鏃ュ織
        self.after(200, self._poll_log_queue)


    # ===================== UI 閺嬪嫬缂?=====================

    # ===================== UI 閺嬪嫬缂?(Modern Layout) =====================

    def tr(self, key):
        """取中文文案（仅支持中文界面）"""
        lang_map = TRANSLATIONS.get("zh", {})
        if key in lang_map:
            return lang_map[key]
        return TRANSLATIONS.get("en", {}).get(key, key)

    def _attach_tooltip(self, widget, key):
        if id(widget) in self._tooltip_widget_ids:
            return
        setattr(widget, "_tooltip_key", key)
        setattr(widget, "_tooltip_disabled_reason", None)
        self._tooltip_widget_ids.add(id(widget))

        def _text_func(w=widget):
            reason = getattr(w, "_tooltip_disabled_reason", None)
            if reason:
                return reason
            return self.tr(getattr(w, "_tooltip_key", ""))

        self._tooltips.append(
            HoverTip(
                widget,
                _text_func,
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
        setattr(widget, "_tooltip_text", text)
        setattr(widget, "_tooltip_disabled_reason", None)
        self._tooltip_widget_ids.add(id(widget))

        def _text_func(w=widget):
            reason = getattr(w, "_tooltip_disabled_reason", None)
            if reason:
                return reason
            return getattr(w, "_tooltip_text", "")

        self._tooltips.append(
            HoverTip(
                widget,
                _text_func,
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
            self.tr("mode_mshelp"): "tip_mode_mshelp",
            self.tr("lbl_sandbox"): "tip_toggle_sandbox",
            self.tr("chk_corpus_manifest"): "tip_toggle_corpus_manifest",
            self.tr("chk_export_markdown"): "tip_toggle_export_markdown",
            self.tr("chk_markdown_strip_header_footer"): "tip_toggle_markdown_strip_header_footer",
            self.tr("chk_markdown_structured_headings"): "tip_toggle_markdown_structured_headings",
            self.tr("chk_markdown_quality_report"): "tip_toggle_markdown_quality_report",
            self.tr("chk_export_records_json"): "tip_toggle_export_records_json",
            self.tr("chk_chromadb_export"): "tip_toggle_chromadb_export",
            self.tr("chk_incremental_mode"): "tip_toggle_incremental_mode",
            self.tr("chk_incremental_verify_hash"): "tip_toggle_incremental_verify_hash",
            self.tr("chk_incremental_reprocess_renamed"): "tip_toggle_incremental_reprocess_renamed",
            self.tr("chk_source_priority_skip_pdf"): "tip_toggle_source_priority_skip_pdf",
            self.tr("chk_global_md5_dedup"): "tip_toggle_global_md5_dedup",
            self.tr("chk_enable_update_package"): "tip_toggle_enable_update_package",
            self.tr("chk_enable_merge"): "tip_toggle_enable_merge",
            self.tr("lbl_filter_date"): "tip_toggle_date_filter",
            self.tr("chk_merge_index"): "tip_toggle_merge_index",
            self.tr("chk_merge_excel"): "tip_toggle_merge_excel",
            self.tr("chk_output_pdf"): "tip_toggle_output_pdf",
            self.tr("chk_output_md"): "tip_toggle_output_md",
            self.tr("chk_output_merged"): "tip_toggle_output_merged",
            self.tr("chk_output_independent"): "tip_toggle_output_independent",
            self.tr("rad_merge_convert_merge_only"): "tip_option_merge_submode_merge_only",
            self.tr("rad_merge_convert_pdf_to_md"): "tip_option_merge_submode_pdf_to_md",
            self.tr("rad_category"): "tip_option_merge_mode_category",
            self.tr("rad_all_in_one"): "tip_option_merge_mode_all_in_one",
            self.tr("rad_src_dir"): "tip_option_merge_source_source",
            self.tr("rad_tgt_dir"): "tip_option_merge_source_target",
            self.tr("chk_tooltip_auto_theme"): "tip_toggle_tooltip_auto_theme",
            self.tr("chk_show_tooltip_advanced"): "tip_toggle_show_tooltip_advanced",
            self.tr("chk_confirm_revert_dirty"): "tip_toggle_confirm_revert_dirty",
            self.tr("btn_task_create"): "tip_task_create",
            self.tr("btn_task_edit"): "tip_task_edit",
            self.tr("btn_task_delete"): "tip_task_delete",
            self.tr("btn_task_refresh"): "tip_task_refresh",
            self.tr("btn_task_load_to_ui"): "tip_task_load_to_ui",
            self.tr("btn_task_run"): "tip_task_run",
            self.tr("btn_task_resume"): "tip_task_resume",
            self.tr("btn_task_stop"): "tip_task_stop",
            self.tr("chk_task_force_full_rebuild"): "tip_task_force_full_rebuild",
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

    def _guess_widget_label_text(self, widget):
        parent = getattr(widget, "master", None)
        if parent is None:
            return ""
        try:
            siblings = parent.winfo_children()
            idx = siblings.index(widget)
        except Exception:
            return ""
        for i in range(idx - 1, -1, -1):
            sib = siblings[i]
            try:
                keys = set(sib.keys())
            except Exception:
                continue
            if "text" not in keys:
                continue
            try:
                text = str(sib.cget("text")).strip()
            except Exception:
                text = ""
            if text and text not in {"...", "?", ">", "+", "-", "C"}:
                return text
        return ""

    def _is_input_like_widget(self, widget):
        cls = str(widget.winfo_class() or "").lower()
        if any(
            token in cls
            for token in (
                "entry",
                "combobox",
                "spinbox",
                "listbox",
                "text",
                "dateentry",
            )
        ):
            return True
        try:
            keys = set(widget.keys())
        except Exception:
            return False
        is_option = "variable" in keys and ("value" in keys or "onvalue" in keys)
        if is_option:
            return False
        return "textvariable" in keys

    def _auto_attach_input_tooltips(self, root, fallback_section_tip_key=None):
        # Prefer specific input tips by nearby label text; fallback to a generic
        # explanation so every config input remains discoverable.
        input_tip_key_by_label_text = {
            self.tr("lbl_source"): "tip_input_source_folder",
            self.tr("lbl_target"): "tip_input_target_folder",
            self.tr("lbl_config"): "tip_input_config_path",
            self.tr("lbl_strategy"): "tip_input_strategy",
            self.tr("lbl_filter_date"): "tip_input_date",
            self.tr("lbl_log_folder"): "tip_input_log_folder",
            self.tr("lbl_sandbox_min_free_gb"): "tip_input_sandbox_min_free_gb",
            self.tr("lbl_gen_timeout"): "tip_input_timeout_seconds",
            self.tr("lbl_pdf_wait"): "tip_input_pdf_wait_seconds",
            self.tr("lbl_ppt_timeout"): "tip_input_ppt_timeout_seconds",
            self.tr("lbl_ppt_wait"): "tip_input_ppt_pdf_wait_seconds",
            self.tr("lbl_office_restart_every"): "tip_input_office_restart_every_n_files",
            self.tr("lbl_max_mb"): "tip_input_max_merge_size_mb",
            self.tr("lbl_mshelp_folder_name"): "tip_input_mshelp_folder_name",
            self.tr("lbl_locator_merged"): "tip_input_locator_merged",
            self.tr("lbl_locator_page"): "tip_input_locator_page",
            self.tr("lbl_locator_id"): "tip_input_locator_short_id",
            self.tr("lbl_excluded"): "tip_input_excluded_folders",
            self.tr("lbl_keywords"): "tip_input_price_keywords",
            self.tr("lbl_tooltip_delay"): "tip_input_tooltip_delay_ms",
            self.tr("lbl_tooltip_font_size"): "tip_input_tooltip_font_size",
            self.tr("lbl_tooltip_bg"): "tip_input_tooltip_bg",
            self.tr("lbl_tooltip_fg"): "tip_input_tooltip_fg",
        }
        special_by_id = {}
        for attr_name, tip_key in (
            ("lst_source_folders", "tip_input_source_folder"),
            ("tree_tasks", "tip_task_list"),
            ("entry_temp_sandbox_root", "tip_input_sandbox_root"),
            ("cb_sandbox_low_space_policy", "tip_input_sandbox_low_space_policy"),
            ("ent_log_folder", "tip_input_log_folder"),
            ("ent_mshelpviewer_folder_name", "tip_input_mshelp_folder_name"),
            ("ent_cfg_mshelpviewer_folder_name", "tip_input_mshelp_folder_name"),
            ("ent_timeout_seconds", "tip_input_timeout_seconds"),
            ("ent_pdf_wait_seconds", "tip_input_pdf_wait_seconds"),
            ("ent_ppt_timeout_seconds", "tip_input_ppt_timeout_seconds"),
            ("ent_ppt_pdf_wait_seconds", "tip_input_ppt_pdf_wait_seconds"),
            ("ent_office_restart_every_n_files", "tip_input_office_restart_every_n_files"),
            ("ent_max_merge_size_mb", "tip_input_max_merge_size_mb"),
            ("ent_tooltip_delay", "tip_input_tooltip_delay_ms"),
            ("ent_tooltip_font_size", "tip_input_tooltip_font_size"),
            ("ent_tooltip_bg", "tip_input_tooltip_bg"),
            ("ent_tooltip_fg", "tip_input_tooltip_fg"),
        ):
            w = getattr(self, attr_name, None)
            if w is not None:
                special_by_id[id(w)] = tip_key

        for child in root.winfo_children():
            self._auto_attach_input_tooltips(child, fallback_section_tip_key)
            if id(child) in self._tooltip_widget_ids:
                continue
            if not self._is_input_like_widget(child):
                continue
            direct_key = special_by_id.get(id(child))
            if direct_key:
                self._attach_tooltip(child, direct_key)
                continue
            label_text = self._guess_widget_label_text(child)
            mapped_key = input_tip_key_by_label_text.get(label_text, "")
            if mapped_key:
                self._attach_tooltip(child, mapped_key)
                continue
            base = self.tr("tip_auto_config_item").format(label_text or child.winfo_class())
            if fallback_section_tip_key:
                base = f"{base}\n{self.tr(fallback_section_tip_key)}"
            self._attach_tooltip_text(child, base)

    def _add_cfg_section_reset_action(self, parent, section_name):
        frame = tb.Frame(parent)
        frame.pack(fill=X, pady=(0, 5))
        btn_save = tb.Button(
            frame,
            text=self.tr("btn_save_cfg_section"),
            command=lambda s=section_name: self._save_specific_config_section(s),
            bootstyle="success-outline",
            width=14,
        )
        btn_save.pack(side=RIGHT, padx=(0, 6))
        self._attach_tooltip(btn_save, "tip_save_config_section")
        btn_reset = tb.Button(
            frame,
            text=self.tr("btn_reset_cfg_section"),
            command=lambda s=section_name: self._reset_config_section_defaults(s),
            bootstyle="warning-outline",
            width=14,
        )
        btn_reset.pack(side=RIGHT)
        self._attach_tooltip(btn_reset, "tip_reset_config_section")
        return frame, btn_save, btn_reset

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

    # ===================== UI 閺嬪嫬缂?=====================

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

        self.var_app_mode = tk.StringVar(value="classic")
        frm_app_mode = tb.Frame(ctrl_frame, bootstyle="light")
        frm_app_mode.pack(side=LEFT, padx=(0, 12))
        tb.Label(frm_app_mode, text=self.tr("app_mode_classic") + " / " + self.tr("app_mode_task") + ":", font=("System", 9)).pack(side=LEFT, padx=(0, 4))
        tb.Radiobutton(
            frm_app_mode,
            text=self.tr("app_mode_classic"),
            variable=self.var_app_mode,
            value="classic",
            bootstyle="toolbutton-outline",
        ).pack(side=LEFT, padx=2)
        tb.Radiobutton(
            frm_app_mode,
            text=self.tr("app_mode_task"),
            variable=self.var_app_mode,
            value="task",
            bootstyle="toolbutton-outline",
        ).pack(side=LEFT, padx=2)
        self._attach_tooltip(frm_app_mode, "tip_app_mode")

        self.var_theme = tk.StringVar(value="cosmo")

        def toggle_theme():
            t = self.var_theme.get()
            new_theme = "superhero" if t == "cosmo" else "cosmo"
            # 无 ttkbootstrap 时 FallbackStyle.theme_use 仅存名称不生效，提示用户安装
            if not HAS_TTKBOOTSTRAP and new_theme == "superhero":
                messagebox.showinfo("主题", "深色主题需要安装 ttkbootstrap：pip install ttkbootstrap")
                return
            try:
                self.style.theme_use(new_theme)
            except Exception:
                pass
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

        # 3. Main body (tabs + log)
        body_frame = tb.Frame(self, padding=10)
        body_frame.pack(fill=BOTH, expand=YES)

        self.config_container = tb.Frame(body_frame)
        self.config_container.pack(fill=BOTH, expand=YES)
        self.main_notebook = tb.Notebook(self.config_container)
        self.main_notebook.pack(fill=BOTH, expand=YES)

        # 鈹€鈹€ 鍗曞眰 tab 缁撴瀯锛氭寜鍔熻兘鍩熷垝鍒嗭紝鍘绘帀銆岃繍琛屽弬鏁?/ 閰嶇疆绠＄悊銆嶄簩鍒嗘硶 鈹€鈹€
        self.tab_run_shared = tb.Frame(self.main_notebook)
        self.tab_run_convert = tb.Frame(self.main_notebook)
        self.tab_run_merge = tb.Frame(self.main_notebook)
        self.tab_run_mshelp = tb.Frame(self.main_notebook)
        self.tab_run_locator = tb.Frame(self.main_notebook)
        self.tab_run_output = tb.Frame(self.main_notebook)
        self.tab_run_tasks = tb.Frame(self.main_notebook)
        self.tab_settings = tb.Frame(self.main_notebook)

        # 椤跺眰 7 涓姛鑳?tab锛?
        # 1) 妯″紡涓庤矾寰? 2) 杞崲閫夐」  3) 鍚堝苟/姊崇悊  4) MSHelp  5) 瀹氫綅  6) 鎴愭灉鏂囦欢  7) 楂樼骇璁剧疆
        self.main_notebook.add(self.tab_run_shared, text=self.tr("grp_shared_runtime"))
        self.main_notebook.add(self.tab_run_convert, text=self.tr("grp_convert_runtime"))
        self.main_notebook.add(
            self.tab_run_merge,
            text=f"{self.tr('grp_merge_runtime')} / {self.tr('grp_collect_runtime')}",
        )
        self.main_notebook.add(self.tab_run_mshelp, text=self.tr("grp_mshelp_runtime"))
        self.main_notebook.add(self.tab_run_locator, text=self.tr("grp_locator_tools"))
        self.main_notebook.add(self.tab_run_output, text=self.tr("grp_output_files"))
        self.main_notebook.add(self.tab_run_tasks, text=self.tr("grp_task_runtime"))
        self.main_notebook.add(self.tab_settings, text=self.tr("tab_config_center"))

        # 记录原始 tab 顺序与标签 key，用于隐藏后恢复时带正确 text
        self._all_tabs = [
            self.tab_run_shared,
            self.tab_run_convert,
            self.tab_run_merge,
            self.tab_run_mshelp,
            self.tab_run_locator,
            self.tab_run_output,
            self.tab_run_tasks,
            self.tab_settings,
        ]
        self._all_tab_text_keys = [
            "grp_shared_runtime",
            "grp_convert_runtime",
            ("grp_merge_runtime", "grp_collect_runtime"),
            "grp_mshelp_runtime",
            "grp_locator_tools",
            "grp_output_files",
            "grp_task_runtime",
            "tab_config_center",
        ]

        # 姣忎釜 tab 鐙珛鍙粴鍔?
        self._scroll_shared = self._create_scrollable_page(self.tab_run_shared)
        self._scroll_convert = self._create_scrollable_page(self.tab_run_convert)
        self._scroll_merge = self._create_scrollable_page(self.tab_run_merge)
        self._scroll_mshelp = self._create_scrollable_page(self.tab_run_mshelp)
        self._scroll_locator = self._create_scrollable_page(self.tab_run_locator)
        self._scroll_output = self._create_scrollable_page(self.tab_run_output)
        self._scroll_tasks = self._create_scrollable_page(self.tab_run_tasks)
        self._scroll_settings = self._create_scrollable_page(self.tab_settings)

        # 涓?_build_config_tab_content 璁剧疆鍒悕锛氶厤缃唴瀹圭洿鎺ヨ拷鍔犲埌瀵瑰簲鍔熻兘 tab
        self.tab_cfg_shared = self._scroll_shared
        self.tab_cfg_convert = self._scroll_convert
        self.tab_cfg_merge = self._scroll_merge
        self.tab_cfg_ai = self._scroll_mshelp
        self.tab_cfg_ui = self._scroll_settings
        self.tab_cfg_rules = self._scroll_settings

        self._build_run_tab_content()
        self._build_config_tab_content(self._scroll_settings)
        self.main_notebook.select(0)

        # 鏃ュ織闈㈡澘锛氬浐瀹氶珮搴︾殑搴曢儴闈㈡澘锛岄粯璁ら殣钘忥紝鐐规寜閽墠鏄剧ず
        self.log_pane = tb.Frame(body_frame)
        self.txt_log = ScrolledText(
            self.log_pane, height=6, font=("Consolas", 9), bootstyle="primary-round"
        )
        self.txt_log.pack(fill=BOTH, expand=YES)
        self.txt_log.text.tag_config("INFO", foreground="#007bff")
        self.txt_log.text.tag_config("SUCCESS", foreground="#28a745")
        self.txt_log.text.tag_config("WARNING", foreground="#ffc107")
        self.txt_log.text.tag_config("ERROR", foreground="#dc3545")
        self.txt_log.text.tag_config("DIM", foreground="#6c757d")
        self._log_visible = False
        # 涓?pack锛屽惎鍔ㄦ椂鏃ュ織瀹屽叏闅愯棌

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

        def _calc_wheel_step(event):
            # Windows uses delta in multiples of 120; macOS often uses small deltas.
            if hasattr(event, "delta") and event.delta:
                if sys.platform == "darwin":
                    return -1 if event.delta > 0 else 1
                return int(-event.delta / 120) or (-1 if event.delta > 0 else 1)
            if getattr(event, "num", None) == 4:
                return -1
            if getattr(event, "num", None) == 5:
                return 1
            return 0

        def _is_descendant(widget, ancestor):
            cur = widget
            while cur is not None:
                if cur is ancestor:
                    return True
                cur = getattr(cur, "master", None)
            return False

        def _widget_can_self_scroll(widget):
            # Give native scrollable controls (list/text/tree/canvas) first chance
            # to consume wheel events, so page-level canvas scrolling does not
            # hijack their behavior on Windows.
            try:
                if isinstance(widget, (tk.Listbox, tk.Text, tk.Canvas, ttk.Treeview)):
                    first, last = widget.yview()
                    first = float(first)
                    last = float(last)
                    return first > 0.0 or last < 1.0
            except Exception:
                return False
            return False

        def _on_mousewheel(event):
            # Keep all-page wheel binding stable, but only scroll when event comes
            # from this canvas subtree. This avoids bind_all/unbind_all conflicts.
            evt_widget = getattr(event, "widget", None)
            if not _is_descendant(evt_widget, canvas):
                return None
            if evt_widget is not None and _widget_can_self_scroll(evt_widget):
                return None
            step = _calc_wheel_step(event)
            if step:
                canvas.yview_scroll(step, "units")
                return "break"
            return None

        canvas.bind("<Configure>", on_canvas_configure)
        self.bind_all("<MouseWheel>", _on_mousewheel, add="+")
        self.bind_all("<Button-4>", _on_mousewheel, add="+")
        self.bind_all("<Button-5>", _on_mousewheel, add="+")
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
            btn_open = tb.Button(f_in, text=">", command=cmd_open, bootstyle="link", width=2)
            btn_open.pack(side=LEFT)
            self._attach_tooltip(btn_open, "tip_open_folder")

    def _add_section_help(self, parent, tip_key):
        row = tb.Frame(parent)
        row.pack(fill=X, pady=(0, 2))
        spacer = tb.Label(row, text="")
        spacer.pack(side=LEFT, fill=X, expand=YES)
        btn = tb.Button(row, text="?", width=2, bootstyle="info-outline")
        btn.pack(side=RIGHT)
        self._attach_tooltip(btn, tip_key)

    def _build_run_tab_content(self):
        parent = self._scroll_shared
        # Section 1: run mode
        lf_mode = tb.Labelframe(parent, text=self.tr("sec_mode"), padding=6)
        lf_mode.pack(fill=X, pady=3)
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
        tb.Radiobutton(
            grid_frame, text=self.tr("mode_mshelp"), variable=self.var_run_mode, value=MODE_MSHELP_ONLY,
            command=self._on_run_mode_change, bootstyle="toolbutton-outline"
        ).grid(row=2, column=0, columnspan=2, sticky="ew", padx=2, pady=2)
        grid_frame.columnconfigure(0, weight=1)
        grid_frame.columnconfigure(1, weight=1)

        # Global output controls (prominent and mode-agnostic)
        lf_output = tb.Labelframe(parent, text=self.tr("grp_output_controls"), padding=6)
        lf_output.pack(fill=X, pady=3)
        self.var_output_enable_pdf = tk.IntVar(value=1)
        self.var_output_enable_md = tk.IntVar(value=1)
        self.var_output_enable_merged = tk.IntVar(value=1)
        self.var_output_enable_independent = tk.IntVar(value=0)
        tb.Checkbutton(
            lf_output,
            text=self.tr("chk_output_pdf"),
            variable=self.var_output_enable_pdf,
        ).pack(anchor="w")
        tb.Checkbutton(
            lf_output,
            text=self.tr("chk_output_md"),
            variable=self.var_output_enable_md,
        ).pack(anchor="w")
        tb.Checkbutton(
            lf_output,
            text=self.tr("chk_output_merged"),
            variable=self.var_output_enable_merged,
        ).pack(anchor="w", pady=(4, 0))
        tb.Checkbutton(
            lf_output,
            text=self.tr("chk_output_independent"),
            variable=self.var_output_enable_independent,
        ).pack(anchor="w")

        # Merge & convert sub-function
        frm_merge_submode = tb.Frame(lf_output)
        frm_merge_submode.pack(fill=X, pady=(6, 0))
        tb.Label(
            frm_merge_submode,
            text=self.tr("lbl_merge_convert_submode"),
            font=("System", 9, "bold"),
        ).pack(anchor="w")
        self.var_merge_convert_submode = tk.StringVar(
            value=MERGE_CONVERT_SUBMODE_MERGE_ONLY
        )
        self.rb_merge_submode_merge = tb.Radiobutton(
            frm_merge_submode,
            text=self.tr("rad_merge_convert_merge_only"),
            variable=self.var_merge_convert_submode,
            value=MERGE_CONVERT_SUBMODE_MERGE_ONLY,
        )
        self.rb_merge_submode_merge.pack(anchor="w")
        self.rb_merge_submode_pdf_to_md = tb.Radiobutton(
            frm_merge_submode,
            text=self.tr("rad_merge_convert_pdf_to_md"),
            variable=self.var_merge_convert_submode,
            value=MERGE_CONVERT_SUBMODE_PDF_TO_MD,
        )
        self.rb_merge_submode_pdf_to_md.pack(anchor="w")

        # Tab 妗嗘灦鍜屾粴鍔ㄩ〉闈㈠凡鍦?_build_ui 涓垱寤猴紝涓嶅啀闇€瑕佸瓙 Notebook

        # 銆屾⒊鐞嗐€嶉€夐」涓庡悎骞堕€夐」鍚堝苟鍒板悓涓€椤碉紝鍑忓皯绌虹櫧
        lf_collect = tb.Labelframe(self._scroll_merge, text=self.tr("grp_collect_runtime"), padding=6)
        lf_collect.pack(fill=X, pady=3)
        tb.Label(lf_collect, text=self.tr("lbl_collect_mode"), font=("System", 9, "bold")).pack(anchor="w")
        self.frm_collect_opts = tb.Frame(lf_collect, padding=(10, 5))
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

        lf_mshelp_runtime = tb.Labelframe(
            self._scroll_mshelp, text=self.tr("grp_mshelp_runtime"), padding=6
        )
        lf_mshelp_runtime.pack(fill=X, pady=3)
        tb.Label(
            lf_mshelp_runtime,
            text=self.tr("lbl_mshelp_folder_name"),
            font=("System", 9, "bold"),
        ).pack(anchor="w")
        self.var_mshelpviewer_folder_name = tk.StringVar(value="MSHelpViewer")
        self.ent_mshelpviewer_folder_name = tb.Entry(
            lf_mshelp_runtime, textvariable=self.var_mshelpviewer_folder_name
        )
        self.ent_mshelpviewer_folder_name.pack(fill=X)
        self._attach_tooltip(
            self.ent_mshelpviewer_folder_name, "tip_input_mshelp_folder_name"
        )
        self.var_enable_mshelp_merge_output = tk.IntVar(value=1)
        self.chk_enable_mshelp_merge_output = tb.Checkbutton(
            lf_mshelp_runtime,
            text=self.tr("chk_mshelp_merge_output"),
            variable=self.var_enable_mshelp_merge_output,
        )
        self.chk_enable_mshelp_merge_output.pack(anchor="w", pady=(6, 0))
        self._attach_tooltip(
            self.chk_enable_mshelp_merge_output, "tip_toggle_mshelp_merge_output"
        )
        self.var_enable_mshelp_output_docx = tk.IntVar(value=0)
        self.chk_enable_mshelp_output_docx = tb.Checkbutton(
            lf_mshelp_runtime,
            text=self.tr("chk_mshelp_output_docx"),
            variable=self.var_enable_mshelp_output_docx,
        )
        self.chk_enable_mshelp_output_docx.pack(anchor="w", pady=(6, 0))
        self._attach_tooltip(
            self.chk_enable_mshelp_output_docx, "tip_toggle_mshelp_output_docx"
        )
        self.var_enable_mshelp_output_pdf = tk.IntVar(value=0)
        self.chk_enable_mshelp_output_pdf = tb.Checkbutton(
            lf_mshelp_runtime,
            text=self.tr("chk_mshelp_output_pdf"),
            variable=self.var_enable_mshelp_output_pdf,
        )
        self.chk_enable_mshelp_output_pdf.pack(anchor="w")
        self._attach_tooltip(
            self.chk_enable_mshelp_output_pdf, "tip_toggle_mshelp_output_pdf"
        )

        # Section 2: paths (runtime only)
        lf_paths = tb.Labelframe(self._scroll_shared, text=self.tr("grp_shared_runtime"), padding=6)
        lf_paths.pack(fill=X, pady=3)
        self._add_section_help(lf_paths, "tip_section_run_paths")

        # Source Folders (Multi-select)
        frm_src = tb.Frame(lf_paths)
        frm_src.pack(fill=X, pady=(5, 0))
        tb.Label(frm_src, text=self.tr("lbl_source"), font=("System", 9, "bold")).pack(anchor="w")

        frm_src_body = tb.Frame(frm_src)
        frm_src_body.pack(fill=X, expand=YES)

        self.lst_source_folders = tk.Listbox(frm_src_body, height=6, selectmode=EXTENDED, font=("System", 9), activestyle="dotbox")
        self.lst_source_folders.pack(side=LEFT, fill=X, expand=YES)
        self._attach_tooltip(self.lst_source_folders, "tip_input_source_folder")

        scr_src = tb.Scrollbar(frm_src_body, orient="vertical", command=self.lst_source_folders.yview)
        scr_src.pack(side=LEFT, fill=Y)
        self.lst_source_folders.configure(yscrollcommand=scr_src.set)
        self.lst_source_folders.bind("<Double-Button-1>", self.open_source_folder)

        frm_src_btns = tb.Frame(frm_src_body)
        frm_src_btns.pack(side=LEFT, fill=Y, padx=(5, 0))

        self.btn_add_src = tb.Button(frm_src_btns, text="+", width=3, command=self.add_source_folder, bootstyle="success-outline")
        self.btn_add_src.pack(pady=1)
        self._attach_tooltip(self.btn_add_src, "tip_add_source_folder")

        self.btn_del_src = tb.Button(frm_src_btns, text="-", width=3, command=self.remove_source_folder, bootstyle="danger-outline")
        self.btn_del_src.pack(pady=1)
        self._attach_tooltip(self.btn_del_src, "tip_remove_source_folder")

        self.btn_clr_src = tb.Button(frm_src_btns, text="C", width=3, command=self.clear_source_folders, bootstyle="secondary-outline")
        self.btn_clr_src.pack(pady=1)
        self._attach_tooltip(self.btn_clr_src, "tip_clear_source_folders")

        # Compatibility
        self.var_source_folder = tk.StringVar()

        self.var_target_folder = tk.StringVar()
        self._create_path_row(lf_paths, "lbl_target", self.var_target_folder, self.browse_target, self.open_target_folder)
        self.var_enable_corpus_manifest = tk.IntVar(value=1)
        self.chk_corpus_manifest = tb.Checkbutton(
            lf_paths,
            text=self.tr("chk_corpus_manifest"),
            variable=self.var_enable_corpus_manifest,
            bootstyle="round-toggle",
        )
        self.chk_corpus_manifest.pack(anchor="w", pady=(6, 0))
        self._attach_tooltip(self.chk_corpus_manifest, "tip_toggle_corpus_manifest")

        # Section 3: feature-specific runtime options锛堣浆鎹㈤€夐」锛氬乏鍙冲弻鍒楀竷灞€锛?
        lf_settings = tb.Labelframe(self._scroll_convert, text=self.tr("grp_convert_runtime"), padding=6)
        lf_settings.pack(fill=BOTH, pady=3)
        self._add_section_help(lf_settings, "tip_section_run_advanced")

        # 鍙屽垪瀹瑰櫒锛氬乏鍒楋紙寮曟搸/娌欑洅/绛涢€夛級锛屽彸鍒楋紙AI 瀵煎嚭 + LLM Hub + 澧為噺锛?
        frm_convert_cols = tb.Frame(lf_settings)
        frm_convert_cols.pack(fill=BOTH, expand=YES)

        col_left = tb.Frame(frm_convert_cols)
        col_left.grid(row=0, column=0, sticky="nsew", padx=(0, 6))

        col_right = tb.Frame(frm_convert_cols)
        col_right.grid(row=0, column=1, sticky="nsew", padx=(6, 0))

        frm_convert_cols.columnconfigure(0, weight=1)
        frm_convert_cols.columnconfigure(1, weight=1)

        lf_convert_runtime = tb.Labelframe(col_left, text=self.tr("grp_convert_runtime"), padding=8)
        lf_convert_runtime.pack(fill=X, pady=(2, 6))
        self.group_exec = tb.Frame(lf_convert_runtime)
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

        frm_sandbox_guard = tb.Frame(self.group_exec)
        frm_sandbox_guard.pack(fill=X, padx=20, pady=(4, 0))
        self.var_sandbox_min_free_gb = tk.StringVar()
        self.entry_sandbox_min_free_gb = tb.Entry(
            frm_sandbox_guard, textvariable=self.var_sandbox_min_free_gb, width=6
        )
        self.entry_sandbox_min_free_gb.pack(side=LEFT)
        lbl_min_free = tb.Label(
            frm_sandbox_guard, text=self.tr("lbl_sandbox_min_free_gb")
        )
        lbl_min_free.pack(side=LEFT, padx=(4, 0))
        self.var_sandbox_low_space_policy = tk.StringVar(value="block")
        self.cb_sandbox_low_space_policy = tb.Combobox(
            frm_sandbox_guard,
            textvariable=self.var_sandbox_low_space_policy,
            values=["block", "confirm", "warn"],
            state="readonly",
            width=8,
        )
        self.cb_sandbox_low_space_policy.pack(side=LEFT, padx=(10, 0))
        self._attach_tooltip(
            self.entry_sandbox_min_free_gb, "tip_input_sandbox_min_free_gb"
        )
        self._attach_tooltip(
            self.cb_sandbox_low_space_policy, "tip_input_sandbox_low_space_policy"
        )

        lf_merge_runtime = tb.Labelframe(self._scroll_merge, text=self.tr("grp_merge_runtime"), padding=8)
        lf_merge_runtime.pack(fill=X, pady=(2, 0))
        self.lbl_merge = tb.Label(lf_merge_runtime, text=self.tr("lbl_merge_logic"), bootstyle="primary")
        self.lbl_merge.pack(anchor="w")
        self.var_enable_merge = tk.IntVar(value=1)
        self.chk_enable_merge = tb.Checkbutton(
            lf_merge_runtime,
            text=self.tr("chk_enable_merge"),
            variable=self.var_enable_merge,
            bootstyle="square-toggle",
        )
        self.chk_enable_merge.pack(anchor="w")
        self.frm_merge_opts = tb.Frame(lf_merge_runtime, padding=(20, 0))
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
        self.frm_merge_src = tb.Frame(self.frm_merge_opts)
        self.frm_merge_src.pack(fill=X)
        tb.Radiobutton(self.frm_merge_src, text=self.tr("rad_src_dir"), variable=self.var_merge_source, value="source").pack(side=LEFT)
        tb.Radiobutton(self.frm_merge_src, text=self.tr("rad_tgt_dir"), variable=self.var_merge_source, value="target").pack(side=LEFT, padx=10)
        tb.Separator(self.frm_merge_opts).pack(fill=X, pady=5)
        self.lbl_merge_output_summary = tb.Label(
            self.frm_merge_opts, text="", bootstyle="secondary", wraplength=420
        )
        self.lbl_merge_output_summary.pack(anchor="w", pady=(4, 0))
        self.lbl_merge_inline_hint = tb.Label(
            self.frm_merge_opts, text="", wraplength=420, font=("System", 9)
        )
        self.lbl_merge_inline_hint.pack(anchor="w", pady=(2, 0))
        try:
            self.var_output_enable_merged.trace_add("write", lambda *a: self.after(0, self._on_merge_output_or_mode_change))
            self.var_merge_mode.trace_add("write", lambda *a: self.after(0, self._on_merge_output_or_mode_change))
            for _v in ("var_output_enable_pdf", "var_output_enable_md", "var_output_enable_independent"):
                if hasattr(self, _v):
                    getattr(self, _v).trace_add("write", lambda *a: self.after(0, self._update_output_summary_label))
        except Exception:
            pass
        self._update_output_summary_label()

        # Section 4: conversion strategy + date filter锛堝乏鍒楋級
        lf_convert_content = tb.Labelframe(
            col_left, text=self.tr("sec_filters"), padding=6
        )
        lf_convert_content.pack(fill=X, pady=5)
        self._add_section_help(lf_convert_content, "tip_section_run_filters")
        self.lbl_strategy = tb.Label(lf_convert_content, text=self.tr("lbl_strategy"))
        self.lbl_strategy.pack(anchor="w")
        self.var_strategy = tk.StringVar(value="standard")
        self.cb_strat = tb.Combobox(
            lf_convert_content,
            textvariable=self.var_strategy,
            values=["standard", "smart_tag", "price_only"],
            state="readonly",
        )
        self.cb_strat.pack(fill=X, pady=(0, 5))

        # Section 5: AI export (convert-specific锛屽彸鍒?
        lf_ai_export = tb.Labelframe(
            col_right, text=self.tr("grp_ai_runtime"), padding=(8, 6)
        )
        lf_ai_export.pack(fill=X, pady=(2, 6))
        frm_ai_export = tb.Frame(lf_ai_export)
        frm_ai_export.pack(fill=X, pady=(2, 6))
        self.var_enable_markdown = tk.IntVar(value=1)
        self.chk_export_markdown = tb.Checkbutton(
            frm_ai_export,
            text=self.tr("chk_export_markdown"),
            variable=self.var_enable_markdown,
            command=self._on_toggle_markdown_master,
        )
        self.chk_export_markdown.pack(anchor="w")
        # Markdown 瀛愰€夐」缂╄繘鏄剧ず锛屽彈涓诲紑鍏宠仈鍔ㄧ伆鍖?
        self._frm_markdown_sub = tb.Frame(frm_ai_export)
        self._frm_markdown_sub.pack(fill=X, padx=(16, 0))
        self.var_markdown_strip_header_footer = tk.IntVar(value=1)
        self.chk_markdown_strip_header_footer = tb.Checkbutton(
            self._frm_markdown_sub,
            text=self.tr("chk_markdown_strip_header_footer"),
            variable=self.var_markdown_strip_header_footer,
        )
        self.chk_markdown_strip_header_footer.pack(anchor="w")
        self.var_markdown_structured_headings = tk.IntVar(value=1)
        self.chk_markdown_structured_headings = tb.Checkbutton(
            self._frm_markdown_sub,
            text=self.tr("chk_markdown_structured_headings"),
            variable=self.var_markdown_structured_headings,
        )
        self.chk_markdown_structured_headings.pack(anchor="w")
        self.var_enable_markdown_quality_report = tk.IntVar(value=1)
        self.chk_markdown_quality_report = tb.Checkbutton(
            self._frm_markdown_sub,
            text=self.tr("chk_markdown_quality_report"),
            variable=self.var_enable_markdown_quality_report,
        )
        self.chk_markdown_quality_report.pack(anchor="w")
        tb.Separator(frm_ai_export, orient="horizontal").pack(fill=X, pady=(4, 4))
        self.var_enable_excel_json = tk.IntVar(value=0)
        self.chk_export_records_json = tb.Checkbutton(
            frm_ai_export,
            text=self.tr("chk_export_records_json"),
            variable=self.var_enable_excel_json,
        )
        self.chk_export_records_json.pack(anchor="w")
        self.var_enable_chromadb_export = tk.IntVar(value=0)
        self.chk_chromadb_export = tb.Checkbutton(
            frm_ai_export,
            text=self.tr("chk_chromadb_export"),
            variable=self.var_enable_chromadb_export,
        )
        self.chk_chromadb_export.pack(anchor="w")
        self._attach_tooltip(self.chk_export_markdown, "tip_toggle_export_markdown")
        self._attach_tooltip(
            self.chk_markdown_strip_header_footer,
            "tip_toggle_markdown_strip_header_footer",
        )
        self._attach_tooltip(
            self.chk_markdown_structured_headings,
            "tip_toggle_markdown_structured_headings",
        )
        self._attach_tooltip(
            self.chk_markdown_quality_report, "tip_toggle_markdown_quality_report"
        )
        self._attach_tooltip(
            self.chk_export_records_json, "tip_toggle_export_records_json"
        )
        self._attach_tooltip(
            self.chk_chromadb_export, "tip_toggle_chromadb_export"
        )
        try:
            self.var_output_enable_md.trace_add(
                "write", lambda *_: self._sync_markdown_master_with_global_output()
            )
        except Exception:
            pass
        self._sync_markdown_master_with_global_output()

        # Section: LLM delivery hub 鈫?moved to Output Files tab
        lf_llm_hub = tb.Labelframe(
            self._scroll_output, text=self.tr("grp_llm_hub_runtime"), padding=(8, 6)
        )
        lf_llm_hub.pack(fill=X, pady=3)
        self.var_enable_llm_delivery_hub = tk.IntVar(value=1)
        self.chk_enable_llm_delivery_hub = tb.Checkbutton(
            lf_llm_hub,
            text=self.tr("chk_enable_llm_delivery_hub"),
            variable=self.var_enable_llm_delivery_hub,
            command=self._on_toggle_llm_hub_master,
        )
        self.chk_enable_llm_delivery_hub.pack(anchor="w")
        # LLM hub 瀛愰€夐」缂╄繘锛屽彈涓诲紑鍏宠仈鍔?
        self._frm_llm_hub_sub = tb.Frame(lf_llm_hub)
        self._frm_llm_hub_sub.pack(fill=X, padx=(16, 0))
        self.var_llm_delivery_root = tk.StringVar()
        frm_llm_root = tb.Frame(self._frm_llm_hub_sub)
        frm_llm_root.pack(fill=X, pady=(2, 2))
        self.entry_llm_delivery_root = tb.Entry(
            frm_llm_root, textvariable=self.var_llm_delivery_root, font=("Consolas", 8)
        )
        self.entry_llm_delivery_root.pack(side=LEFT, fill=X, expand=YES)
        self.btn_llm_delivery_root_reset = tb.Button(
            frm_llm_root,
            text=self.tr("btn_llm_delivery_root_reset"),
            bootstyle="outline",
            width=10,
            command=self._on_reset_llm_delivery_root,
        )
        self.btn_llm_delivery_root_reset.pack(side=LEFT, padx=4)
        self.var_llm_delivery_flatten = tk.IntVar(value=0)
        self.chk_llm_delivery_flatten = tb.Checkbutton(
            self._frm_llm_hub_sub,
            text=self.tr("chk_llm_delivery_flatten"),
            variable=self.var_llm_delivery_flatten,
        )
        self.chk_llm_delivery_flatten.pack(anchor="w")
        self.var_llm_delivery_include_pdf = tk.IntVar(value=0)
        self.chk_llm_delivery_include_pdf = tb.Checkbutton(
            self._frm_llm_hub_sub,
            text=self.tr("chk_llm_delivery_include_pdf"),
            variable=self.var_llm_delivery_include_pdf,
        )
        self.chk_llm_delivery_include_pdf.pack(anchor="w")
        self._attach_tooltip(
            self.chk_enable_llm_delivery_hub, "tip_toggle_enable_llm_delivery_hub"
        )
        self._attach_tooltip(
            self.entry_llm_delivery_root, "tip_input_llm_delivery_root"
        )
        self._attach_tooltip(
            self.chk_llm_delivery_flatten, "tip_toggle_llm_delivery_flatten"
        )
        self._attach_tooltip(
            self.chk_llm_delivery_include_pdf, "tip_toggle_llm_delivery_include_pdf"
        )

        # Section: Upload manifest settings
        lf_upload_manifest = tb.Labelframe(
            self._scroll_output, text=self.tr("grp_upload_manifest"), padding=(8, 6)
        )
        lf_upload_manifest.pack(fill=X, pady=3)
        self._add_section_help(lf_upload_manifest, "tip_section_upload_manifest")

        self.var_enable_upload_readme = tk.IntVar(value=1)
        self.chk_enable_upload_readme = tb.Checkbutton(
            lf_upload_manifest,
            text=self.tr("chk_enable_upload_readme"),
            variable=self.var_enable_upload_readme,
        )
        self.chk_enable_upload_readme.pack(anchor="w")
        self._attach_tooltip(self.chk_enable_upload_readme, "tip_toggle_upload_readme")

        self.var_enable_upload_json_manifest = tk.IntVar(value=1)
        self.chk_enable_upload_json_manifest = tb.Checkbutton(
            lf_upload_manifest,
            text=self.tr("chk_enable_upload_json_manifest"),
            variable=self.var_enable_upload_json_manifest,
        )
        self.chk_enable_upload_json_manifest.pack(anchor="w")
        self._attach_tooltip(self.chk_enable_upload_json_manifest, "tip_toggle_upload_json_manifest")

        # Section: Upload dedup policy
        lf_upload_dedup = tb.Labelframe(
            self._scroll_output, text=self.tr("grp_upload_dedup"), padding=(8, 6)
        )
        lf_upload_dedup.pack(fill=X, pady=3)
        self._add_section_help(lf_upload_dedup, "tip_section_upload_dedup")

        self.var_upload_dedup_merged = tk.IntVar(value=1)
        self.chk_upload_dedup_merged = tb.Checkbutton(
            lf_upload_dedup,
            text=self.tr("chk_upload_dedup_merged"),
            variable=self.var_upload_dedup_merged,
        )
        self.chk_upload_dedup_merged.pack(anchor="w")
        self._attach_tooltip(self.chk_upload_dedup_merged, "tip_toggle_upload_dedup_merged")

        # Section 7: incremental / dedup (convert-specific锛屽彸鍒?
        lf_incremental = tb.Labelframe(
            col_right, text=self.tr("grp_incremental_runtime"), padding=(8, 6)
        )
        lf_incremental.pack(fill=X, pady=(2, 6))
        self.var_enable_incremental_mode = tk.IntVar(value=0)
        self.chk_incremental_mode = tb.Checkbutton(
            lf_incremental,
            text=self.tr("chk_incremental_mode"),
            variable=self.var_enable_incremental_mode,
            command=self._on_toggle_incremental_mode,
        )
        self.chk_incremental_mode.pack(anchor="w")
        # 澧為噺瀛愰€夐」缂╄繘锛屽彈涓诲紑鍏宠仈鍔?
        self._frm_incremental_sub = tb.Frame(lf_incremental)
        self._frm_incremental_sub.pack(fill=X, padx=(16, 0))
        self.var_incremental_verify_hash = tk.IntVar(value=0)
        self.chk_incremental_verify_hash = tb.Checkbutton(
            self._frm_incremental_sub,
            text=self.tr("chk_incremental_verify_hash"),
            variable=self.var_incremental_verify_hash,
        )
        self.chk_incremental_verify_hash.pack(anchor="w")
        self.var_incremental_reprocess_renamed = tk.IntVar(value=0)
        self.chk_incremental_reprocess_renamed = tb.Checkbutton(
            self._frm_incremental_sub,
            text=self.tr("chk_incremental_reprocess_renamed"),
            variable=self.var_incremental_reprocess_renamed,
        )
        self.chk_incremental_reprocess_renamed.pack(anchor="w")
        self.var_source_priority_skip_same_name_pdf = tk.IntVar(value=0)
        self.chk_source_priority_skip_pdf = tb.Checkbutton(
            lf_incremental,
            text=self.tr("chk_source_priority_skip_pdf"),
            variable=self.var_source_priority_skip_same_name_pdf,
        )
        self.chk_source_priority_skip_pdf.pack(anchor="w")
        self.var_global_md5_dedup = tk.IntVar(value=0)
        self.chk_global_md5_dedup = tb.Checkbutton(
            lf_incremental,
            text=self.tr("chk_global_md5_dedup"),
            variable=self.var_global_md5_dedup,
        )
        self.chk_global_md5_dedup.pack(anchor="w")
        self.var_enable_update_package = tk.IntVar(value=1)
        self.chk_enable_update_package = tb.Checkbutton(
            lf_incremental,
            text=self.tr("chk_enable_update_package"),
            variable=self.var_enable_update_package,
        )
        self.chk_enable_update_package.pack(anchor="w")
        self._attach_tooltip(self.chk_incremental_mode, "tip_toggle_incremental_mode")
        self._attach_tooltip(
            self.chk_incremental_verify_hash, "tip_toggle_incremental_verify_hash"
        )
        self._attach_tooltip(
            self.chk_incremental_reprocess_renamed,
            "tip_toggle_incremental_reprocess_renamed",
        )
        self._attach_tooltip(
            self.chk_source_priority_skip_pdf, "tip_toggle_source_priority_skip_pdf"
        )
        self._attach_tooltip(
            self.chk_global_md5_dedup, "tip_toggle_global_md5_dedup"
        )
        self._attach_tooltip(
            self.chk_enable_update_package, "tip_toggle_enable_update_package"
        )

        self.var_enable_date_filter = tk.IntVar(value=0)
        self.chk_date_filter = tb.Checkbutton(
            lf_convert_content,
            text=self.tr("lbl_filter_date"),
            variable=self.var_enable_date_filter,
            command=self._on_toggle_date_filter,
        )
        self.chk_date_filter.pack(anchor="w", pady=(5, 0))
        self.frm_date = tb.Frame(lf_convert_content)
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

        # Section 5: NotebookLM locator (runtime) 鈥斺€?涓?MSHelp 鍚堝苟鍒板悓涓€ tab
        lf_locator = tb.Labelframe(self._scroll_locator, text=self.tr("sec_locator"), padding=10)
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
        self._auto_attach_action_tooltips(lf_collect)
        self._auto_attach_action_tooltips(lf_mshelp_runtime)
        self._auto_attach_action_tooltips(lf_paths)
        self._auto_attach_action_tooltips(lf_settings)
        self._auto_attach_action_tooltips(lf_convert_content)
        self._auto_attach_action_tooltips(lf_ai_export)
        self._auto_attach_action_tooltips(lf_incremental)
        self._auto_attach_action_tooltips(lf_locator)
        self._auto_attach_input_tooltips(lf_mode, "tip_section_run_mode")
        self._auto_attach_input_tooltips(lf_paths, "tip_section_run_paths")
        self._auto_attach_input_tooltips(lf_settings, "tip_section_run_advanced")
        self._auto_attach_input_tooltips(lf_merge_runtime, "tip_section_run_advanced")
        self._auto_attach_input_tooltips(lf_collect, "tip_section_run_advanced")
        self._auto_attach_input_tooltips(lf_mshelp_runtime, "tip_mode_mshelp")
        self._auto_attach_input_tooltips(lf_locator, "tip_section_run_locator")
        self._build_task_tab_content()
        self._attach_tooltip(self.entry_temp_sandbox_root, "tip_input_sandbox_root")
        self._attach_tooltip(self.cb_strat, "tip_input_strategy")
        self._attach_tooltip(self.ent_date, "tip_input_date")
        self._bind_var_validation(self.var_locator_page, lambda: self._normalize_then_validate(self.var_locator_page, self._normalize_numeric_var, "locator"))
        self._bind_var_validation(self.var_locator_short_id, lambda: self._normalize_then_validate(self.var_locator_short_id, self._normalize_short_id_var, "locator"))
        self._bind_var_validation(self.var_date_str, lambda: self._normalize_then_validate(self.var_date_str, self._normalize_date_var, "run"))
        self._bind_var_validation(self.var_enable_date_filter, lambda: self.validate_runtime_inputs(silent=False, scope="run"))

    def _build_task_tab_content(self):
        parent = self._scroll_tasks
        lf_tasks = tb.Labelframe(parent, text=self.tr("grp_task_runtime"), padding=8)
        lf_tasks.pack(fill=BOTH, expand=YES, pady=3)

        cols = ("name", "source", "target", "status", "last_run")
        self.tree_tasks = ttk.Treeview(
            lf_tasks, columns=cols, show="headings", height=12, selectmode="browse"
        )
        for c, key in zip(cols, ("col_task_name", "col_task_source", "col_task_target", "col_task_status", "col_task_last_run")):
            self.tree_tasks.heading(c, text=self.tr(key))
        self.tree_tasks.column("name", width=160)
        self.tree_tasks.column("source", width=180)
        self.tree_tasks.column("target", width=180)
        self.tree_tasks.column("status", width=56)
        self.tree_tasks.column("last_run", width=100)
        self.tree_tasks.pack(fill=BOTH, expand=YES, side=LEFT)
        self._attach_tooltip(self.tree_tasks, "tip_task_list")
        scr = tb.Scrollbar(lf_tasks, orient="vertical", command=self.tree_tasks.yview)
        scr.pack(side=LEFT, fill=Y)
        self.tree_tasks.configure(yscrollcommand=scr.set)
        self.tree_tasks.bind("<<TreeviewSelect>>", lambda _e: self._on_task_select())

        btn_col = tb.Frame(lf_tasks)
        btn_col.pack(side=LEFT, fill=Y, padx=(8, 0))

        self.btn_task_create = tb.Button(
            btn_col,
            text=self.tr("btn_task_create"),
            command=self._on_click_task_create,
            bootstyle="success-outline",
            width=12,
        )
        self.btn_task_create.pack(pady=2, fill=X)

        self.btn_task_edit = tb.Button(
            btn_col,
            text=self.tr("btn_task_edit"),
            command=self._on_click_task_edit,
            bootstyle="secondary-outline",
            width=12,
        )
        self.btn_task_edit.pack(pady=2, fill=X)

        self.btn_task_delete = tb.Button(
            btn_col,
            text=self.tr("btn_task_delete"),
            command=self._on_click_task_delete,
            bootstyle="danger-outline",
            width=12,
        )
        self.btn_task_delete.pack(pady=2, fill=X)

        self.btn_task_refresh = tb.Button(
            btn_col,
            text=self.tr("btn_task_refresh"),
            command=self._refresh_task_list_ui,
            bootstyle="info-outline",
            width=12,
        )
        self.btn_task_refresh.pack(pady=(2, 8), fill=X)

        self.btn_task_load_to_ui = tb.Button(
            btn_col,
            text=self.tr("btn_task_load_to_ui"),
            command=self._on_click_task_load_to_ui,
            bootstyle="secondary-outline",
            width=12,
        )
        self.btn_task_load_to_ui.pack(pady=2, fill=X)

        self.btn_task_run = tb.Button(
            btn_col,
            text=self.tr("btn_task_run"),
            command=lambda: self._on_click_task_run(resume=False),
            bootstyle="success",
            width=12,
        )
        self.btn_task_run.pack(pady=2, fill=X)

        self.btn_task_resume = tb.Button(
            btn_col,
            text=self.tr("btn_task_resume"),
            command=lambda: self._on_click_task_run(resume=True),
            bootstyle="warning-outline",
            width=12,
        )
        self.btn_task_resume.pack(pady=2, fill=X)

        self.btn_task_stop = tb.Button(
            btn_col,
            text=self.tr("btn_task_stop"),
            command=self._on_click_stop,
            bootstyle="danger-outline",
            width=12,
            state="disabled",
        )
        self.btn_task_stop.pack(pady=(2, 8), fill=X)

        self.var_task_force_full_rebuild = tk.IntVar(value=0)
        self.chk_task_force_full_rebuild = tb.Checkbutton(
            btn_col,
            text=self.tr("chk_task_force_full_rebuild"),
            variable=self.var_task_force_full_rebuild,
            bootstyle="round-toggle",
        )
        self.chk_task_force_full_rebuild.pack(anchor="w", pady=(4, 2))

        tk.Label(parent, text=self.tr("lbl_task_config_section"), font=("System", 9, "bold")).pack(anchor=W, pady=(6, 2))
        self.txt_task_detail = ScrolledText(parent, height=40, wrap=tk.WORD, font=("Consolas", 9))
        self.txt_task_detail.pack(fill=BOTH, expand=YES, pady=(0, 4))
        try:
            self.txt_task_detail.insert(tk.END, self.tr("msg_task_none_selected"))
            self.txt_task_detail.configure(state="disabled")
        except Exception:
            pass

        self._auto_attach_action_tooltips(lf_tasks)
        self._auto_attach_input_tooltips(lf_tasks, "tip_section_run_mode")
        if not getattr(self, "_task_tab_app_mode_trace_done", False) and hasattr(self, "var_app_mode"):
            self.var_app_mode.trace_add("write", lambda *a: self.after(0, self._on_app_mode_change_for_task_tab))
            self._task_tab_app_mode_trace_done = True
        self._refresh_task_list_ui()

    def _get_selected_task_id(self):
        if not hasattr(self, "tree_tasks"):
            return None
        sel = self.tree_tasks.selection()
        if not sel:
            return None
        return str(sel[0])

    def _short_path(self, path, max_len=36):
        p = str(path or "").strip()
        if len(p) <= max_len:
            return p
        return "..." + p[-(max_len - 3) :]

    def _refresh_task_list_ui(self):
        if not hasattr(self, "tree_tasks"):
            return
        selected_id = self._get_selected_task_id()
        for iid in self.tree_tasks.get_children():
            self.tree_tasks.delete(iid)
        tasks = self.task_store.list_tasks()
        for task in tasks:
            task_id = str(task.get("id", ""))
            name = str(task.get("name", ""))[:48]
            source = self._short_path(task.get("source_folder", ""))
            target = self._short_path(task.get("target_folder", ""))
            status = str(task.get("status", "idle"))
            last_run = (task.get("last_run_at") or "")[:16]
            self.tree_tasks.insert("", END, iid=task_id, values=(name, source, target, status, last_run))
        if selected_id and self.tree_tasks.exists(selected_id):
            self.tree_tasks.selection_set(selected_id)
            self.tree_tasks.focus(selected_id)
        self._on_task_select()

    def _on_task_select(self):
        task_id = self._get_selected_task_id()
        task = self.task_store.get_task(task_id) if task_id else None
        if not task:
            if hasattr(self, "txt_task_detail"):
                try:
                    self.txt_task_detail.configure(state="normal")
                    self.txt_task_detail.delete("1.0", tk.END)
                    self.txt_task_detail.insert(tk.END, self.tr("msg_task_none_selected"))
                    self.txt_task_detail.configure(state="disabled")
                except Exception:
                    pass
            if hasattr(self, "btn_task_resume"):
                self.btn_task_resume.configure(state="disabled")
            self._update_task_tab_for_app_mode()
            return

        cp = self.task_store.load_checkpoint(task_id)
        planned = len((cp or {}).get("planned_files", []) or [])
        done = len((cp or {}).get("completed_files", []) or [])
        run_mode = ""
        try:
            runtime_preview = build_task_runtime_config(
                self._load_config_for_write(), task, force_full_rebuild=False
            )
            run_mode = str(runtime_preview.get("run_mode", ""))
        except Exception:
            run_mode = str(task.get("config_overrides", {}).get("run_mode", ""))

        # 下方展示任务完整配置（非仅摘要）
        full_lines = [
            self.tr("msg_task_detail").format(
                task.get("name", ""),
                task.get("source_folder", ""),
                task.get("target_folder", ""),
                run_mode or "-",
                "ON" if task.get("run_incremental", True) else "OFF",
                task.get("status", "idle"),
                done,
                planned,
            ),
            "",
            "--- " + (self.tr("lbl_task_full_config") or "Task full config") + " ---",
            "name: " + str(task.get("name", "")),
            "description: " + str((task.get("description") or "").replace("\n", " ")[:200]),
            "source_folders: " + str(task.get("source_folders", []) or [task.get("source_folder", "")]),
            "target_folder: " + str(task.get("target_folder", "")),
            "run_incremental: " + str(task.get("run_incremental", True)),
            "status: " + str(task.get("status", "idle")),
            "last_run_at: " + str(task.get("last_run_at") or ""),
            "checkpoint: " + str(done) + " / " + str(planned),
        ]
        overrides = task.get("config_overrides") or {}
        if overrides:
            full_lines.append("")
            full_lines.append("config_overrides:")
            for k, v in sorted(overrides.items()):
                full_lines.append("  " + str(k) + ": " + str(v))
        full_text = "\n".join(full_lines)

        if hasattr(self, "txt_task_detail"):
            try:
                self.txt_task_detail.configure(state="normal")
                self.txt_task_detail.delete("1.0", tk.END)
                self.txt_task_detail.insert(tk.END, full_text)
                self.txt_task_detail.configure(state="disabled")
            except Exception:
                pass
        can_resume = bool(cp and done < planned)
        self.btn_task_resume.configure(state="normal" if can_resume else "disabled")
        self._update_task_tab_for_app_mode()

    def _update_task_tab_for_app_mode(self):
        """In classic mode hide task tab entirely; in task mode show it and enable Run/Resume (design 7.1)."""
        if not hasattr(self, "var_app_mode"):
            return
        is_classic = self.var_app_mode.get() == "classic"
        if hasattr(self, "main_notebook") and hasattr(self, "tab_run_tasks"):
            self._set_run_tab_state(self.tab_run_tasks, "hidden" if is_classic else "normal")
        # 打开任务模式后任务管理标签页显示为绿色
        self._set_task_tab_highlight(not is_classic)
        if not hasattr(self, "btn_task_run"):
            return
        if is_classic:
            self.btn_task_run.configure(state="disabled")
            if hasattr(self, "btn_task_resume"):
                self.btn_task_resume.configure(state="disabled")
        else:
            self.btn_task_run.configure(state="normal")
            # resume state left as set by _on_task_select

    def _on_app_mode_change_for_task_tab(self):
        self._update_task_tab_for_app_mode()
        if getattr(self, "var_app_mode", None) and self.var_app_mode.get() == "task":
            self._on_task_select()

    def _normalize_task_cfg_for_compare(self, cfg):
        normalized = dict(cfg if isinstance(cfg, dict) else {})
        if "output_enable_md" not in normalized:
            normalized["output_enable_md"] = bool(
                normalized.get("enable_markdown", True)
            )
        normalized["enable_markdown"] = bool(normalized.get("output_enable_md", True))
        if normalized.get("run_mode") == MODE_CONVERT_THEN_MERGE:
            normalized["merge_source"] = "target"
        return normalized

    def _build_task_overrides_from_ui(self, project_cfg=None, only_diff=True):
        mode = self.var_run_mode.get()
        merge_source = (
            "target" if mode == MODE_CONVERT_THEN_MERGE else self.var_merge_source.get()
        )
        current = {
            "run_mode": mode,
            "collect_mode": self.var_collect_mode.get(),
            "content_strategy": self.var_strategy.get(),
            "default_engine": self.var_engine.get(),
            "kill_process_mode": self.var_kill_mode.get(),
            "enable_merge": bool(self.var_enable_merge.get()),
            "output_enable_pdf": bool(self.var_output_enable_pdf.get()),
            "output_enable_md": bool(self.var_output_enable_md.get()),
            "output_enable_merged": bool(self.var_output_enable_merged.get()),
            "output_enable_independent": bool(self.var_output_enable_independent.get()),
            "merge_convert_submode": self.var_merge_convert_submode.get(),
            "merge_mode": self.var_merge_mode.get(),
            "merge_source": merge_source,
            "enable_merge_index": bool(self.var_enable_merge_index.get()),
            "enable_merge_excel": bool(self.var_enable_merge_excel.get()),
            "enable_sandbox": bool(self.var_enable_sandbox.get()),
            "sandbox_min_free_gb": self._safe_positive_int(
                self.var_sandbox_min_free_gb.get(), 10
            ),
            "sandbox_low_space_policy": self.var_sandbox_low_space_policy.get()
            or "block",
            "enable_markdown_quality_report": bool(
                self.var_enable_markdown_quality_report.get()
            ),
            "enable_excel_json": bool(self.var_enable_excel_json.get()),
            "enable_chromadb_export": bool(self.var_enable_chromadb_export.get()),
            "incremental_verify_hash": bool(self.var_incremental_verify_hash.get()),
            "incremental_reprocess_renamed": bool(
                self.var_incremental_reprocess_renamed.get()
            ),
            "source_priority_skip_same_name_pdf": bool(
                self.var_source_priority_skip_same_name_pdf.get()
            ),
            "global_md5_dedup": bool(self.var_global_md5_dedup.get()),
            "enable_update_package": bool(self.var_enable_update_package.get()),
        }
        current = self._normalize_task_cfg_for_compare(current)
        if not only_diff:
            return current

        base = self._normalize_task_cfg_for_compare(
            project_cfg if isinstance(project_cfg, dict) else self._load_config_for_write()
        )
        overrides = {}
        for key, value in current.items():
            if base.get(key) != value:
                overrides[key] = value
        return overrides

    def _new_task_id(self):
        return datetime.now().strftime("task_%Y%m%d_%H%M%S_%f")

    def _on_click_task_create(self):
        self._open_task_wizard()

    def _open_task_wizard(self):
        win = tk.Toplevel(self)
        win.title(self.tr("win_task_wizard"))
        win.minsize(680, 560)
        win.geometry("700x580")
        win.transient(self)
        try:
            win.configure(bg="#f0f0f0")
        except Exception:
            pass
        data = {
            "name": "",
            "description": "",
            "source_folder": "",
            "source_folders": [],
            "target_folder": "",
            "run_incremental": True,
            "run_mode": MODE_CONVERT_THEN_MERGE,
            "output_enable_pdf": True,
            "output_enable_md": True,
            "output_enable_merged": True,
            "output_enable_independent": False,
            "output_choice": "both",
            "merge_how": MERGE_MODE_CATEGORY,
            "max_merge_size_mb": 80,
            "merge_filename_pattern": (
                getattr(self, "var_merge_filename_pattern", None)
                and self.var_merge_filename_pattern.get().strip()
                or "Merged_{category}_{timestamp}_{idx}"
            ),
        }
        win._wizard_data = data
        win._wizard_step = 1
        # 向导窗口内全部使用纯 tk 控件，避免 ttk/ttkbootstrap 在 Toplevel 上不渲染导致空白
        _bg = "#f0f0f0"
        nav = tk.Frame(win, bg=_bg, padx=10, pady=10)
        nav.pack(side=BOTTOM, fill=X)
        content_holder = tk.Frame(win, bg=_bg, padx=15, pady=15)
        content_holder.pack(fill=BOTH, expand=True)
        # 可滚动区域，避免第 3 步配置项被挡住
        canvas = tk.Canvas(content_holder, bg=_bg, highlightthickness=0)
        scrollbar = tk.Scrollbar(content_holder, orient="vertical", command=canvas.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        container = tk.Frame(canvas, bg=_bg)
        canvas_window = canvas.create_window(0, 0, window=container, anchor=tk.NW)

        def _on_container_configure(_ev=None):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def _on_canvas_configure(ev):
            if container.winfo_reqwidth() != ev.width:
                canvas.itemconfig(canvas_window, width=ev.width)

        container.bind("<Configure>", _on_container_configure)
        canvas.bind("<Configure>", _on_canvas_configure)

        def _on_mousewheel(ev):
            canvas.yview_scroll(int(-1 * (ev.delta / 120)), "units")

        def _unbind_mousewheel(e):
            if e.widget == win:
                try:
                    canvas.unbind_all("<MouseWheel>")
                except Exception:
                    pass

        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        win.bind("<Destroy>", _unbind_mousewheel)
        f1 = tk.Frame(container, bg=_bg)
        f2 = tk.Frame(container, bg=_bg)
        f3 = tk.Frame(container, bg=_bg)
        f4 = tk.Frame(container, bg=_bg)
        tk.Label(f1, text=self.tr("wizard_step1"), font=("System", 11, "bold"), bg=_bg).pack(anchor=W)
        tk.Label(f1, text=self.tr("msg_task_input_name"), bg=_bg).pack(anchor=W)
        ent_name = tk.Entry(f1, width=50)
        ent_name.pack(fill=X, pady=(0, 8))
        tk.Label(f1, text=self.tr("lbl_task_description"), bg=_bg).pack(anchor=W)
        txt_desc = tk.Text(f1, height=3, width=50)
        txt_desc.pack(fill=X, pady=(0, 8))
        tk.Label(f2, text=self.tr("wizard_step2"), font=("System", 11, "bold"), bg=_bg).pack(anchor=W)
        tk.Label(f2, text=self.tr("lbl_source"), bg=_bg).pack(anchor=W)
        f2_src = tk.Frame(f2, bg=_bg)
        f2_src.pack(fill=X)
        lst_src = tk.Listbox(f2_src, height=4, selectmode=SINGLE, font=("Consolas", 9))
        lst_src.pack(side=LEFT, fill=BOTH, expand=True, padx=(0, 4))
        scr_src = tk.Scrollbar(f2_src, orient="vertical", command=lst_src.yview)
        scr_src.pack(side=LEFT, fill=Y)
        lst_src.configure(yscrollcommand=scr_src.set)
        f2_src_btns = tk.Frame(f2_src, bg=_bg)
        f2_src_btns.pack(side=LEFT, fill=Y)
        def add_src():
            # Ask user if they want to select multiple folders
            result = messagebox.askyesno(
                self.tr("msg_task_pick_source"),
                self.tr("msg_multi_select_folders"),
                icon="question",
                parent=win
            )
            if result:
                # Multi-select mode - open multi-folder dialog
                self._open_task_multi_folder_dialog(win, lst_src)
            else:
                # Single-select mode
                p = filedialog.askdirectory(title=self.tr("msg_task_pick_source"), parent=win)
                if p and p not in lst_src.get(0, END):
                    lst_src.insert(END, p)
        def remove_src():
            sel = lst_src.curselection()
            if sel:
                lst_src.delete(sel[0])
        tk.Button(f2_src_btns, text="+", command=add_src, width=3).pack(pady=2)
        tk.Button(f2_src_btns, text="-", command=remove_src, width=3).pack(pady=2)
        tk.Label(f2, text=self.tr("lbl_target"), bg=_bg).pack(anchor=W, pady=(8, 0))
        f2_tgt = tk.Frame(f2, bg=_bg)
        f2_tgt.pack(fill=X)
        ent_tgt = tk.Entry(f2_tgt, width=40)
        ent_tgt.pack(side=LEFT, fill=X, expand=True, padx=(0, 4))
        def pick_tgt():
            p = filedialog.askdirectory(title=self.tr("msg_task_pick_target"), parent=win)
            if p:
                ent_tgt.delete(0, END)
                ent_tgt.insert(0, p)
        tk.Button(f2_tgt, text=self.tr("btn_browse"), command=pick_tgt, width=4).pack(side=LEFT)
        var_inc = tk.IntVar(value=1)
        tk.Checkbutton(f2, text=self.tr("chk_incremental_mode"), variable=var_inc, bg=_bg).pack(anchor=W, pady=(8, 0))
        tk.Label(f3, text=self.tr("wizard_step3"), font=("System", 11, "bold"), bg=_bg).pack(anchor=W)
        var_mode = tk.StringVar(value=MODE_CONVERT_THEN_MERGE)
        for val, key in [
            (MODE_CONVERT_ONLY, "mode_convert"),
            (MODE_CONVERT_THEN_MERGE, "mode_convert_merge"),
            (MODE_MERGE_ONLY, "mode_merge"),
            (MODE_COLLECT_ONLY, "mode_collect"),
            (MODE_MSHELP_ONLY, "mode_mshelp"),
        ]:
            tk.Radiobutton(f3, text=self.tr(key), variable=var_mode, value=val, bg=_bg).pack(anchor=W)
        tk.Label(f3, text=self.tr("grp_output_controls"), font=("System", 9, "bold"), bg=_bg).pack(anchor=W, pady=(8, 0))
        var_output_choice = tk.StringVar(value="both")
        tk.Radiobutton(f3, text=self.tr("wizard_output_only_independent"), variable=var_output_choice, value="only_independent", bg=_bg).pack(anchor=W)
        tk.Radiobutton(f3, text=self.tr("wizard_output_only_merged"), variable=var_output_choice, value="only_merged", bg=_bg).pack(anchor=W)
        tk.Radiobutton(f3, text=self.tr("wizard_output_both"), variable=var_output_choice, value="both", bg=_bg).pack(anchor=W)
        f3_merge = tk.Frame(f3, bg=_bg)
        f3_merge.pack(fill=X, pady=(8, 0))
        var_merge_how = tk.StringVar(value=MERGE_MODE_CATEGORY)
        tk.Radiobutton(f3_merge, text=self.tr("wizard_merge_single_file"), variable=var_merge_how, value=MERGE_MODE_ALL_IN_ONE, bg=_bg).pack(anchor=W)
        tk.Radiobutton(f3_merge, text=self.tr("wizard_merge_by_size"), variable=var_merge_how, value=MERGE_MODE_CATEGORY, bg=_bg).pack(anchor=W)
        f3_mb = tk.Frame(f3_merge, bg=_bg)
        f3_mb.pack(fill=X, pady=(2, 0))
        tk.Label(f3_mb, text=self.tr("wizard_merge_size_mb"), bg=_bg).pack(side=LEFT, padx=(0, 4))
        var_mb = tk.StringVar(value="80")
        ent_mb = tk.Entry(f3_mb, width=6, textvariable=var_mb)
        ent_mb.pack(side=LEFT)
        lbl_mb_tip = tk.Label(f3_merge, text="", bg=_bg, fg="#666")
        lbl_mb_tip.pack(anchor=W, pady=(2, 0))

        def _wizard_update_merge_ui():
            need_merge = var_output_choice.get() in ("only_merged", "both")
            if need_merge:
                f3_merge.pack(fill=X, pady=(8, 0))
                is_split = var_merge_how.get() == MERGE_MODE_CATEGORY
                if is_split:
                    ent_mb.configure(state="normal")
                    lbl_mb_tip.configure(text="")
                else:
                    ent_mb.configure(state="disabled")
                    lbl_mb_tip.configure(text=self.tr("tip_wizard_merge_size_disabled"))
            else:
                f3_merge.pack_forget()
        var_output_choice.trace_add("write", lambda *a: _wizard_update_merge_ui())
        var_merge_how.trace_add("write", lambda *a: _wizard_update_merge_ui())
        _wizard_update_merge_ui()
        var_pdf = tk.IntVar(value=1)
        var_md = tk.IntVar(value=1)
        tk.Checkbutton(f3, text=self.tr("chk_output_pdf"), variable=var_pdf, bg=_bg).pack(anchor=W, pady=(4, 0))
        tk.Checkbutton(f3, text=self.tr("chk_output_md"), variable=var_md, bg=_bg).pack(anchor=W)
        tk.Label(f4, text=self.tr("wizard_step4"), font=("System", 11, "bold"), bg=_bg).pack(anchor=W)
        lbl_summary = tk.Label(f4, text="", justify=LEFT, wraplength=450, bg=_bg)
        lbl_summary.pack(anchor=W, pady=(8, 0))
        var_run_after_save = tk.IntVar(value=0)
        chk_run_after = tk.Checkbutton(f4, text=self.tr("chk_wizard_run_after_save"), variable=var_run_after_save, bg=_bg)
        chk_run_after.pack(anchor=W, pady=(8, 0))

        def _show(step):
            win._wizard_step = step
            f1.pack_forget()
            f2.pack_forget()
            f3.pack_forget()
            f4.pack_forget()
            (f1, f2, f3, f4)[step - 1].pack(fill=BOTH, expand=True)
            btn_back.configure(state="normal" if step > 1 else "disabled")
            btn_next.configure(state="normal" if step < 4 else "disabled")
            btn_save.pack_forget()
            btn_next.pack_forget()
            if step == 4:
                btn_save.pack(side=LEFT, padx=4)
                _refresh_summary()
            else:
                btn_next.pack(side=LEFT, padx=4)
            try:
                canvas.yview_moveto(0)
                _on_container_configure()
            except Exception:
                pass

        def _merge_summary_text(d):
            if not d.get("output_enable_merged"):
                return self.tr("wizard_merge_summary_none")
            if d.get("merge_mode") == MERGE_MODE_ALL_IN_ONE:
                return self.tr("wizard_merge_summary_single")
            return self.tr("wizard_merge_summary_split").format(d.get("max_merge_size_mb", 80))

        def _refresh_summary():
            d = win._wizard_data
            src_display = d["source_folder"] if len(d.get("source_folders", [])) <= 1 else f"{d['source_folder']} (+{len(d['source_folders'])-1})"
            lbl_summary.configure(
                text=f"{self.tr('msg_task_input_name')} {d['name']}\n"
                     f"{self.tr('lbl_source')} {src_display}\n"
                     f"{self.tr('lbl_target')} {d['target_folder']}\n"
                     f"{self.tr('chk_incremental_mode')}: {'Y' if d['run_incremental'] else 'N'}\n"
                     f"Run mode: {d['run_mode']}\n"
                     f"{_merge_summary_text(d)}"
            )

        def _collect():
            data["name"] = ent_name.get().strip()
            data["description"] = txt_desc.get("1.0", END).strip()
            data["source_folders"] = list(lst_src.get(0, END))
            data["source_folder"] = (data["source_folders"] or [""])[0]
            data["target_folder"] = ent_tgt.get().strip()
            data["run_incremental"] = bool(var_inc.get())
            data["run_mode"] = var_mode.get()
            data["output_enable_pdf"] = bool(var_pdf.get())
            data["output_enable_md"] = bool(var_md.get())
            choice = var_output_choice.get()
            data["output_choice"] = choice
            if choice == "only_independent":
                data["output_enable_merged"] = False
                data["output_enable_independent"] = True
            elif choice == "only_merged":
                data["output_enable_merged"] = True
                data["output_enable_independent"] = False
            else:
                data["output_enable_merged"] = True
                data["output_enable_independent"] = True
            data["merge_how"] = var_merge_how.get()
            data["max_merge_size_mb"] = self._safe_positive_int(var_mb.get(), 80)
            data["merge_mode"] = data["merge_how"] if data["output_enable_merged"] else MERGE_MODE_ALL_IN_ONE

        def _go(delta):
            _collect()
            step = win._wizard_step + delta
            step = max(1, min(4, step))
            _show(step)

        def _save():
            _collect()
            d = win._wizard_data
            if not d["name"]:
                messagebox.showwarning(self.tr("win_task_wizard"), self.tr("msg_task_input_name"), parent=win)
                return
            name_trim = d["name"].strip()
            for t in (self.task_store.list_tasks() or []):
                if isinstance(t, dict) and (t.get("name") or "").strip() == name_trim:
                    messagebox.showwarning(
                        self.tr("win_task_wizard"), self.tr("msg_task_name_duplicate"), parent=win
                    )
                    return
            source_folders = [p.strip() for p in (d.get("source_folders") or []) if p and str(p).strip()]
            if not source_folders:
                messagebox.showwarning(self.tr("win_task_wizard"), self.tr("msg_source_folder_required"), parent=win)
                return
            for p in source_folders:
                if not os.path.isdir(p):
                    messagebox.showwarning(self.tr("win_task_wizard"), self.tr("msg_source_folder_required"), parent=win)
                    return
            if not d["target_folder"] or not os.path.isdir(d["target_folder"]):
                messagebox.showwarning(self.tr("win_task_wizard"), self.tr("msg_target_folder_required"), parent=win)
                return
            project_cfg = self._load_config_for_write()
            overrides = self._build_task_overrides_from_ui(project_cfg=project_cfg, only_diff=True)
            overrides["run_mode"] = d["run_mode"]
            overrides["output_enable_pdf"] = d["output_enable_pdf"]
            overrides["output_enable_md"] = d["output_enable_md"]
            overrides["output_enable_merged"] = d["output_enable_merged"]
            overrides["output_enable_independent"] = d["output_enable_independent"]
            overrides["merge_mode"] = d.get("merge_mode", MERGE_MODE_CATEGORY)
            overrides["max_merge_size_mb"] = d.get("max_merge_size_mb", 80)
            overrides["merge_filename_pattern"] = (
                d.get("merge_filename_pattern") or "Merged_{category}_{timestamp}_{idx}"
            )
            task = {
                "id": self._new_task_id(),
                "name": d["name"],
                "description": d.get("description", ""),
                "source_folders": source_folders,
                "source_folder": (source_folders or [""])[0],
                "target_folder": d["target_folder"],
                "run_incremental": d["run_incremental"],
                "config_overrides": overrides,
                "status": "idle",
            }
            try:
                self.task_store.save_task(task)
            except Exception as e:
                messagebox.showerror(self.tr("win_task_create"), str(e), parent=win)
                return
            saved_id = task["id"]
            run_after_save = bool(var_run_after_save.get())
            self._refresh_task_list_ui()
            win.destroy()
            if run_after_save and saved_id and hasattr(self, "tree_tasks") and self.tree_tasks.exists(saved_id):
                self.tree_tasks.selection_set(saved_id)
                self.tree_tasks.focus(saved_id)
                self.after(200, lambda: self._on_click_task_run(False))

        # 在 _go/_save 定义之后再创建按钮，避免 Python 3.12 闭包 "free variable not associated" 错误
        btn_back = tk.Button(nav, text=self.tr("btn_wizard_back"), state="disabled", command=lambda: _go(-1))
        btn_back.pack(side=LEFT, padx=4)
        btn_next = tk.Button(nav, text=self.tr("btn_wizard_next"), command=lambda: _go(1))
        btn_next.pack(side=LEFT, padx=4)
        btn_save = tk.Button(nav, text=self.tr("btn_wizard_save"), command=_save)
        btn_save.pack(side=LEFT, padx=4)

        # 直接显示第一步（纯 tk 控件在 Toplevel 上无需等待主题/布局，可立即渲染）
        _show(1)
        try:
            win.update_idletasks()
            win.update()
        except Exception:
            pass

    def _on_click_task_edit(self):
        task_id = self._get_selected_task_id()
        if not task_id:
            messagebox.showinfo(
                self.tr("grp_task_runtime"), self.tr("msg_task_select_required")
            )
            return
        task = self.task_store.get_task(task_id)
        if not task:
            return
        name = simpledialog.askstring(
            self.tr("win_task_edit"),
            self.tr("msg_task_input_name"),
            initialvalue=task.get("name", ""),
            parent=self,
        )
        if not name:
            return
        name = name.strip()
        if not name:
            messagebox.showinfo(self.tr("win_task_edit"), self.tr("msg_task_input_name"))
            return
        source = task.get("source_folder", "")
        target = task.get("target_folder", "")
        if messagebox.askyesno(
            self.tr("win_task_edit"),
            self.tr("msg_task_change_source"),
            parent=self,
        ):
            picked = filedialog.askdirectory(title=self.tr("msg_task_pick_source"))
            if picked:
                source = picked
        if messagebox.askyesno(
            self.tr("win_task_edit"),
            self.tr("msg_task_change_target"),
            parent=self,
        ):
            picked = filedialog.askdirectory(title=self.tr("msg_task_pick_target"))
            if picked:
                target = picked
        run_incremental = messagebox.askyesno(
            self.tr("win_task_edit"),
            self.tr("msg_task_incremental_prompt"),
            parent=self,
        )
        project_cfg = self._load_config_for_write()
        task["name"] = name
        task["source_folder"] = source
        task["source_folders"] = [source]
        task["target_folder"] = target
        task["run_incremental"] = bool(run_incremental)
        task["config_overrides"] = self._build_task_overrides_from_ui(
            project_cfg=project_cfg, only_diff=True
        )
        try:
            self.task_store.save_task(task)
        except Exception as e:
            messagebox.showerror(self.tr("win_task_edit"), str(e))
            return
        self._refresh_task_list_ui()

    def _on_click_task_delete(self):
        task_id = self._get_selected_task_id()
        if not task_id:
            messagebox.showinfo(
                self.tr("grp_task_runtime"), self.tr("msg_task_select_required")
            )
            return
        task = self.task_store.get_task(task_id) or {}
        if not messagebox.askyesno(
            self.tr("btn_task_delete"),
            self.tr("msg_task_delete_confirm").format(task.get("name", task_id)),
            parent=self,
        ):
            return
        self.task_store.delete_task(task_id)
        self._refresh_task_list_ui()

    def _on_click_task_load_to_ui(self):
        """Load selected task's effective config into main UI (design 7.1: use task as template in classic mode)."""
        task_id = self._get_selected_task_id()
        if not task_id:
            messagebox.showinfo(
                self.tr("grp_task_runtime"), self.tr("msg_task_select_required")
            )
            return
        task = self.task_store.get_task(task_id)
        if not task:
            return
        try:
            cfg = build_task_runtime_config(self._load_config_for_write(), task, force_full_rebuild=False)
        except Exception as e:
            messagebox.showerror(self.tr("btn_task_load_to_ui"), str(e))
            return
        self._suspend_cfg_dirty = True
        try:
            src_list = cfg.get("source_folders") or []
            if not src_list and cfg.get("source_folder"):
                src_list = [cfg.get("source_folder")]
            self.source_folders_list = list(src_list)
            self.lst_source_folders.delete(0, END)
            for p in self.source_folders_list:
                self.lst_source_folders.insert(END, p)
            self.var_source_folder.set(self.source_folders_list[0] if self.source_folders_list else "")
            self.var_target_folder.set(cfg.get("target_folder") or "")
            self.var_run_mode.set(cfg.get("run_mode", MODE_CONVERT_THEN_MERGE))
            self.var_output_enable_pdf.set(1 if cfg.get("output_enable_pdf", True) else 0)
            self.var_output_enable_md.set(1 if cfg.get("output_enable_md", True) else 0)
            self.var_output_enable_merged.set(1 if cfg.get("output_enable_merged", True) else 0)
            self.var_output_enable_independent.set(1 if cfg.get("output_enable_independent", False) else 0)
            self.var_merge_mode.set(cfg.get("merge_mode", MERGE_MODE_CATEGORY))
            self.var_max_merge_size_mb.set(str(cfg.get("max_merge_size_mb", 80)))
            self.var_merge_filename_pattern.set(
                cfg.get("merge_filename_pattern") or "Merged_{category}_{timestamp}_{idx}"
            )
        finally:
            self._suspend_cfg_dirty = False
        messagebox.showinfo(self.tr("btn_task_load_to_ui"), self.tr("msg_task_load_to_ui_done"), parent=self)

    def _apply_runtime_cfg_to_converter(self, converter, runtime_cfg):
        converter.config = dict(runtime_cfg)
        converter._apply_config_defaults()
        converter._init_paths_from_config()
        converter.run_mode = converter.config.get("run_mode", converter.run_mode)
        converter.collect_mode = converter.config.get("collect_mode", converter.collect_mode)
        converter.content_strategy = converter.config.get(
            "content_strategy", converter.content_strategy
        )
        converter.merge_mode = converter.config.get("merge_mode", converter.merge_mode)
        converter.enable_merge_index = bool(
            converter.config.get("enable_merge_index", converter.enable_merge_index)
        )
        converter.enable_merge_excel = bool(
            converter.config.get("enable_merge_excel", converter.enable_merge_excel)
        )
        converter.engine_type = converter.config.get("default_engine", self.var_engine.get())

    def _on_click_task_run(self, resume=False):
        if self.worker_thread and self.worker_thread.is_alive():
            messagebox.showinfo(self.tr("btn_start"), self.tr("msg_task_already_running"))
            return
        task_id = self._get_selected_task_id()
        if not task_id:
            messagebox.showinfo(
                self.tr("grp_task_runtime"), self.tr("msg_task_select_required")
            )
            return
        task = self.task_store.get_task(task_id)
        if not task:
            return

        checkpoint = self.task_store.load_checkpoint(task_id)
        resume_list = None
        if resume:
            if not checkpoint:
                messagebox.showinfo(
                    self.tr("btn_task_resume"), self.tr("msg_task_resume_not_ready")
                )
                return
            planned = checkpoint.get("planned_files", []) or []
            completed = set(checkpoint.get("completed_files", []) or [])
            resume_list = [p for p in planned if p not in completed]
            if not resume_list:
                self.task_store.clear_checkpoint(task_id)
                self.task_store.update_task_runtime(task_id, status="idle")
                self._refresh_task_list_ui()
                messagebox.showinfo(
                    self.tr("btn_task_resume"), self.tr("msg_task_resume_empty")
                )
                return

        project_cfg = self._load_config_for_write()
        force_full_rebuild = bool(self.var_task_force_full_rebuild.get()) and not resume
        runtime_cfg = build_task_runtime_config(
            project_cfg, task, force_full_rebuild=force_full_rebuild
        )
        base_mode = runtime_cfg.get("run_mode", "")
        coercion_msgs = self._sanitize_runtime_config_for_mode(runtime_cfg, base_mode)
        if coercion_msgs:
            self._log_coercion_summary(coercion_msgs, show_dialog=True)
        if force_full_rebuild:
            remove_task_registry_if_exists(task_id, runtime_cfg.get("target_folder", ""))
            self.task_store.clear_checkpoint(task_id)

        self.stop_requested = False
        self.current_task_id = task_id
        self.current_run_context = "task"
        self.task_store.update_task_runtime(task_id, status="running", last_error="")

        def worker():
            try:
                converter = GUIOfficeConverter(self.config_path)
                converter.progress_callback = self.on_progress_update
                self._apply_runtime_cfg_to_converter(converter, runtime_cfg)

                def on_plan(files):
                    if resume:
                        return
                    cp = create_checkpoint(task_id, files)
                    self.task_store.save_checkpoint(task_id, cp)

                def on_done(record):
                    cp = self.task_store.load_checkpoint(task_id)
                    if not cp:
                        seed = resume_list if resume else []
                        cp = create_checkpoint(task_id, seed)
                    cp = mark_checkpoint_file_done(cp, record.get("source_path", ""))
                    self.task_store.save_checkpoint(task_id, cp)

                converter.file_plan_callback = on_plan
                converter.file_done_callback = on_done
                self.current_converter = converter
                converter.run(resume_file_list=resume_list)

                now = datetime.now().isoformat(timespec="seconds")
                if self.stop_requested:
                    self.task_store.update_task_runtime(task_id, status="paused")
                else:
                    self.task_store.clear_checkpoint(task_id)
                    self.task_store.update_task_runtime(
                        task_id, status="idle", last_run_at=now, last_error=""
                    )
            except Exception as e:
                self.task_store.update_task_runtime(
                    task_id,
                    status="paused" if self.stop_requested else "error",
                    last_error=str(e),
                )
                self.after(
                    0,
                    lambda err=str(e): messagebox.showerror(
                        self.tr("msg_runtime_error_title"),
                        self.tr("msg_runtime_error_body").format(err),
                    ),
                )
            finally:
                self.current_converter = None
                self.current_task_id = None
                self.current_run_context = "manual"
                self.stop_requested = False
                self.after(0, lambda: self._set_running_ui_state(False))
                self.after(0, self._refresh_task_list_ui)

        self.worker_thread = threading.Thread(target=worker, daemon=True)
        self.worker_thread.start()
        self._set_running_ui_state(True)

    def _build_config_tab_content(self, parent):
        self.lbl_cfg_defaults_hint = tb.Label(
            parent,
            text=self.tr("lbl_cfg_defaults_hint"),
            bootstyle="secondary",
        )
        self.lbl_cfg_defaults_hint.pack(anchor="w", pady=(0, 2))
        self.lbl_cfg_dirty_state = tb.Label(
            parent,
            text=self.tr("lbl_cfg_dirty_clean"),
            bootstyle="success",
        )
        self.lbl_cfg_dirty_state.pack(anchor="w", pady=(0, 2))
        dirty_row = tb.Frame(parent)
        dirty_row.pack(fill=X, pady=(0, 4))
        self.frm_cfg_dirty_summary = tb.Frame(dirty_row)
        self.frm_cfg_dirty_summary.pack(side=LEFT, fill=X, expand=YES)
        self.lbl_cfg_dirty_sections = tb.Label(
            self.frm_cfg_dirty_summary,
            text=self.tr("lbl_cfg_dirty_none"),
            bootstyle="secondary",
        )
        self.lbl_cfg_dirty_sections.pack(side=LEFT)
        self.btn_save_cfg_dirty = tb.Button(
            dirty_row,
            text=self.tr("btn_save_cfg_dirty"),
            command=self._save_dirty_config_sections,
            bootstyle="success-outline",
            width=18,
            state="disabled",
        )
        self.btn_save_cfg_dirty.pack(side=RIGHT, padx=(0, 6))
        self._attach_tooltip(self.btn_save_cfg_dirty, "tip_save_config_dirty")
        self._set_config_dirty(False)

        # tab_cfg_* 宸插湪 _build_ui 涓涓哄埆鍚嶏紙鎸囧悜瀵瑰簲鐨勫姛鑳?tab 婊氬姩椤甸潰锛?
        # cfg_tabs 瀛?Notebook 宸茬Щ闄わ紝閰嶇疆鍐呭鐩存帴杩藉姞鍒板搴斿姛鑳?tab
        self._cfg_tab_meta = [
            ("shared", self.tab_run_shared, "grp_shared_runtime"),
            ("convert", self.tab_run_convert, "grp_convert_runtime"),
            ("ai", self.tab_run_mshelp, "grp_mshelp_runtime"),
            ("incremental", self.tab_run_convert, "grp_convert_runtime"),
            ("merge", self.tab_run_merge, "grp_merge_runtime"),
            ("ui", self.tab_settings, "tab_config_center"),
            ("rules", self.tab_settings, "tab_config_center"),
        ]
        self._update_config_tab_dirty_markers({})

        # 楂樼骇璁剧疆 tab 閲囩敤宸﹀彸鍙屽垪甯冨眬锛?
        # 宸﹀垪锛氬叡浜厤缃紙璺緞 / 杩涚▼ / 鏃ュ織锛? 杞崲瓒呮椂
        # 鍙冲垪锛氬悎骞惰緭鍑?+ MSHelp + UI + 瑙勫垯
        settings_cols = tb.Frame(parent)
        settings_cols.pack(fill=BOTH, expand=YES, pady=(4, 0))

        settings_left = tb.Frame(settings_cols)
        settings_left.grid(row=0, column=0, sticky="nsew", padx=(0, 6))

        settings_right = tb.Frame(settings_cols)
        settings_right.grid(row=0, column=1, sticky="nsew", padx=(6, 0))

        settings_cols.columnconfigure(0, weight=1)
        settings_cols.columnconfigure(1, weight=1)

        # Shared defaults: paths锛堝乏鍒楋級
        lf_cfg_path = tb.Labelframe(settings_left, text=self.tr("sec_paths"), padding=6)
        lf_cfg_path.pack(fill=X, pady=3)
        self._add_section_help(lf_cfg_path, "tip_section_cfg_paths")
        self.var_config_path = tk.StringVar(value=self.config_path)
        self._create_path_row(lf_cfg_path, "lbl_config", self.var_config_path, self.open_config_folder, None)

        # Shared defaults: process strategy锛堝乏鍒楋級
        lf_proc_shared = tb.Labelframe(settings_left, text=self.tr("grp_cfg_shared_process"), padding=6)
        lf_proc_shared.pack(fill=X, pady=3)
        self._add_section_help(lf_proc_shared, "tip_section_cfg_process")
        tb.Label(lf_proc_shared, text=self.tr("lbl_kill_mode"), font=("System", 9)).pack(anchor="w")
        self.var_kill_mode = tk.StringVar(value=KILL_MODE_AUTO)
        frm_kill = tb.Frame(lf_proc_shared)
        frm_kill.pack(fill=X)
        tb.Radiobutton(frm_kill, text=self.tr("rad_auto_kill"), variable=self.var_kill_mode, value=KILL_MODE_AUTO).pack(side=LEFT)
        tb.Radiobutton(frm_kill, text=self.tr("rad_keep_running"), variable=self.var_kill_mode, value=KILL_MODE_KEEP).pack(side=LEFT, padx=10)

        # Shared defaults: log output锛堝乏鍒楋級
        lf_cfg_log = tb.Labelframe(settings_left, text=self.tr("grp_cfg_shared_log"), padding=6)
        lf_cfg_log.pack(fill=X, pady=3)
        tb.Label(lf_cfg_log, text=self.tr("lbl_log_folder"), font=("System", 9)).pack(anchor="w", pady=(0, 0))
        frm_log = tb.Frame(lf_cfg_log)
        frm_log.pack(fill=X)
        self.var_log_folder = tk.StringVar(value="./logs")
        self.ent_log_folder = tb.Entry(frm_log, textvariable=self.var_log_folder)
        self.ent_log_folder.pack(side=LEFT, fill=X, expand=YES)
        self._attach_tooltip(self.ent_log_folder, "tip_input_log_folder")
        self.btn_log_folder = tb.Button(frm_log, text=self.tr("btn_browse"), command=self.browse_log_folder, bootstyle="outline", width=3)
        self.btn_log_folder.pack(side=LEFT, padx=2)
        self._attach_tooltip(self.btn_log_folder, "tip_choose_log")
        (
            frm_cfg_shared_actions,
            self.btn_save_cfg_shared,
            self.btn_reset_cfg_shared,
        ) = self._add_cfg_section_reset_action(settings_left, "shared")

        # Convert defaults锛堝乏鍒楋級
        lf_proc_convert = tb.Labelframe(settings_left, text=self.tr("grp_cfg_convert"), padding=6)
        lf_proc_convert.pack(fill=X, pady=3)
        self._add_section_help(lf_proc_convert, "tip_section_cfg_process")
        frm_time = tb.Frame(lf_proc_convert)
        frm_time.pack(fill=X, pady=5)
        self.var_timeout_seconds = tk.StringVar(value="60")
        tb.Label(frm_time, text=self.tr("lbl_gen_timeout")).grid(row=0, column=0, sticky="e")
        spinbox_cls = getattr(tb, "Spinbox", None)
        if spinbox_cls is None:
            from tkinter import Spinbox as spinbox_cls  # type: ignore
        self.ent_timeout_seconds = spinbox_cls(
            frm_time,
            textvariable=self.var_timeout_seconds,
            width=5,
            from_=1,
            to=9999,
        )
        self.ent_timeout_seconds.grid(row=0, column=1, sticky="w", padx=5)
        self._attach_tooltip(self.ent_timeout_seconds, "tip_input_timeout_seconds")
        self.var_pdf_wait_seconds = tk.StringVar(value="15")
        tb.Label(frm_time, text=self.tr("lbl_pdf_wait")).grid(row=0, column=2, sticky="e")
        self.ent_pdf_wait_seconds = spinbox_cls(
            frm_time,
            textvariable=self.var_pdf_wait_seconds,
            width=5,
            from_=1,
            to=9999,
        )
        self.ent_pdf_wait_seconds.grid(row=0, column=3, sticky="w", padx=5)
        self._attach_tooltip(self.ent_pdf_wait_seconds, "tip_input_pdf_wait_seconds")
        self.var_ppt_timeout_seconds = tk.StringVar(value="180")
        tb.Label(frm_time, text=self.tr("lbl_ppt_timeout")).grid(row=1, column=0, sticky="e")
        self.ent_ppt_timeout_seconds = spinbox_cls(
            frm_time,
            textvariable=self.var_ppt_timeout_seconds,
            width=5,
            from_=1,
            to=9999,
        )
        self.ent_ppt_timeout_seconds.grid(row=1, column=1, sticky="w", padx=5)
        self._attach_tooltip(self.ent_ppt_timeout_seconds, "tip_input_ppt_timeout_seconds")
        self.var_ppt_pdf_wait_seconds = tk.StringVar(value="30")
        tb.Label(frm_time, text=self.tr("lbl_ppt_wait")).grid(row=1, column=2, sticky="e")
        self.ent_ppt_pdf_wait_seconds = spinbox_cls(
            frm_time,
            textvariable=self.var_ppt_pdf_wait_seconds,
            width=5,
            from_=1,
            to=9999,
        )
        self.ent_ppt_pdf_wait_seconds.grid(row=1, column=3, sticky="w", padx=5)
        self._attach_tooltip(self.ent_ppt_pdf_wait_seconds, "tip_input_ppt_pdf_wait_seconds")
        self.var_office_reuse_app = tk.IntVar(value=1)
        self.chk_office_reuse_app = tb.Checkbutton(
            frm_time,
            text=self.tr("chk_office_reuse_app"),
            variable=self.var_office_reuse_app,
        )
        self.chk_office_reuse_app.grid(row=2, column=0, columnspan=2, sticky="w", pady=(4, 0))
        self._attach_tooltip(self.chk_office_reuse_app, "tip_toggle_office_reuse_app")
        self.var_office_restart_every_n_files = tk.StringVar(value="25")
        tb.Label(frm_time, text=self.tr("lbl_office_restart_every")).grid(
            row=2, column=2, sticky="e", pady=(4, 0)
        )
        self.ent_office_restart_every_n_files = spinbox_cls(
            frm_time,
            textvariable=self.var_office_restart_every_n_files,
            width=5,
            from_=1,
            to=9999,
        )
        self.ent_office_restart_every_n_files.grid(
            row=2, column=3, sticky="w", padx=5, pady=(4, 0)
        )
        self._attach_tooltip(
            self.ent_office_restart_every_n_files,
            "tip_input_office_restart_every_n_files",
        )
        (
            frm_cfg_convert_actions,
            self.btn_save_cfg_convert,
            self.btn_reset_cfg_convert,
        ) = self._add_cfg_section_reset_action(settings_left, "convert")

        # 鍚堝苟琛屼负绫诲紑鍏冲凡鍦ㄨ繍琛岃缃腑缁熶竴鎺у埗锛岃繖閲屼粎淇濈暀閮ㄥ垎杈撳嚭鐩稿叧榛樿鍊硷紝閬垮厤閲嶅鎺т欢锛堝彸鍒楋級
        lf_proc_merge_output = tb.Labelframe(
            settings_right, text=self.tr("grp_cfg_merge_output"), padding=6
        )
        lf_proc_merge_output.pack(fill=X, pady=5)
        tb.Checkbutton(
            lf_proc_merge_output,
            text=self.tr("chk_merge_index"),
            variable=self.var_enable_merge_index,
        ).pack(anchor="w")
        tb.Checkbutton(
            lf_proc_merge_output,
            text=self.tr("chk_merge_excel"),
            variable=self.var_enable_merge_excel,
        ).pack(anchor="w")
        self.var_max_merge_size_mb = tk.StringVar(value="80")
        frm_merge_cfg = tb.Frame(lf_proc_merge_output)
        frm_merge_cfg.pack(fill=X, pady=(4, 0))
        tb.Label(frm_merge_cfg, text=self.tr("lbl_max_mb")).pack(side=LEFT)
        self.ent_max_merge_size_mb = spinbox_cls(
            frm_merge_cfg,
            textvariable=self.var_max_merge_size_mb,
            width=5,
            from_=1,
            to=9999,
        )
        self.ent_max_merge_size_mb.pack(side=LEFT, padx=(5, 0))
        self._attach_tooltip(self.ent_max_merge_size_mb, "tip_input_max_merge_size_mb")
        try:
            self.var_max_merge_size_mb.trace_add("write", lambda *a: self.after(0, self._update_output_summary_label))
        except Exception:
            pass
        tb.Label(lf_proc_merge_output, text=self.tr("lbl_merge_filename_pattern"), font=("System", 9)).pack(anchor="w", pady=(8, 0))
        self.var_merge_filename_pattern = tk.StringVar(value="Merged_{category}_{timestamp}_{idx}")
        self.ent_merge_filename_pattern = tb.Entry(
            lf_proc_merge_output,
            textvariable=self.var_merge_filename_pattern,
            width=45,
        )
        self.ent_merge_filename_pattern.pack(fill=X, pady=(2, 0))
        self._attach_tooltip(self.ent_merge_filename_pattern, "tip_input_merge_filename_pattern")
        (
            frm_cfg_merge_actions,
            self.btn_save_cfg_merge,
            self.btn_reset_cfg_merge,
        ) = self._add_cfg_section_reset_action(settings_right, "merge")

        # MSHelp 閰嶇疆璁剧疆锛堝彸鍒楋級
        lf_cfg_ai_mshelp = tb.Labelframe(
            settings_right, text=self.tr("grp_mshelp_runtime"), padding=6
        )
        lf_cfg_ai_mshelp.pack(fill=X, pady=5)
        tb.Label(
            lf_cfg_ai_mshelp,
            text=self.tr("lbl_mshelp_folder_name"),
            font=("System", 9),
        ).pack(anchor="w")
        self.ent_cfg_mshelpviewer_folder_name = tb.Entry(
            lf_cfg_ai_mshelp, textvariable=self.var_mshelpviewer_folder_name
        )
        self.ent_cfg_mshelpviewer_folder_name.pack(fill=X)
        self._attach_tooltip(
            self.ent_cfg_mshelpviewer_folder_name, "tip_input_mshelp_folder_name"
        )
        tb.Checkbutton(
            lf_cfg_ai_mshelp,
            text=self.tr("chk_mshelp_merge_output"),
            variable=self.var_enable_mshelp_merge_output,
        ).pack(anchor="w", pady=(6, 0))
        tb.Checkbutton(
            lf_cfg_ai_mshelp,
            text=self.tr("chk_mshelp_output_docx"),
            variable=self.var_enable_mshelp_output_docx,
        ).pack(anchor="w", pady=(6, 0))
        tb.Checkbutton(
            lf_cfg_ai_mshelp,
            text=self.tr("chk_mshelp_output_pdf"),
            variable=self.var_enable_mshelp_output_pdf,
        ).pack(anchor="w")
        (
            frm_cfg_ai_actions,
            self.btn_save_cfg_ai,
            self.btn_reset_cfg_ai,
        ) = self._add_cfg_section_reset_action(settings_right, "ai")

        # 澧為噺閰嶇疆鍦ㄨ繍琛屽弬鏁拌浆鎹?tab 涓凡灞曠ず锛屼笉鍐嶅崟鐙崰鐢ㄩ〉闈?

        # UI / tooltip 閰嶇疆锛堝彸鍒楋級
        lf_proc_ui = tb.Labelframe(settings_right, text=self.tr("grp_cfg_ui"), padding=6)
        lf_proc_ui.pack(fill=X, pady=3)
        self._add_section_help(lf_proc_ui, "tip_section_cfg_process")
        tb.Label(lf_proc_ui, text=self.tr("lbl_tooltip_cfg"), font=("System", 9, "bold")).pack(anchor="w")
        # Tooltip 楂樼骇璁剧疆鎶樺彔鍖哄煙
        self.var_show_tooltip_advanced = tk.IntVar(value=0)
        frm_tip_toggle = tb.Frame(lf_proc_ui)
        frm_tip_toggle.pack(fill=X, pady=(4, 0))
        self.chk_show_tooltip_advanced = tb.Checkbutton(
            frm_tip_toggle,
            text=self.tr("chk_show_tooltip_advanced"),
            variable=self.var_show_tooltip_advanced,
            command=self._toggle_tooltip_advanced,
        )
        self.chk_show_tooltip_advanced.pack(anchor="w")
        self._attach_tooltip(
            self.chk_show_tooltip_advanced, "tip_toggle_show_tooltip_advanced"
        )
        frm_tip = tb.Frame(lf_proc_ui)
        self.frm_tooltip_advanced = frm_tip
        self.var_tooltip_auto_theme = tk.IntVar(value=1)
        self.chk_tooltip_auto_theme = tb.Checkbutton(frm_tip, text=self.tr("chk_tooltip_auto_theme"), variable=self.var_tooltip_auto_theme)
        self.chk_tooltip_auto_theme.grid(row=0, column=0, sticky="w", padx=(0, 8))
        self._attach_tooltip(self.chk_tooltip_auto_theme, "tip_toggle_tooltip_auto_theme")
        self.var_confirm_revert_dirty = tk.IntVar(value=1)
        self.chk_confirm_revert_dirty = tb.Checkbutton(
            frm_tip,
            text=self.tr("chk_confirm_revert_dirty"),
            variable=self.var_confirm_revert_dirty,
        )
        self.chk_confirm_revert_dirty.grid(row=0, column=5, sticky="w", padx=(8, 0))
        self._attach_tooltip(
            self.chk_confirm_revert_dirty, "tip_toggle_confirm_revert_dirty"
        )
        self.var_tooltip_delay_ms = tk.StringVar(value="300")
        tb.Label(frm_tip, text=self.tr("lbl_tooltip_delay")).grid(row=0, column=1, sticky="e")
        # tooltip 鏁板€艰緭鍏ヤ篃閲囩敤 Spinbox
        self.ent_tooltip_delay = spinbox_cls(
            frm_tip,
            textvariable=self.var_tooltip_delay_ms,
            width=6,
            from_=0,
            to=10000,
        )
        self.ent_tooltip_delay.grid(row=0, column=2, sticky="w", padx=4)
        self._attach_tooltip(self.ent_tooltip_delay, "tip_input_tooltip_delay_ms")
        self.var_tooltip_font_size = tk.StringVar(value="9")
        tb.Label(frm_tip, text=self.tr("lbl_tooltip_font_size")).grid(row=0, column=3, sticky="e")
        self.ent_tooltip_font_size = spinbox_cls(
            frm_tip,
            textvariable=self.var_tooltip_font_size,
            width=6,
            from_=6,
            to=48,
        )
        self.ent_tooltip_font_size.grid(row=0, column=4, sticky="w", padx=4)
        self._attach_tooltip(
            self.ent_tooltip_font_size, "tip_input_tooltip_font_size"
        )
        self.var_tooltip_bg = tk.StringVar(value="#FFF7D6")
        tb.Label(frm_tip, text=self.tr("lbl_tooltip_bg")).grid(row=1, column=1, sticky="e")
        self.ent_tooltip_bg = tb.Entry(frm_tip, textvariable=self.var_tooltip_bg, width=10)
        self.ent_tooltip_bg.grid(row=1, column=2, sticky="w", padx=4)
        self._attach_tooltip(self.ent_tooltip_bg, "tip_input_tooltip_bg")
        self.btn_pick_tooltip_bg = tb.Button(frm_tip, text="...", width=3, command=lambda: self.pick_tooltip_color("bg"))
        self.btn_pick_tooltip_bg.grid(row=1, column=2, sticky="e", padx=(0, 0))
        self._attach_tooltip(self.btn_pick_tooltip_bg, "tip_pick_color")
        self.var_tooltip_fg = tk.StringVar(value="#202124")
        tb.Label(frm_tip, text=self.tr("lbl_tooltip_fg")).grid(row=1, column=3, sticky="e")
        self.ent_tooltip_fg = tb.Entry(frm_tip, textvariable=self.var_tooltip_fg, width=10)
        self.ent_tooltip_fg.grid(row=1, column=4, sticky="w", padx=4)
        self._attach_tooltip(self.ent_tooltip_fg, "tip_input_tooltip_fg")
        self.btn_pick_tooltip_fg = tb.Button(frm_tip, text="...", width=3, command=lambda: self.pick_tooltip_color("fg"))
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
        (
            frm_cfg_ui_actions,
            self.btn_save_cfg_ui,
            self.btn_reset_cfg_ui,
        ) = self._add_cfg_section_reset_action(settings_right, "ui")

        # Rules defaults: excluded folders锛堝彸鍒楋級
        lf_rules_excluded = tb.Labelframe(
            settings_right, text=self.tr("grp_cfg_rules_excluded"), padding=6
        )
        lf_rules_excluded.pack(fill=X, pady=5)
        self._add_section_help(lf_rules_excluded, "tip_section_cfg_lists")
        tb.Label(lf_rules_excluded, text=self.tr("lbl_excluded")).pack(anchor="w")
        self.txt_excluded_folders = ScrolledText(
            lf_rules_excluded, height=4, font=("Consolas", 8), bootstyle="default"
        )
        self.txt_excluded_folders.pack(fill=X, pady=(0, 5))

        # Rules defaults: keyword strategy锛堝彸鍒楋級
        lf_rules_keywords = tb.Labelframe(
            settings_right, text=self.tr("grp_cfg_rules_keywords"), padding=6
        )
        lf_rules_keywords.pack(fill=X, pady=5)
        tb.Label(lf_rules_keywords, text=self.tr("lbl_keywords")).pack(anchor="w")
        self.txt_price_keywords = ScrolledText(
            lf_rules_keywords, height=3, font=("Consolas", 8), bootstyle="default"
        )
        self.txt_price_keywords.pack(fill=X)
        (
            frm_cfg_rules_actions,
            self.btn_save_cfg_rules,
            self.btn_reset_cfg_rules,
        ) = self._add_cfg_section_reset_action(settings_right, "rules")

        # Emphasized save in config tab锛堝簳閮ㄦí鍚戞寜閽尯锛岃法涓ゅ垪锛?
        cfg_actions = tb.Frame(parent)
        cfg_actions.pack(fill=X, pady=(8, 12))
        self.btn_save_cfg_tab = tb.Button(
            cfg_actions,
            text=self.tr("btn_save_cfg"),
            command=self.open_save_profile_dialog,
            bootstyle="success",
            width=14,
        )
        self.btn_save_cfg_tab.pack(side=LEFT)
        self._attach_tooltip(self.btn_save_cfg_tab, "tip_save_config")
        self.btn_load_cfg_tab = tb.Button(
            cfg_actions,
            text=self.tr("btn_load_cfg"),
            command=self.open_load_profile_dialog,
            bootstyle="secondary-outline",
            width=14,
        )
        self.btn_load_cfg_tab.pack(side=LEFT, padx=(8, 0))
        self._attach_tooltip(self.btn_load_cfg_tab, "tip_load_config")
        self.btn_manage_profiles = tb.Button(
            cfg_actions,
            text=self.tr("btn_manage_profiles"),
            command=self.open_profile_manager_window,
            bootstyle="info-outline",
            width=16,
        )
        self.btn_manage_profiles.pack(side=LEFT, padx=(8, 0))
        self._attach_tooltip(self.btn_manage_profiles, "tip_manage_profiles")
        self._auto_attach_action_tooltips(lf_cfg_path)
        self._auto_attach_action_tooltips(lf_proc_shared)
        self._auto_attach_action_tooltips(lf_cfg_log)
        self._auto_attach_action_tooltips(lf_proc_convert)
        self._auto_attach_action_tooltips(lf_cfg_ai_mshelp)
        self._auto_attach_action_tooltips(lf_proc_merge_output)
        self._auto_attach_action_tooltips(lf_proc_ui)
        self._auto_attach_action_tooltips(lf_rules_excluded)
        self._auto_attach_action_tooltips(lf_rules_keywords)
        self._auto_attach_action_tooltips(frm_cfg_shared_actions)
        self._auto_attach_action_tooltips(frm_cfg_convert_actions)
        self._auto_attach_action_tooltips(frm_cfg_ai_actions)
        self._auto_attach_action_tooltips(frm_cfg_merge_actions)
        self._auto_attach_action_tooltips(frm_cfg_ui_actions)
        self._auto_attach_action_tooltips(frm_cfg_rules_actions)
        self._auto_attach_action_tooltips(cfg_actions)
        self._auto_attach_input_tooltips(lf_cfg_path, "tip_section_cfg_paths")
        self._auto_attach_input_tooltips(lf_proc_shared, "tip_section_cfg_process")
        self._auto_attach_input_tooltips(lf_cfg_log, "tip_section_cfg_process")
        self._auto_attach_input_tooltips(lf_proc_convert, "tip_section_cfg_process")
        self._auto_attach_input_tooltips(lf_proc_merge_output, "tip_section_cfg_process")
        self._auto_attach_input_tooltips(lf_cfg_ai_mshelp, "tip_mode_mshelp")
        self._auto_attach_input_tooltips(lf_proc_ui, "tip_section_cfg_process")
        self._auto_attach_input_tooltips(lf_rules_excluded, "tip_section_cfg_lists")
        self._auto_attach_input_tooltips(lf_rules_keywords, "tip_section_cfg_lists")
        self._attach_tooltip(self.txt_excluded_folders, "tip_input_excluded_folders")
        self._attach_tooltip(self.txt_price_keywords, "tip_input_price_keywords")
        self._bind_var_validation(self.var_timeout_seconds, lambda: self._normalize_then_validate(self.var_timeout_seconds, self._normalize_numeric_var, "config"))
        self._bind_var_validation(self.var_pdf_wait_seconds, lambda: self._normalize_then_validate(self.var_pdf_wait_seconds, self._normalize_numeric_var, "config"))
        self._bind_var_validation(self.var_ppt_timeout_seconds, lambda: self._normalize_then_validate(self.var_ppt_timeout_seconds, self._normalize_numeric_var, "config"))
        self._bind_var_validation(self.var_ppt_pdf_wait_seconds, lambda: self._normalize_then_validate(self.var_ppt_pdf_wait_seconds, self._normalize_numeric_var, "config"))
        self._bind_var_validation(self.var_office_restart_every_n_files, lambda: self._normalize_then_validate(self.var_office_restart_every_n_files, self._normalize_numeric_var, "config"))
        self._bind_var_validation(self.var_max_merge_size_mb, lambda: self._normalize_then_validate(self.var_max_merge_size_mb, self._normalize_numeric_var, "config"))
        self._bind_config_dirty_text(self.txt_excluded_folders)
        self._bind_config_dirty_text(self.txt_price_keywords)
        for _dirty_var in (
            self.var_kill_mode,
            self.var_log_folder,
            self.var_timeout_seconds,
            self.var_pdf_wait_seconds,
            self.var_ppt_timeout_seconds,
            self.var_ppt_pdf_wait_seconds,
            self.var_office_reuse_app,
            self.var_office_restart_every_n_files,
            self.var_enable_merge,
            self.var_output_enable_pdf,
            self.var_output_enable_md,
            self.var_output_enable_merged,
            self.var_output_enable_independent,
            self.var_merge_convert_submode,
            self.var_merge_mode,
            self.var_merge_source,
            self.var_enable_merge_index,
            self.var_enable_merge_excel,
            self.var_max_merge_size_mb,
            self.var_enable_corpus_manifest,
            self.var_enable_markdown,
            self.var_markdown_strip_header_footer,
            self.var_markdown_structured_headings,
            self.var_enable_markdown_quality_report,
            self.var_enable_excel_json,
            self.var_enable_chromadb_export,
            self.var_mshelpviewer_folder_name,
            self.var_enable_mshelp_merge_output,
            self.var_enable_mshelp_output_docx,
            self.var_enable_mshelp_output_pdf,
            self.var_enable_incremental_mode,
            self.var_incremental_verify_hash,
            self.var_incremental_reprocess_renamed,
            self.var_source_priority_skip_same_name_pdf,
            self.var_global_md5_dedup,
            self.var_enable_update_package,
            self.var_tooltip_auto_theme,
            self.var_confirm_revert_dirty,
            self.var_tooltip_delay_ms,
            self.var_tooltip_font_size,
            self.var_tooltip_bg,
            self.var_tooltip_fg,
        ):
            self._bind_config_dirty_var(_dirty_var)

    def _build_footer(self, parent):
        """Build footer actions and status widgets."""
        parent.columnconfigure(1, weight=1)  # Spacer between status and buttons

        # Status Label (Left)
        if not hasattr(self, "var_status"):
            self.var_status = tk.StringVar(value=self.tr("status_ready"))

        tb.Label(parent, textvariable=self.var_status, bootstyle="secondary").grid(
            row=0, column=0, padx=10, sticky="w"
        )

        # 涓儴鎸夐挳缁勶細淇濆瓨 / 鍔犺浇 / 鏃ュ織鍒囨崲锛堝叏灞€鍙锛?
        mid_btn_frame = tb.Frame(parent)
        mid_btn_frame.grid(row=0, column=2, padx=5)

        self.btn_save_cfg = tb.Button(
            mid_btn_frame,
            text=self.tr("btn_save_cfg"),
            command=self.open_save_profile_dialog,
            bootstyle="success-outline",
        )
        self.btn_save_cfg.pack(side=LEFT, padx=(0, 4))
        self._attach_tooltip(self.btn_save_cfg, "tip_save_config")

        self.btn_load_cfg = tb.Button(
            mid_btn_frame,
            text=self.tr("btn_load_cfg"),
            command=self.open_load_profile_dialog,
            bootstyle="secondary-outline",
        )
        self.btn_load_cfg.pack(side=LEFT, padx=(0, 4))
        self._attach_tooltip(self.btn_load_cfg, "tip_load_config")

        self.btn_toggle_logs = tb.Button(
            mid_btn_frame,
            text=self.tr("btn_toggle_logs"),
            command=self._toggle_logs,
            bootstyle="info-outline",
        )
        self.btn_toggle_logs.pack(side=LEFT)
        self._attach_tooltip(self.btn_toggle_logs, "tip_toggle_logs")

        # Start 鈥?瑙嗚鏈€绐佸嚭锛屽姞瀹?
        self.btn_start = tb.Button(
            parent,
            text=self.tr("btn_start"),
            command=self._on_click_start,
            bootstyle="success",
            width=24,
        )
        self.btn_start.grid(row=0, column=3, padx=(10, 5))
        self._attach_tooltip(self.btn_start, "tip_start_task")

        # Stop
        self.btn_stop = tb.Button(
            parent,
            text=self.tr("btn_stop"),
            command=self._on_click_stop,
            bootstyle="danger-outline",
            state="disabled",
        )
        self.btn_stop.grid(row=0, column=4, padx=5)
        self._attach_tooltip(self.btn_stop, "tip_stop_task")
        self._auto_attach_action_tooltips(parent)

        # 娉細淇濆瓨/鍔犺浇/绠＄悊鏂规鎸夐挳宸茬Щ鑷抽厤缃鐞?tab 搴曢儴锛屼笉鍐嶅湪 Footer 閲嶅鏄剧ず

    def _toggle_logs(self):
        """Toggle log pane visibility (pack/pack_forget)."""
        if self._log_visible:
            self.log_pane.pack_forget()
            self._log_visible = False
        else:
            self.log_pane.pack(side=BOTTOM, fill=X, before=self.config_container)
            self._log_visible = True

    # ===================== UI state sync (Adapt for new structure) =====================

    def _set_widget_tree_state(self, root, state):
        for child in root.winfo_children():
            try:
                child.configure(state=state)
            except Exception:
                pass
            self._set_widget_tree_state(child, state)

    def _set_disabled_reason_in_tree(self, root, reason_str):
        """Set _tooltip_disabled_reason on all descendants that have tooltip key/text (for gray-reason tooltip)."""
        if reason_str is None:
            return
        for child in root.winfo_children():
            if getattr(child, "_tooltip_key", None) is not None or getattr(child, "_tooltip_text", None) is not None:
                try:
                    if str(child.cget("state")) == "disabled":
                        setattr(child, "_tooltip_disabled_reason", reason_str)
                except Exception:
                    pass
            self._set_disabled_reason_in_tree(child, reason_str)

    def _clear_disabled_reason_in_tree(self, root):
        """Clear _tooltip_disabled_reason on all descendants."""
        for child in root.winfo_children():
            if hasattr(child, "_tooltip_disabled_reason"):
                setattr(child, "_tooltip_disabled_reason", None)
            self._clear_disabled_reason_in_tree(child)

    def _set_run_tab_state(self, tab, state):
        """Set run tab visibility state. Supports ttkbootstrap (hide/restore) and standard ttk (tab state)."""
        try:
            if state in ("disabled", "hidden"):
                if hasattr(self.main_notebook, "hide"):
                    self.main_notebook.hide(tab)
                else:
                    # 标准 ttk.Notebook 无 hide，用 tab(state="hidden")
                    try:
                        self.main_notebook.tab(tab, state="hidden")
                    except Exception:
                        pass
            else:
                if hasattr(self.main_notebook, "hide"):
                    try:
                        self.main_notebook.index(tab)
                        # tab 仍在 notebook 中（如 ttkbootstrap 用 state 隐藏）：显式设为 normal
                        try:
                            self.main_notebook.tab(tab, state="normal")
                        except Exception:
                            pass
                    except Exception:
                        self._restore_tab(tab)
                else:
                    try:
                        # 标准 ttk：先直接用 widget 设 state="normal"，再尝试用 tab_id
                        try:
                            self.main_notebook.tab(tab, state="normal")
                        except Exception:
                            if hasattr(self, "_all_tabs") and tab in self._all_tabs:
                                idx = self._all_tabs.index(tab)
                                tab_list = list(self.main_notebook.tabs())
                                if idx < len(tab_list):
                                    self.main_notebook.tab(tab_list[idx], state="normal")
                            else:
                                for tab_id in self.main_notebook.tabs():
                                    try:
                                        if self.main_notebook.nametowidget(tab_id) is tab:
                                            self.main_notebook.tab(tab_id, state="normal")
                                            break
                                    except Exception:
                                        continue
                    except Exception:
                        pass
        except Exception:
            pass

    def _set_task_tab_highlight(self, on):
        """任务模式下把「任务管理」标签页设为绿色（前面加 🟢），传统模式下恢复默认。"""
        if not hasattr(self, "main_notebook") or not hasattr(self, "tab_run_tasks"):
            return
        try:
            base_text = self.tr("grp_task_runtime")
            new_text = f"🟢 {base_text}" if on else base_text
            try:
                self.main_notebook.tab(self.tab_run_tasks, text=new_text)
            except Exception:
                for tab_id in self.main_notebook.tabs():
                    try:
                        if self.main_notebook.nametowidget(tab_id) is self.tab_run_tasks:
                            self.main_notebook.tab(tab_id, text=new_text)
                            break
                    except Exception:
                        continue
        except Exception:
            pass

    def _restore_tab(self, tab):
        """Restore a hidden tab to its original index, with correct tab text."""
        if not hasattr(self, "_all_tabs"):
            return
        target_idx = 0
        tab_text = None
        for i, t in enumerate(self._all_tabs):
            if t is tab:
                if hasattr(self, "_all_tab_text_keys") and i < len(self._all_tab_text_keys):
                    key = self._all_tab_text_keys[i]
                    if isinstance(key, tuple):
                        tab_text = f"{self.tr(key[0])} / {self.tr(key[1])}"
                    else:
                        tab_text = self.tr(key)
                break
            try:
                self.main_notebook.index(t)
                target_idx += 1
            except Exception:
                pass
        if tab_text is not None:
            self.main_notebook.insert(target_idx, tab, text=tab_text)
        else:
            self.main_notebook.insert(target_idx, tab)

    def _set_cfg_tab_state(self, tab, state):
        """Backward-compatible wrapper for config tab visibility."""
        self._set_run_tab_state(tab, state)

    def _on_run_mode_change(self):
        mode = self.var_run_mode.get()
        is_collect = mode == MODE_COLLECT_ONLY
        is_mshelp = mode == MODE_MSHELP_ONLY
        is_convert = mode in (MODE_CONVERT_ONLY, MODE_CONVERT_THEN_MERGE)
        is_merge_related = mode in (MODE_CONVERT_THEN_MERGE, MODE_MERGE_ONLY)
        is_rules_related = is_convert or is_collect
        if mode == MODE_CONVERT_THEN_MERGE:
            self.var_merge_source.set("target")

        # Enable only tabs relevant to the current mode锛堝悎骞?姊崇悊銆丮SHelp/瀹氫綅 涓鸿仛鍚?tab锛?
        self._set_run_tab_state(self.tab_run_shared, "normal")
        self._set_run_tab_state(self.tab_run_convert, "normal" if is_convert else "disabled")
        # 銆屽悎骞?/ 姊崇悊銆嶏細浠绘剰鍚堝苟鐩稿叧鎴栨⒊鐞嗘ā寮忔椂鏄剧ず
        self._set_run_tab_state(
            self.tab_run_merge,
            "normal" if (is_merge_related or is_collect) else "disabled",
        )
        # MSHelp锛氫粎鍦?mshelp 妯″紡鏄剧ず
        self._set_run_tab_state(self.tab_run_mshelp, "normal" if is_mshelp else "disabled")
        # 瀹氫綅宸ュ叿锛氬缁堝彲鐢?
        self._set_run_tab_state(self.tab_run_locator, "normal")
        # 鎴愭灉鏂囦欢锛氬缁堝彲鐢?
        self._set_run_tab_state(self.tab_run_output, "normal")
        # 任务管理页：传统模式下隐藏，任务模式下可见（由 _update_task_tab_for_app_mode 统一决定）
        if hasattr(self, "_update_task_tab_for_app_mode"):
            self._update_task_tab_for_app_mode()

        # Collect options
        if is_collect:
            self._set_widget_tree_state(self.frm_collect_opts, "normal")
        else:
            self._set_widget_tree_state(self.frm_collect_opts, "disabled")

        # 鑷姩閫変腑瀵瑰簲鐨勯《灞?tab锛堝悎骞朵簡姊崇悊銆丮SHelp/瀹氫綅锛?
        try:
            if is_collect:
                # 姊崇悊妯″紡浣跨敤銆屽悎骞?/ 姊崇悊銆嶈仛鍚堥〉
                self.main_notebook.select(self.tab_run_merge)
            elif is_mshelp:
                self.main_notebook.select(self.tab_run_mshelp)
            elif mode == MODE_MERGE_ONLY:
                self.main_notebook.select(self.tab_run_merge)
            else:
                self.main_notebook.select(self.tab_run_convert)
        except Exception:
            pass

        # Engine & Sandbox (Enable only if converting)
        state_exec = "normal" if is_convert else "disabled"
        self._set_widget_tree_state(self.group_exec, state_exec)

        # Convert-only strategy controls
        try:
            self.lbl_strategy.configure(state=state_exec)
        except Exception:
            pass
        try:
            self.cb_strat.configure(state="readonly" if is_convert else "disabled")
        except Exception:
            pass
        try:
            self.chk_date_filter.configure(state=state_exec)
        except Exception:
            pass
        try:
            self._sync_markdown_master_with_global_output()
        except Exception:
            pass
        try:
            self.chk_markdown_strip_header_footer.configure(state=state_exec)
        except Exception:
            pass
        try:
            self.chk_markdown_structured_headings.configure(state=state_exec)
        except Exception:
            pass
        try:
            self.chk_markdown_quality_report.configure(state=state_exec)
        except Exception:
            pass
        try:
            self.chk_export_records_json.configure(state=state_exec)
        except Exception:
            pass
        try:
            self.chk_chromadb_export.configure(state=state_exec)
        except Exception:
            pass
        self._on_toggle_incremental_mode()

        # Trigger sandbox toggle to refresh sub-widgets
        self._on_toggle_sandbox()

        if not is_convert:
            for child in self.frm_date.winfo_children():
                try: child.configure(state="disabled")
                except: pass
            try:
                self.ent_date.configure(state="disabled")
            except Exception:
                pass
        else:
            self._on_toggle_date_filter()

        # Merge Options
        state_merge = "normal" if is_merge_related else "disabled"
        self.lbl_merge.configure(state=state_merge)
        self.chk_enable_merge.configure(state=state_merge)
        self._set_widget_tree_state(self.frm_merge_opts, state_merge)
        merge_submode_state = "normal" if mode == MODE_MERGE_ONLY else "disabled"
        try:
            self.rb_merge_submode_merge.configure(state=merge_submode_state)
            self.rb_merge_submode_pdf_to_md.configure(state=merge_submode_state)
        except Exception:
            pass
        # Merge sub-controls: when merged output off, gray merge opts; when all_in_one, gray max_mb only
        if is_merge_related:
            if not bool(self.var_output_enable_merged.get()):
                self._set_widget_tree_state(self.frm_merge_opts, "disabled")
            else:
                self._set_widget_tree_state(self.frm_merge_opts, "normal")
                try:
                    if self.var_merge_mode.get() == MERGE_MODE_ALL_IN_ONE:
                        self.ent_max_merge_size_mb.configure(state="disabled")
                    else:
                        self.ent_max_merge_size_mb.configure(state="normal")
                except Exception:
                    pass
            # Convert+Merge forces merge_source=target: make merge source radios read-only
            try:
                if mode == MODE_CONVERT_THEN_MERGE:
                    self._set_widget_tree_state(self.frm_merge_src, "disabled")
                else:
                    self._set_widget_tree_state(self.frm_merge_src, "normal")
            except Exception:
                pass
        self._apply_disabled_reason_tooltips()
        self._update_output_summary_label()

    def _apply_disabled_reason_tooltips(self):
        """Set or clear _tooltip_disabled_reason on mode-sensitive widgets so gray controls show reason."""
        mode = self.var_run_mode.get()
        mode_labels = {
            MODE_CONVERT_ONLY: self.tr("mode_convert"),
            MODE_MERGE_ONLY: self.tr("mode_merge"),
            MODE_CONVERT_THEN_MERGE: self.tr("mode_convert_merge"),
            MODE_COLLECT_ONLY: self.tr("mode_collect"),
            MODE_MSHELP_ONLY: self.tr("mode_mshelp"),
        }
        mode_label = mode_labels.get(mode, mode)
        reason_by_mode = self.tr("tip_disabled_by_mode").format(mode_label)
        reason_merged_off = self.tr("tip_disabled_merged_off")
        reason_all_in_one = self.tr("tip_disabled_merge_all_in_one")

        def _try_state(w):
            try:
                return str(w.cget("state"))
            except Exception:
                return ""

        is_convert = mode in (MODE_CONVERT_ONLY, MODE_CONVERT_THEN_MERGE)
        is_merge_related = mode in (MODE_CONVERT_THEN_MERGE, MODE_MERGE_ONLY)
        merged_on = bool(self.var_output_enable_merged.get())
        all_in_one = self.var_merge_mode.get() == MERGE_MODE_ALL_IN_ONE

        # Convert-related trees
        for root, reason in [
            (self.group_exec, reason_by_mode if not is_convert else None),
            (self.frm_date, reason_by_mode if not is_convert else None),
        ]:
            if root is None:
                continue
            try:
                if reason and _try_state(root) == "disabled":
                    self._set_disabled_reason_in_tree(root, reason)
                else:
                    self._clear_disabled_reason_in_tree(root)
            except Exception:
                pass

        # Single widgets that are disabled by mode
        for attr in ("lbl_strategy", "cb_strat", "chk_date_filter", "lbl_merge", "chk_enable_merge",
                     "chk_markdown_strip_header_footer", "chk_markdown_structured_headings",
                     "chk_markdown_quality_report", "chk_export_records_json", "chk_chromadb_export",
                     "chk_incremental_mode", "chk_source_priority_skip_pdf", "chk_global_md5_dedup",
                     "chk_enable_update_package", "chk_incremental_verify_hash", "chk_incremental_reprocess_renamed"):
            w = getattr(self, attr, None)
            if w is None:
                continue
            try:
                if _try_state(w) == "disabled":
                    setattr(w, "_tooltip_disabled_reason", reason_by_mode)
                else:
                    setattr(w, "_tooltip_disabled_reason", None)
            except Exception:
                pass

        # Sandbox group
        for attr in ("chk_enable_sandbox", "entry_temp_sandbox_root", "btn_temp_sandbox_root",
                     "entry_sandbox_min_free_gb", "cb_sandbox_low_space_policy"):
            w = getattr(self, attr, None)
            if w is None:
                continue
            try:
                if _try_state(w) == "disabled":
                    setattr(w, "_tooltip_disabled_reason", reason_by_mode)
                else:
                    setattr(w, "_tooltip_disabled_reason", None)
            except Exception:
                pass

        # Merge opts tree: reason by mode, or merged_off, or (for max_mb only) all_in_one
        try:
            self._clear_disabled_reason_in_tree(self.frm_merge_opts)
            if not is_merge_related:
                self._set_disabled_reason_in_tree(self.frm_merge_opts, reason_by_mode)
            elif not merged_on:
                self._set_disabled_reason_in_tree(self.frm_merge_opts, reason_merged_off)
            elif all_in_one and hasattr(self, "ent_max_merge_size_mb"):
                setattr(self.ent_max_merge_size_mb, "_tooltip_disabled_reason",
                        reason_all_in_one if _try_state(self.ent_max_merge_size_mb) == "disabled" else None)
        except Exception:
            pass

        # Merge source frame (disabled when convert_then_merge)
        try:
            self._clear_disabled_reason_in_tree(self.frm_merge_src)
            if mode == MODE_CONVERT_THEN_MERGE:
                self._set_disabled_reason_in_tree(self.frm_merge_src, reason_by_mode)
        except Exception:
            pass

    def _update_output_summary_label(self):
        """Update the fixed output summary line (design 5.6.1) and inline hint bar (design 5.6.2)."""
        if not hasattr(self, "lbl_merge_output_summary"):
            return
        merged = bool(self.var_output_enable_merged.get())
        indep = bool(self.var_output_enable_independent.get())
        if not merged and not indep:
            part = self.tr("summary_merged_off")
        elif merged and not indep:
            part = self.tr("summary_merged_only")
        elif indep and not merged:
            part = self.tr("summary_independent_only")
        else:
            part = self.tr("summary_both")
        if merged:
            mode = getattr(self, "var_merge_mode", None) and self.var_merge_mode.get()
            if mode == MERGE_MODE_ALL_IN_ONE:
                part += " " + self.tr("summary_merge_single")
            else:
                mb = 80
                if hasattr(self, "var_max_merge_size_mb") and self.var_max_merge_size_mb.get():
                    mb = self._safe_positive_int(self.var_max_merge_size_mb.get(), 80)
                part += " " + self.tr("summary_merge_split_mb").format(mb)
        prefix = self.tr("lbl_output_summary")
        self.lbl_merge_output_summary.configure(text=prefix + part)
        # Inline hint bar (5.6.2): yellow / green / gray
        if hasattr(self, "lbl_merge_inline_hint"):
            mode = getattr(self, "var_run_mode", None) and self.var_run_mode.get()
            if mode not in (MODE_CONVERT_THEN_MERGE, MODE_MERGE_ONLY):
                self.lbl_merge_inline_hint.configure(text="", bootstyle="secondary")
            elif not merged:
                self.lbl_merge_inline_hint.configure(
                    text=self.tr("hint_merge_inline_merged_off"), bootstyle="secondary"
                )
            elif getattr(self, "var_merge_mode", None) and self.var_merge_mode.get() == MERGE_MODE_ALL_IN_ONE:
                self.lbl_merge_inline_hint.configure(
                    text=self.tr("hint_merge_inline_all_in_one"), bootstyle="warning"
                )
            else:
                self.lbl_merge_inline_hint.configure(
                    text=self.tr("hint_merge_inline_split"), bootstyle="success"
                )

    def _on_merge_output_or_mode_change(self):
        """Update merge sub-controls state when Merged Output or Merge mode (all_in_one/category) changes."""
        mode = self.var_run_mode.get()
        is_merge_related = mode in (MODE_CONVERT_THEN_MERGE, MODE_MERGE_ONLY)
        if not is_merge_related:
            return
        merged_on = bool(self.var_output_enable_merged.get())
        if not merged_on:
            self._set_widget_tree_state(self.frm_merge_opts, "disabled")
        else:
            self._set_widget_tree_state(self.frm_merge_opts, "normal")
            try:
                if self.var_merge_mode.get() == MERGE_MODE_ALL_IN_ONE:
                    self.ent_max_merge_size_mb.configure(state="disabled")
                else:
                    self.ent_max_merge_size_mb.configure(state="normal")
            except Exception:
                pass
        self._apply_disabled_reason_tooltips()
        self._update_output_summary_label()

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
        is_disabled_globally = mode in (MODE_COLLECT_ONLY, MODE_MERGE_ONLY, MODE_MSHELP_ONLY)
        
        # If group is disabled, sandbox should look disabled
        if is_disabled_globally:
            self.chk_enable_sandbox.configure(state="disabled")
            self.entry_temp_sandbox_root.configure(state="disabled")
            self.btn_temp_sandbox_root.configure(state="disabled")
            self.entry_sandbox_min_free_gb.configure(state="disabled")
            self.cb_sandbox_low_space_policy.configure(state="disabled")
            return

        # Otherwise standard toggle logic
        self.chk_enable_sandbox.configure(state="normal")
        enabled = bool(self.var_enable_sandbox.get())
        state = "normal" if enabled else "disabled"
        self.entry_temp_sandbox_root.configure(state=state)
        self.btn_temp_sandbox_root.configure(state=state)
        self.entry_sandbox_min_free_gb.configure(state=state)
        self.cb_sandbox_low_space_policy.configure(state=state)

    def _on_toggle_incremental_mode(self):
        mode = self.var_run_mode.get()
        is_convert = mode in (MODE_CONVERT_ONLY, MODE_CONVERT_THEN_MERGE)
        master_state = "normal" if is_convert else "disabled"
        verify_state = (
            "normal"
            if is_convert and bool(self.var_enable_incremental_mode.get())
            else "disabled"
        )
        for widget in (
            self.chk_incremental_mode,
            self.chk_source_priority_skip_pdf,
            self.chk_global_md5_dedup,
            self.chk_enable_update_package,
        ):
            try:
                widget.configure(state=master_state)
            except Exception:
                pass
        try:
            self.chk_incremental_verify_hash.configure(state=verify_state)
        except Exception:
            pass
        try:
            self.chk_incremental_reprocess_renamed.configure(state=verify_state)
        except Exception:
            pass

    def _on_toggle_markdown_master(self):
        """Toggle state of markdown sub-options."""
        enabled = bool(self.var_enable_markdown.get())
        state = "normal" if enabled else "disabled"
        if hasattr(self, "_frm_markdown_sub"):
            self._set_widget_tree_state(self._frm_markdown_sub, state)

    def _sync_markdown_master_with_global_output(self):
        """Keep legacy markdown master aligned with global MD output switch."""
        if not hasattr(self, "var_output_enable_md") or not hasattr(
            self, "var_enable_markdown"
        ):
            return
        md_enabled = bool(self.var_output_enable_md.get())
        try:
            self.var_enable_markdown.set(1 if md_enabled else 0)
        except Exception:
            pass
        if hasattr(self, "chk_export_markdown"):
            try:
                self.chk_export_markdown.configure(state="disabled")
            except Exception:
                pass
        self._on_toggle_markdown_master()

    def _on_toggle_llm_hub_master(self):
        """Toggle state of LLM hub sub-options."""
        enabled = bool(self.var_enable_llm_delivery_hub.get())
        state = "normal" if enabled else "disabled"
        if hasattr(self, "_frm_llm_hub_sub"):
            self._set_widget_tree_state(self._frm_llm_hub_sub, state)

    # ===================== 閻╊喖缍?鎸夐挳鍔ㄤ綔 =====================

    def add_source_folder(self):
        """Add source folder with option to select multiple folders."""
        # Show dialog to choose single or multi select
        if not hasattr(self, "_multi_folder_dialog"):
            self._multi_folder_dialog = None

        # Ask user if they want to select multiple folders
        result = messagebox.askyesno(
            self.tr("tip_add_source_folder"),
            self.tr("msg_multi_select_folders"),
            icon="question"
        )

        if result:
            # Multi-select mode
            self._open_multi_folder_dialog()
        else:
            # Single-select mode
            path = filedialog.askdirectory(title=self.tr("tip_add_source_folder"))
            if path:
                if sys.platform == "win32":
                    path = path.replace("/", "\\")
                if path not in self.source_folders_list:
                    self.source_folders_list.append(path)
                    self.lst_source_folders.insert(END, path)
                    self.cfg_dirty = True
            # Sync to hidden var for compatibility
            if self.source_folders_list:
                self.var_source_folder.set(self.source_folders_list[0])

    def _open_multi_folder_dialog(self):
        """Open a custom dialog for selecting multiple folders."""
        if hasattr(self, "_multi_folder_dialog") and self._multi_folder_dialog is not None:
            try:
                if self._multi_folder_dialog.winfo_exists():
                    self._multi_folder_dialog.lift()
                    self._multi_folder_dialog.focus_force()
                    return
            except Exception:
                pass

        dlg = tk.Toplevel(self)
        dlg.title(self.tr("msg_multi_select_title"))
        self._place_dialog_in_main(dlg, 720, 650)
        dlg.minsize(650, 550)
        dlg.transient(self)
        dlg.grab_set()
        dlg.lift()
        dlg.protocol("WM_DELETE_WINDOW", self._close_multi_folder_dialog)
        self._multi_folder_dialog = dlg

        # Main frame
        frame = tb.Frame(dlg, padding=12)
        frame.pack(fill=BOTH, expand=YES)

        # Instructions
        tb.Label(frame, text="方式1：扫描父文件夹（推荐，一次添加多个）", font=("System", 9, "bold"), foreground="#007bff").pack(anchor="w", pady=(0, 5))
        
        # Parent folder selection section
        parent_frame = tb.Frame(frame)
        parent_frame.pack(fill=X, pady=(0, 10))
        
        tb.Label(parent_frame, text="选择父文件夹：", font=("System", 9)).pack(side=LEFT, padx=(0, 8))
        ent_parent = tb.Entry(parent_frame, width=50)
        ent_parent.pack(side=LEFT, fill=X, expand=YES)
        
        def pick_parent():
            p = filedialog.askdirectory(title="选择父文件夹")
            if p:
                if sys.platform == "win32":
                    p = p.replace("/", "\\")
                ent_parent.delete(0, END)
                ent_parent.insert(0, p)
        
        def scan_subfolders():
            parent = ent_parent.get().strip()
            if not parent or not os.path.isdir(parent):
                messagebox.showwarning("提示", "请选择有效的父文件夹！", parent=dlg)
                return
            
            # Clear existing
            folder_tree.delete(*folder_tree.get_children())
            
            # Scan subfolders
            try:
                subfolders = []
                for item in os.listdir(parent):
                    item_path = os.path.join(parent, item)
                    if os.path.isdir(item_path):
                        subfolders.append(item_path)
                
                # Sort folders
                subfolders.sort()
                
                # Add to tree (all checked by default)
                for subfolder in subfolders:
                    item = folder_tree.insert("", "end", values=(subfolder,))
                    folder_tree.item(item, tags=("checked",))
                    folder_tree.set(item, "selected", "1")  # Custom flag
                
                # Update count
                self._update_folder_count(folder_tree)
                
                # Update display
                total = len(subfolders)
                count_label.config(text=f"已扫描 {total} 个子文件夹")
                
                messagebox.showinfo("扫描完成", f"已从父文件夹扫描到 {total} 个子文件夹！", parent=dlg)
                
            except Exception as e:
                messagebox.showerror("错误", f"扫描失败：{str(e)}", parent=dlg)
        
        tb.Button(parent_frame, text="浏览", command=pick_parent, bootstyle="info", width=8).pack(side=LEFT, padx=(8, 4))
        tb.Button(parent_frame, text="扫描子文件夹", command=scan_subfolders, bootstyle="success", width=12).pack(side=LEFT, padx=(4, 0))

        # Manual add section
        tb.Label(frame, text="方式2：手动逐个添加", font=("System", 9, "bold"), foreground="#666").pack(anchor="w", pady=(10, 5))

        # Treeview for folder selection with checkboxes
        left_frame = tb.Frame(frame)
        left_frame.pack(fill=BOTH, expand=YES)

        # Use a custom checkbox listbox approach
        cols = ("selected", "path",)
        folder_tree = ttk.Treeview(left_frame, columns=cols, show="headings", selectmode="extended")
        folder_tree.heading("selected", text="", width=30)
        folder_tree.column("selected", width=30, anchor="center")
        folder_tree.heading("path", text="文件夹路径")
        folder_tree.column("path", width=550, anchor="w")
        folder_tree.pack(side=LEFT, fill=BOTH, expand=YES)

        # Scrollbar for treeview
        tree_scroll = tb.Scrollbar(left_frame, orient=VERTICAL, command=folder_tree.yview)
        folder_tree.configure(yscrollcommand=tree_scroll.set)
        tree_scroll.pack(side=RIGHT, fill=Y)

        # Right: Buttons
        btn_frame = tb.Frame(frame)
        btn_frame.pack(side=RIGHT, fill=Y, padx=(10, 0))

        btn_browse = tb.Button(btn_frame, text="+ 浏览", command=lambda: self._browse_folder_to_dialog(folder_tree), bootstyle="info", width=10)
        btn_browse.pack(pady=2)

        tb.Label(btn_frame, text="", height=1).pack()  # Spacer

        btn_remove = tb.Button(btn_frame, text="- 移除", command=lambda: self._remove_folder_from_tree(folder_tree), bootstyle="danger-outline", width=10)
        btn_remove.pack(pady=2)

        btn_clear = tb.Button(btn_frame, text="C 清空", command=lambda: self._clear_folder_tree(folder_tree), bootstyle="secondary-outline", width=10)
        btn_clear.pack(pady=2)

        # Bottom section: Controls
        bottom_frame = tb.Frame(frame)
        bottom_frame.pack(fill=X, pady=(10, 0))

        # Count display
        self._multi_folder_count_var = tk.StringVar()
        count_label = tb.Label(bottom_frame, textvariable=self._multi_folder_count_var, font=("System", 9))
        count_label.pack(anchor="w")

        # Checkbox controls
        check_frame = tb.Frame(bottom_frame)
        check_frame.pack(anchor="w", pady=(8, 0))

        def select_all():
            for item in folder_tree.get_children():
                folder_tree.item(item, tags=("checked",))

        def deselect_all():
            for item in folder_tree.get_children():
                folder_tree.item(item, tags=())

        def invert_selection():
            for item in folder_tree.get_children():
                tags = folder_tree.item(item, "tags")
                if tags and "checked" in tags:
                    folder_tree.item(item, tags=())
                else:
                    folder_tree.item(item, tags=("checked",))

        tb.Button(check_frame, text="全选", command=select_all, bootstyle="info-outline", width=8).pack(side=LEFT, padx=(0, 4))
        tb.Button(check_frame, text="全不选", command=deselect_all, bootstyle="secondary-outline", width=8).pack(side=LEFT, padx=4)
        tb.Button(check_frame, text="反选", command=invert_selection, bootstyle="secondary-outline", width=8).pack(side=LEFT, padx=4)

        # Initialize
        self._update_folder_count(folder_tree)

        # Action buttons
        action_btn_frame = tb.Frame(frame)
        action_btn_frame.pack(fill=X, pady=(10, 0))

        btn_add = tb.Button(action_btn_frame, text="添加到列表", command=lambda: self._confirm_multi_folder_selection(folder_tree), bootstyle="success", width=12)
        btn_add.pack(side=LEFT)

        btn_cancel = tb.Button(action_btn_frame, text="取消", command=self._close_multi_folder_dialog, bootstyle="secondary-outline", width=10)
        btn_cancel.pack(side=RIGHT)


    def _close_multi_folder_dialog(self):
        """Close the multi-folder selection dialog."""
        if hasattr(self, "_multi_folder_dialog") and self._multi_folder_dialog is not None:
            try:
                if self._multi_folder_dialog.grab_current() == self._multi_folder_dialog:
                    self._multi_folder_dialog.grab_release()
            except Exception:
                pass
            try:
                self._multi_folder_dialog.destroy()
            except Exception:
                pass
        self._multi_folder_dialog = None

    def _browse_folder_to_dialog(self, tree):
        """Browse for a single folder and add to tree."""
        path = filedialog.askdirectory(title=self.tr("tip_add_source_folder"))
        if path:
            if sys.platform == "win32":
                path = path.replace("/", "\\")
            # Check if already exists
            exists = False
            for item in tree.get_children():
                if tree.get(item, "path")[0] == path:
                    exists = True
                    break
            if not exists:
                tree.insert("", "end", values=(path,))
                self._update_folder_count(tree)

    def _browse_multiple_to_dialog(self, tree):
        """Browse for multiple folders using multiple askdirectory calls."""
        # Since askdirectory doesn't support multi-select natively,
        # we'll use a simple loop approach with a counter
        base_dir = filedialog.askdirectory(title=self.tr("tip_add_source_folder") + " (选择第一个文件夹)")
        if not base_dir:
            return

        if sys.platform == "win32":
            base_dir = base_dir.replace("/", "\\")

        # Add first folder
        exists = False
        for item in tree.get_children():
            if tree.get(item, "path")[0] == base_dir:
                exists = True
                break
        if not exists:
            tree.insert("", "end", values=(base_dir,))
        self._update_folder_count(tree)

        # Ask if user wants to add more folders
        while True:
            result = messagebox.askyesno(
                "继续添加",
                "是否要继续添加更多文件夹？",
                icon="question"
            )
            if not result:
                break

            next_dir = filedialog.askdirectory(title=self.tr("tip_add_source_folder") + " (选择下一个文件夹)")
            if not next_dir:
                break

            if sys.platform == "win32":
                next_dir = next_dir.replace("/", "\\")

            # Check if already exists
            exists = False
            for item in tree.get_children():
                if tree.get(item, "path")[0] == next_dir:
                    exists = True
                    break
            if not exists:
                tree.insert("", "end", values=(next_dir,))
            self._update_folder_count(tree)

    def _remove_folder_from_tree(self, tree):
        """Remove selected folders from the tree."""
        selection = tree.selection()
        for item in reversed(selection):
            tree.delete(item)
        self._update_folder_count(tree)

    def _clear_folder_tree(self, tree):
        """Clear all folders from the tree."""
        tree.delete(*tree.get_children())
        self._update_folder_count(tree)

    def _update_folder_count(self, tree):
        """Update the folder count display."""
        count = len(tree.get_children())
        if hasattr(self, "_multi_folder_count_var"):
            if count == 0:
                self._multi_folder_count_var.set("已选择 0 个文件夹")
            elif count == 1:
                self._multi_folder_count_var.set("已选择 1 个文件夹")
            else:
                self._multi_folder_count_var.set(f"已选择 {count} 个文件夹")

    def _confirm_multi_folder_selection(self, tree):
        """Add selected folders to the main source folders list."""
        added_count = 0
        for item in tree.get_children():
            tags = tree.item(item, "tags")
            if tags and "checked" in tags:  # Only add checked items
                path = tree.get(item, "path")[0]
                if path not in self.source_folders_list:
                    self.source_folders_list.append(path)
                    self.lst_source_folders.insert(END, path)
                    added_count += 1
        # Sync to hidden var for compatibility
        if self.source_folders_list:
            self.var_source_folder.set(self.source_folders_list[0])
        if added_count > 0:
            self.cfg_dirty = True
        self._close_multi_folder_dialog()

    def _open_task_multi_folder_dialog(self, parent_win, target_listbox):
        """Open a simplified multi-folder dialog for task wizard."""
        # Check if dialog already exists
        if hasattr(self, "_task_multi_folder_dialog") and self._task_multi_folder_dialog is not None:
            try:
                if self._task_multi_folder_dialog.winfo_exists():
                    self._task_multi_folder_dialog.lift()
                    self._task_multi_folder_dialog.focus_force()
                    return
            except Exception:
                pass

        dlg = tk.Toplevel(parent_win)
        dlg.title(self.tr("msg_multi_select_title"))
        dlg.minsize(620, 600)
        dlg.geometry("660x620")
        dlg.transient(parent_win)
        dlg.grab_set()
        dlg.lift()
        dlg.protocol("WM_DELETE_WINDOW", lambda: self._close_task_multi_folder_dialog(dlg))
        self._task_multi_folder_dialog = dlg

        # Main frame with simple style
        frame = tk.Frame(dlg, padx=12, pady=12)
        frame.pack(fill=BOTH, expand=YES)

        # Instructions
        tk.Label(frame, text="方式1：扫描父文件夹（推荐，一次添加多个）", font=("System", 9, "bold"), fg="blue").pack(anchor="w", pady=(0, 5))
        
        # Parent folder selection section
        parent_frame = tk.Frame(frame)
        parent_frame.pack(fill=X, pady=(0, 8))
        
        tk.Label(parent_frame, text="选择父文件夹：", font=("System", 9)).pack(side=LEFT, padx=(0, 8))
        ent_parent = tk.Entry(parent_frame, width=50)
        ent_parent.pack(side=LEFT, fill=X, expand=YES)
        
        def pick_parent():
            p = filedialog.askdirectory(title="选择父文件夹")
            if p:
                if sys.platform == "win32":
                    p = p.replace("/", "\\")
                ent_parent.delete(0, END)
                ent_parent.insert(0, p)
        
        def scan_subfolders():
            parent = ent_parent.get().strip()
            if not parent or not os.path.isdir(parent):
                messagebox.showwarning("提示", "请选择有效的父文件夹！", parent=dlg)
                return
            
            # Clear existing
            folder_listbox.delete(0, END)
            
            # Scan subfolders
            try:
                subfolders = []
                for item in os.listdir(parent):
                    item_path = os.path.join(parent, item)
                    if os.path.isdir(item_path):
                        subfolders.append(item_path)
                
                # Sort folders
                subfolders.sort()
                
                # Add to listbox (all selected by default)
                for subfolder in subfolders:
                    folder_listbox.insert(END, subfolder)
                
                # Update count
                update_task_count()
                
                messagebox.showinfo("扫描完成", f"已从父文件夹扫描到 {len(subfolders)} 个子文件夹！", parent=dlg)
                
            except Exception as e:
                messagebox.showerror("错误", f"扫描失败：{str(e)}", parent=dlg)
        
        tk.Button(parent_frame, text="浏览", command=pick_parent, width=8).pack(side=LEFT, padx=(8, 4))
        tk.Button(parent_frame, text="扫描子文件夹", command=scan_subfolders, bg="green", fg="white", width=12).pack(side=LEFT, padx=(4, 0))

        # Manual add section
        tk.Label(frame, text="方式2：手动逐个添加", font=("System", 9, "bold"), fg="gray").pack(anchor="w", pady=(8, 5))

        # Top section with split layout
        top_frame = tk.Frame(frame)
        top_frame.pack(fill=BOTH, expand=YES, pady=(0, 10))

        # Left: Listbox for folder selection
        left_frame = tk.Frame(top_frame)
        left_frame.pack(side=LEFT, fill=BOTH, expand=YES)

        folder_listbox = tk.Listbox(left_frame, selectmode=EXTENDED, font=("Consolas", 9))
        folder_listbox.pack(side=LEFT, fill=BOTH, expand=YES)

        # Scrollbar for listbox
        list_scroll = tk.Scrollbar(left_frame, orient=VERTICAL, command=folder_listbox.yview)
        folder_listbox.configure(yscrollcommand=list_scroll.set)
        list_scroll.pack(side=RIGHT, fill=Y)

        # Right: Buttons for add/remove
        btn_frame = tk.Frame(top_frame)
        btn_frame.pack(side=RIGHT, fill=Y, padx=(10, 0))

        tk.Button(btn_frame, text="+ 浏览", width=8, command=lambda: self._task_browse_folder_to_dialog(folder_listbox)).pack(pady=2)
        
        tk.Label(btn_frame, text="", height=1).pack()  # Spacer
        
        tk.Button(btn_frame, text="- 移除", width=8, command=lambda: self._task_remove_from_listbox(folder_listbox)).pack(pady=2)
        tk.Button(btn_frame, text="C 清空", width=8, command=lambda: self._task_clear_listbox(folder_listbox)).pack(pady=2)

        # Bottom section: Current selected folders info
        bottom_frame = tk.Frame(frame)
        bottom_frame.pack(fill=X, pady=(8, 0))

        task_count_var = tk.StringVar()
        task_count_label = tk.Label(bottom_frame, textvariable=task_count_var, font=("System", 9))
        task_count_label.pack(anchor="w")

        # Selection controls
        check_frame = tk.Frame(bottom_frame)
        check_frame.pack(anchor="w", pady=(6, 0))

        def select_all():
            folder_listbox.selection_set(0, END)
        def deselect_all():
            folder_listbox.selection_clear(0, END)

        tk.Button(check_frame, text="全选", command=select_all, width=8).pack(side=LEFT, padx=(0, 4))
        tk.Button(check_frame, text="全不选", command=deselect_all, width=8).pack(side=LEFT, padx=4)

        # Update initial count
        def update_task_count():
            count = folder_listbox.size()
            if count == 0:
                task_count_var.set("已选择 0 个文件夹")
            elif count == 1:
                task_count_var.set("已选择 1 个文件夹")
            else:
                task_count_var.set(f"已选择 {count} 个文件夹")
        update_task_count()

        # Action buttons
        action_btn_frame = tk.Frame(frame)
        action_btn_frame.pack(fill=X, pady=(10, 0))

        tk.Button(action_btn_frame, text="添加", width=10, command=lambda: self._task_confirm_selection(folder_listbox, target_listbox)).pack(side=LEFT)
        tk.Button(action_btn_frame, text="取消", width=10, command=lambda: self._close_task_multi_folder_dialog(dlg)).pack(side=RIGHT)

    def _close_task_multi_folder_dialog(self, dlg):
        """Close the task multi-folder selection dialog."""
        try:
            if dlg.grab_current() == dlg:
                dlg.grab_release()
        except Exception:
            pass
        try:
            dlg.destroy()
        except Exception:
            pass
        self._task_multi_folder_dialog = None

    def _task_browse_folder_to_dialog(self, listbox):
        """Browse for a single folder and add to listbox."""
        p = filedialog.askdirectory(title=self.tr("msg_task_pick_source"))
        if p:
            if sys.platform == "win32":
                p = p.replace("/", "\\")
            # Check if already exists
            exists = False
            for i in range(listbox.size()):
                if listbox.get(i) == p:
                    exists = True
                    break
            if not exists:
                listbox.insert(END, p)

    def _task_browse_multiple_to_dialog(self, listbox):
        """Browse for multiple folders using multiple askdirectory calls."""
        base_dir = filedialog.askdirectory(title=self.tr("msg_task_pick_source") + " (选择第一个文件夹)")
        if not base_dir:
            return

        if sys.platform == "win32":
            base_dir = base_dir.replace("/", "\\")

        # Add first folder
        exists = False
        for i in range(listbox.size()):
            if listbox.get(i) == base_dir:
                exists = True
                break
        if not exists:
            listbox.insert(END, base_dir)

        # Ask if user wants to add more folders
        while True:
            result = messagebox.askyesno(
                "继续添加",
                "是否要继续添加更多文件夹？",
                icon="question"
            )
            if not result:
                break

            next_dir = filedialog.askdirectory(title=self.tr("msg_task_pick_source") + " (选择下一个文件夹)")
            if not next_dir:
                break

            if sys.platform == "win32":
                next_dir = next_dir.replace("/", "\\")

            # Check if already exists
            exists = False
            for i in range(listbox.size()):
                if listbox.get(i) == next_dir:
                    exists = True
                    break
            if not exists:
                listbox.insert(END, next_dir)

    def _task_remove_from_listbox(self, listbox):
        """Remove selected folder from listbox."""
        sel = listbox.curselection()
        if sel:
            listbox.delete(sel[0])

    def _task_clear_listbox(self, listbox):
        """Clear all folders from listbox."""
        listbox.delete(0, END)

    def _task_confirm_selection(self, source_listbox, target_listbox):
        """Add all folders from source listbox to target listbox."""
        added_count = 0
        current_folders = list(target_listbox.get(0, END))
        for i in range(source_listbox.size()):
            path = source_listbox.get(i)
            if path not in current_folders:
                target_listbox.insert(END, path)
                added_count += 1
        if added_count > 0:
            self._close_task_multi_folder_dialog(self._task_multi_folder_dialog)



    def remove_source_folder(self):
        selection = self.lst_source_folders.curselection()
        if not selection:
            return
        for index in reversed(selection):
            path = self.lst_source_folders.get(index)
            self.lst_source_folders.delete(index)
            if path in self.source_folders_list:
                self.source_folders_list.remove(path)
        # Sync
        if self.source_folders_list:
            self.var_source_folder.set(self.source_folders_list[0])
        else:
            self.var_source_folder.set("")


    def clear_source_folders(self):
        self.source_folders_list = []
        self.lst_source_folders.delete(0, END)
        self.var_source_folder.set("")


    def browse_target(self):
        path = filedialog.askdirectory(title="閫夋嫨鐩爣鐩綍")
        if path:
            self.var_target_folder.set(path)
            self.refresh_locator_maps()

    def open_source_folder(self, event=None):
        # Try to get selection from listbox first
        if hasattr(self, "lst_source_folders"):
            selection = self.lst_source_folders.curselection()
            if selection:
                path = self.lst_source_folders.get(selection[0])
                self._open_path(path)
                return

        # Fallback to var (first item usually)
        path = self.var_source_folder.get().strip()
        self._open_path(path)

    def open_target_folder(self):
        path = self.var_target_folder.get().strip()
        self._open_path(path)

    def open_config_folder(self):
        folder = os.path.dirname(self.config_path)
        self._open_path(folder)

    def _toggle_tooltip_advanced(self):
        """Show or hide advanced tooltip settings."""
        if not hasattr(self, "frm_tooltip_advanced"):
            return
        # 鍏堢Щ闄わ紝鍐嶆寜闇€瑕侀噸鏂?pack
        try:
            self.frm_tooltip_advanced.pack_forget()
        except Exception:
            pass
        if getattr(self, "var_show_tooltip_advanced", None) and self.var_show_tooltip_advanced.get():
            self.frm_tooltip_advanced.pack(fill=X, pady=(4, 0))

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
        path = filedialog.askdirectory(title="Select temp sandbox root")
        if path:
            self.var_temp_sandbox_root.set(path)

    def _on_reset_llm_delivery_root(self):
        self.var_llm_delivery_root.set("")

    def browse_log_folder(self):
        path = filedialog.askdirectory(title="閫夋嫨鏃ュ織鐩綍")
        if path:
            self.var_log_folder.set(path)

    def _profiles_dir(self):
        return os.path.join(self.script_dir, "config_profiles")

    def _profiles_index_path(self):
        return os.path.join(self._profiles_dir(), "profiles_index.json")

    def _profile_abs_path(self, file_name):
        safe_file = os.path.basename(str(file_name or "").strip())
        return os.path.join(self._profiles_dir(), safe_file)

    def _profile_timestamp(self):
        return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    def _profile_file_mtime(self, file_path):
        try:
            dt = datetime.fromtimestamp(os.path.getmtime(file_path))
            return dt.strftime("%Y-%m-%d %H:%M:%S")
        except Exception:
            return ""

    def _sanitize_profile_stem(self, name):
        stem = str(name or "").strip()
        stem = re.sub(r'[\\/:*?"<>|]+', "_", stem)
        stem = stem.strip().strip(".")
        return stem[:80] or "profile"

    def _next_profile_filename(self, name, records, exclude_file=None):
        stem = self._sanitize_profile_stem(name)
        existing = {
            str(rec.get("file", "")).strip().lower()
            for rec in (records or [])
            if isinstance(rec, dict)
        }
        exclude_lower = str(exclude_file or "").strip().lower()
        idx = 1
        while True:
            suffix = "" if idx == 1 else f"_{idx}"
            candidate = f"{stem}{suffix}.json"
            lower = candidate.lower()
            if lower == exclude_lower or lower not in existing:
                return candidate
            idx += 1

    def _load_profile_records(self):
        os.makedirs(self._profiles_dir(), exist_ok=True)
        index_path = self._profiles_index_path()
        index_name = os.path.basename(index_path).lower()
        records = []
        payload = {}
        if os.path.exists(index_path):
            try:
                with open(index_path, "r", encoding="utf-8") as f:
                    payload = json.load(f)
            except Exception:
                payload = {}

        raw_records = payload.get("profiles", []) if isinstance(payload, dict) else []
        for rec in raw_records:
            if not isinstance(rec, dict):
                continue
            file_name = os.path.basename(str(rec.get("file", "")).strip())
            if not file_name or file_name.lower() == index_name:
                continue
            abs_path = self._profile_abs_path(file_name)
            if not os.path.isfile(abs_path):
                continue
            name = str(rec.get("name", "")).strip() or os.path.splitext(file_name)[0]
            note = str(rec.get("note", "")).strip()
            updated_at = (
                str(rec.get("updated_at", "")).strip()
                or self._profile_file_mtime(abs_path)
            )
            records.append(
                {
                    "id": str(rec.get("id", "")).strip()
                    or datetime.now().strftime("%Y%m%d%H%M%S%f"),
                    "name": name,
                    "file": file_name,
                    "note": note,
                    "updated_at": updated_at,
                }
            )

        known_files = {str(rec.get("file", "")).strip().lower() for rec in records}
        for entry in sorted(os.listdir(self._profiles_dir())):
            lower = entry.lower()
            if not lower.endswith(".json"):
                continue
            if lower == index_name or lower in known_files:
                continue
            abs_path = self._profile_abs_path(entry)
            if not os.path.isfile(abs_path):
                continue
            records.append(
                {
                    "id": datetime.now().strftime("%Y%m%d%H%M%S%f"),
                    "name": os.path.splitext(entry)[0],
                    "file": entry,
                    "note": "",
                    "updated_at": self._profile_file_mtime(abs_path),
                }
            )
        return records

    def _save_profile_records(self, records):
        os.makedirs(self._profiles_dir(), exist_ok=True)
        index_path = self._profiles_index_path()
        out_records = []
        for rec in records or []:
            if not isinstance(rec, dict):
                continue
            file_name = os.path.basename(str(rec.get("file", "")).strip())
            if not file_name:
                continue
            out_records.append(
                {
                    "id": str(rec.get("id", "")).strip()
                    or datetime.now().strftime("%Y%m%d%H%M%S%f"),
                    "name": str(rec.get("name", "")).strip()
                    or os.path.splitext(file_name)[0],
                    "file": file_name,
                    "note": str(rec.get("note", "")).strip(),
                    "updated_at": str(rec.get("updated_at", "")).strip(),
                }
            )
        with open(index_path, "w", encoding="utf-8") as f:
            json.dump({"version": 1, "profiles": out_records}, f, indent=4, ensure_ascii=False)

    def _get_selected_profile_record(self):
        if self.profile_tree is None or not self.profile_tree.winfo_exists():
            return None
        selection = self.profile_tree.selection()
        if not selection:
            return None
        return self._profile_tree_rows.get(selection[0])

    def _update_profile_manager_controls(self):
        if self.profile_manager_win is None or not self.profile_manager_win.winfo_exists():
            return
        has_selected = self._get_selected_profile_record() is not None
        base_state = "disabled" if self._ui_running else "normal"
        selected_state = "disabled" if (self._ui_running or not has_selected) else "normal"
        for btn_name in ("btn_profile_new", "btn_profile_refresh", "btn_profile_open_dir"):
            btn = getattr(self, btn_name, None)
            if btn is not None and btn.winfo_exists():
                btn.configure(state=base_state)
        for btn_name in (
            "btn_profile_load",
            "btn_profile_save",
            "btn_profile_rename",
            "btn_profile_note",
            "btn_profile_delete",
        ):
            btn = getattr(self, btn_name, None)
            if btn is not None and btn.winfo_exists():
                btn.configure(state=selected_state)

    def _on_profile_tree_select(self, _event=None):
        self._update_profile_manager_controls()

    def _refresh_profile_tree(self, select_file=None):
        if self.profile_tree is None or not self.profile_tree.winfo_exists():
            return
        records = self._load_profile_records()
        self._profile_tree_rows = {}
        self.profile_tree.delete(*self.profile_tree.get_children())
        active_path = os.path.abspath(self.config_path)
        target_file = str(select_file or "").strip()
        target_iid = None
        for rec in records:
            abs_path = os.path.abspath(self._profile_abs_path(rec.get("file", "")))
            iid = self.profile_tree.insert(
                "",
                "end",
                values=(
                    rec.get("name", ""),
                    rec.get("file", ""),
                    rec.get("note", ""),
                    rec.get("updated_at", ""),
                ),
            )
            row = dict(rec)
            row["abs_path"] = abs_path
            self._profile_tree_rows[iid] = row
            if target_file and rec.get("file", "") == target_file:
                target_iid = iid
            elif not target_file and abs_path == active_path:
                target_iid = iid
        if target_iid is not None:
            self.profile_tree.selection_set(target_iid)
            self.profile_tree.focus(target_iid)
            self.profile_tree.see(target_iid)
        self.var_profile_active_path.set(self.config_path)
        self._update_profile_manager_controls()

    def _close_profile_manager_window(self):
        if self.profile_manager_win is not None and self.profile_manager_win.winfo_exists():
            try:
                self.profile_manager_win.destroy()
            except Exception:
                pass
        self.profile_manager_win = None
        self.profile_tree = None
        self._profile_tree_rows = {}

    def open_profile_manager_window(self):
        if self.profile_manager_win is not None and self.profile_manager_win.winfo_exists():
            self.profile_manager_win.lift()
            self.profile_manager_win.focus_force()
            self._refresh_profile_tree()
            return
        os.makedirs(self._profiles_dir(), exist_ok=True)
        self.profile_manager_win = tk.Toplevel(self)
        self.profile_manager_win.title(self.tr("win_profile_manager"))
        self.profile_manager_win.geometry("980x520")
        self.profile_manager_win.minsize(840, 420)
        self.profile_manager_win.protocol(
            "WM_DELETE_WINDOW", self._close_profile_manager_window
        )

        root = tb.Frame(self.profile_manager_win, padding=10)
        root.pack(fill=BOTH, expand=YES)

        top_row = tb.Frame(root)
        top_row.pack(fill=X, pady=(0, 8))
        tb.Label(
            top_row,
            text=self.tr("lbl_profile_active_config"),
            font=("System", 9, "bold"),
        ).pack(side=LEFT)
        tb.Label(
            top_row,
            textvariable=self.var_profile_active_path,
            bootstyle="secondary",
        ).pack(side=LEFT, fill=X, expand=YES, padx=(8, 0))

        tree_frame = tb.Frame(root)
        tree_frame.pack(fill=BOTH, expand=YES)
        cols = ("name", "file", "note", "updated")
        self.profile_tree = ttk.Treeview(
            tree_frame, columns=cols, show="headings", selectmode="browse"
        )
        self.profile_tree.heading("name", text=self.tr("col_profile_name"))
        self.profile_tree.heading("file", text=self.tr("col_profile_file"))
        self.profile_tree.heading("note", text=self.tr("col_profile_note"))
        self.profile_tree.heading("updated", text=self.tr("col_profile_updated"))
        self.profile_tree.column("name", width=180, anchor="w")
        self.profile_tree.column("file", width=220, anchor="w")
        self.profile_tree.column("note", width=360, anchor="w")
        self.profile_tree.column("updated", width=160, anchor="center")
        scr_y = tb.Scrollbar(
            tree_frame, orient=VERTICAL, command=self.profile_tree.yview
        )
        scr_x = tb.Scrollbar(
            tree_frame, orient=HORIZONTAL, command=self.profile_tree.xview
        )
        self.profile_tree.configure(
            yscrollcommand=scr_y.set, xscrollcommand=scr_x.set
        )
        self.profile_tree.pack(side=LEFT, fill=BOTH, expand=YES)
        scr_y.pack(side=RIGHT, fill=Y)
        scr_x.pack(side=BOTTOM, fill=X)
        self.profile_tree.bind("<<TreeviewSelect>>", self._on_profile_tree_select)

        btn_row = tb.Frame(root)
        btn_row.pack(fill=X, pady=(8, 0))
        self.btn_profile_new = tb.Button(
            btn_row,
            text=self.tr("btn_profile_new"),
            command=self._create_profile_from_current,
            bootstyle="success-outline",
            width=14,
        )
        self.btn_profile_new.pack(side=LEFT)
        self.btn_profile_load = tb.Button(
            btn_row,
            text=self.tr("btn_profile_load"),
            command=self._load_selected_profile,
            bootstyle="primary",
            width=12,
        )
        self.btn_profile_load.pack(side=LEFT, padx=(6, 0))
        self.btn_profile_save = tb.Button(
            btn_row,
            text=self.tr("btn_profile_save"),
            command=self._save_current_to_selected_profile,
            bootstyle="success",
            width=14,
        )
        self.btn_profile_save.pack(side=LEFT, padx=(6, 0))
        self.btn_profile_rename = tb.Button(
            btn_row,
            text=self.tr("btn_profile_rename"),
            command=self._rename_selected_profile,
            bootstyle="secondary-outline",
            width=10,
        )
        self.btn_profile_rename.pack(side=LEFT, padx=(6, 0))
        self.btn_profile_note = tb.Button(
            btn_row,
            text=self.tr("btn_profile_note"),
            command=self._edit_selected_profile_note,
            bootstyle="secondary-outline",
            width=10,
        )
        self.btn_profile_note.pack(side=LEFT, padx=(6, 0))
        self.btn_profile_delete = tb.Button(
            btn_row,
            text=self.tr("btn_profile_delete"),
            command=self._delete_selected_profile,
            bootstyle="danger-outline",
            width=10,
        )
        self.btn_profile_delete.pack(side=LEFT, padx=(6, 0))
        self.btn_profile_open_dir = tb.Button(
            btn_row,
            text=self.tr("btn_profile_open_dir"),
            command=self._open_profiles_folder,
            bootstyle="info-outline",
            width=10,
        )
        self.btn_profile_open_dir.pack(side=RIGHT)
        self.btn_profile_refresh = tb.Button(
            btn_row,
            text=self.tr("btn_profile_refresh"),
            command=self._refresh_profile_tree,
            bootstyle="info-outline",
            width=10,
        )
        self.btn_profile_refresh.pack(side=RIGHT, padx=(0, 6))

        self._refresh_profile_tree()
        self._update_profile_manager_controls()

    def _build_default_profile_name(self):
        return datetime.now().strftime("config_%Y%m%d_%H%M%S")

    def _save_profile_with_meta(self, name, note="", show_msg=True):
        profile_name = str(name or "").strip()
        if not profile_name:
            if show_msg:
                messagebox.showwarning(
                    self.tr("btn_save_cfg"), self.tr("msg_profile_name_required")
                )
            return False
        profile_note = str(note or "").strip()
        records = self._load_profile_records()
        file_name = self._next_profile_filename(profile_name, records)
        payload = self._compose_config_from_ui(self._load_config_for_write(), scope="all")
        profile_path = self._profile_abs_path(file_name)
        try:
            with open(profile_path, "w", encoding="utf-8") as f:
                json.dump(payload, f, indent=4, ensure_ascii=False)
            records.insert(
                0,
                {
                    "id": datetime.now().strftime("%Y%m%d%H%M%S%f"),
                    "name": profile_name,
                    "file": file_name,
                    "note": profile_note,
                    "updated_at": self._profile_timestamp(),
                },
            )
            self._save_profile_records(records)
            if self.profile_manager_win is not None and self.profile_manager_win.winfo_exists():
                self._refresh_profile_tree(select_file=file_name)
            if self.load_profile_dialog is not None and self.load_profile_dialog.winfo_exists():
                self._refresh_load_profile_tree(select_file=file_name)
            msg = self.tr("msg_profile_create_ok").format(profile_name)
            if show_msg:
                messagebox.showinfo(self.tr("btn_save_cfg"), msg)
            if hasattr(self, "var_status"):
                self.var_status.set(msg)
            if hasattr(self, "var_locator_result"):
                self.var_locator_result.set(msg)
            return True
        except Exception as e:
            if show_msg:
                messagebox.showerror(self.tr("btn_save_cfg"), self.tr("msg_save_fail").format(e))
            return False

    def _load_profile_record(self, rec, confirm_dirty=True, show_msg=False):
        if not isinstance(rec, dict):
            return False
        if confirm_dirty and self.cfg_dirty:
            confirm = messagebox.askyesno(
                self.tr("btn_load_cfg"),
                self.tr("msg_confirm_load_config_dirty"),
            )
            if not confirm:
                return False
        profile_path = str(rec.get("abs_path", "")).strip()
        if not os.path.isfile(profile_path):
            messagebox.showerror(
                self.tr("btn_load_cfg"),
                self.tr("msg_profile_file_missing").format(profile_path),
            )
            return False
        try:
            with open(profile_path, "r", encoding="utf-8") as f:
                profile_cfg = json.load(f)
            if not isinstance(profile_cfg, dict):
                raise ValueError("profile config must be an object")
            self.config_path = profile_path
            if hasattr(self, "var_config_path"):
                self.var_config_path.set(self.config_path)
            self.var_profile_active_path.set(self.config_path)
            self._load_config_to_ui()
            if self.profile_manager_win is not None and self.profile_manager_win.winfo_exists():
                self._refresh_profile_tree(select_file=rec.get("file", ""))
            if self.load_profile_dialog is not None and self.load_profile_dialog.winfo_exists():
                self._refresh_load_profile_tree(select_file=rec.get("file", ""))
            msg = self.tr("msg_profile_load_ok").format(rec.get("name", ""))
            if show_msg:
                messagebox.showinfo(self.tr("btn_load_cfg"), msg)
            if hasattr(self, "var_status"):
                self.var_status.set(msg)
            if hasattr(self, "var_locator_result"):
                self.var_locator_result.set(msg)
            return True
        except Exception as e:
            messagebox.showerror(
                self.tr("btn_load_cfg"), self.tr("msg_save_fail").format(e)
            )
            return False

    def _create_profile_from_current(self):
        parent = (
            self.profile_manager_win
            if self.profile_manager_win is not None and self.profile_manager_win.winfo_exists()
            else self
        )
        default_name = self._build_default_profile_name()
        name = simpledialog.askstring(
            self.tr("btn_profile_new"),
            self.tr("msg_profile_name_prompt"),
            parent=parent,
            initialvalue=default_name,
        )
        if name is None:
            return
        name = str(name).strip()
        if not name:
            messagebox.showwarning(
                self.tr("btn_profile_new"), self.tr("msg_profile_name_required")
            )
            return
        note = simpledialog.askstring(
            self.tr("btn_profile_note"),
            self.tr("msg_profile_note_prompt"),
            parent=parent,
            initialvalue="",
        )
        note = "" if note is None else note
        self._save_profile_with_meta(name, note, show_msg=False)

    def _load_selected_profile(self):
        rec = self._get_selected_profile_record()
        if rec is None:
            messagebox.showwarning(
                self.tr("btn_profile_load"), self.tr("msg_profile_select_required")
            )
            return
        self._load_profile_record(rec, confirm_dirty=True, show_msg=False)

    def _save_current_to_selected_profile(self):
        rec = self._get_selected_profile_record()
        if rec is None:
            messagebox.showwarning(
                self.tr("btn_profile_save"), self.tr("msg_profile_select_required")
            )
            return
        profile_path = rec.get("abs_path", "")
        if not profile_path:
            return
        payload = self._compose_config_from_ui(self._load_config_for_write(), scope="all")
        records = self._load_profile_records()
        try:
            with open(profile_path, "w", encoding="utf-8") as f:
                json.dump(payload, f, indent=4, ensure_ascii=False)
            for item in records:
                if item.get("file", "") == rec.get("file", ""):
                    item["updated_at"] = self._profile_timestamp()
                    break
            self._save_profile_records(records)
            self._refresh_profile_tree(select_file=rec.get("file", ""))
            msg = self.tr("msg_profile_save_ok").format(rec.get("name", ""))
            if hasattr(self, "var_status"):
                self.var_status.set(msg)
            if hasattr(self, "var_locator_result"):
                self.var_locator_result.set(msg)
        except Exception as e:
            messagebox.showerror(
                self.tr("btn_profile_save"), self.tr("msg_save_fail").format(e)
            )

    def _rename_selected_profile(self):
        rec = self._get_selected_profile_record()
        if rec is None:
            messagebox.showwarning(
                self.tr("btn_profile_rename"), self.tr("msg_profile_select_required")
            )
            return
        parent = (
            self.profile_manager_win
            if self.profile_manager_win is not None and self.profile_manager_win.winfo_exists()
            else self
        )
        new_name = simpledialog.askstring(
            self.tr("btn_profile_rename"),
            self.tr("msg_profile_rename_prompt"),
            parent=parent,
            initialvalue=rec.get("name", ""),
        )
        if new_name is None:
            return
        new_name = str(new_name).strip()
        if not new_name:
            messagebox.showwarning(
                self.tr("btn_profile_rename"), self.tr("msg_profile_name_required")
            )
            return
        records = self._load_profile_records()
        old_file = rec.get("file", "")
        new_file = self._next_profile_filename(new_name, records, exclude_file=old_file)
        old_path = self._profile_abs_path(old_file)
        new_path = self._profile_abs_path(new_file)
        try:
            if new_file != old_file:
                os.replace(old_path, new_path)
            for item in records:
                if item.get("file", "") == old_file:
                    item["name"] = new_name
                    item["file"] = new_file
                    item["updated_at"] = self._profile_timestamp()
                    break
            self._save_profile_records(records)
            if os.path.abspath(self.config_path) == os.path.abspath(old_path):
                self.config_path = new_path
                self.var_profile_active_path.set(self.config_path)
                if hasattr(self, "var_config_path"):
                    self.var_config_path.set(self.config_path)
            self._refresh_profile_tree(select_file=new_file)
            msg = self.tr("msg_profile_rename_ok").format(new_name)
            if hasattr(self, "var_status"):
                self.var_status.set(msg)
            if hasattr(self, "var_locator_result"):
                self.var_locator_result.set(msg)
        except Exception as e:
            messagebox.showerror(
                self.tr("btn_profile_rename"), self.tr("msg_save_fail").format(e)
            )

    def _edit_selected_profile_note(self):
        rec = self._get_selected_profile_record()
        if rec is None:
            messagebox.showwarning(
                self.tr("btn_profile_note"), self.tr("msg_profile_select_required")
            )
            return
        parent = (
            self.profile_manager_win
            if self.profile_manager_win is not None and self.profile_manager_win.winfo_exists()
            else self
        )
        note = simpledialog.askstring(
            self.tr("btn_profile_note"),
            self.tr("msg_profile_note_prompt"),
            parent=parent,
            initialvalue=rec.get("note", ""),
        )
        if note is None:
            return
        records = self._load_profile_records()
        for item in records:
            if item.get("file", "") == rec.get("file", ""):
                item["note"] = str(note).strip()
                item["updated_at"] = self._profile_timestamp()
                break
        try:
            self._save_profile_records(records)
            self._refresh_profile_tree(select_file=rec.get("file", ""))
            msg = self.tr("msg_profile_note_ok")
            if hasattr(self, "var_status"):
                self.var_status.set(msg)
            if hasattr(self, "var_locator_result"):
                self.var_locator_result.set(msg)
        except Exception as e:
            messagebox.showerror(
                self.tr("btn_profile_note"), self.tr("msg_save_fail").format(e)
            )

    def _delete_selected_profile(self):
        rec = self._get_selected_profile_record()
        if rec is None:
            messagebox.showwarning(
                self.tr("btn_profile_delete"), self.tr("msg_profile_select_required")
            )
            return
        profile_path = rec.get("abs_path", "")
        if os.path.abspath(profile_path) == os.path.abspath(self.config_path):
            messagebox.showwarning(
                self.tr("btn_profile_delete"),
                self.tr("msg_profile_delete_active_blocked"),
            )
            return
        confirm = messagebox.askyesno(
            self.tr("btn_profile_delete"),
            self.tr("msg_profile_delete_confirm").format(rec.get("name", "")),
        )
        if not confirm:
            return
        records = self._load_profile_records()
        try:
            if os.path.exists(profile_path):
                os.remove(profile_path)
            records = [
                item
                for item in records
                if item.get("file", "") != rec.get("file", "")
            ]
            self._save_profile_records(records)
            self._refresh_profile_tree()
            msg = self.tr("msg_profile_delete_ok").format(rec.get("name", ""))
            if hasattr(self, "var_status"):
                self.var_status.set(msg)
            if hasattr(self, "var_locator_result"):
                self.var_locator_result.set(msg)
        except Exception as e:
            messagebox.showerror(
                self.tr("btn_profile_delete"), self.tr("msg_save_fail").format(e)
            )

    def _open_profiles_folder(self):
        os.makedirs(self._profiles_dir(), exist_ok=True)
        self._open_path(self._profiles_dir())

    def _close_save_profile_dialog(self):
        if self.save_profile_dialog is not None and self.save_profile_dialog.winfo_exists():
            try:
                if self.save_profile_dialog.grab_current() == self.save_profile_dialog:
                    self.save_profile_dialog.grab_release()
            except Exception:
                pass
            try:
                self.save_profile_dialog.destroy()
            except Exception:
                pass
        self.save_profile_dialog = None

    def _place_dialog_in_main(self, dialog, width, height):
        try:
            self.update_idletasks()
            base_x = int(self.winfo_rootx())
            base_y = int(self.winfo_rooty())
            base_w = max(int(self.winfo_width()), 200)
            base_h = max(int(self.winfo_height()), 200)
            x = base_x + max((base_w - int(width)) // 2, 20)
            y = base_y + max((base_h - int(height)) // 2, 20)
            dialog.geometry(f"{int(width)}x{int(height)}+{x}+{y}")
        except Exception:
            pass

    def open_save_profile_dialog(self):
        if self._ui_running:
            return
        if self.save_profile_dialog is not None and self.save_profile_dialog.winfo_exists():
            self._place_dialog_in_main(self.save_profile_dialog, 520, 220)
            self.save_profile_dialog.lift()
            self.save_profile_dialog.focus_force()
            return
        dlg = tk.Toplevel(self)
        dlg.title(self.tr("win_save_config"))
        self._place_dialog_in_main(dlg, 520, 220)
        dlg.minsize(460, 210)
        dlg.transient(self)
        dlg.grab_set()
        dlg.lift()
        dlg.protocol("WM_DELETE_WINDOW", self._close_save_profile_dialog)
        self.save_profile_dialog = dlg

        frame = tb.Frame(dlg, padding=12)
        frame.pack(fill=BOTH, expand=YES)
        tb.Label(frame, text=self.tr("lbl_profile_name"), font=("System", 9, "bold")).pack(anchor="w")
        self.var_save_profile_name = tk.StringVar(value=self._build_default_profile_name())
        ent_name = tb.Entry(frame, textvariable=self.var_save_profile_name)
        ent_name.pack(fill=X, pady=(2, 10))
        ent_name.focus_set()
        ent_name.selection_range(0, END)

        tb.Label(frame, text=self.tr("lbl_profile_note"), font=("System", 9, "bold")).pack(anchor="w")
        self.var_save_profile_note = tk.StringVar(value="")
        ent_note = tb.Entry(frame, textvariable=self.var_save_profile_note)
        ent_note.pack(fill=X, pady=(2, 12))

        btn_row = tb.Frame(frame)
        btn_row.pack(fill=X)
        self.btn_save_profile_confirm = tb.Button(
            btn_row,
            text=self.tr("btn_confirm_save_profile"),
            command=self._confirm_save_profile_dialog,
            bootstyle="success",
            width=14,
        )
        self.btn_save_profile_confirm.pack(side=LEFT)
        self.btn_save_profile_cancel = tb.Button(
            btn_row,
            text=self.tr("btn_cancel"),
            command=self._close_save_profile_dialog,
            bootstyle="secondary-outline",
            width=10,
        )
        self.btn_save_profile_cancel.pack(side=LEFT, padx=(8, 0))
        self._update_profile_dialog_controls()

    def _confirm_save_profile_dialog(self):
        name = self.var_save_profile_name.get().strip() if hasattr(self, "var_save_profile_name") else ""
        note = self.var_save_profile_note.get().strip() if hasattr(self, "var_save_profile_note") else ""
        if not name:
            messagebox.showwarning(self.tr("btn_save_cfg"), self.tr("msg_profile_name_required"))
            return
        if self._save_profile_with_meta(name, note, show_msg=True):
            self._close_save_profile_dialog()

    def _close_load_profile_dialog(self):
        if self.load_profile_dialog is not None and self.load_profile_dialog.winfo_exists():
            try:
                if self.load_profile_dialog.grab_current() == self.load_profile_dialog:
                    self.load_profile_dialog.grab_release()
            except Exception:
                pass
            try:
                self.load_profile_dialog.destroy()
            except Exception:
                pass
        self.load_profile_dialog = None
        self.load_profile_tree = None
        self._load_profile_tree_rows = {}

    def _get_selected_load_profile_record(self):
        if self.load_profile_tree is None or not self.load_profile_tree.winfo_exists():
            return None
        selection = self.load_profile_tree.selection()
        if not selection:
            return None
        return self._load_profile_tree_rows.get(selection[0])

    def _refresh_load_profile_tree(self, select_file=None):
        if self.load_profile_tree is None or not self.load_profile_tree.winfo_exists():
            return
        records = self._load_profile_records()
        self._load_profile_tree_rows = {}
        self.load_profile_tree.delete(*self.load_profile_tree.get_children())
        target_file = str(select_file or "").strip()
        active_path = os.path.abspath(self.config_path)
        target_iid = None
        for rec in records:
            abs_path = os.path.abspath(self._profile_abs_path(rec.get("file", "")))
            iid = self.load_profile_tree.insert(
                "",
                "end",
                values=(
                    rec.get("name", ""),
                    rec.get("file", ""),
                    rec.get("note", ""),
                    rec.get("updated_at", ""),
                ),
            )
            row = dict(rec)
            row["abs_path"] = abs_path
            self._load_profile_tree_rows[iid] = row
            if target_file and rec.get("file", "") == target_file:
                target_iid = iid
            elif not target_file and abs_path == active_path:
                target_iid = iid
        if target_iid is not None:
            self.load_profile_tree.selection_set(target_iid)
            self.load_profile_tree.focus(target_iid)
            self.load_profile_tree.see(target_iid)
        self._update_profile_dialog_controls()

    def open_load_profile_dialog(self):
        if self._ui_running:
            return
        if self.load_profile_dialog is not None and self.load_profile_dialog.winfo_exists():
            self._place_dialog_in_main(self.load_profile_dialog, 860, 460)
            self.load_profile_dialog.lift()
            self.load_profile_dialog.focus_force()
            self._refresh_load_profile_tree()
            return
        dlg = tk.Toplevel(self)
        dlg.title(self.tr("win_load_config"))
        self._place_dialog_in_main(dlg, 860, 460)
        dlg.minsize(760, 380)
        dlg.transient(self)
        dlg.grab_set()
        dlg.lift()
        dlg.protocol("WM_DELETE_WINDOW", self._close_load_profile_dialog)
        self.load_profile_dialog = dlg

        root = tb.Frame(dlg, padding=10)
        root.pack(fill=BOTH, expand=YES)
        tb.Label(root, text=self.tr("lbl_profile_select"), font=("System", 9, "bold")).pack(anchor="w", pady=(0, 6))

        tree_frame = tb.Frame(root)
        tree_frame.pack(fill=BOTH, expand=YES)
        cols = ("name", "file", "note", "updated")
        self.load_profile_tree = ttk.Treeview(
            tree_frame, columns=cols, show="headings", selectmode="browse"
        )
        self.load_profile_tree.heading("name", text=self.tr("col_profile_name"))
        self.load_profile_tree.heading("file", text=self.tr("col_profile_file"))
        self.load_profile_tree.heading("note", text=self.tr("col_profile_note"))
        self.load_profile_tree.heading("updated", text=self.tr("col_profile_updated"))
        self.load_profile_tree.column("name", width=180, anchor="w")
        self.load_profile_tree.column("file", width=220, anchor="w")
        self.load_profile_tree.column("note", width=300, anchor="w")
        self.load_profile_tree.column("updated", width=140, anchor="center")
        scr_y = tb.Scrollbar(tree_frame, orient=VERTICAL, command=self.load_profile_tree.yview)
        scr_x = tb.Scrollbar(tree_frame, orient=HORIZONTAL, command=self.load_profile_tree.xview)
        self.load_profile_tree.configure(yscrollcommand=scr_y.set, xscrollcommand=scr_x.set)
        self.load_profile_tree.pack(side=LEFT, fill=BOTH, expand=YES)
        scr_y.pack(side=RIGHT, fill=Y)
        scr_x.pack(side=BOTTOM, fill=X)
        self.load_profile_tree.bind("<<TreeviewSelect>>", lambda _e: self._update_profile_dialog_controls())

        btn_row = tb.Frame(root)
        btn_row.pack(fill=X, pady=(8, 0))
        self.btn_load_profile_confirm = tb.Button(
            btn_row,
            text=self.tr("btn_confirm_load_profile"),
            command=self._confirm_load_profile_dialog,
            bootstyle="primary",
            width=14,
        )
        self.btn_load_profile_confirm.pack(side=LEFT)
        self.btn_load_profile_refresh = tb.Button(
            btn_row,
            text=self.tr("btn_profile_refresh"),
            command=self._refresh_load_profile_tree,
            bootstyle="info-outline",
            width=10,
        )
        self.btn_load_profile_refresh.pack(side=LEFT, padx=(8, 0))
        self.btn_load_profile_cancel = tb.Button(
            btn_row,
            text=self.tr("btn_cancel"),
            command=self._close_load_profile_dialog,
            bootstyle="secondary-outline",
            width=10,
        )
        self.btn_load_profile_cancel.pack(side=RIGHT)

        self._refresh_load_profile_tree()
        self._update_profile_dialog_controls()

    def _confirm_load_profile_dialog(self):
        rec = self._get_selected_load_profile_record()
        if rec is None:
            messagebox.showwarning(
                self.tr("btn_load_cfg"),
                self.tr("msg_profile_select_required"),
            )
            return
        if self._load_profile_record(rec, confirm_dirty=True, show_msg=True):
            self._close_load_profile_dialog()

    def _update_profile_dialog_controls(self):
        state_base = "disabled" if self._ui_running else "normal"
        if self.save_profile_dialog is not None and self.save_profile_dialog.winfo_exists():
            for btn_name in ("btn_save_profile_confirm", "btn_save_profile_cancel"):
                btn = getattr(self, btn_name, None)
                if btn is not None and btn.winfo_exists():
                    btn.configure(state=state_base)
        if self.load_profile_dialog is not None and self.load_profile_dialog.winfo_exists():
            selected = self._get_selected_load_profile_record() is not None
            confirm_state = "disabled" if (self._ui_running or not selected) else "normal"
            btn = getattr(self, "btn_load_profile_confirm", None)
            if btn is not None and btn.winfo_exists():
                btn.configure(state=confirm_state)
            for btn_name in ("btn_load_profile_refresh", "btn_load_profile_cancel"):
                btn = getattr(self, btn_name, None)
                if btn is not None and btn.winfo_exists():
                    btn.configure(state=state_base)

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
            alt = ", ".join(
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

    def _set_config_dirty(self, dirty=True):
        self.cfg_dirty = bool(dirty)
        if hasattr(self, "lbl_cfg_dirty_state"):
            if self.cfg_dirty:
                self.lbl_cfg_dirty_state.configure(
                    text=self.tr("lbl_cfg_dirty_unsaved"),
                    bootstyle="warning",
                )
            else:
                self.lbl_cfg_dirty_state.configure(
                    text=self.tr("lbl_cfg_dirty_clean"),
                    bootstyle="success",
                )
        if not self.cfg_dirty:
            self._update_config_dirty_summary({})

    def _update_config_tab_dirty_markers(self, section_dirty=None):
        if not hasattr(self, "main_notebook"):
            return
        if not self._cfg_tab_meta:
            return
        section_dirty = section_dirty or {}
        # 澶氫釜 section 鍙兘鏄犲皠鍒板悓涓€涓墿鐞?tab锛岃仛鍚?dirty 鐘舵€?
        tab_dirty = {}
        tab_labels = {}
        for section_name, tab_widget, label_key in self._cfg_tab_meta:
            if id(tab_widget) not in tab_dirty:
                tab_dirty[id(tab_widget)] = False
                tab_labels[id(tab_widget)] = (tab_widget, label_key)
            if section_dirty.get(section_name, False):
                tab_dirty[id(tab_widget)] = True
        for tid, is_dirty in tab_dirty.items():
            tab_widget, label_key = tab_labels[tid]
            try:
                title = self.tr(label_key)
                if is_dirty:
                    title = f"{title} *"
                self.main_notebook.tab(tab_widget, text=title)
            except Exception:
                pass

    def _compute_section_dirty(self, ui_snapshot, base_snapshot):
        section_fields = {
            "shared": ["kill_process_mode", "log_folder"],
            "convert": [
                "timeout_seconds",
                "pdf_wait_seconds",
                "ppt_timeout_seconds",
                "ppt_pdf_wait_seconds",
                "office_reuse_app",
                "office_restart_every_n_files",
                "enable_sandbox",
                "temp_sandbox_root",
                "sandbox_min_free_gb",
                "sandbox_low_space_policy",
            ],
            "ai": [
                "enable_corpus_manifest",
                "enable_markdown",
                "markdown_strip_header_footer",
                "markdown_structured_headings",
                "enable_markdown_quality_report",
                "enable_excel_json",
                "enable_chromadb_export",
                "enable_llm_delivery_hub",
                "llm_delivery_root",
                "llm_delivery_flatten",
                "llm_delivery_include_pdf",
                "mshelpviewer_folder_name",
                "enable_mshelp_merge_output",
                "enable_mshelp_output_docx",
                "enable_mshelp_output_pdf",
            ],
            "incremental": [
                "enable_incremental_mode",
                "incremental_verify_hash",
                "incremental_reprocess_renamed",
                "source_priority_skip_same_name_pdf",
                "global_md5_dedup",
                "enable_update_package",
            ],
            "merge": [
                "enable_merge",
                "output_enable_pdf",
                "output_enable_md",
                "output_enable_merged",
                "output_enable_independent",
                "merge_convert_submode",
                "merge_mode",
                "merge_source",
                "enable_merge_index",
                "enable_merge_excel",
                "max_merge_size_mb",
            ],
            "rules": ["excluded_folders", "price_keywords"],
        }
        section_dirty = {}
        for section_name, keys in section_fields.items():
            section_dirty[section_name] = any(
                ui_snapshot.get(k) != base_snapshot.get(k) for k in keys
            )
        section_dirty["ui"] = ui_snapshot.get("ui", {}) != base_snapshot.get("ui", {})
        return section_dirty

    def _update_config_dirty_summary(self, section_dirty):
        if not hasattr(self, "lbl_cfg_dirty_sections"):
            return
        section_dirty = section_dirty or {}
        self._last_section_dirty = dict(section_dirty)
        dirty_names = []
        for section_name, _, label_key in self._cfg_tab_meta:
            if section_dirty.get(section_name, False):
                dirty_names.append(self.tr(label_key))
        if dirty_names:
            self.lbl_cfg_dirty_sections.configure(
                text=self.tr("lbl_cfg_dirty_sections").format(", ".join(dirty_names)),
                bootstyle="warning",
            )
        else:
            self.lbl_cfg_dirty_sections.configure(
                text=self.tr("lbl_cfg_dirty_none"),
                bootstyle="secondary",
            )
        can_act = bool(dirty_names) and (not self._ui_running)
        if hasattr(self, "btn_save_cfg_dirty"):
            count = len(dirty_names)
            if count > 0:
                self.btn_save_cfg_dirty.configure(
                    text=self.tr("btn_save_cfg_dirty_count").format(count)
                )
            else:
                self.btn_save_cfg_dirty.configure(text=self.tr("btn_save_cfg_dirty"))
            self.btn_save_cfg_dirty.configure(
                state=("normal" if can_act else "disabled")
            )

    def _focus_dirty_section(self, section_name):
        if not hasattr(self, "main_notebook"):
            return
        if not self._cfg_tab_meta:
            return
        for section, tab_widget, _ in self._cfg_tab_meta:
            if section == section_name:
                try:
                    self.main_notebook.select(tab_widget)
                except Exception:
                    pass
                return

    def _get_cfg_section_titles(self, section_names):
        titles = []
        section_set = set(section_names or [])
        for section_name, _, label_key in self._cfg_tab_meta:
            if section_name in section_set:
                titles.append(self.tr(label_key))
        return titles

    def _apply_snapshot_sections_to_ui(self, snapshot, section_names):
        snapshot = snapshot or {}
        sections = set(section_names or [])
        if "shared" in sections:
            self.var_kill_mode.set(snapshot.get("kill_process_mode", KILL_MODE_AUTO))
            self.var_log_folder.set(snapshot.get("log_folder", "./logs"))
        if "convert" in sections:
            self.var_timeout_seconds.set(str(snapshot.get("timeout_seconds", 60)))
            self.var_pdf_wait_seconds.set(str(snapshot.get("pdf_wait_seconds", 15)))
            self.var_ppt_timeout_seconds.set(
                str(snapshot.get("ppt_timeout_seconds", 180))
            )
            self.var_ppt_pdf_wait_seconds.set(
                str(snapshot.get("ppt_pdf_wait_seconds", 30))
            )
            self.var_office_reuse_app.set(
                1 if snapshot.get("office_reuse_app", True) else 0
            )
            self.var_office_restart_every_n_files.set(
                str(snapshot.get("office_restart_every_n_files", 25))
            )
            self.var_enable_sandbox.set(
                1 if snapshot.get("enable_sandbox", True) else 0
            )
            self.var_temp_sandbox_root.set(
                snapshot.get("temp_sandbox_root", "") or ""
            )
            self.var_sandbox_min_free_gb.set(
                str(snapshot.get("sandbox_min_free_gb", 10))
            )
            self.var_sandbox_low_space_policy.set(
                snapshot.get("sandbox_low_space_policy", "block") or "block"
            )
            self._on_toggle_sandbox()
        if "ai" in sections:
            self.var_enable_corpus_manifest.set(
                1 if snapshot.get("enable_corpus_manifest", True) else 0
            )
            self.var_enable_markdown.set(1 if snapshot.get("enable_markdown", True) else 0)
            self.var_markdown_strip_header_footer.set(
                1 if snapshot.get("markdown_strip_header_footer", True) else 0
            )
            self.var_markdown_structured_headings.set(
                1 if snapshot.get("markdown_structured_headings", True) else 0
            )
            self.var_enable_markdown_quality_report.set(
                1 if snapshot.get("enable_markdown_quality_report", True) else 0
            )
            self.var_enable_excel_json.set(
                1 if snapshot.get("enable_excel_json", False) else 0
            )
            self.var_enable_chromadb_export.set(
                1 if snapshot.get("enable_chromadb_export", False) else 0
            )
            self.var_enable_llm_delivery_hub.set(
                1 if snapshot.get("enable_llm_delivery_hub", True) else 0
            )
            self.var_llm_delivery_root.set(
                snapshot.get("llm_delivery_root", "") or ""
            )
            self.var_llm_delivery_flatten.set(
                1 if snapshot.get("llm_delivery_flatten", False) else 0
            )
            self.var_llm_delivery_include_pdf.set(
                1 if snapshot.get("llm_delivery_include_pdf", False) else 0
            )
            self.var_enable_upload_readme.set(
                1 if snapshot.get("enable_upload_readme", True) else 0
            )
            self.var_enable_upload_json_manifest.set(
                1 if snapshot.get("enable_upload_json_manifest", True) else 0
            )
            self.var_upload_dedup_merged.set(
                1 if snapshot.get("upload_dedup_merged", True) else 0
            )
            self.var_mshelpviewer_folder_name.set(
                str(snapshot.get("mshelpviewer_folder_name", "MSHelpViewer") or "MSHelpViewer")
            )
            self.var_enable_mshelp_merge_output.set(
                1 if snapshot.get("enable_mshelp_merge_output", True) else 0
            )
            self.var_enable_mshelp_output_docx.set(
                1 if snapshot.get("enable_mshelp_output_docx", False) else 0
            )
            self.var_enable_mshelp_output_pdf.set(
                1 if snapshot.get("enable_mshelp_output_pdf", False) else 0
            )
        if "incremental" in sections:
            self.var_enable_incremental_mode.set(
                1 if snapshot.get("enable_incremental_mode", False) else 0
            )
            self.var_incremental_verify_hash.set(
                1 if snapshot.get("incremental_verify_hash", False) else 0
            )
            self.var_incremental_reprocess_renamed.set(
                1 if snapshot.get("incremental_reprocess_renamed", False) else 0
            )
            self.var_source_priority_skip_same_name_pdf.set(
                1 if snapshot.get("source_priority_skip_same_name_pdf", False) else 0
            )
            self.var_global_md5_dedup.set(
                1 if snapshot.get("global_md5_dedup", False) else 0
            )
            self.var_enable_update_package.set(
                1 if snapshot.get("enable_update_package", True) else 0
            )
        if "merge" in sections:
            self.var_enable_merge.set(1 if snapshot.get("enable_merge", True) else 0)
            self.var_output_enable_pdf.set(
                1 if snapshot.get("output_enable_pdf", True) else 0
            )
            self.var_output_enable_md.set(
                1 if snapshot.get("output_enable_md", True) else 0
            )
            self.var_output_enable_merged.set(
                1 if snapshot.get("output_enable_merged", True) else 0
            )
            self.var_output_enable_independent.set(
                1 if snapshot.get("output_enable_independent", False) else 0
            )
            self.var_merge_convert_submode.set(
                snapshot.get(
                    "merge_convert_submode",
                    MERGE_CONVERT_SUBMODE_MERGE_ONLY,
                )
            )
            self.var_merge_mode.set(snapshot.get("merge_mode", MERGE_MODE_CATEGORY))
            self.var_merge_source.set(snapshot.get("merge_source", "source"))
            self.var_enable_merge_index.set(
                1 if snapshot.get("enable_merge_index", False) else 0
            )
            self.var_enable_merge_excel.set(
                1 if snapshot.get("enable_merge_excel", False) else 0
            )
            self.var_max_merge_size_mb.set(str(snapshot.get("max_merge_size_mb", 80)))
            self.var_merge_filename_pattern.set(
                snapshot.get("merge_filename_pattern") or "Merged_{category}_{timestamp}_{idx}"
            )
        if "rules" in sections:
            self._set_text_widget_lines(
                self.txt_excluded_folders, snapshot.get("excluded_folders", [])
            )
            self._set_text_widget_lines(
                self.txt_price_keywords, snapshot.get("price_keywords", [])
            )
        if "ui" in sections:
            ui_snapshot = snapshot.get("ui", {}) if isinstance(snapshot.get("ui"), dict) else {}
            self.var_tooltip_delay_ms.set(
                str(ui_snapshot.get("tooltip_delay_ms", self.TOOLTIP_DEFAULTS["tooltip_delay_ms"]))
            )
            self.var_tooltip_font_size.set(
                str(ui_snapshot.get("tooltip_font_size", self.TOOLTIP_DEFAULTS["tooltip_font_size"]))
            )
            self.var_tooltip_bg.set(
                ui_snapshot.get("tooltip_bg", self.TOOLTIP_DEFAULTS["tooltip_bg"])
            )
            self.var_tooltip_fg.set(
                ui_snapshot.get("tooltip_fg", self.TOOLTIP_DEFAULTS["tooltip_fg"])
            )
            self.var_tooltip_auto_theme.set(
                1
                if ui_snapshot.get(
                    "tooltip_auto_theme", self.TOOLTIP_DEFAULTS["tooltip_auto_theme"]
                )
                else 0
            )
            self.var_confirm_revert_dirty.set(
                1 if ui_snapshot.get("confirm_revert_dirty", True) else 0
            )
            self.apply_tooltip_settings(silent=True)
            self.validate_tooltip_inputs(silent=True)
        self._sync_markdown_master_with_global_output()

    def _revert_dirty_config_sections(self, show_msg=True):
        self._refresh_config_dirty_from_file()
        dirty_sections = []
        for section_name, _, _ in self._cfg_tab_meta:
            if self._last_section_dirty.get(section_name, False):
                dirty_sections.append(section_name)
        if not dirty_sections:
            msg = self.tr("msg_revert_dirty_none")
            if show_msg:
                messagebox.showinfo(self.tr("btn_revert_cfg_dirty"), msg)
            if hasattr(self, "var_status"):
                self.var_status.set(msg)
            if hasattr(self, "var_locator_result"):
                self.var_locator_result.set(msg)
            return
        section_titles = self._get_cfg_section_titles(dirty_sections)
        section_text = ", ".join(section_titles)
        need_confirm = (
            hasattr(self, "var_confirm_revert_dirty")
            and bool(self.var_confirm_revert_dirty.get())
        )
        if show_msg and need_confirm:
            confirm = messagebox.askyesno(
                self.tr("btn_revert_cfg_dirty"),
                self.tr("msg_confirm_revert_dirty").format(section_text),
            )
            if not confirm:
                return
        self._suspend_cfg_dirty = True
        try:
            self._apply_snapshot_sections_to_ui(
                self._baseline_config_snapshot, dirty_sections
            )
            self._on_toggle_incremental_mode()
            self.validate_runtime_inputs(silent=True, scope="config")
        finally:
            self._suspend_cfg_dirty = False
        self._refresh_config_dirty_state()
        msg = self.tr("msg_revert_dirty_sections").format(section_text)
        if show_msg:
            messagebox.showinfo(self.tr("btn_revert_cfg_dirty"), msg)
        if hasattr(self, "var_status"):
            self.var_status.set(msg)
        if hasattr(self, "var_locator_result"):
            self.var_locator_result.set(msg)

    def _focus_first_dirty_section(self):
        if not hasattr(self, "main_notebook"):
            return
        if not self._cfg_tab_meta:
            return
        for section_name, tab_widget, _ in self._cfg_tab_meta:
            if self._last_section_dirty.get(section_name, False):
                self._focus_dirty_section(section_name)
                return

    def _refresh_config_dirty_state(self):
        if self._suspend_cfg_dirty:
            return
        if not isinstance(self._baseline_config_snapshot, dict):
            self._baseline_config_snapshot = {}
        ui_snapshot = self._build_config_snapshot_from_ui()
        section_dirty = self._compute_section_dirty(
            ui_snapshot, self._baseline_config_snapshot
        )
        self._set_config_dirty(any(section_dirty.values()))
        self._update_config_tab_dirty_markers(section_dirty)
        self._update_config_dirty_summary(section_dirty)

    def _mark_config_dirty(self):
        if self._suspend_cfg_dirty:
            return
        self._refresh_config_dirty_state()

    def _bind_config_dirty_var(self, var):
        if var is None:
            return
        try:
            var.trace_add("write", lambda *_: self._mark_config_dirty())
        except Exception:
            pass

    def _bind_config_dirty_text(self, widget):
        if widget is None:
            return
        for event_name in ("<KeyRelease>", "<<Paste>>", "<<Cut>>"):
            try:
                widget.bind(event_name, lambda *_: self._mark_config_dirty(), add="+")
            except Exception:
                pass

    @staticmethod
    def _safe_positive_int(raw, default):
        try:
            value = int(str(raw).strip())
            return value if value > 0 else default
        except Exception:
            return default

    @staticmethod
    def _normalize_lines(lines):
        return [str(x).strip() for x in (lines or []) if str(x).strip()]

    def _read_text_lines(self, widget):
        if widget is None:
            return []
        try:
            raw = widget.get("1.0", "end").strip()
        except Exception:
            return []
        return self._normalize_lines(raw.splitlines())

    def _build_config_snapshot_from_ui(self):
        return {
            "kill_process_mode": self.var_kill_mode.get(),
            "log_folder": self.var_log_folder.get().strip() or "./logs",
            "timeout_seconds": self._safe_positive_int(
                self.var_timeout_seconds.get(), 60
            ),
            "pdf_wait_seconds": self._safe_positive_int(
                self.var_pdf_wait_seconds.get(), 15
            ),
            "ppt_timeout_seconds": self._safe_positive_int(
                self.var_ppt_timeout_seconds.get(), 180
            ),
            "ppt_pdf_wait_seconds": self._safe_positive_int(
                self.var_ppt_pdf_wait_seconds.get(), 30
            ),
            "office_reuse_app": bool(self.var_office_reuse_app.get()),
            "office_restart_every_n_files": self._safe_positive_int(
                self.var_office_restart_every_n_files.get(), 25
            ),
            "enable_sandbox": bool(self.var_enable_sandbox.get()),
            "temp_sandbox_root": self.var_temp_sandbox_root.get().strip(),
            "sandbox_min_free_gb": self._safe_positive_int(
                self.var_sandbox_min_free_gb.get(), 10
            ),
            "sandbox_low_space_policy": self.var_sandbox_low_space_policy.get() or "block",
            "enable_merge": bool(self.var_enable_merge.get()),
            "output_enable_pdf": bool(self.var_output_enable_pdf.get()),
            "output_enable_md": bool(self.var_output_enable_md.get()),
            "output_enable_merged": bool(self.var_output_enable_merged.get()),
            "output_enable_independent": bool(
                self.var_output_enable_independent.get()
            ),
            "merge_convert_submode": self.var_merge_convert_submode.get(),
            "merge_mode": self.var_merge_mode.get(),
            "merge_source": self.var_merge_source.get(),
            "enable_merge_index": bool(self.var_enable_merge_index.get()),
            "enable_merge_excel": bool(self.var_enable_merge_excel.get()),
            "max_merge_size_mb": self._safe_positive_int(
                self.var_max_merge_size_mb.get(), 80
            ),
            "merge_filename_pattern": self.var_merge_filename_pattern.get().strip()
            or "Merged_{category}_{timestamp}_{idx}",
            "enable_corpus_manifest": bool(self.var_enable_corpus_manifest.get()),
            "enable_markdown": bool(self.var_output_enable_md.get()),
            "markdown_strip_header_footer": bool(
                self.var_markdown_strip_header_footer.get()
            ),
            "markdown_structured_headings": bool(
                self.var_markdown_structured_headings.get()
            ),
            "enable_markdown_quality_report": bool(
                self.var_enable_markdown_quality_report.get()
            ),
            "enable_excel_json": bool(self.var_enable_excel_json.get()),
            "enable_chromadb_export": bool(self.var_enable_chromadb_export.get()),
            "enable_llm_delivery_hub": bool(self.var_enable_llm_delivery_hub.get()),
            "llm_delivery_root": self.var_llm_delivery_root.get().strip(),
            "llm_delivery_flatten": bool(self.var_llm_delivery_flatten.get()),
            "llm_delivery_include_pdf": bool(self.var_llm_delivery_include_pdf.get()),
            "mshelpviewer_folder_name": str(
                self.var_mshelpviewer_folder_name.get()
            ).strip()
            or "MSHelpViewer",
            "enable_mshelp_merge_output": bool(
                self.var_enable_mshelp_merge_output.get()
            ),
            "enable_mshelp_output_docx": bool(
                self.var_enable_mshelp_output_docx.get()
            ),
            "enable_mshelp_output_pdf": bool(
                self.var_enable_mshelp_output_pdf.get()
            ),
            "enable_incremental_mode": bool(self.var_enable_incremental_mode.get()),
            "incremental_verify_hash": bool(self.var_incremental_verify_hash.get()),
            "incremental_reprocess_renamed": bool(
                self.var_incremental_reprocess_renamed.get()
            ),
            "source_priority_skip_same_name_pdf": bool(
                self.var_source_priority_skip_same_name_pdf.get()
            ),
            "global_md5_dedup": bool(self.var_global_md5_dedup.get()),
            "enable_update_package": bool(self.var_enable_update_package.get()),
            "excluded_folders": self._read_text_lines(self.txt_excluded_folders),
            "price_keywords": self._read_text_lines(self.txt_price_keywords),
            "ui": {
                "tooltip_delay_ms": self._safe_positive_int(
                    self.var_tooltip_delay_ms.get(),
                    self.TOOLTIP_DEFAULTS["tooltip_delay_ms"],
                ),
                "tooltip_bg": str(self.var_tooltip_bg.get()).strip().upper()
                or self.TOOLTIP_DEFAULTS["tooltip_bg"],
                "tooltip_fg": str(self.var_tooltip_fg.get()).strip().upper()
                or self.TOOLTIP_DEFAULTS["tooltip_fg"],
                "tooltip_font_size": self._safe_positive_int(
                    self.var_tooltip_font_size.get(),
                    self.TOOLTIP_DEFAULTS["tooltip_font_size"],
                ),
                "tooltip_auto_theme": bool(self.var_tooltip_auto_theme.get()),
                "confirm_revert_dirty": bool(self.var_confirm_revert_dirty.get()),
            },
        }

    def _build_config_snapshot_from_cfg(self, cfg):
        ui_cfg = cfg.get("ui", {}) if isinstance(cfg.get("ui"), dict) else {}
        return {
            "kill_process_mode": cfg.get("kill_process_mode", KILL_MODE_AUTO),
            "log_folder": str(cfg.get("log_folder", "./logs")).strip() or "./logs",
            "timeout_seconds": self._safe_positive_int(cfg.get("timeout_seconds", 60), 60),
            "pdf_wait_seconds": self._safe_positive_int(
                cfg.get("pdf_wait_seconds", 15), 15
            ),
            "ppt_timeout_seconds": self._safe_positive_int(
                cfg.get("ppt_timeout_seconds", 180), 180
            ),
            "ppt_pdf_wait_seconds": self._safe_positive_int(
                cfg.get("ppt_pdf_wait_seconds", 30), 30
            ),
            "office_reuse_app": bool(cfg.get("office_reuse_app", True)),
            "office_restart_every_n_files": self._safe_positive_int(
                cfg.get("office_restart_every_n_files", 25), 25
            ),
            "enable_sandbox": bool(cfg.get("enable_sandbox", True)),
            "temp_sandbox_root": str(cfg.get("temp_sandbox_root", "")).strip(),
            "sandbox_min_free_gb": self._safe_positive_int(
                cfg.get("sandbox_min_free_gb", 10), 10
            ),
            "sandbox_low_space_policy": str(cfg.get("sandbox_low_space_policy", "block")).strip() or "block",
            "enable_merge": bool(cfg.get("enable_merge", True)),
            "output_enable_pdf": bool(cfg.get("output_enable_pdf", True)),
            "output_enable_md": bool(cfg.get("output_enable_md", True)),
            "output_enable_merged": bool(cfg.get("output_enable_merged", True)),
            "output_enable_independent": bool(
                cfg.get("output_enable_independent", False)
            ),
            "merge_convert_submode": str(
                cfg.get(
                    "merge_convert_submode",
                    MERGE_CONVERT_SUBMODE_MERGE_ONLY,
                )
            ),
            "merge_mode": cfg.get("merge_mode", MERGE_MODE_CATEGORY),
            "merge_source": cfg.get("merge_source", "source"),
            "enable_merge_index": bool(cfg.get("enable_merge_index", False)),
            "enable_merge_excel": bool(cfg.get("enable_merge_excel", False)),
            "max_merge_size_mb": self._safe_positive_int(
                cfg.get("max_merge_size_mb", 80), 80
            ),
            "merge_filename_pattern": cfg.get("merge_filename_pattern")
            or "Merged_{category}_{timestamp}_{idx}",
            "enable_corpus_manifest": bool(cfg.get("enable_corpus_manifest", True)),
            "enable_markdown": bool(cfg.get("output_enable_md", cfg.get("enable_markdown", True))),
            "markdown_strip_header_footer": bool(
                cfg.get("markdown_strip_header_footer", True)
            ),
            "markdown_structured_headings": bool(
                cfg.get("markdown_structured_headings", True)
            ),
            "enable_markdown_quality_report": bool(
                cfg.get("enable_markdown_quality_report", True)
            ),
            "enable_excel_json": bool(cfg.get("enable_excel_json", False)),
            "enable_chromadb_export": bool(cfg.get("enable_chromadb_export", False)),
            "enable_llm_delivery_hub": bool(cfg.get("enable_llm_delivery_hub", True)),
            "llm_delivery_root": cfg.get("llm_delivery_root", "") or "",
            "llm_delivery_flatten": bool(cfg.get("llm_delivery_flatten", False)),
            "llm_delivery_include_pdf": bool(cfg.get("llm_delivery_include_pdf", False)),
            "mshelpviewer_folder_name": str(
                cfg.get("mshelpviewer_folder_name", "MSHelpViewer")
            ).strip()
            or "MSHelpViewer",
            "enable_mshelp_merge_output": bool(
                cfg.get("enable_mshelp_merge_output", True)
            ),
            "enable_mshelp_output_docx": bool(
                cfg.get("enable_mshelp_output_docx", False)
            ),
            "enable_mshelp_output_pdf": bool(
                cfg.get("enable_mshelp_output_pdf", False)
            ),
            "enable_incremental_mode": bool(cfg.get("enable_incremental_mode", False)),
            "incremental_verify_hash": bool(cfg.get("incremental_verify_hash", False)),
            "incremental_reprocess_renamed": bool(
                cfg.get("incremental_reprocess_renamed", False)
            ),
            "source_priority_skip_same_name_pdf": bool(
                cfg.get("source_priority_skip_same_name_pdf", False)
            ),
            "global_md5_dedup": bool(cfg.get("global_md5_dedup", False)),
            "enable_update_package": bool(cfg.get("enable_update_package", True)),
            "excluded_folders": self._normalize_lines(cfg.get("excluded_folders", [])),
            "price_keywords": self._normalize_lines(cfg.get("price_keywords", [])),
            "ui": {
                "tooltip_delay_ms": self._safe_positive_int(
                    ui_cfg.get("tooltip_delay_ms", self.TOOLTIP_DEFAULTS["tooltip_delay_ms"]),
                    self.TOOLTIP_DEFAULTS["tooltip_delay_ms"],
                ),
                "tooltip_bg": str(
                    ui_cfg.get("tooltip_bg", self.TOOLTIP_DEFAULTS["tooltip_bg"])
                ).strip().upper(),
                "tooltip_fg": str(
                    ui_cfg.get("tooltip_fg", self.TOOLTIP_DEFAULTS["tooltip_fg"])
                ).strip().upper(),
                "tooltip_font_size": self._safe_positive_int(
                    ui_cfg.get("tooltip_font_size", self.TOOLTIP_DEFAULTS["tooltip_font_size"]),
                    self.TOOLTIP_DEFAULTS["tooltip_font_size"],
                ),
                "tooltip_auto_theme": bool(
                    ui_cfg.get("tooltip_auto_theme", self.TOOLTIP_DEFAULTS["tooltip_auto_theme"])
                ),
                "confirm_revert_dirty": bool(ui_cfg.get("confirm_revert_dirty", True)),
            },
        }

    def _refresh_config_dirty_from_file(self):
        if self._suspend_cfg_dirty:
            return
        if not os.path.exists(self.config_path):
            return
        try:
            with open(self.config_path, "r", encoding="utf-8") as f:
                cfg = json.load(f)
        except Exception:
            return
        self._baseline_config_snapshot = self._build_config_snapshot_from_cfg(cfg)
        self._refresh_config_dirty_state()

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
                (
                    "var_office_restart_every_n_files",
                    "ent_office_restart_every_n_files",
                    "lbl_office_restart_every",
                ),
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

    def _set_text_widget_lines(self, widget, lines):
        if widget is None:
            return
        widget.delete("1.0", "end")
        if lines:
            widget.insert("1.0", "\n".join(lines))

    def _reset_config_section_defaults(self, section_name):
        section = str(section_name or "").strip().lower()
        section_title_key = None

        if section == "shared":
            self.var_kill_mode.set(KILL_MODE_AUTO)
            self.var_log_folder.set("./logs")
            section_title_key = "grp_cfg_shared"
        elif section == "convert":
            self.var_timeout_seconds.set("60")
            self.var_pdf_wait_seconds.set("15")
            self.var_ppt_timeout_seconds.set("180")
            self.var_ppt_pdf_wait_seconds.set("30")
            self.var_office_reuse_app.set(1)
            self.var_office_restart_every_n_files.set("25")
            self.var_enable_sandbox.set(1)
            self.var_temp_sandbox_root.set("")
            self.var_sandbox_min_free_gb.set("10")
            self.var_sandbox_low_space_policy.set("block")
            self._on_toggle_sandbox()
            section_title_key = "grp_cfg_convert"
        elif section == "ai":
            self.var_enable_corpus_manifest.set(1)
            self.var_enable_markdown.set(1)
            self.var_markdown_strip_header_footer.set(1)
            self.var_markdown_structured_headings.set(1)
            self.var_enable_markdown_quality_report.set(1)
            self.var_enable_excel_json.set(0)
            self.var_enable_chromadb_export.set(0)
            self.var_enable_llm_delivery_hub.set(1)
            self.var_llm_delivery_root.set("")
            self.var_llm_delivery_flatten.set(0)
            self.var_llm_delivery_include_pdf.set(0)
            self.var_enable_upload_readme.set(1)
            self.var_enable_upload_json_manifest.set(1)
            self.var_upload_dedup_merged.set(1)
            self.var_mshelpviewer_folder_name.set("MSHelpViewer")
            self.var_enable_mshelp_merge_output.set(1)
            self.var_enable_mshelp_output_docx.set(0)
            self.var_enable_mshelp_output_pdf.set(0)
            section_title_key = "grp_cfg_ai"
        elif section == "incremental":
            self.var_enable_incremental_mode.set(0)
            self.var_incremental_verify_hash.set(0)
            self.var_incremental_reprocess_renamed.set(0)
            self.var_source_priority_skip_same_name_pdf.set(0)
            self.var_global_md5_dedup.set(0)
            self.var_enable_update_package.set(1)
            section_title_key = "grp_cfg_incremental"
        elif section == "merge":
            self.var_enable_merge.set(1)
            self.var_output_enable_pdf.set(1)
            self.var_output_enable_md.set(1)
            self.var_output_enable_merged.set(1)
            self.var_output_enable_independent.set(0)
            self.var_merge_convert_submode.set(MERGE_CONVERT_SUBMODE_MERGE_ONLY)
            self.var_merge_mode.set(MERGE_MODE_CATEGORY)
            self.var_merge_source.set("source")
            self.var_enable_merge_index.set(0)
            self.var_enable_merge_excel.set(0)
            self.var_max_merge_size_mb.set("80")
            section_title_key = "grp_cfg_merge"
        elif section == "ui":
            self.reset_tooltip_settings()
            section_title_key = "grp_cfg_ui"
        elif section == "rules":
            self._set_text_widget_lines(
                self.txt_excluded_folders, ["temp", "backup", "archive"]
            )
            self._set_text_widget_lines(
                self.txt_price_keywords, ["报价", "价格表", "Price", "Quotation"]
            )
            section_title_key = "grp_cfg_rules"
        else:
            return

        self._on_run_mode_change()
        self._on_toggle_sandbox()
        self._on_toggle_incremental_mode()
        self.validate_runtime_inputs(silent=True, scope="all")
        self._refresh_config_dirty_state()

        section_title = self.tr(section_title_key) if section_title_key else section
        reset_message = self.tr("msg_cfg_section_reset").format(section_title)
        if hasattr(self, "var_status"):
            self.var_status.set(reset_message)
        if hasattr(self, "var_locator_result"):
            self.var_locator_result.set(reset_message)

    def reset_tooltip_settings(self):
        self.var_tooltip_delay_ms.set(str(self.TOOLTIP_DEFAULTS["tooltip_delay_ms"]))
        self.var_tooltip_font_size.set(str(self.TOOLTIP_DEFAULTS["tooltip_font_size"]))
        self.var_tooltip_bg.set(self.TOOLTIP_DEFAULTS["tooltip_bg"])
        self.var_tooltip_fg.set(self.TOOLTIP_DEFAULTS["tooltip_fg"])
        self.var_tooltip_auto_theme.set(1 if self.TOOLTIP_DEFAULTS["tooltip_auto_theme"] else 0)
        if hasattr(self, "var_confirm_revert_dirty"):
            self.var_confirm_revert_dirty.set(1)
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

    # ===================== Config read/write =====================

    def _load_config_to_ui(self):
        """Load values from config.json into UI controls."""
        self._suspend_cfg_dirty = True
        if hasattr(self, "var_profile_active_path"):
            self.var_profile_active_path.set(self.config_path)
        if not os.path.exists(self.config_path):
            self._suspend_cfg_dirty = False
            return
        try:
            with open(self.config_path, "r", encoding="utf-8") as f:
                cfg = json.load(f)
        except Exception:
            self._suspend_cfg_dirty = False
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
        if hasattr(self, "var_confirm_revert_dirty"):
            self.var_confirm_revert_dirty.set(
                1 if ui_cfg.get("confirm_revert_dirty", True) else 0
            )
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

        # 鎭㈠绐楀彛灏哄涓庝綅缃紙濡傛灉宸蹭繚瀛橈級
        try:
            win_geo = ui_cfg.get("window_geometry")
            if isinstance(win_geo, str) and win_geo:
                self.geometry(win_geo)
            win_state = ui_cfg.get("window_state")
            if isinstance(win_state, str) and win_state:
                self.state(win_state)
        except Exception:
            pass

        if hasattr(self, "var_app_mode"):
            self.var_app_mode.set(cfg.get("app_mode", "classic"))

        # Runtime parameters
        
        is_win = (sys.platform == "win32")
        is_mac = (sys.platform == "darwin")
        
        def _get_os_path(key_base):
            if is_win:
                val = cfg.get(f"{key_base}_win")
            elif is_mac:
                val = cfg.get(f"{key_base}_mac")
            else:
                val = None
            if not val:
                 val = cfg.get(key_base)
            return val
            
        src_val = _get_os_path("source_folder")
        src_list_raw = _get_os_path("source_folders") # Assume list if present
        
        src_list = []
        if src_list_raw and isinstance(src_list_raw, list):
             src_list = src_list_raw
        elif src_val:
             src_list = [src_val]
        else:
             # Fallback to generic if OS specific failed
             src_list = cfg.get("source_folders", [])
             if not src_list:
                  single = cfg.get("source_folder", "")
                  if single:
                      src_list = [single]
        
        self.source_folders_list = src_list
        self.lst_source_folders.delete(0, END)
        for p in self.source_folders_list:
            self.lst_source_folders.insert(END, p)
        
        if self.source_folders_list:
            self.var_source_folder.set(self.source_folders_list[0])
        else:
            self.var_source_folder.set("")
            
        self.var_target_folder.set(_get_os_path("target_folder") or "")
        self.var_enable_sandbox.set(1 if cfg.get("enable_sandbox", True) else 0)
        self.var_temp_sandbox_root.set(_get_os_path("temp_sandbox_root") or "")
        self.var_sandbox_min_free_gb.set(str(cfg.get("sandbox_min_free_gb", 10)))
        self.var_sandbox_low_space_policy.set(cfg.get("sandbox_low_space_policy", "block"))
        self.var_enable_llm_delivery_hub.set(
            1 if cfg.get("enable_llm_delivery_hub", True) else 0
        )
        self.var_llm_delivery_root.set(_get_os_path("llm_delivery_root") or "")
        self.var_llm_delivery_flatten.set(
            1 if cfg.get("llm_delivery_flatten", False) else 0
        )
        self.var_llm_delivery_include_pdf.set(
            1 if cfg.get("llm_delivery_include_pdf", False) else 0
        )
        self.var_enable_upload_readme.set(
            1 if cfg.get("enable_upload_readme", True) else 0
        )
        self.var_enable_upload_json_manifest.set(
            1 if cfg.get("enable_upload_json_manifest", True) else 0
        )
        self.var_upload_dedup_merged.set(
            1 if cfg.get("upload_dedup_merged", True) else 0
        )
        self.var_enable_corpus_manifest.set(
            1 if cfg.get("enable_corpus_manifest", True) else 0
        )
        self.var_enable_markdown.set(
            1 if cfg.get("output_enable_md", cfg.get("enable_markdown", True)) else 0
        )
        self.var_markdown_strip_header_footer.set(
            1 if cfg.get("markdown_strip_header_footer", True) else 0
        )
        self.var_markdown_structured_headings.set(
            1 if cfg.get("markdown_structured_headings", True) else 0
        )
        self.var_enable_markdown_quality_report.set(
            1 if cfg.get("enable_markdown_quality_report", True) else 0
        )
        self.var_enable_excel_json.set(
            1 if cfg.get("enable_excel_json", False) else 0
        )
        self.var_enable_chromadb_export.set(
            1 if cfg.get("enable_chromadb_export", False) else 0
        )
        self._sync_markdown_master_with_global_output()
        self.var_mshelpviewer_folder_name.set(
            str(cfg.get("mshelpviewer_folder_name", "MSHelpViewer") or "MSHelpViewer")
        )
        self.var_enable_mshelp_merge_output.set(
            1 if cfg.get("enable_mshelp_merge_output", True) else 0
        )
        self.var_enable_mshelp_output_docx.set(
            1 if cfg.get("enable_mshelp_output_docx", False) else 0
        )
        self.var_enable_mshelp_output_pdf.set(
            1 if cfg.get("enable_mshelp_output_pdf", False) else 0
        )
        self.var_enable_incremental_mode.set(
            1 if cfg.get("enable_incremental_mode", False) else 0
        )
        self.var_incremental_verify_hash.set(
            1 if cfg.get("incremental_verify_hash", False) else 0
        )
        self.var_incremental_reprocess_renamed.set(
            1 if cfg.get("incremental_reprocess_renamed", False) else 0
        )
        self.var_source_priority_skip_same_name_pdf.set(
            1 if cfg.get("source_priority_skip_same_name_pdf", False) else 0
        )
        self.var_global_md5_dedup.set(
            1 if cfg.get("global_md5_dedup", False) else 0
        )
        self.var_enable_update_package.set(
            1 if cfg.get("enable_update_package", True) else 0
        )

        self.var_enable_merge.set(1 if cfg.get("enable_merge", True) else 0)
        self.var_output_enable_pdf.set(1 if cfg.get("output_enable_pdf", True) else 0)
        self.var_output_enable_md.set(1 if cfg.get("output_enable_md", True) else 0)
        self.var_output_enable_merged.set(
            1 if cfg.get("output_enable_merged", True) else 0
        )
        self.var_output_enable_independent.set(
            1 if cfg.get("output_enable_independent", False) else 0
        )
        self.var_merge_convert_submode.set(
            cfg.get(
                "merge_convert_submode", MERGE_CONVERT_SUBMODE_MERGE_ONLY
            )
        )
        self.var_merge_mode.set(cfg.get("merge_mode", MERGE_MODE_CATEGORY))
        self.var_merge_source.set(cfg.get("merge_source", "source"))
        self.var_enable_merge_index.set(1 if cfg.get("enable_merge_index", False) else 0)
        self.var_enable_merge_excel.set(1 if cfg.get("enable_merge_excel", False) else 0)

        # 杩愯妯″紡 / 瀛愭ā寮?/ 绛栫暐锛堜綔涓洪粯璁わ級
        self.var_run_mode.set(cfg.get("run_mode", MODE_CONVERT_THEN_MERGE))
        self.var_collect_mode.set(cfg.get("collect_mode", COLLECT_MODE_COPY_AND_INDEX))
        self.var_strategy.set(cfg.get("content_strategy", "standard"))

        # 寮曟搸 & 杩涚▼绛栫暐
        default_engine = cfg.get("default_engine", ENGINE_WPS)
        if default_engine not in (ENGINE_WPS, ENGINE_MS):
            default_engine = ENGINE_WPS
        self.var_engine.set(default_engine)

        kill_mode = cfg.get("kill_process_mode", KILL_MODE_AUTO)
        if kill_mode not in (KILL_MODE_AUTO, KILL_MODE_KEEP):
            kill_mode = KILL_MODE_AUTO
        self.var_kill_mode.set(kill_mode)

        # 閰嶇疆绠＄悊椤?
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
        self.var_office_reuse_app.set(1 if cfg.get("office_reuse_app", True) else 0)
        self.var_office_restart_every_n_files.set(
            str(cfg.get("office_restart_every_n_files", 25))
        )
        self.var_max_merge_size_mb.set(str(cfg.get("max_merge_size_mb", 80)))
        self.var_merge_filename_pattern.set(
            cfg.get("merge_filename_pattern") or "Merged_{category}_{timestamp}_{idx}"
        )

        # 鑱斿姩鍒锋柊
        self._on_run_mode_change()
        self._on_toggle_sandbox()
        self._on_toggle_incremental_mode()
        self.refresh_locator_maps()
        self._normalize_numeric_var(self.var_timeout_seconds)
        self._normalize_numeric_var(self.var_pdf_wait_seconds)
        self._normalize_numeric_var(self.var_ppt_timeout_seconds)
        self._normalize_numeric_var(self.var_ppt_pdf_wait_seconds)
        self._normalize_numeric_var(self.var_office_restart_every_n_files)
        self._normalize_numeric_var(self.var_max_merge_size_mb)
        self._normalize_short_id_var(self.var_locator_short_id)
        self._normalize_date_var(self.var_date_str)
        self.validate_runtime_inputs(silent=True, scope="all")
        self._suspend_cfg_dirty = False
        self._refresh_config_dirty_from_file()
        if self.profile_manager_win is not None and self.profile_manager_win.winfo_exists():
            self._refresh_profile_tree()
        if self.load_profile_dialog is not None and self.load_profile_dialog.winfo_exists():
            self._refresh_load_profile_tree()
        if hasattr(self, "_update_task_tab_for_app_mode"):
            self._update_task_tab_for_app_mode()

    def _load_config_for_write(self):
        cfg = {}
        if os.path.exists(self.config_path):
            try:
                with open(self.config_path, "r", encoding="utf-8") as f:
                    cfg = json.load(f)
            except Exception:
                cfg = {}
        return cfg

    def _write_config_sections_to_cfg(self, cfg, section_names):
        sections = set(section_names or [])
        if "shared" in sections:
            cfg["kill_process_mode"] = self.var_kill_mode.get()
            cfg["log_folder"] = self.var_log_folder.get().strip() or "./logs"

        if "convert" in sections:
            cfg["timeout_seconds"] = self._safe_positive_int(
                self.var_timeout_seconds.get(), 60
            )
            cfg["pdf_wait_seconds"] = self._safe_positive_int(
                self.var_pdf_wait_seconds.get(), 15
            )
            cfg["ppt_timeout_seconds"] = self._safe_positive_int(
                self.var_ppt_timeout_seconds.get(), 180
            )
            cfg["ppt_pdf_wait_seconds"] = self._safe_positive_int(
                self.var_ppt_pdf_wait_seconds.get(), 30
            )
            cfg["office_reuse_app"] = bool(self.var_office_reuse_app.get())
            cfg["office_restart_every_n_files"] = self._safe_positive_int(
                self.var_office_restart_every_n_files.get(), 25
            )
            cfg["enable_sandbox"] = bool(self.var_enable_sandbox.get())
            cfg["temp_sandbox_root"] = self.var_temp_sandbox_root.get().strip()
            cfg["sandbox_min_free_gb"] = self._safe_positive_int(
                self.var_sandbox_min_free_gb.get(), 10
            )
            cfg["sandbox_low_space_policy"] = self.var_sandbox_low_space_policy.get() or "block"

        if "ai" in sections:
            cfg["enable_corpus_manifest"] = bool(self.var_enable_corpus_manifest.get())
            cfg.pop("enable_markdown", None)
            cfg["markdown_strip_header_footer"] = bool(
                self.var_markdown_strip_header_footer.get()
            )
            cfg["markdown_structured_headings"] = bool(
                self.var_markdown_structured_headings.get()
            )
            cfg["enable_markdown_quality_report"] = bool(
                self.var_enable_markdown_quality_report.get()
            )
            cfg["enable_excel_json"] = bool(self.var_enable_excel_json.get())
            cfg["enable_chromadb_export"] = bool(self.var_enable_chromadb_export.get())
            cfg["mshelpviewer_folder_name"] = (
                self.var_mshelpviewer_folder_name.get().strip() or "MSHelpViewer"
            )
            cfg["enable_mshelp_merge_output"] = bool(
                self.var_enable_mshelp_merge_output.get()
            )
            cfg["enable_mshelp_output_docx"] = bool(
                self.var_enable_mshelp_output_docx.get()
            )
            cfg["enable_mshelp_output_pdf"] = bool(
                self.var_enable_mshelp_output_pdf.get()
            )

        if "incremental" in sections:
            cfg["enable_incremental_mode"] = bool(self.var_enable_incremental_mode.get())
            cfg["incremental_verify_hash"] = bool(self.var_incremental_verify_hash.get())
            cfg["incremental_reprocess_renamed"] = bool(
                self.var_incremental_reprocess_renamed.get()
            )
            cfg["source_priority_skip_same_name_pdf"] = bool(
                self.var_source_priority_skip_same_name_pdf.get()
            )
            cfg["global_md5_dedup"] = bool(self.var_global_md5_dedup.get())
            cfg["enable_update_package"] = bool(self.var_enable_update_package.get())

        if "merge" in sections:
            cfg["enable_merge"] = bool(self.var_enable_merge.get())
            cfg["output_enable_pdf"] = bool(self.var_output_enable_pdf.get())
            cfg["output_enable_md"] = bool(self.var_output_enable_md.get())
            cfg["output_enable_merged"] = bool(
                self.var_output_enable_merged.get()
            )
            cfg["output_enable_independent"] = bool(
                self.var_output_enable_independent.get()
            )
            cfg["merge_convert_submode"] = self.var_merge_convert_submode.get()
            cfg["merge_mode"] = self.var_merge_mode.get()
            cfg["merge_source"] = (
                "target"
                if self.var_run_mode.get() == MODE_CONVERT_THEN_MERGE
                else self.var_merge_source.get()
            )
            cfg["enable_merge_index"] = bool(self.var_enable_merge_index.get())
            cfg["enable_merge_excel"] = bool(self.var_enable_merge_excel.get())
            cfg["max_merge_size_mb"] = self._safe_positive_int(
                self.var_max_merge_size_mb.get(), 80
            )
            cfg["merge_filename_pattern"] = (
                self.var_merge_filename_pattern.get().strip()
                or "Merged_{category}_{timestamp}_{idx}"
            )

        if "rules" in sections:
            cfg["excluded_folders"] = self._read_text_lines(self.txt_excluded_folders)
            cfg["price_keywords"] = self._read_text_lines(self.txt_price_keywords)

        if "ui" in sections:
            self.apply_tooltip_settings(silent=True)
            cfg["ui"] = {
                "tooltip_delay_ms": self.tooltip_delay_ms,
                "tooltip_bg": self.tooltip_bg,
                "tooltip_fg": self.tooltip_fg,
                "tooltip_font_family": self.tooltip_font_family,
                "tooltip_font_size": self.tooltip_font_size,
                "tooltip_auto_theme": self.tooltip_auto_theme,
                "confirm_revert_dirty": bool(self.var_confirm_revert_dirty.get()),
                # 璁板繂绐楀彛灏哄涓庝綅缃?
                "window_geometry": self.geometry(),
                "window_state": self.state(),
            }
        return cfg

    def _save_specific_config_section(self, section_name, show_msg=True):
        section = str(section_name or "").strip().lower()
        valid_sections = {
            "shared",
            "convert",
            "ai",
            "incremental",
            "merge",
            "ui",
            "rules",
        }
        if section not in valid_sections:
            return
        cfg = self._load_config_for_write()
        self._write_config_sections_to_cfg(cfg, [section])
        try:
            with open(self.config_path, "w", encoding="utf-8") as f:
                json.dump(cfg, f, indent=4, ensure_ascii=False)
            self._baseline_config_snapshot = self._build_config_snapshot_from_cfg(cfg)
            self._refresh_config_dirty_state()
            section_title = ", ".join(self._get_cfg_section_titles([section]))
            msg = self.tr("msg_cfg_section_saved").format(section_title)
            if show_msg:
                messagebox.showinfo(self.tr("btn_save_cfg_section"), msg)
            if hasattr(self, "var_status"):
                self.var_status.set(msg)
            if hasattr(self, "var_locator_result"):
                self.var_locator_result.set(msg)
        except Exception as e:
            if show_msg:
                messagebox.showerror(
                    self.tr("btn_save_cfg_section"),
                    self.tr("msg_save_fail").format(e),
                )

    def _on_close_main_window(self):
        """Persist UI geometry settings and close the window."""
        try:
            cfg = self._load_config_for_write()
            self._write_config_sections_to_cfg(cfg, ["ui"])
            with open(self.config_path, "w", encoding="utf-8") as f:
                json.dump(cfg, f, indent=4, ensure_ascii=False)
        except Exception:
            pass
        try:
            self.destroy()
        except Exception:
            pass

    def _save_dirty_config_sections(self, show_msg=True):
        dirty_sections = []
        for section_name, _, _ in self._cfg_tab_meta:
            if self._last_section_dirty.get(section_name, False):
                dirty_sections.append(section_name)
        if not dirty_sections:
            self._refresh_config_dirty_state()
            msg = self.tr("msg_save_dirty_none")
            if show_msg:
                messagebox.showinfo(self.tr("btn_save_cfg_dirty"), msg)
            if hasattr(self, "var_status"):
                self.var_status.set(msg)
            if hasattr(self, "var_locator_result"):
                self.var_locator_result.set(msg)
            return

        cfg = self._load_config_for_write()
        self._write_config_sections_to_cfg(cfg, dirty_sections)

        try:
            with open(self.config_path, "w", encoding="utf-8") as f:
                json.dump(cfg, f, indent=4, ensure_ascii=False)
            self._baseline_config_snapshot = self._build_config_snapshot_from_cfg(cfg)
            self._refresh_config_dirty_state()
            saved_sections_text = ", ".join(self._get_cfg_section_titles(dirty_sections))
            msg = self.tr("msg_save_dirty_sections").format(saved_sections_text)
            if show_msg:
                messagebox.showinfo(self.tr("btn_save_cfg_dirty"), msg)
            if hasattr(self, "var_status"):
                self.var_status.set(msg)
            if hasattr(self, "var_locator_result"):
                self.var_locator_result.set(msg)
        except Exception as e:
            if show_msg:
                messagebox.showerror(
                    self.tr("btn_save_cfg_dirty"),
                    self.tr("msg_save_fail").format(e),
                )

    def _compose_config_from_ui(self, cfg, scope="all"):
        cfg = cfg if isinstance(cfg, dict) else {}
        scope = "mode" if str(scope).lower() == "mode" else "all"
        mode = self.var_run_mode.get()
        write_convert = scope == "all" or mode in (MODE_CONVERT_ONLY, MODE_CONVERT_THEN_MERGE)
        write_merge = scope == "all" or mode in (MODE_CONVERT_THEN_MERGE, MODE_MERGE_ONLY)
        write_collect = scope == "all" or mode == MODE_COLLECT_ONLY
        write_mshelp = scope == "all" or mode == MODE_MSHELP_ONLY
        write_rules = scope == "all" or mode in (
            MODE_CONVERT_ONLY,
            MODE_CONVERT_THEN_MERGE,
            MODE_COLLECT_ONLY,
        )

        is_win = (sys.platform == "win32")
        is_mac = (sys.platform == "darwin")

        if is_win:
             cfg["source_folders_win"] = self.source_folders_list
             cfg["source_folder_win"] = self.source_folders_list[0] if self.source_folders_list else ""
             cfg["target_folder_win"] = self.var_target_folder.get().strip()
             if write_convert:
                 cfg["temp_sandbox_root_win"] = self.var_temp_sandbox_root.get().strip()
                 cfg["llm_delivery_root_win"] = self.var_llm_delivery_root.get().strip()
        elif is_mac:
             cfg["source_folders_mac"] = self.source_folders_list
             cfg["source_folder_mac"] = self.source_folders_list[0] if self.source_folders_list else ""
             cfg["target_folder_mac"] = self.var_target_folder.get().strip()
             if write_convert:
                 cfg["temp_sandbox_root_mac"] = self.var_temp_sandbox_root.get().strip()
                 cfg["llm_delivery_root_mac"] = self.var_llm_delivery_root.get().strip()

        if hasattr(self, "var_app_mode"):
            cfg["app_mode"] = self.var_app_mode.get()
        else:
            cfg["app_mode"] = "classic"
        cfg["source_folders"] = self.source_folders_list
        cfg["source_folder"] = self.source_folders_list[0] if self.source_folders_list else ""
        cfg["target_folder"] = self.var_target_folder.get().strip()
        cfg["output_enable_pdf"] = bool(self.var_output_enable_pdf.get())
        cfg["output_enable_md"] = bool(self.var_output_enable_md.get())
        cfg["output_enable_merged"] = bool(self.var_output_enable_merged.get())
        cfg["output_enable_independent"] = bool(
            self.var_output_enable_independent.get()
        )
        cfg["merge_convert_submode"] = self.var_merge_convert_submode.get()
        cfg["enable_corpus_manifest"] = bool(self.var_enable_corpus_manifest.get())
        cfg.pop("enable_markdown", None)
        cfg["markdown_strip_header_footer"] = bool(
            self.var_markdown_strip_header_footer.get()
        )
        cfg["markdown_structured_headings"] = bool(
            self.var_markdown_structured_headings.get()
        )
        cfg["enable_markdown_quality_report"] = bool(
            self.var_enable_markdown_quality_report.get()
        )
        cfg["enable_excel_json"] = bool(self.var_enable_excel_json.get())
        cfg["enable_chromadb_export"] = bool(self.var_enable_chromadb_export.get())
        if write_mshelp:
            cfg["mshelpviewer_folder_name"] = (
                self.var_mshelpviewer_folder_name.get().strip() or "MSHelpViewer"
            )
            cfg["enable_mshelp_merge_output"] = bool(
                self.var_enable_mshelp_merge_output.get()
            )
            cfg["enable_mshelp_output_docx"] = bool(
                self.var_enable_mshelp_output_docx.get()
            )
            cfg["enable_mshelp_output_pdf"] = bool(
                self.var_enable_mshelp_output_pdf.get()
            )
        if write_convert:
            cfg["enable_sandbox"] = bool(self.var_enable_sandbox.get())
            cfg["temp_sandbox_root"] = self.var_temp_sandbox_root.get().strip()
            cfg["sandbox_min_free_gb"] = self._safe_positive_int(
                self.var_sandbox_min_free_gb.get(), 10
            )
            cfg["sandbox_low_space_policy"] = self.var_sandbox_low_space_policy.get() or "block"
            cfg["enable_llm_delivery_hub"] = bool(self.var_enable_llm_delivery_hub.get())
            cfg["llm_delivery_root"] = self.var_llm_delivery_root.get().strip()
            cfg["llm_delivery_flatten"] = bool(self.var_llm_delivery_flatten.get())
            cfg["llm_delivery_include_pdf"] = bool(self.var_llm_delivery_include_pdf.get())
            cfg["enable_upload_readme"] = bool(self.var_enable_upload_readme.get())
            cfg["enable_upload_json_manifest"] = bool(self.var_enable_upload_json_manifest.get())
            cfg["upload_dedup_merged"] = bool(self.var_upload_dedup_merged.get())
            cfg["enable_incremental_mode"] = bool(
                self.var_enable_incremental_mode.get()
            )
            cfg["incremental_verify_hash"] = bool(
                self.var_incremental_verify_hash.get()
            )
            cfg["incremental_reprocess_renamed"] = bool(
                self.var_incremental_reprocess_renamed.get()
            )
            cfg["source_priority_skip_same_name_pdf"] = bool(
                self.var_source_priority_skip_same_name_pdf.get()
            )
            cfg["global_md5_dedup"] = bool(self.var_global_md5_dedup.get())
            cfg["enable_update_package"] = bool(self.var_enable_update_package.get())

        if write_merge:
            cfg["enable_merge"] = bool(self.var_enable_merge.get())
            cfg["output_enable_pdf"] = bool(self.var_output_enable_pdf.get())
            cfg["output_enable_md"] = bool(self.var_output_enable_md.get())
            cfg["output_enable_merged"] = bool(
                self.var_output_enable_merged.get()
            )
            cfg["output_enable_independent"] = bool(
                self.var_output_enable_independent.get()
            )
            cfg["merge_convert_submode"] = self.var_merge_convert_submode.get()
            cfg["merge_mode"] = self.var_merge_mode.get()
            cfg["merge_source"] = (
                "target"
                if mode == MODE_CONVERT_THEN_MERGE
                else self.var_merge_source.get()
            )
            cfg["enable_merge_index"] = bool(self.var_enable_merge_index.get())
            cfg["enable_merge_excel"] = bool(self.var_enable_merge_excel.get())
            cfg["max_merge_size_mb"] = self._safe_positive_int(
                self.var_max_merge_size_mb.get(), 80
            )
            cfg["merge_filename_pattern"] = (
                self.var_merge_filename_pattern.get().strip()
                or "Merged_{category}_{timestamp}_{idx}"
            )

        cfg["run_mode"] = self.var_run_mode.get()
        if write_collect:
            cfg["collect_mode"] = self.var_collect_mode.get()
        if write_convert:
            cfg["content_strategy"] = self.var_strategy.get()

        if write_convert:
            cfg["default_engine"] = self.var_engine.get()
        cfg["kill_process_mode"] = self.var_kill_mode.get()
        cfg["log_folder"] = self.var_log_folder.get().strip() or "./logs"

        if write_rules:
            excluded_text = self.txt_excluded_folders.get("1.0", "end").strip()
            cfg["excluded_folders"] = [
                line.strip() for line in excluded_text.splitlines() if line.strip()
            ]

            kw_text = self.txt_price_keywords.get("1.0", "end").strip()
            cfg["price_keywords"] = [
                line.strip() for line in kw_text.splitlines() if line.strip()
            ]

        def _to_int(var, default):
            try:
                v = int(var.get().strip())
                return v if v > 0 else default
            except Exception:
                return default

        if write_convert:
            cfg["timeout_seconds"] = _to_int(self.var_timeout_seconds, 60)
            cfg["pdf_wait_seconds"] = _to_int(self.var_pdf_wait_seconds, 15)
            cfg["ppt_timeout_seconds"] = _to_int(self.var_ppt_timeout_seconds, 180)
            cfg["ppt_pdf_wait_seconds"] = _to_int(self.var_ppt_pdf_wait_seconds, 30)
            cfg["office_reuse_app"] = bool(self.var_office_reuse_app.get())
            cfg["office_restart_every_n_files"] = _to_int(
                self.var_office_restart_every_n_files, 25
            )
        if write_merge:
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
            "confirm_revert_dirty": bool(self.var_confirm_revert_dirty.get()),
        }
        return cfg

    def _save_settings_to_file(self, show_msg=True, scope="all"):
        """Save current UI values into the active config file."""
        cfg = self._compose_config_from_ui(self._load_config_for_write(), scope=scope)
        try:
            with open(self.config_path, "w", encoding="utf-8") as f:
                json.dump(cfg, f, indent=4, ensure_ascii=False)
            self._baseline_config_snapshot = self._build_config_snapshot_from_cfg(cfg)
            self._refresh_config_dirty_state()
            if show_msg:
                messagebox.showinfo(self.tr("btn_save_cfg"), self.tr("msg_save_ok"))
        except Exception as e:
            if show_msg:
                messagebox.showerror(
                    self.tr("btn_save_cfg"),
                    self.tr("msg_save_fail").format(e),
                )

    def _reload_config_from_file(self, show_msg=True, confirm_dirty=True):
        if confirm_dirty and self.cfg_dirty:
            confirm = messagebox.askyesno(
                self.tr("btn_load_cfg"),
                self.tr("msg_confirm_load_config_dirty"),
            )
            if not confirm:
                return
        if not os.path.exists(self.config_path):
            if show_msg:
                messagebox.showerror(
                    self.tr("btn_load_cfg"),
                    self.tr("msg_load_missing").format(self.config_path),
                )
            return
        try:
            with open(self.config_path, "r", encoding="utf-8") as f:
                json.load(f)
            self._load_config_to_ui()
            if show_msg:
                messagebox.showinfo(self.tr("btn_load_cfg"), self.tr("msg_load_ok"))
            if hasattr(self, "var_status"):
                self.var_status.set(self.tr("msg_load_ok"))
            if hasattr(self, "var_locator_result"):
                self.var_locator_result.set(self.tr("msg_load_ok"))
        except Exception as e:
            if show_msg:
                messagebox.showerror(
                    self.tr("btn_load_cfg"),
                    self.tr("msg_load_fail").format(e),
                )

    # ===================== 鏃ュ織 & 鐘舵€?====================
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
        self._ui_running = bool(running)
        if running:
            if hasattr(self, "btn_start"): self.btn_start.configure(state="disabled")
            if hasattr(self, "btn_stop"): self.btn_stop.configure(state="normal")
            if hasattr(self, "btn_task_stop"): self.btn_task_stop.configure(state="normal")
            if hasattr(self, "btn_save_cfg"): self.btn_save_cfg.configure(state="disabled")
            if hasattr(self, "btn_load_cfg"): self.btn_load_cfg.configure(state="disabled")
            if hasattr(self, "btn_manage_profiles"): self.btn_manage_profiles.configure(state="disabled")
            if hasattr(self, "btn_manage_profiles"): self.btn_manage_profiles.configure(state="disabled")
            for btn_name in (
                "btn_task_create",
                "btn_task_edit",
                "btn_task_delete",
                "btn_task_run",
                "btn_task_resume",
                "btn_task_refresh",
                "btn_task_load_to_ui",
            ):
                if hasattr(self, btn_name):
                    getattr(self, btn_name).configure(state="disabled")
            if hasattr(self, "chk_task_force_full_rebuild"):
                self.chk_task_force_full_rebuild.configure(state="disabled")
            for btn_name in (
                "btn_save_cfg_shared",
                "btn_save_cfg_convert",
                "btn_save_cfg_ai",
                "btn_save_cfg_merge",
                "btn_save_cfg_ui",
                "btn_save_cfg_rules",
                "btn_reset_cfg_shared",
                "btn_reset_cfg_convert",
                "btn_reset_cfg_ai",
                "btn_reset_cfg_merge",
                "btn_reset_cfg_ui",
                "btn_reset_cfg_rules",
                "btn_save_cfg_dirty",
            ):
                if hasattr(self, btn_name):
                    getattr(self, btn_name).configure(state="disabled")
            self.progress["mode"] = "determinate"
            self.progress["value"] = 0
            self.var_status.set(self.tr("status_init") if hasattr(self, "tr") else "Initializing...")
        else:
            if hasattr(self, "btn_start"): self.btn_start.configure(state="normal")
            if hasattr(self, "btn_stop"): self.btn_stop.configure(state="disabled")
            if hasattr(self, "btn_task_stop"): self.btn_task_stop.configure(state="disabled")
            if hasattr(self, "btn_save_cfg"): self.btn_save_cfg.configure(state="normal")
            if hasattr(self, "btn_load_cfg"): self.btn_load_cfg.configure(state="normal")
            if hasattr(self, "btn_manage_profiles"): self.btn_manage_profiles.configure(state="normal")
            for btn_name in (
                "btn_task_create",
                "btn_task_edit",
                "btn_task_delete",
                "btn_task_run",
                "btn_task_refresh",
                "btn_task_load_to_ui",
            ):
                if hasattr(self, btn_name):
                    getattr(self, btn_name).configure(state="normal")
            if hasattr(self, "chk_task_force_full_rebuild"):
                self.chk_task_force_full_rebuild.configure(state="normal")
            for btn_name in (
                "btn_save_cfg_shared",
                "btn_save_cfg_convert",
                "btn_save_cfg_ai",
                "btn_save_cfg_merge",
                "btn_save_cfg_ui",
                "btn_save_cfg_rules",
                "btn_reset_cfg_shared",
                "btn_reset_cfg_convert",
                "btn_reset_cfg_ai",
                "btn_reset_cfg_merge",
                "btn_reset_cfg_ui",
                "btn_reset_cfg_rules",
            ):
                if hasattr(self, btn_name):
                    getattr(self, btn_name).configure(state="normal")
            self._update_config_dirty_summary(getattr(self, "_last_section_dirty", {}))
            if hasattr(self, "_on_task_select"):
                self._on_task_select()
            if hasattr(self, "_update_task_tab_for_app_mode"):
                self._update_task_tab_for_app_mode()
            self.progress.stop()
            self.progress["value"] = 100
            self.var_status.set(self.tr("status_ready") if hasattr(self, "tr") else "Ready")
        self._update_profile_manager_controls()
        self._update_profile_dialog_controls()

    def on_progress_update(self, current, total):
        """Thread-safe callback invoked from converter worker thread."""

        def _update():
            if total > 0:
                pct = (current / total) * 100
                self.progress["value"] = pct
                self.var_status.set(
                    self.tr("status_processing").format(current, total, pct)
                )
            else:
                self.progress["mode"] = "indeterminate"
                self.progress.start(20)
                self.var_status.set(
                    self.tr("status_processing_unknown").format(current)
                )

        # Thread-safe marshal to main UI loop.
        self.after(0, _update)

    def _build_artifact_summary_text(self, converter, step_index, total_steps):
        if converter is None:
            return ""

        converted_count = len(getattr(converter, "generated_pdfs", []) or [])
        merged_count = len(getattr(converter, "generated_merge_outputs", []) or [])
        merged_md_count = len(
            getattr(converter, "generated_merge_markdown_outputs", []) or []
        )
        map_count = len(getattr(converter, "generated_map_outputs", []) or [])
        markdown_count = len(
            getattr(converter, "generated_markdown_outputs", []) or []
        )
        markdown_quality_count = len(
            getattr(converter, "generated_markdown_quality_outputs", []) or []
        )
        excel_json_count = len(
            getattr(converter, "generated_excel_json_outputs", []) or []
        )
        records_json_count = len(
            getattr(converter, "generated_records_json_outputs", []) or []
        )
        chromadb_count = len(
            getattr(converter, "generated_chromadb_outputs", []) or []
        )
        mshelp_count = len(getattr(converter, "generated_mshelp_outputs", []) or [])

        lines = [
            self.tr("log_artifacts_title").format(step_index, total_steps),
            self.tr("log_artifacts_counts").format(
                converted_count, merged_count + merged_md_count, map_count
            ),
            self.tr("log_artifacts_ai_counts").format(
                markdown_count, excel_json_count, records_json_count
            ),
            self.tr("log_artifacts_ai_quality").format(markdown_quality_count),
            self.tr("log_artifacts_ai_vector").format(chromadb_count),
        ]
        if mshelp_count:
            lines.append(self.tr("log_artifacts_mshelp").format(mshelp_count))

        manifest_path = getattr(converter, "corpus_manifest_path", "")
        if manifest_path:
            lines.append(self.tr("log_artifacts_manifest").format(manifest_path))

        convert_index = getattr(converter, "convert_index_path", "")
        if convert_index:
            lines.append(self.tr("log_artifacts_convert_index").format(convert_index))

        collect_index = getattr(converter, "collect_index_path", "")
        if collect_index:
            lines.append(self.tr("log_artifacts_collect_index").format(collect_index))

        merge_excel = getattr(converter, "merge_excel_path", "")
        if merge_excel:
            lines.append(self.tr("log_artifacts_merge_excel").format(merge_excel))

        for md_path in (getattr(converter, "generated_markdown_outputs", []) or [])[:2]:
            lines.append(self.tr("log_artifacts_markdown").format(md_path))
        for q_path in (
            getattr(converter, "generated_markdown_quality_outputs", []) or []
        )[:2]:
            lines.append(self.tr("log_artifacts_markdown_quality").format(q_path))
        for excel_json_path in (
            getattr(converter, "generated_excel_json_outputs", []) or []
        )[:2]:
            lines.append(self.tr("log_artifacts_excel_json").format(excel_json_path))
        for js_path in (
            getattr(converter, "generated_records_json_outputs", []) or []
        )[:2]:
            lines.append(self.tr("log_artifacts_records_json").format(js_path))
        for vec_path in (getattr(converter, "generated_chromadb_outputs", []) or [])[:2]:
            lines.append(self.tr("log_artifacts_chromadb").format(vec_path))
        for mshelp_path in (getattr(converter, "generated_mshelp_outputs", []) or [])[:2]:
            lines.append(self.tr("log_artifacts_markdown").format(mshelp_path))
        for merged_md_path in (
            getattr(converter, "generated_merge_markdown_outputs", []) or []
        )[:2]:
            lines.append(self.tr("log_artifacts_markdown").format(merged_md_path))
        update_manifest = getattr(converter, "update_package_manifest_path", "")
        if update_manifest:
            lines.append(self.tr("log_artifacts_update_package").format(update_manifest))
        llm_hub_root = getattr(converter, "llm_hub_root", "")
        if llm_hub_root:
            lines.append(self.tr("log_artifacts_llm_hub").format(llm_hub_root))
        inc_ctx = getattr(converter, "_incremental_context", None) or {}
        if inc_ctx.get("enabled"):
            lines.append(
                self.tr("log_artifacts_incremental").format(
                    inc_ctx.get("added_count", 0),
                    inc_ctx.get("modified_count", 0),
                    inc_ctx.get("renamed_count", 0),
                    inc_ctx.get("unchanged_count", 0),
                    inc_ctx.get("deleted_count", 0),
                )
            )

        return "\n".join(lines)

    def _scan_first_file_with_ext(self, roots, ext):
        ext = str(ext or "").lower()
        for root in roots or []:
            if not root or not os.path.isdir(root):
                continue
            for cur, _, files in os.walk(root):
                for name in files:
                    if name.lower().endswith(ext):
                        return os.path.join(cur, name)
        return ""

    def _should_continue_when_md_merge_missing(self, clean_sources, target):
        if self.var_run_mode.get() != MODE_MERGE_ONLY:
            return True
        if self.var_merge_convert_submode.get() != MERGE_CONVERT_SUBMODE_MERGE_ONLY:
            return True
        if not bool(self.var_output_enable_merged.get()):
            return True
        if not bool(self.var_output_enable_md.get()):
            return True
        roots = [target] if self.var_merge_source.get() == "target" else list(clean_sources or [])
        if self._scan_first_file_with_ext(roots, ".md"):
            return True
        msg = (
            "已选择“仅合并 + MD合并”，但未找到可合并的 .md 文件。\n\n"
            "选择“是”将继续执行并跳过MD合并；选择“否”将退出本次任务。"
        )
        return bool(messagebox.askyesno(self.tr("btn_start"), msg))

    def _sanitize_runtime_config_for_mode(self, cfg, run_mode):
        """Apply mode-driven coercion rules in place; return list of message strings."""
        messages = []
        if run_mode == MODE_CONVERT_THEN_MERGE:
            prev = cfg.get("merge_source", "source")
            if prev != "target":
                cfg["merge_source"] = "target"
                messages.append(self.tr("msg_coercion_merge_source").format(prev))
        merge_mode = cfg.get("merge_mode", MERGE_MODE_CATEGORY)
        if merge_mode == MERGE_MODE_ALL_IN_ONE:
            messages.append(self.tr("msg_coercion_max_mb_ignored"))
        elif merge_mode == MERGE_MODE_CATEGORY:
            try:
                mb = int(cfg.get("max_merge_size_mb", 80))
                if mb < 1:
                    cfg["max_merge_size_mb"] = 80
                    messages.append(self.tr("msg_coercion_max_mb_default").format(80))
            except (TypeError, ValueError):
                cfg["max_merge_size_mb"] = 80
                messages.append(self.tr("msg_coercion_max_mb_default").format(80))
        return messages

    def _log_coercion_summary(self, coercion_messages, show_dialog=False):
        """Write coercion summary to log and optionally show dialog."""
        if not coercion_messages:
            return
        header = self.tr("log_coercion_header")
        block = "\n".join(f"  - {m}" for m in coercion_messages)
        self.txt_log.insert("end", f"{header}\n{block}\n")
        self.txt_log.see("end")
        if show_dialog:
            body = self.tr("msg_coercion_body").format(block)
            messagebox.showinfo(self.tr("msg_coercion_title"), body)

    # ===================== 浠诲姟鎺у埗 =====================
    def _on_click_start(self):
        if self.worker_thread and self.worker_thread.is_alive():
            messagebox.showinfo(self.tr("btn_start"), self.tr("msg_task_already_running"))
            return
        if getattr(self, "var_app_mode", None) and self.var_app_mode.get() == "task":
            task_id = self._get_selected_task_id()
            if not task_id:
                messagebox.showinfo(self.tr("btn_start"), self.tr("msg_task_select_required"))
                return
            self._on_click_task_run(resume=False)
            return
        if not self.validate_runtime_inputs(silent=False, scope="all"):
            messagebox.showerror(self.tr("btn_start"), self.tr("msg_validation_fix_before_run"))
            return

        clean_sources = []
        for s in self.source_folders_list:
            s = s.strip().strip('"').strip("'")
            if os.path.isdir(s):
                clean_sources.append(s)

        if not clean_sources:
            fallback = self.var_source_folder.get().strip().strip('"').strip("'")
            if fallback and os.path.isdir(fallback):
                clean_sources.append(fallback)

        target = self.var_target_folder.get().strip().strip('"').strip("'")
        self.var_target_folder.set(target)

        if not clean_sources:
            messagebox.showerror(self.tr("btn_start"), self.tr("msg_source_folder_required"))
            return
        if not target:
            messagebox.showerror(self.tr("btn_start"), self.tr("msg_target_folder_required"))
            return
        if not self._should_continue_when_md_merge_missing(clean_sources, target):
            return

        self.stop_requested = False
        self.current_task_id = None
        self.current_run_context = "manual"
        self.txt_log.insert("end", f"\n========== {self.tr('log_start')} ==========\n")
        self.txt_log.see("end")

        def worker():
            try:
                base_mode = self.var_run_mode.get()
                steps = []

                if base_mode == MODE_COLLECT_ONLY:
                    for src in clean_sources:
                        steps.append(
                            {
                                "source": src,
                                "mode": MODE_COLLECT_ONLY,
                                "desc": f"Collect: {src}",
                            }
                        )

                elif base_mode == MODE_MSHELP_ONLY:
                    for src in clean_sources:
                        steps.append(
                            {
                                "source": src,
                                "mode": MODE_MSHELP_ONLY,
                                "desc": f"MSHelp: {src}",
                            }
                        )

                elif base_mode == MODE_MERGE_ONLY:
                    m_src = self.var_merge_source.get()
                    if m_src == "target":
                        steps.append(
                            {
                                "source": clean_sources[0],
                                "mode": MODE_MERGE_ONLY,
                                "desc": "Merge (target-based)",
                            }
                        )
                    else:
                        for src in clean_sources:
                            steps.append(
                                {
                                    "source": src,
                                    "mode": MODE_MERGE_ONLY,
                                    "desc": f"Merge (source-based: {src})",
                                }
                            )

                else:
                    for src in clean_sources:
                        steps.append(
                            {
                                "source": src,
                                "mode": MODE_CONVERT_ONLY,
                                "desc": f"Convert: {src}",
                            }
                        )

                    if (
                        base_mode == MODE_CONVERT_THEN_MERGE
                        and self.var_enable_merge.get()
                        and self.var_output_enable_merged.get()
                    ):
                        steps.append(
                            {
                                "source": clean_sources[0],
                                "mode": MODE_MERGE_ONLY,
                                "desc": "Merge (target-based)",
                            }
                        )

                total_steps = len(steps)
                print(f"[GUI] total steps: {total_steps}")

                for idx, step in enumerate(steps, 1):
                    if self.stop_requested:
                        print("[GUI] stop request accepted; remaining steps skipped.")
                        break

                    step_desc = step["desc"]
                    print(f"\n[GUI] >>> step {idx}/{total_steps}: {step_desc}")
                    self.txt_log.insert("end", f"\n>>> step {idx}/{total_steps}: {step_desc}\n")
                    self.txt_log.see("end")

                    print(f"[GUI] using config file: {self.config_path}")
                    converter = GUIOfficeConverter(self.config_path)
                    converter.progress_callback = self.on_progress_update
                    self.current_converter = converter

                    cfg = converter.config
                    cfg["source_folder"] = step["source"]
                    cfg["source_folders"] = [step["source"]]
                    cfg["target_folder"] = target
                    cfg["enable_sandbox"] = bool(self.var_enable_sandbox.get())
                    cfg["temp_sandbox_root"] = self.var_temp_sandbox_root.get().strip()
                    cfg["sandbox_min_free_gb"] = self._safe_positive_int(
                        self.var_sandbox_min_free_gb.get(), 10
                    )
                    cfg["sandbox_low_space_policy"] = self.var_sandbox_low_space_policy.get() or "block"
                    cfg["enable_llm_delivery_hub"] = bool(self.var_enable_llm_delivery_hub.get())
                    cfg["llm_delivery_root"] = self.var_llm_delivery_root.get().strip()
                    cfg["llm_delivery_flatten"] = bool(self.var_llm_delivery_flatten.get())
                    cfg["llm_delivery_include_pdf"] = bool(self.var_llm_delivery_include_pdf.get())
                    cfg["enable_upload_readme"] = bool(self.var_enable_upload_readme.get())
                    cfg["enable_upload_json_manifest"] = bool(self.var_enable_upload_json_manifest.get())
                    cfg["upload_dedup_merged"] = bool(self.var_upload_dedup_merged.get())
                    cfg["enable_merge"] = bool(self.var_enable_merge.get())
                    cfg["output_enable_pdf"] = bool(self.var_output_enable_pdf.get())
                    cfg["output_enable_md"] = bool(self.var_output_enable_md.get())
                    cfg["output_enable_merged"] = bool(self.var_output_enable_merged.get())
                    cfg["output_enable_independent"] = bool(
                        self.var_output_enable_independent.get()
                    )
                    cfg["merge_convert_submode"] = self.var_merge_convert_submode.get()
                    cfg["merge_mode"] = self.var_merge_mode.get()
                    cfg["merge_source"] = (
                        "target"
                        if base_mode == MODE_CONVERT_THEN_MERGE
                        else self.var_merge_source.get()
                    )
                    cfg["enable_merge_index"] = bool(self.var_enable_merge_index.get())
                    cfg["enable_merge_excel"] = bool(self.var_enable_merge_excel.get())
                    cfg["max_merge_size_mb"] = self._safe_positive_int(
                        self.var_max_merge_size_mb.get(), 80
                    )
                    cfg["merge_filename_pattern"] = (
                        self.var_merge_filename_pattern.get().strip()
                        or "Merged_{category}_{timestamp}_{idx}"
                    )
                    cfg["enable_corpus_manifest"] = bool(self.var_enable_corpus_manifest.get())
                    cfg.pop("enable_markdown", None)
                    cfg["markdown_strip_header_footer"] = bool(self.var_markdown_strip_header_footer.get())
                    cfg["markdown_structured_headings"] = bool(self.var_markdown_structured_headings.get())
                    cfg["enable_markdown_quality_report"] = bool(self.var_enable_markdown_quality_report.get())
                    cfg["enable_excel_json"] = bool(self.var_enable_excel_json.get())
                    cfg["enable_chromadb_export"] = bool(self.var_enable_chromadb_export.get())
                    cfg["mshelpviewer_folder_name"] = (
                        self.var_mshelpviewer_folder_name.get().strip() or "MSHelpViewer"
                    )
                    cfg["enable_mshelp_merge_output"] = bool(
                        self.var_enable_mshelp_merge_output.get()
                    )
                    cfg["enable_mshelp_output_docx"] = bool(
                        self.var_enable_mshelp_output_docx.get()
                    )
                    cfg["enable_mshelp_output_pdf"] = bool(
                        self.var_enable_mshelp_output_pdf.get()
                    )
                    cfg["enable_incremental_mode"] = bool(self.var_enable_incremental_mode.get())
                    cfg["incremental_verify_hash"] = bool(self.var_incremental_verify_hash.get())
                    cfg["incremental_reprocess_renamed"] = bool(self.var_incremental_reprocess_renamed.get())
                    cfg["source_priority_skip_same_name_pdf"] = bool(self.var_source_priority_skip_same_name_pdf.get())
                    cfg["global_md5_dedup"] = bool(self.var_global_md5_dedup.get())
                    cfg["enable_update_package"] = bool(self.var_enable_update_package.get())
                    cfg["kill_process_mode"] = self.var_kill_mode.get()
                    cfg["default_engine"] = self.var_engine.get()
                    cfg["office_reuse_app"] = bool(self.var_office_reuse_app.get())
                    cfg["office_restart_every_n_files"] = self._safe_positive_int(
                        self.var_office_restart_every_n_files.get(), 25
                    )
                    coercion_msgs = self._sanitize_runtime_config_for_mode(cfg, base_mode)
                    if idx == 1 and coercion_msgs:
                        self.after(0, lambda m=coercion_msgs: self._log_coercion_summary(m, show_dialog=True))

                    converter.run_mode = step["mode"]
                    converter.collect_mode = self.var_collect_mode.get()
                    converter.content_strategy = self.var_strategy.get()
                    converter.merge_mode = self.var_merge_mode.get()
                    converter.engine_type = self.var_engine.get()
                    converter.enable_merge_index = bool(self.var_enable_merge_index.get())
                    converter.enable_merge_excel = bool(self.var_enable_merge_excel.get())

                    if self.var_enable_date_filter.get():
                        date_str = self.var_date_str.get().strip()
                        try:
                            converter.filter_date = datetime.strptime(date_str, "%Y-%m-%d")
                            converter.filter_mode = self.var_filter_mode.get()
                        except ValueError:
                            pass

                    temp_root = cfg.get("temp_sandbox_root", "").strip()
                    if temp_root:
                        if not os.path.isabs(temp_root):
                            temp_root = os.path.abspath(os.path.join(get_app_path(), temp_root))
                    else:
                        temp_root = tempfile.gettempdir()
                    converter.temp_sandbox_root = temp_root
                    converter.temp_sandbox = os.path.join(temp_root, "OfficeToPDF_Sandbox")
                    os.makedirs(converter.temp_sandbox, exist_ok=True)

                    converter.failed_dir = os.path.join(cfg["target_folder"], "_FAILED_FILES")
                    os.makedirs(converter.failed_dir, exist_ok=True)
                    converter.merge_output_dir = os.path.join(cfg["target_folder"], "_MERGED")
                    os.makedirs(converter.merge_output_dir, exist_ok=True)

                    converter.run()
                    artifact_summary = self._build_artifact_summary_text(converter, idx, total_steps)
                    if artifact_summary:
                        self.txt_log.insert("end", f"{artifact_summary}\n")
                        self.txt_log.see("end")

                print("[GUI] all tasks completed.")
                self.txt_log.insert("end", f"\n========== {self.tr('log_stop')} ==========\n")
                self.txt_log.see("end")

            except Exception as e:
                print(f"[GUI] runtime error: {e}")
                print(traceback.format_exc())
                messagebox.showerror(
                    self.tr("msg_runtime_error_title"),
                    self.tr("msg_runtime_error_body").format(e),
                )
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
        if messagebox.askyesno(self.tr("btn_stop"), self.tr("msg_confirm_stop")):
            self.stop_requested = True
            self.current_converter.is_running = False
            print("[GUI] stop requested; waiting for current step to finish...")
            self.var_status.set(self.tr("status_stop_wait"))

    # ===================== 绋嬪簭鍏ュ彛 =====================


if __name__ == "__main__":
    try:
        app = OfficeGUI()
        app.mainloop()
    except Exception:
        traceback.print_exc()
