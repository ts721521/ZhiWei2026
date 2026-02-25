# -*- coding: utf-8 -*-
"""UI shell/build methods extracted from OfficeGUI."""

import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.constants import *
from office_converter import create_default_config

try:
    import ttkbootstrap as tb
    from ttkbootstrap.widgets.scrolled import ScrolledText
    HAS_TTKBOOTSTRAP = True
except ModuleNotFoundError:
    from tkinter.scrolledtext import ScrolledText as _TkScrolledText
    HAS_TTKBOOTSTRAP = False

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
    class _FallbackSeparator(_BootstyleMixin, ttk.Separator):
        pass
    class _FallbackRadiobutton(_BootstyleMixin, ttk.Radiobutton):
        pass
    class _TBNamespace:
        Frame = _FallbackFrame
        Label = _FallbackLabel
        Button = _FallbackButton
        Checkbutton = _FallbackCheckbutton
        Progressbar = _FallbackProgressbar
        Notebook = _FallbackNotebook
        Scrollbar = _FallbackScrollbar
        Entry = _FallbackEntry
        Labelframe = _FallbackLabelframe
        Separator = _FallbackSeparator
        Radiobutton = _FallbackRadiobutton
    tb = _TBNamespace()

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


class UIShellMixin:
    def _finish_init(self):
        """延迟执行：创建默认配置、构建主 UI、加载配置。保证主窗口已显示后再做重活。"""
        try:
            if not self.winfo_exists():
                return
        except Exception:
            return
        try:
            self._loading_frame.destroy()
        except Exception:
            pass
        self._loading_frame = None
        if not os.path.exists(self.config_path):
            success = create_default_config(self.config_path)
            if success:
                info_title = "提示"
                messagebox.showinfo(info_title, self.tr("msg_no_config"))
        try:
            if getattr(self, "task_store", None) is not None:
                self.task_store.migrate_legacy_tasks()
        except Exception:
            pass
        self._build_ui()
        self._load_config_to_ui()
        self.locator_short_id_index = {}
        try:
            self.protocol("WM_DELETE_WINDOW", self._on_close_main_window)
        except Exception:
            pass
        self._after_force_refresh_ids = []
        self._after_poll_log_id = self.after(200, self._poll_log_queue)
        self.update_idletasks()
        self._force_refresh_all_canvases()
        self._after_force_refresh_ids.append(
            self.after(100, self._force_refresh_all_canvases)
        )
        self._after_force_refresh_ids.append(
            self.after(500, self._force_refresh_all_canvases)
        )
        self.main_notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)
        self.lift()

    def _force_refresh_all_canvases(self):
        try:
            if not self.winfo_exists():
                return
        except Exception:
            return
        for tab_frame in getattr(self, "_all_tabs", []):
            for child in tab_frame.winfo_children():
                if isinstance(child, tk.Canvas):
                    child.update_idletasks()
                    w = child.winfo_width()
                    if w > 1:
                        for item_id in child.find_all():
                            child.itemconfig(item_id, width=w)
                        child.configure(scrollregion=child.bbox("all"))

    def _on_tab_changed(self, _event=None):
        try:
            ids = getattr(self, "_after_force_refresh_ids", None)
            if ids is None:
                ids = []
                self._after_force_refresh_ids = ids
            ids.append(self.after(10, self._force_refresh_all_canvases))
            ids.append(self.after(100, self._force_refresh_all_canvases))
        except Exception:
            pass

    # ===================== UI 閺嬪嫬缂?=====================

    # ===================== UI 閺嬪嫬缂?(Modern Layout) =====================

    def _create_scrollable_page(self, parent):
        canvas = tk.Canvas(parent, highlightthickness=0)
        scrollbar = tb.Scrollbar(parent, orient="vertical", command=canvas.yview)
        content = tb.Frame(canvas)
        content.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")),
        )
        win_id = canvas.create_window((0, 0), window=content, anchor="nw")

        def _sync_width(retries=10):
            w = canvas.winfo_width()
            if w > 1:
                canvas.itemconfig(win_id, width=w)
                canvas.configure(scrollregion=canvas.bbox("all"))
            elif retries > 0:
                canvas.after(100, lambda: _sync_width(retries - 1))

        def on_canvas_configure(e):
            w = e.width if e.width > 1 else canvas.winfo_width()
            if w > 1:
                canvas.itemconfig(win_id, width=w)
            canvas.configure(scrollregion=canvas.bbox("all"))

        canvas.bind("<Map>", lambda e: canvas.after(20, _sync_width))
        canvas.bind("<Expose>", lambda e: canvas.after(20, _sync_width))

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
        tb.Label(f, text=self.tr(label_key), font=("System", 9, "bold")).pack(
            anchor="w"
        )

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
        btn_browse = tb.Button(
            f_in, text="...", command=cmd_browse, bootstyle="outline", width=3
        )
        btn_browse.pack(side=LEFT, padx=(2, 0))
        self._attach_tooltip(btn_browse, "tip_browse_folder")
        if cmd_open:
            btn_open = tb.Button(
                f_in, text=">", command=cmd_open, bootstyle="link", width=2
            )
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

        if not hasattr(self, "var_app_mode"):
            self.var_app_mode = tk.StringVar(value="classic")
        frm_app_mode = tb.Frame(ctrl_frame, bootstyle="light")
        frm_app_mode.pack(side=LEFT, padx=(0, 12))
        tb.Label(
            frm_app_mode,
            text=self.tr("app_mode_classic") + " / " + self.tr("app_mode_task") + ":",
            font=("System", 9),
        ).pack(side=LEFT, padx=(0, 4))
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
                messagebox.showinfo(
                    "主题", "深色主题需要安装 ttkbootstrap：pip install ttkbootstrap"
                )
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
        self.main_notebook.add(
            self.tab_run_convert, text=self.tr("grp_convert_runtime")
        )
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

    def _build_task_tab_content(self):
        parent = self._scroll_tasks
        lf_tasks = tb.Labelframe(parent, text=self.tr("grp_task_runtime"), padding=8)
        lf_tasks.pack(fill=BOTH, expand=YES, pady=3)

        if not hasattr(self, "var_task_filter_text"):
            self.var_task_filter_text = tk.StringVar(value="")
        if not hasattr(self, "var_task_status_filter"):
            self.var_task_status_filter = tk.StringVar(value="all")
        if not hasattr(self, "var_task_sort_by"):
            self.var_task_sort_by = tk.StringVar(value="updated_desc")
        if not hasattr(self, "var_task_scope_current_config_only"):
            self.var_task_scope_current_config_only = tk.IntVar(value=1)

        list_col = tb.Frame(lf_tasks)
        list_col.pack(side=LEFT, fill=BOTH, expand=YES)

        filter_row = tb.Frame(list_col)
        filter_row.pack(fill=X, pady=(0, 6))
        tb.Label(filter_row, text=self.tr("lbl_task_filter")).pack(side=LEFT)
        self.ent_task_filter = tb.Entry(
            filter_row, textvariable=self.var_task_filter_text, width=18
        )
        self.ent_task_filter.pack(side=LEFT, padx=(4, 8))

        tb.Label(filter_row, text=self.tr("lbl_task_status_filter")).pack(side=LEFT)
        self.cb_task_status_filter = tb.Combobox(
            filter_row,
            textvariable=self.var_task_status_filter,
            values=("all", "idle", "running", "failed", "completed"),
            state="readonly",
            width=10,
        )
        self.cb_task_status_filter.pack(side=LEFT, padx=(4, 8))

        tb.Label(filter_row, text=self.tr("lbl_task_sort")).pack(side=LEFT)
        self.cb_task_sort = tb.Combobox(
            filter_row,
            textvariable=self.var_task_sort_by,
            values=("updated_desc", "last_run_desc", "name_asc", "name_desc"),
            state="readonly",
            width=14,
        )
        self.cb_task_sort.pack(side=LEFT, padx=(4, 8))

        self.chk_task_scope_current_config_only = tb.Checkbutton(
            filter_row,
            text=self.tr("chk_task_scope_current_config_only"),
            variable=self.var_task_scope_current_config_only,
            bootstyle="round-toggle",
        )
        self.chk_task_scope_current_config_only.pack(side=LEFT, padx=(0, 8))
        self._attach_tooltip(
            self.chk_task_scope_current_config_only,
            "tip_task_scope_current_config_only",
        )

        self.btn_task_filter_clear = tb.Button(
            filter_row,
            text=self.tr("btn_task_filter_clear"),
            width=8,
            bootstyle="secondary-outline",
            command=self._reset_task_list_filters,
        )
        self.btn_task_filter_clear.pack(side=LEFT)

        tree_row = tb.Frame(list_col)
        tree_row.pack(fill=BOTH, expand=YES)
        cols = ("name", "source", "target", "config", "binding", "status", "last_run")
        self.tree_tasks = ttk.Treeview(
            tree_row, columns=cols, show="headings", height=12, selectmode="browse"
        )
        col_meta = (
            ("name", "col_task_name", "Name"),
            ("source", "col_task_source", "Source"),
            ("target", "col_task_target", "Target"),
            ("config", "col_task_config", "Config"),
            ("binding", "col_task_binding", "Binding"),
            ("status", "col_task_status", "Status"),
            ("last_run", "col_task_last_run", "Last run"),
        )
        for c, key, fallback in col_meta:
            text = self.tr(key)
            self.tree_tasks.heading(c, text=fallback if text == key else text)
        self.tree_tasks.column("name", width=160)
        self.tree_tasks.column("source", width=150)
        self.tree_tasks.column("target", width=150)
        self.tree_tasks.column("config", width=150)
        self.tree_tasks.column("binding", width=110)
        self.tree_tasks.column("status", width=70)
        self.tree_tasks.column("last_run", width=100)
        self.tree_tasks.pack(fill=BOTH, expand=YES, side=LEFT)
        self._attach_tooltip(self.tree_tasks, "tip_task_list")
        scr = tb.Scrollbar(tree_row, orient="vertical", command=self.tree_tasks.yview)
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

        tk.Label(
            parent, text=self.tr("lbl_task_config_section"), font=("System", 9, "bold")
        ).pack(anchor=W, pady=(6, 2))
        self.txt_task_detail = ScrolledText(
            parent, height=40, wrap=tk.WORD, font=("Consolas", 9)
        )
        self.txt_task_detail.pack(fill=BOTH, expand=YES, pady=(0, 4))
        try:
            self.txt_task_detail.insert(tk.END, self.tr("msg_task_none_selected"))
            self.txt_task_detail.configure(state="disabled")
        except Exception:
            pass

        self._auto_attach_action_tooltips(lf_tasks)
        self._auto_attach_input_tooltips(lf_tasks, "tip_section_run_mode")
        if not getattr(self, "_task_tab_filters_trace_done", False):
            self.var_task_filter_text.trace_add(
                "write", lambda *a: self.after(0, self._refresh_task_list_ui)
            )
            self.var_task_status_filter.trace_add(
                "write", lambda *a: self.after(0, self._refresh_task_list_ui)
            )
            self.var_task_sort_by.trace_add(
                "write", lambda *a: self.after(0, self._refresh_task_list_ui)
            )
            self.var_task_scope_current_config_only.trace_add(
                "write", lambda *a: self.after(0, self._refresh_task_list_ui)
            )
            self._task_tab_filters_trace_done = True
        if not getattr(self, "_task_tab_app_mode_trace_done", False) and hasattr(
            self, "var_app_mode"
        ):
            self.var_app_mode.trace_add(
                "write", lambda *a: self.after(0, self._on_app_mode_change_for_task_tab)
            )
            self._task_tab_app_mode_trace_done = True
        self._refresh_task_list_ui()

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
            if (
                getattr(child, "_tooltip_key", None) is not None
                or getattr(child, "_tooltip_text", None) is not None
            ):
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
                                    self.main_notebook.tab(
                                        tab_list[idx], state="normal"
                                    )
                            else:
                                for tab_id in self.main_notebook.tabs():
                                    try:
                                        if (
                                            self.main_notebook.nametowidget(tab_id)
                                            is tab
                                        ):
                                            self.main_notebook.tab(
                                                tab_id, state="normal"
                                            )
                                            break
                                    except Exception:
                                        continue
                    except Exception:
                        pass
        except Exception:
            pass

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
        base_dir = filedialog.askdirectory(
            title=self.tr("msg_task_pick_source") + " (选择第一个文件夹)"
        )
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
                "继续添加", "是否要继续添加更多文件夹？", icon="question"
            )
            if not result:
                break

            next_dir = filedialog.askdirectory(
                title=self.tr("msg_task_pick_source") + " (选择下一个文件夹)"
            )
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

