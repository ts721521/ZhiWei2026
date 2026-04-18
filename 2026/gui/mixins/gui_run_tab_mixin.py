# -*- coding: utf-8 -*-
"""Run tab UI methods extracted from OfficeGUI to reduce file size and complexity."""

from datetime import datetime
import tkinter as tk
from tkinter import ttk
from tkinter.constants import *

try:
    import ttkbootstrap as tb
except ModuleNotFoundError:
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

    class _FallbackSpinbox(_BootstyleMixin, ttk.Spinbox):
        pass

    class _FallbackDateEntry(_FallbackEntry):
        def __init__(self, *args, **kwargs):
            kwargs.pop("dateformat", None)
            kwargs.pop("firstweekday", None)
            kwargs.pop("startdate", None)
            super().__init__(*args, **kwargs)
            self.entry = self

    class _TBNamespace:
        Frame = _FallbackFrame
        Label = _FallbackLabel
        Button = _FallbackButton
        Checkbutton = _FallbackCheckbutton
        Scrollbar = _FallbackScrollbar
        Entry = _FallbackEntry
        Labelframe = _FallbackLabelframe
        Radiobutton = _FallbackRadiobutton
        Separator = _FallbackSeparator
        Combobox = _FallbackCombobox
        Spinbox = _FallbackSpinbox
        DateEntry = _FallbackDateEntry

    tb = _TBNamespace()

from office_converter import (
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
)


class RunTabUIMixin:
    def _build_run_tab_content(self):
        parent = self._scroll_shared
        # 运行参数页整体提示：仅用于编辑/预览，实际执行必须从任务中心启动。
        hint_lbl = tb.Label(
            parent,
            text=self.tr("run_tab_hint_task_only"),
            wraplength=900,
            justify=LEFT,
        )
        hint_lbl.pack(fill=X, pady=(2, 8))
        # Section 1: run mode
        lf_mode = tb.Labelframe(parent, text=self.tr("sec_mode"), padding=6)
        lf_mode.pack(fill=X, pady=3)
        self._add_section_help(lf_mode, "tip_section_run_mode")
        self.var_run_mode = tk.StringVar(value=MODE_CONVERT_THEN_MERGE)
        grid_frame = tb.Frame(lf_mode)
        grid_frame.pack(fill=X)
        tb.Radiobutton(
            grid_frame,
            text=self.tr("mode_convert"),
            variable=self.var_run_mode,
            value=MODE_CONVERT_ONLY,
            command=self._on_run_mode_change,
            bootstyle="toolbutton-outline",
        ).grid(row=0, column=0, sticky="ew", padx=2, pady=2)
        tb.Radiobutton(
            grid_frame,
            text=self.tr("mode_merge"),
            variable=self.var_run_mode,
            value=MODE_MERGE_ONLY,
            command=self._on_run_mode_change,
            bootstyle="toolbutton-outline",
        ).grid(row=0, column=1, sticky="ew", padx=2, pady=2)
        tb.Radiobutton(
            grid_frame,
            text=self.tr("mode_convert_merge"),
            variable=self.var_run_mode,
            value=MODE_CONVERT_THEN_MERGE,
            command=self._on_run_mode_change,
            bootstyle="toolbutton-outline",
        ).grid(row=1, column=0, sticky="ew", padx=2, pady=2)
        tb.Radiobutton(
            grid_frame,
            text=self.tr("mode_collect"),
            variable=self.var_run_mode,
            value=MODE_COLLECT_ONLY,
            command=self._on_run_mode_change,
            bootstyle="toolbutton-outline",
        ).grid(row=1, column=1, sticky="ew", padx=2, pady=2)
        # MSHelp 模式已剥离为独立 CLI（tools/mshelp_run.py），界面不再露出。
        grid_frame.columnconfigure(0, weight=1)
        grid_frame.columnconfigure(1, weight=1)

        # Global output controls (prominent and mode-agnostic)
        lf_output = tb.Labelframe(
            parent, text=self.tr("grp_output_controls"), padding=6
        )
        lf_output.pack(fill=X, pady=3)
        self.var_output_enable_pdf = tk.IntVar(value=1)
        self.var_output_enable_md = tk.IntVar(value=1)
        self.var_output_enable_merged = tk.IntVar(value=1)
        self.var_output_enable_independent = tk.IntVar(value=0)
        self.chk_output_enable_pdf = tb.Checkbutton(
            lf_output,
            text=self.tr("chk_output_pdf"),
            variable=self.var_output_enable_pdf,
        )
        self.chk_output_enable_pdf.pack(anchor="w")
        self.chk_output_enable_md = tb.Checkbutton(
            lf_output,
            text=self.tr("chk_output_md"),
            variable=self.var_output_enable_md,
        )
        self.chk_output_enable_md.pack(anchor="w")
        self.chk_output_enable_merged = tb.Checkbutton(
            lf_output,
            text=self.tr("chk_output_merged"),
            variable=self.var_output_enable_merged,
        )
        self.chk_output_enable_merged.pack(anchor="w", pady=(4, 0))
        self.chk_output_enable_independent = tb.Checkbutton(
            lf_output,
            text=self.tr("chk_output_independent"),
            variable=self.var_output_enable_independent,
        )
        self.chk_output_enable_independent.pack(anchor="w")

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

        # Tab 濡楀棙鐏﹂崪灞剧泊閸斻劑銆夐棃銏犲嚒閸?_build_ui 娑擃厼鍨卞鐚寸礉娑撳秴鍟€闂団偓鐟曚礁鐡?Notebook

        # 閵嗗本鈷婇悶鍡愨偓宥夆偓澶愩€嶆稉搴℃値楠炲爼鈧銆嶉崥鍫濊嫙閸掓澘鎮撴稉鈧い纰夌礉閸戝繐鐨粚铏规
        lf_collect = tb.Labelframe(
            self._scroll_merge, text=self.tr("grp_collect_runtime"), padding=6
        )
        lf_collect.pack(fill=X, pady=3)
        tb.Label(
            lf_collect, text=self.tr("lbl_collect_mode"), font=("System", 9, "bold")
        ).pack(anchor="w")
        self.frm_collect_opts = tb.Frame(lf_collect, padding=(10, 5))
        self.frm_collect_opts.pack(fill=X)
        self.var_collect_mode = tk.StringVar(value=COLLECT_MODE_COPY_AND_INDEX)
        tb.Radiobutton(
            self.frm_collect_opts,
            text="Copy + Index",
            variable=self.var_collect_mode,
            value=COLLECT_MODE_COPY_AND_INDEX,
        ).pack(anchor="w")
        tb.Radiobutton(
            self.frm_collect_opts,
            text="Index Only",
            variable=self.var_collect_mode,
            value=COLLECT_MODE_INDEX_ONLY,
        ).pack(anchor="w")

        # MSHelp 配置块已下线（独立 CLI 见 tools/mshelp_run.py），
        # 但保留 IntVar/StringVar 占位，以便 config 加载/保存/对比逻辑无需改动。
        self.var_mshelpviewer_folder_name = tk.StringVar(value="MSHelpViewer")
        self.var_enable_mshelp_merge_output = tk.IntVar(value=1)
        self.var_enable_mshelp_output_docx = tk.IntVar(value=0)
        self.var_enable_mshelp_output_pdf = tk.IntVar(value=0)

        # Section 2: paths (runtime only)
        lf_paths = tb.Labelframe(
            self._scroll_shared, text=self.tr("grp_shared_runtime"), padding=6
        )
        lf_paths.pack(fill=X, pady=3)
        self._add_section_help(lf_paths, "tip_section_run_paths")

        # Source Folders (Multi-select)
        frm_src = tb.Frame(lf_paths)
        frm_src.pack(fill=X, pady=(5, 0))
        tb.Label(frm_src, text=self.tr("lbl_source"), font=("System", 9, "bold")).pack(
            anchor="w"
        )

        frm_src_body = tb.Frame(frm_src)
        frm_src_body.pack(fill=X, expand=YES)

        self.lst_source_folders = tk.Listbox(
            frm_src_body,
            height=6,
            selectmode=EXTENDED,
            font=("System", 9),
            activestyle="dotbox",
        )
        self.lst_source_folders.pack(side=LEFT, fill=X, expand=YES)
        self._attach_tooltip(self.lst_source_folders, "tip_input_source_folder")

        scr_src = tb.Scrollbar(
            frm_src_body, orient="vertical", command=self.lst_source_folders.yview
        )
        scr_src.pack(side=LEFT, fill=Y)
        self.lst_source_folders.configure(yscrollcommand=scr_src.set)
        self.lst_source_folders.bind("<Double-Button-1>", self.open_source_folder)

        frm_src_btns = tb.Frame(frm_src_body)
        frm_src_btns.pack(side=LEFT, fill=Y, padx=(5, 0))

        self.btn_add_src = tb.Button(
            frm_src_btns,
            text="+",
            width=3,
            command=self.add_source_folder,
            bootstyle="success-outline",
        )
        self.btn_add_src.pack(pady=1)
        self._attach_tooltip(self.btn_add_src, "tip_add_source_folder")

        self.btn_del_src = tb.Button(
            frm_src_btns,
            text="-",
            width=3,
            command=self.remove_source_folder,
            bootstyle="danger-outline",
        )
        self.btn_del_src.pack(pady=1)
        self._attach_tooltip(self.btn_del_src, "tip_remove_source_folder")

        self.btn_clr_src = tb.Button(
            frm_src_btns,
            text="C",
            width=3,
            command=self.clear_source_folders,
            bootstyle="secondary-outline",
        )
        self.btn_clr_src.pack(pady=1)
        self._attach_tooltip(self.btn_clr_src, "tip_clear_source_folders")

        # Compatibility
        self.var_source_folder = tk.StringVar()

        self.var_target_folder = tk.StringVar()
        self._create_path_row(
            lf_paths,
            "lbl_target",
            self.var_target_folder,
            self.browse_target,
            self.open_target_folder,
        )
        self.var_enable_corpus_manifest = tk.IntVar(value=1)
        self.chk_corpus_manifest = tb.Checkbutton(
            lf_paths,
            text=self.tr("chk_corpus_manifest"),
            variable=self.var_enable_corpus_manifest,
            bootstyle="round-toggle",
        )
        self.chk_corpus_manifest.pack(anchor="w", pady=(6, 0))
        self._attach_tooltip(self.chk_corpus_manifest, "tip_toggle_corpus_manifest")

        # Section 3: feature-specific runtime options閿涘牐娴嗛幑銏も偓澶愩€嶉敍姘箯閸欏啿寮婚崚妤€绔风仦鈧敍?
        lf_settings = tb.Labelframe(
            self._scroll_convert, text=self.tr("grp_convert_runtime"), padding=6
        )
        lf_settings.pack(fill=BOTH, pady=3)
        self._add_section_help(lf_settings, "tip_section_run_advanced")

        # 閸欏苯鍨€圭懓娅掗敍姘箯閸掓绱欏鏇熸惛/濞屾瑧娲?缁涙盯鈧绱氶敍灞藉礁閸掓绱橝I 鐎电厧鍤?+ LLM Hub + 婢х偤鍣洪敍?
        frm_convert_cols = tb.Frame(lf_settings)
        frm_convert_cols.pack(fill=BOTH, expand=YES)

        col_left = tb.Frame(frm_convert_cols)
        col_left.grid(row=0, column=0, sticky="nsew", padx=(0, 6))

        col_right = tb.Frame(frm_convert_cols)
        col_right.grid(row=0, column=1, sticky="nsew", padx=(6, 0))

        frm_convert_cols.columnconfigure(0, weight=1)
        frm_convert_cols.columnconfigure(1, weight=1)

        lf_convert_runtime = tb.Labelframe(
            col_left, text=self.tr("grp_convert_runtime"), padding=8
        )
        lf_convert_runtime.pack(fill=X, pady=(2, 6))
        self.group_exec = tb.Frame(lf_convert_runtime)
        self.group_exec.pack(fill=X, pady=5)
        tb.Label(self.group_exec, text=self.tr("lbl_engine"), bootstyle="primary").pack(
            anchor="w"
        )
        self.var_engine = tk.StringVar(value=ENGINE_WPS)
        frm_eng = tb.Frame(self.group_exec)
        frm_eng.pack(anchor="w")
        tb.Radiobutton(
            frm_eng, text="WPS Office", variable=self.var_engine, value=ENGINE_WPS
        ).pack(side=LEFT, padx=5)
        tb.Radiobutton(
            frm_eng, text="MS Office", variable=self.var_engine, value=ENGINE_MS
        ).pack(side=LEFT, padx=5)
        self.var_enable_sandbox = tk.IntVar(value=1)
        self.chk_enable_sandbox = tb.Checkbutton(
            self.group_exec,
            text=self.tr("lbl_sandbox"),
            variable=self.var_enable_sandbox,
            bootstyle="success-round-toggle",
            command=self._on_toggle_sandbox,
        )
        self.chk_enable_sandbox.pack(anchor="w", pady=(10, 2))
        self.frm_sandbox_path = tb.Frame(self.group_exec)
        self.frm_sandbox_path.pack(fill=X, padx=20)
        self.var_temp_sandbox_root = tk.StringVar()
        self.entry_temp_sandbox_root = tb.Entry(
            self.frm_sandbox_path,
            textvariable=self.var_temp_sandbox_root,
            font=("Consolas", 8),
        )
        self.entry_temp_sandbox_root.pack(side=LEFT, fill=X, expand=YES)
        self.btn_temp_sandbox_root = tb.Button(
            self.frm_sandbox_path,
            text=self.tr("btn_browse"),
            command=self.browse_temp_sandbox_root,
            bootstyle="outline",
            width=3,
        )
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

        lf_merge_runtime = tb.Labelframe(
            self._scroll_merge, text=self.tr("grp_merge_runtime"), padding=8
        )
        lf_merge_runtime.pack(fill=X, pady=(2, 0))
        self.lbl_merge = tb.Label(
            lf_merge_runtime, text=self.tr("lbl_merge_logic"), bootstyle="primary"
        )
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
        tb.Radiobutton(
            self.frm_merge_opts,
            text=self.tr("rad_category"),
            variable=self.var_merge_mode,
            value=MERGE_MODE_CATEGORY,
        ).pack(anchor="w")
        tb.Radiobutton(
            self.frm_merge_opts,
            text=self.tr("rad_all_in_one"),
            variable=self.var_merge_mode,
            value=MERGE_MODE_ALL_IN_ONE,
        ).pack(anchor="w")
        tb.Separator(self.frm_merge_opts).pack(fill=X, pady=5)
        self.var_enable_merge_index = tk.IntVar(value=0)
        self.chk_merge_index = tb.Checkbutton(
            self.frm_merge_opts,
            text=self.tr("chk_merge_index"),
            variable=self.var_enable_merge_index,
        )
        self.chk_merge_index.pack(anchor="w")
        self.var_enable_merge_excel = tk.IntVar(value=0)
        self.chk_merge_excel = tb.Checkbutton(
            self.frm_merge_opts,
            text=self.tr("chk_merge_excel"),
            variable=self.var_enable_merge_excel,
        )
        self.chk_merge_excel.pack(anchor="w")
        tb.Separator(self.frm_merge_opts).pack(fill=X, pady=5)
        self.lbl_m_src = tb.Label(
            self.frm_merge_opts, text=self.tr("lbl_merge_src"), font=("System", 9)
        )
        self.lbl_m_src.pack(anchor="w")
        self.var_merge_source = tk.StringVar(value="source")
        self.frm_merge_src = tb.Frame(self.frm_merge_opts)
        self.frm_merge_src.pack(fill=X)
        tb.Radiobutton(
            self.frm_merge_src,
            text=self.tr("rad_src_dir"),
            variable=self.var_merge_source,
            value="source",
        ).pack(side=LEFT)
        tb.Radiobutton(
            self.frm_merge_src,
            text=self.tr("rad_tgt_dir"),
            variable=self.var_merge_source,
            value="target",
        ).pack(side=LEFT, padx=10)
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
            self.var_output_enable_merged.trace_add(
                "write", lambda *a: self.after(0, self._on_merge_output_or_mode_change)
            )
            self.var_merge_mode.trace_add(
                "write", lambda *a: self.after(0, self._on_merge_output_or_mode_change)
            )
            for _v in (
                "var_output_enable_pdf",
                "var_output_enable_md",
                "var_output_enable_independent",
            ):
                if hasattr(self, _v):
                    getattr(self, _v).trace_add(
                        "write",
                        lambda *a: self.after(0, self._update_output_summary_label),
                    )
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

        # Section 5: AI export (convert-specific锛屽彸鍒楋級
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
        tb.Separator(frm_ai_export, orient="horizontal").pack(fill=X, pady=(4, 4))
        self.var_enable_fast_md_engine = tk.IntVar(value=0)
        self.chk_enable_fast_md_engine = tb.Checkbutton(
            frm_ai_export,
            text=self.tr("chk_enable_fast_md_engine"),
            variable=self.var_enable_fast_md_engine,
        )
        self.chk_enable_fast_md_engine.pack(anchor="w")
        self.var_enable_traceability_anchor_and_map = tk.IntVar(value=1)
        self.chk_enable_traceability_anchor_and_map = tb.Checkbutton(
            frm_ai_export,
            text=self.tr("chk_enable_traceability_anchor_and_map"),
            variable=self.var_enable_traceability_anchor_and_map,
        )
        self.chk_enable_traceability_anchor_and_map.pack(anchor="w")
        self.var_enable_markdown_image_manifest = tk.IntVar(value=1)
        self.chk_enable_markdown_image_manifest = tb.Checkbutton(
            frm_ai_export,
            text="Markdown image manifest (MD<->PDF map)",
            variable=self.var_enable_markdown_image_manifest,
        )
        self.chk_enable_markdown_image_manifest.pack(anchor="w")
        self.var_enable_prompt_wrapper = tk.IntVar(value=0)
        self.chk_enable_prompt_wrapper = tb.Checkbutton(
            frm_ai_export,
            text=self.tr("chk_enable_prompt_wrapper"),
            variable=self.var_enable_prompt_wrapper,
        )
        self.chk_enable_prompt_wrapper.pack(anchor="w")
        frm_prompt_template = tb.Frame(frm_ai_export)
        frm_prompt_template.pack(fill=X, pady=(2, 0))
        tb.Label(frm_prompt_template, text=self.tr("lbl_prompt_template_type")).pack(
            side=LEFT
        )
        self.var_prompt_template_type = tk.StringVar(value="new_solution")
        self.cb_prompt_template_type = tb.Combobox(
            frm_prompt_template,
            textvariable=self.var_prompt_template_type,
            values=["new_solution", "tech_clarification", "device_extract"],
            state="readonly",
            width=22,
        )
        self.cb_prompt_template_type.pack(side=LEFT, padx=(8, 0))
        self.var_short_id_prefix = tk.StringVar(value="ZW-")
        frm_short_id_prefix = tb.Frame(frm_ai_export)
        frm_short_id_prefix.pack(fill=X, pady=(2, 0))
        tb.Label(frm_short_id_prefix, text=self.tr("lbl_short_id_prefix")).pack(side=LEFT)
        self.ent_short_id_prefix = tb.Entry(
            frm_short_id_prefix,
            textvariable=self.var_short_id_prefix,
            width=10,
        )
        self.ent_short_id_prefix.pack(side=LEFT, padx=(8, 0))
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
        self._attach_tooltip(self.chk_chromadb_export, "tip_toggle_chromadb_export")
        self._attach_tooltip(
            self.chk_enable_fast_md_engine, "tip_toggle_enable_fast_md_engine"
        )
        self._attach_tooltip(
            self.chk_enable_traceability_anchor_and_map,
            "tip_toggle_enable_traceability_anchor_and_map",
        )
        self._attach_tooltip(
            self.chk_enable_prompt_wrapper, "tip_toggle_enable_prompt_wrapper"
        )
        self._attach_tooltip(
            self.cb_prompt_template_type, "tip_input_prompt_template_type"
        )
        self._attach_tooltip(self.ent_short_id_prefix, "tip_input_short_id_prefix")
        try:
            self.var_output_enable_md.trace_add(
                "write", lambda *_: self._sync_markdown_master_with_global_output()
            )
            self.var_enable_fast_md_engine.trace_add(
                "write", lambda *_: self.after(0, self._on_toggle_fast_md_engine)
            )
        except Exception:
            pass
        self._sync_markdown_master_with_global_output()

        # Section: LLM delivery hub -> moved to Output Files tab
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
        self._attach_tooltip(
            self.chk_enable_upload_json_manifest, "tip_toggle_upload_json_manifest"
        )

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
        self._attach_tooltip(
            self.chk_upload_dedup_merged, "tip_toggle_upload_dedup_merged"
        )

        # Google Drive 上传暂不暴露：构建到一个孤立 Frame 中，
        # 控件/状态变量保留，便于 GDriveMixin、配置 IO/save/dirty/compose 链路无需改动。
        self._frm_gdrive_hidden = tb.Frame(self)
        lf_gdrive = tb.Labelframe(
            self._frm_gdrive_hidden, text=self.tr("grp_gdrive_upload"), padding=(8, 6)
        )
        self.var_enable_gdrive_upload = tk.IntVar(value=0)
        self.chk_enable_gdrive_upload = tb.Checkbutton(
            lf_gdrive,
            text=self.tr("chk_enable_gdrive_upload"),
            variable=self.var_enable_gdrive_upload,
        )
        self.chk_enable_gdrive_upload.pack(anchor="w")
        self._frm_gdrive_sub = tb.Frame(lf_gdrive)
        self._frm_gdrive_sub.pack(fill=X, padx=(16, 0))
        frm_gdrive_secrets = tb.Frame(self._frm_gdrive_sub)
        frm_gdrive_secrets.pack(fill=X, pady=(2, 2))
        frm_gdrive_secrets_label = tb.Frame(frm_gdrive_secrets)
        frm_gdrive_secrets_label.pack(fill=X)
        tb.Label(
            frm_gdrive_secrets_label,
            text=self.tr("lbl_gdrive_client_secrets_path"),
            font=("System", 9),
        ).pack(side=LEFT, anchor="w")
        self.btn_gdrive_open_console = tb.Button(
            frm_gdrive_secrets_label,
            text=self.tr("link_gdrive_get_secrets"),
            bootstyle="link",
            command=self._on_open_gdrive_console,
        )
        self.btn_gdrive_open_console.pack(side=LEFT, padx=(8, 0))
        self._attach_tooltip(self.btn_gdrive_open_console, "tip_gdrive_open_console")
        self.btn_gdrive_enable_api = tb.Button(
            frm_gdrive_secrets_label,
            text=self.tr("link_gdrive_enable_api"),
            bootstyle="link",
            command=self._on_open_gdrive_enable_api,
        )
        self.btn_gdrive_enable_api.pack(side=LEFT, padx=(8, 0))
        self._attach_tooltip(self.btn_gdrive_enable_api, "tip_gdrive_enable_api")
        frm_gdrive_secrets_row = tb.Frame(frm_gdrive_secrets)
        frm_gdrive_secrets_row.pack(fill=X, pady=(2, 0))
        self.var_gdrive_client_secrets_path = tk.StringVar()
        self.entry_gdrive_client_secrets_path = tb.Entry(
            frm_gdrive_secrets_row,
            textvariable=self.var_gdrive_client_secrets_path,
            font=("Consolas", 8),
        )
        self.entry_gdrive_client_secrets_path.pack(side=LEFT, fill=X, expand=YES)
        self.btn_browse_gdrive_secrets = tb.Button(
            frm_gdrive_secrets_row,
            text=self.tr("btn_browse"),
            width=6,
            command=self._on_browse_gdrive_secrets,
        )
        self.btn_browse_gdrive_secrets.pack(side=LEFT, padx=4)
        self._attach_tooltip(
            self.entry_gdrive_client_secrets_path, "tip_gdrive_client_secrets_path"
        )
        tb.Label(
            self._frm_gdrive_sub,
            text=self.tr("lbl_gdrive_folder_id"),
            font=("System", 9),
        ).pack(anchor="w", pady=(4, 0))
        self.var_gdrive_folder_id = tk.StringVar()
        self.entry_gdrive_folder_id = tb.Entry(
            self._frm_gdrive_sub,
            textvariable=self.var_gdrive_folder_id,
            font=("Consolas", 8),
        )
        self.entry_gdrive_folder_id.pack(fill=X, pady=(2, 0))
        self._attach_tooltip(self.entry_gdrive_folder_id, "tip_gdrive_folder_id")
        frm_gdrive_btns = tb.Frame(self._frm_gdrive_sub)
        frm_gdrive_btns.pack(anchor="w", pady=(8, 0))
        self.btn_install_gdrive_deps = tb.Button(
            frm_gdrive_btns,
            text=self.tr("btn_install_gdrive_deps"),
            bootstyle="secondary-outline",
            command=self._on_install_gdrive_deps,
        )
        self.btn_install_gdrive_deps.pack(side=LEFT, padx=(0, 8))
        self._attach_tooltip(self.btn_install_gdrive_deps, "tip_install_gdrive_deps")
        self.btn_upload_to_gdrive = tb.Button(
            frm_gdrive_btns,
            text=self.tr("btn_upload_llm_to_gdrive"),
            bootstyle="primary-outline",
            command=self._on_upload_llm_to_gdrive,
        )
        self.btn_upload_to_gdrive.pack(side=LEFT)
        self.btn_fetch_gdrive_structure = tb.Button(
            frm_gdrive_btns,
            text=self.tr("btn_fetch_gdrive_structure"),
            bootstyle="secondary-outline",
            command=self._on_fetch_gdrive_structure,
        )
        self.btn_fetch_gdrive_structure.pack(side=LEFT, padx=(8, 0))
        self._attach_tooltip(
            self.btn_fetch_gdrive_structure, "tip_fetch_gdrive_structure"
        )
        try:
            import gdrive_upload as _gd

            if not getattr(_gd, "HAS_GDEPEND", True):
                for w in (
                    self.chk_enable_gdrive_upload,
                    self.entry_gdrive_client_secrets_path,
                    self.btn_gdrive_open_console,
                    self.btn_gdrive_enable_api,
                    self.btn_browse_gdrive_secrets,
                    self.entry_gdrive_folder_id,
                    self.btn_upload_to_gdrive,
                    self.btn_fetch_gdrive_structure,
                ):
                    w.configure(state="disabled")
                self._attach_tooltip(lf_gdrive, "msg_gdrive_no_deps")
        except Exception:
            for w in (
                self.chk_enable_gdrive_upload,
                self.entry_gdrive_client_secrets_path,
                self.btn_gdrive_open_console,
                self.btn_gdrive_enable_api,
                self.btn_browse_gdrive_secrets,
                self.entry_gdrive_folder_id,
                self.btn_upload_to_gdrive,
                self.btn_fetch_gdrive_structure,
            ):
                w.configure(state="disabled")
            self._attach_tooltip(lf_gdrive, "msg_gdrive_no_deps")

        # Section 7: incremental / dedup (convert-specific锛屽彸鍒楋級
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
        self._attach_tooltip(self.chk_global_md5_dedup, "tip_toggle_global_md5_dedup")
        self._attach_tooltip(
            self.chk_enable_update_package, "tip_toggle_enable_update_package"
        )

        # ========== 骞跺彂澶勭悊鍜屾柇鐐圭画浼犻厤缃?==========
        # 骞跺彂杞崲寮€鍏?
        self.var_enable_parallel_conversion = tk.IntVar(value=0)
        self.chk_enable_parallel_conversion = tb.Checkbutton(
            lf_incremental,
            text=self.tr("chk_enable_parallel_conversion"),
            variable=self.var_enable_parallel_conversion,
            command=self._on_toggle_parallel_conversion,
        )
        self.chk_enable_parallel_conversion.pack(anchor="w")
        # 骞跺彂鏁伴厤缃紙鍙楀紑鍏虫帶鍒讹級
        self._frm_parallel_sub = tb.Frame(lf_incremental)
        self._frm_parallel_sub.pack(fill=X, padx=(16, 0))
        frm_workers = tb.Frame(self._frm_parallel_sub)
        frm_workers.pack(fill=X)
        tb.Label(frm_workers, text=self.tr("lbl_parallel_workers")).pack(side=LEFT)
        self.var_parallel_workers = tk.StringVar(value="4")
        self.spn_parallel_workers = tb.Spinbox(
            frm_workers,
            from_=1,
            to=16,
            width=5,
            textvariable=self.var_parallel_workers,
        )
        self.spn_parallel_workers.pack(side=LEFT, padx=(6, 0))

        # 鏂偣缁紶寮€鍏?
        self.var_enable_checkpoint = tk.IntVar(value=1)
        self.chk_enable_checkpoint = tb.Checkbutton(
            lf_incremental,
            text=self.tr("chk_enable_checkpoint"),
            variable=self.var_enable_checkpoint,
        )
        self.chk_enable_checkpoint.pack(anchor="w")
        # 鑷姩鎭㈠鏂偣
        self.var_checkpoint_auto_resume = tk.IntVar(value=1)
        self.chk_checkpoint_auto_resume = tb.Checkbutton(
            lf_incremental,
            text=self.tr("chk_checkpoint_auto_resume"),
            variable=self.var_checkpoint_auto_resume,
        )
        self.chk_checkpoint_auto_resume.pack(anchor="w", padx=(16, 0))

        # 缁戝畾鎻愮ず
        self._attach_tooltip(
            self.chk_enable_parallel_conversion, "tip_toggle_parallel_conversion"
        )
        self._attach_tooltip(self.spn_parallel_workers, "tip_parallel_workers")
        self._attach_tooltip(self.chk_enable_checkpoint, "tip_toggle_checkpoint")
        self._attach_tooltip(
            self.chk_checkpoint_auto_resume, "tip_checkpoint_auto_resume"
        )

        # 鍒濆鐘舵€侊細骞跺彂瀛愰€夐」榛樿闅愯棌
        self._on_toggle_parallel_conversion()

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
            self.ent_date = tb.DateEntry(
                self.frm_date,
                dateformat="%Y-%m-%d",
                firstweekday=0,
                startdate=datetime.now(),
            )
            self.ent_date.entry.configure(textvariable=self.var_date_str)
            self.ent_date.pack(fill=X)
        except Exception:
            self.ent_date = tb.Entry(self.frm_date, textvariable=self.var_date_str)
            self.ent_date.pack(fill=X)
        self.var_filter_mode = tk.StringVar(value="after")
        frm_dt_mode = tb.Frame(self.frm_date)
        frm_dt_mode.pack(fill=X)
        self.rb_filter_after = tb.Radiobutton(
            frm_dt_mode,
            text=self.tr("rad_after"),
            variable=self.var_filter_mode,
            value="after",
        )
        self.rb_filter_after.pack(side=LEFT)
        self.rb_filter_before = tb.Radiobutton(
            frm_dt_mode,
            text=self.tr("rad_before"),
            variable=self.var_filter_mode,
            value="before",
        )
        self.rb_filter_before.pack(side=LEFT, padx=10)

        # Section 5: NotebookLM locator (runtime) 鈥斺€?涓?MSHelp 鍚堝苟鍒板悓涓€ tab
        lf_locator = tb.Labelframe(
            self._scroll_locator, text=self.tr("sec_locator"), padding=10
        )
        lf_locator.pack(fill=X, pady=5)
        self._add_section_help(lf_locator, "tip_section_run_locator")
        tb.Label(lf_locator, text=self.tr("lbl_locator_merged")).pack(anchor="w")
        self.var_locator_merged = tk.StringVar()
        self.cb_locator_merged = tb.Combobox(
            lf_locator,
            textvariable=self.var_locator_merged,
            state="readonly",
            values=[],
        )
        self.cb_locator_merged.pack(fill=X, pady=(0, 4))
        self._attach_tooltip(self.cb_locator_merged, "tip_input_locator_merged")
        row_locator = tb.Frame(lf_locator)
        row_locator.pack(fill=X)
        tb.Label(row_locator, text=self.tr("lbl_locator_page")).pack(side=LEFT)
        self.var_locator_page = tk.StringVar()
        self.ent_locator_page = tb.Entry(
            row_locator, textvariable=self.var_locator_page, width=8
        )
        self.ent_locator_page.pack(side=LEFT, padx=(6, 12))
        self._attach_tooltip(self.ent_locator_page, "tip_input_locator_page")
        tb.Label(row_locator, text=self.tr("lbl_locator_id")).pack(side=LEFT)
        self.var_locator_short_id = tk.StringVar()
        self.ent_locator_short_id = tb.Entry(
            row_locator, textvariable=self.var_locator_short_id, width=14
        )
        self.ent_locator_short_id.pack(side=LEFT, padx=(6, 0))
        self._attach_tooltip(self.ent_locator_short_id, "tip_input_locator_short_id")
        row_locator_btn = tb.Labelframe(
            lf_locator, text=self.tr("lbl_locator_group_actions"), padding=(8, 6)
        )
        row_locator_btn.pack(fill=X, pady=(6, 0))
        self.btn_locator_refresh = tb.Button(
            row_locator_btn,
            text=self.tr("btn_locator_refresh"),
            command=self.refresh_locator_maps,
            bootstyle="secondary-outline",
            width=12,
        )
        self.btn_locator_refresh.pack(side=LEFT, padx=2)
        self._attach_tooltip(self.btn_locator_refresh, "tip_locator_refresh")
        self.btn_locator_locate = tb.Button(
            row_locator_btn,
            text=self.tr("btn_locator_locate"),
            command=self.run_locator_query,
            bootstyle="primary",
            width=10,
        )
        self.btn_locator_locate.pack(side=LEFT, padx=2)
        self._attach_tooltip(self.btn_locator_locate, "tip_locator_locate")
        self.btn_locator_open_file = tb.Button(
            row_locator_btn,
            text=self.tr("btn_locator_open_file"),
            command=self.open_locator_file,
            bootstyle="success-outline",
            width=10,
        )
        self.btn_locator_open_file.pack(side=LEFT, padx=2)
        self._attach_tooltip(self.btn_locator_open_file, "tip_locator_open_file")
        self.btn_locator_open_dir = tb.Button(
            row_locator_btn,
            text=self.tr("btn_locator_open_dir"),
            command=self.open_locator_folder,
            bootstyle="info-outline",
            width=10,
        )
        self.btn_locator_open_dir.pack(side=LEFT, padx=2)
        self._attach_tooltip(self.btn_locator_open_dir, "tip_locator_open_dir")
        row_locator_btn2 = tb.Labelframe(
            lf_locator, text=self.tr("lbl_locator_group_external"), padding=(8, 6)
        )
        row_locator_btn2.pack(fill=X, pady=(4, 0))
        self.btn_locator_everything = tb.Button(
            row_locator_btn2,
            text=self.tr("btn_locator_everything"),
            command=self.search_with_everything,
            bootstyle="warning-outline",
            width=14,
        )
        self.btn_locator_everything.pack(side=LEFT, padx=2)
        self._attach_tooltip(self.btn_locator_everything, "tip_locator_everything")
        self.btn_locator_copy_listary = tb.Button(
            row_locator_btn2,
            text=self.tr("btn_locator_copy_listary"),
            command=self.copy_listary_query,
            bootstyle="dark-outline",
            width=16,
        )
        self.btn_locator_copy_listary.pack(side=LEFT, padx=2)
        self._attach_tooltip(self.btn_locator_copy_listary, "tip_locator_listary")
        self.var_locator_result = tk.StringVar(value=self.tr("msg_locator_waiting"))
        tb.Label(
            lf_locator,
            textvariable=self.var_locator_result,
            bootstyle="secondary",
            wraplength=880,
            justify=LEFT,
        ).pack(anchor="w", pady=(6, 0))
        self.last_locate_record = None
        self._set_locator_action_state(False)
        self._auto_attach_action_tooltips(lf_mode)
        self._auto_attach_action_tooltips(lf_collect)
        # MSHelp Labelframe 已删除。
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
        # MSHelp Labelframe 已删除，不再附加 tip_mode_mshelp。
        self._auto_attach_input_tooltips(lf_locator, "tip_section_run_locator")
        self._build_task_tab_content()
        self._attach_tooltip(self.entry_temp_sandbox_root, "tip_input_sandbox_root")
        self._attach_tooltip(self.cb_strat, "tip_input_strategy")
        self._attach_tooltip(self.ent_date, "tip_input_date")
        self._bind_var_validation(
            self.var_locator_page,
            lambda: self._normalize_then_validate(
                self.var_locator_page, self._normalize_numeric_var, "locator"
            ),
        )
        self._bind_var_validation(
            self.var_locator_short_id,
            lambda: self._normalize_then_validate(
                self.var_locator_short_id, self._normalize_short_id_var, "locator"
            ),
        )
        self._bind_var_validation(
            self.var_date_str,
            lambda: self._normalize_then_validate(
                self.var_date_str, self._normalize_date_var, "run"
            ),
        )
        self._bind_var_validation(
            self.var_enable_date_filter,
            lambda: self.validate_runtime_inputs(silent=False, scope="run"),
        )

