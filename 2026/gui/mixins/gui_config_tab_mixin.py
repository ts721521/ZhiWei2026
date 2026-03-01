# -*- coding: utf-8 -*-
"""Config tab UI methods extracted from OfficeGUI for maintainability."""

import tkinter as tk
from tkinter import ttk
from tkinter.constants import *
from office_converter import KILL_MODE_AUTO, KILL_MODE_KEEP

try:
    import ttkbootstrap as tb
    from ttkbootstrap.widgets.scrolled import ScrolledText
except ModuleNotFoundError:
    from tkinter.scrolledtext import ScrolledText as _TkScrolledText

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


class ConfigTabUIMixin:
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
        self._create_path_row(
            lf_cfg_path,
            "lbl_config",
            self.var_config_path,
            self.open_config_folder,
            None,
        )

        # Shared defaults: process strategy锛堝乏鍒楋級
        lf_proc_shared = tb.Labelframe(
            settings_left, text=self.tr("grp_cfg_shared_process"), padding=6
        )
        lf_proc_shared.pack(fill=X, pady=3)
        self._add_section_help(lf_proc_shared, "tip_section_cfg_process")
        tb.Label(
            lf_proc_shared, text=self.tr("lbl_kill_mode"), font=("System", 9)
        ).pack(anchor="w")
        self.var_kill_mode = tk.StringVar(value=KILL_MODE_AUTO)
        frm_kill = tb.Frame(lf_proc_shared)
        frm_kill.pack(fill=X)
        tb.Radiobutton(
            frm_kill,
            text=self.tr("rad_auto_kill"),
            variable=self.var_kill_mode,
            value=KILL_MODE_AUTO,
        ).pack(side=LEFT)
        tb.Radiobutton(
            frm_kill,
            text=self.tr("rad_keep_running"),
            variable=self.var_kill_mode,
            value=KILL_MODE_KEEP,
        ).pack(side=LEFT, padx=10)

        # Shared defaults: log output锛堝乏鍒楋級
        lf_cfg_log = tb.Labelframe(
            settings_left, text=self.tr("grp_cfg_shared_log"), padding=6
        )
        lf_cfg_log.pack(fill=X, pady=3)
        tb.Label(lf_cfg_log, text=self.tr("lbl_log_folder"), font=("System", 9)).pack(
            anchor="w", pady=(0, 0)
        )
        frm_log = tb.Frame(lf_cfg_log)
        frm_log.pack(fill=X)
        self.var_log_folder = tk.StringVar(value="./logs")
        self.ent_log_folder = tb.Entry(frm_log, textvariable=self.var_log_folder)
        self.ent_log_folder.pack(side=LEFT, fill=X, expand=YES)
        self._attach_tooltip(self.ent_log_folder, "tip_input_log_folder")
        self.btn_log_folder = tb.Button(
            frm_log,
            text=self.tr("btn_browse"),
            command=self.browse_log_folder,
            bootstyle="outline",
            width=3,
        )
        self.btn_log_folder.pack(side=LEFT, padx=2)
        self._attach_tooltip(self.btn_log_folder, "tip_choose_log")
        (
            frm_cfg_shared_actions,
            self.btn_save_cfg_shared,
            self.btn_reset_cfg_shared,
        ) = self._add_cfg_section_reset_action(settings_left, "shared")

        # Convert defaults锛堝乏鍒楋級
        lf_proc_convert = tb.Labelframe(
            settings_left, text=self.tr("grp_cfg_convert"), padding=6
        )
        lf_proc_convert.pack(fill=X, pady=3)
        self._add_section_help(lf_proc_convert, "tip_section_cfg_process")
        frm_time = tb.Frame(lf_proc_convert)
        frm_time.pack(fill=X, pady=5)
        self.var_timeout_seconds = tk.StringVar(value="60")
        tb.Label(frm_time, text=self.tr("lbl_gen_timeout")).grid(
            row=0, column=0, sticky="e"
        )
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
        tb.Label(frm_time, text=self.tr("lbl_pdf_wait")).grid(
            row=0, column=2, sticky="e"
        )
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
        tb.Label(frm_time, text=self.tr("lbl_ppt_timeout")).grid(
            row=1, column=0, sticky="e"
        )
        self.ent_ppt_timeout_seconds = spinbox_cls(
            frm_time,
            textvariable=self.var_ppt_timeout_seconds,
            width=5,
            from_=1,
            to=9999,
        )
        self.ent_ppt_timeout_seconds.grid(row=1, column=1, sticky="w", padx=5)
        self._attach_tooltip(
            self.ent_ppt_timeout_seconds, "tip_input_ppt_timeout_seconds"
        )
        self.var_ppt_pdf_wait_seconds = tk.StringVar(value="30")
        tb.Label(frm_time, text=self.tr("lbl_ppt_wait")).grid(
            row=1, column=2, sticky="e"
        )
        self.ent_ppt_pdf_wait_seconds = spinbox_cls(
            frm_time,
            textvariable=self.var_ppt_pdf_wait_seconds,
            width=5,
            from_=1,
            to=9999,
        )
        self.ent_ppt_pdf_wait_seconds.grid(row=1, column=3, sticky="w", padx=5)
        self._attach_tooltip(
            self.ent_ppt_pdf_wait_seconds, "tip_input_ppt_pdf_wait_seconds"
        )
        self.var_office_reuse_app = tk.IntVar(value=1)
        self.chk_office_reuse_app = tb.Checkbutton(
            frm_time,
            text=self.tr("chk_office_reuse_app"),
            variable=self.var_office_reuse_app,
        )
        self.chk_office_reuse_app.grid(
            row=2, column=0, columnspan=2, sticky="w", pady=(4, 0)
        )
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
            self.var_max_merge_size_mb.trace_add(
                "write", lambda *a: self.after(0, self._update_output_summary_label)
            )
        except Exception:
            pass
        tb.Label(
            lf_proc_merge_output,
            text=self.tr("lbl_merge_filename_pattern"),
            font=("System", 9),
        ).pack(anchor="w", pady=(8, 0))
        self.var_merge_filename_pattern = tk.StringVar(
            value="Merged_{category}_{timestamp}_{idx}"
        )
        self.ent_merge_filename_pattern = tb.Entry(
            lf_proc_merge_output,
            textvariable=self.var_merge_filename_pattern,
            width=45,
        )
        self.ent_merge_filename_pattern.pack(fill=X, pady=(2, 0))
        self._attach_tooltip(
            self.ent_merge_filename_pattern, "tip_input_merge_filename_pattern"
        )
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
        tb.Checkbutton(
            lf_cfg_ai_mshelp,
            text="Markdown image manifest (MD<->PDF map)",
            variable=self.var_enable_markdown_image_manifest,
        ).pack(anchor="w", pady=(6, 0))
        (
            frm_cfg_ai_actions,
            self.btn_save_cfg_ai,
            self.btn_reset_cfg_ai,
        ) = self._add_cfg_section_reset_action(settings_right, "ai")

        # 澧為噺閰嶇疆鍦ㄨ繍琛屽弬鏁拌浆鎹?tab 涓凡灞曠ず锛屼笉鍐嶅崟鐙崰鐢ㄩ〉闈?

        # UI / tooltip 閰嶇疆锛堝彸鍒楋級
        lf_proc_ui = tb.Labelframe(
            settings_right, text=self.tr("grp_cfg_ui"), padding=6
        )
        lf_proc_ui.pack(fill=X, pady=3)
        self._add_section_help(lf_proc_ui, "tip_section_cfg_process")
        tb.Label(
            lf_proc_ui, text=self.tr("lbl_tooltip_cfg"), font=("System", 9, "bold")
        ).pack(anchor="w")
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
        self.chk_tooltip_auto_theme = tb.Checkbutton(
            frm_tip,
            text=self.tr("chk_tooltip_auto_theme"),
            variable=self.var_tooltip_auto_theme,
        )
        self.chk_tooltip_auto_theme.grid(row=0, column=0, sticky="w", padx=(0, 8))
        self._attach_tooltip(
            self.chk_tooltip_auto_theme, "tip_toggle_tooltip_auto_theme"
        )
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
        tb.Label(frm_tip, text=self.tr("lbl_tooltip_delay")).grid(
            row=0, column=1, sticky="e"
        )
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
        tb.Label(frm_tip, text=self.tr("lbl_tooltip_font_size")).grid(
            row=0, column=3, sticky="e"
        )
        self.ent_tooltip_font_size = spinbox_cls(
            frm_tip,
            textvariable=self.var_tooltip_font_size,
            width=6,
            from_=6,
            to=48,
        )
        self.ent_tooltip_font_size.grid(row=0, column=4, sticky="w", padx=4)
        self._attach_tooltip(self.ent_tooltip_font_size, "tip_input_tooltip_font_size")
        self.var_tooltip_bg = tk.StringVar(value="#FFF7D6")
        tb.Label(frm_tip, text=self.tr("lbl_tooltip_bg")).grid(
            row=1, column=1, sticky="e"
        )
        self.ent_tooltip_bg = tb.Entry(
            frm_tip, textvariable=self.var_tooltip_bg, width=10
        )
        self.ent_tooltip_bg.grid(row=1, column=2, sticky="w", padx=4)
        self._attach_tooltip(self.ent_tooltip_bg, "tip_input_tooltip_bg")
        self.btn_pick_tooltip_bg = tb.Button(
            frm_tip, text="...", width=3, command=lambda: self.pick_tooltip_color("bg")
        )
        self.btn_pick_tooltip_bg.grid(row=1, column=2, sticky="e", padx=(0, 0))
        self._attach_tooltip(self.btn_pick_tooltip_bg, "tip_pick_color")
        self.var_tooltip_fg = tk.StringVar(value="#202124")
        tb.Label(frm_tip, text=self.tr("lbl_tooltip_fg")).grid(
            row=1, column=3, sticky="e"
        )
        self.ent_tooltip_fg = tb.Entry(
            frm_tip, textvariable=self.var_tooltip_fg, width=10
        )
        self.ent_tooltip_fg.grid(row=1, column=4, sticky="w", padx=4)
        self._attach_tooltip(self.ent_tooltip_fg, "tip_input_tooltip_fg")
        self.btn_pick_tooltip_fg = tb.Button(
            frm_tip, text="...", width=3, command=lambda: self.pick_tooltip_color("fg")
        )
        self.btn_pick_tooltip_fg.grid(row=1, column=4, sticky="e", padx=(0, 0))
        self._attach_tooltip(self.btn_pick_tooltip_fg, "tip_pick_color")
        self.lbl_tooltip_bg_preview = tb.Label(
            frm_tip, text=self.tr("lbl_tooltip_preview_bg"), width=12, anchor="center"
        )
        self.lbl_tooltip_bg_preview.grid(
            row=2, column=1, columnspan=2, sticky="w", pady=(4, 0)
        )
        self.lbl_tooltip_fg_preview = tb.Label(
            frm_tip, text=self.tr("lbl_tooltip_preview_fg"), width=12, anchor="center"
        )
        self.lbl_tooltip_fg_preview.grid(
            row=2, column=3, columnspan=2, sticky="w", pady=(4, 0)
        )
        self.lbl_tooltip_sample_preview = tb.Label(
            frm_tip,
            text=self.tr("lbl_tooltip_preview_sample"),
            anchor="center",
            padding=(8, 4),
        )
        self.lbl_tooltip_sample_preview.grid(
            row=3, column=1, columnspan=4, sticky="ew", pady=(4, 0)
        )
        self.btn_apply_tooltip = tb.Button(
            frm_tip,
            text=self.tr("btn_apply_tooltip"),
            command=self.apply_tooltip_settings,
            bootstyle="secondary-outline",
        )
        self.btn_apply_tooltip.grid(row=1, column=0, sticky="w", pady=(4, 0))
        self._attach_tooltip(self.btn_apply_tooltip, "tip_apply_tooltip")
        self.btn_reset_tooltip = tb.Button(
            frm_tip,
            text=self.tr("btn_reset_tooltip"),
            command=self.reset_tooltip_settings,
            bootstyle="secondary",
        )
        self.btn_reset_tooltip.grid(row=2, column=0, sticky="w", pady=(4, 0))
        self._attach_tooltip(self.btn_reset_tooltip, "tip_reset_tooltip")
        for v in (
            self.var_tooltip_delay_ms,
            self.var_tooltip_font_size,
            self.var_tooltip_bg,
            self.var_tooltip_fg,
            self.var_tooltip_auto_theme,
        ):
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
        self._auto_attach_input_tooltips(
            lf_proc_merge_output, "tip_section_cfg_process"
        )
        self._auto_attach_input_tooltips(lf_cfg_ai_mshelp, "tip_mode_mshelp")
        self._auto_attach_input_tooltips(lf_proc_ui, "tip_section_cfg_process")
        self._auto_attach_input_tooltips(lf_rules_excluded, "tip_section_cfg_lists")
        self._auto_attach_input_tooltips(lf_rules_keywords, "tip_section_cfg_lists")
        self._attach_tooltip(self.txt_excluded_folders, "tip_input_excluded_folders")
        self._attach_tooltip(self.txt_price_keywords, "tip_input_price_keywords")
        self._bind_var_validation(
            self.var_timeout_seconds,
            lambda: self._normalize_then_validate(
                self.var_timeout_seconds, self._normalize_numeric_var, "config"
            ),
        )
        self._bind_var_validation(
            self.var_pdf_wait_seconds,
            lambda: self._normalize_then_validate(
                self.var_pdf_wait_seconds, self._normalize_numeric_var, "config"
            ),
        )
        self._bind_var_validation(
            self.var_ppt_timeout_seconds,
            lambda: self._normalize_then_validate(
                self.var_ppt_timeout_seconds, self._normalize_numeric_var, "config"
            ),
        )
        self._bind_var_validation(
            self.var_ppt_pdf_wait_seconds,
            lambda: self._normalize_then_validate(
                self.var_ppt_pdf_wait_seconds, self._normalize_numeric_var, "config"
            ),
        )
        self._bind_var_validation(
            self.var_office_restart_every_n_files,
            lambda: self._normalize_then_validate(
                self.var_office_restart_every_n_files,
                self._normalize_numeric_var,
                "config",
            ),
        )
        self._bind_var_validation(
            self.var_max_merge_size_mb,
            lambda: self._normalize_then_validate(
                self.var_max_merge_size_mb, self._normalize_numeric_var, "config"
            ),
        )
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
            self.var_enable_fast_md_engine,
            self.var_enable_traceability_anchor_and_map,
            self.var_enable_markdown_image_manifest,
            self.var_enable_prompt_wrapper,
            self.var_prompt_template_type,
            self.var_short_id_prefix,
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
            self.var_enable_parallel_conversion,
            self.var_parallel_workers,
            self.var_enable_checkpoint,
            self.var_checkpoint_auto_resume,
            self.var_tooltip_auto_theme,
            self.var_confirm_revert_dirty,
            self.var_tooltip_delay_ms,
            self.var_tooltip_font_size,
            self.var_tooltip_bg,
            self.var_tooltip_fg,
        ):
            self._bind_config_dirty_var(_dirty_var)

