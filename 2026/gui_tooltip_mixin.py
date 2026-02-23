# -*- coding: utf-8 -*-
"""Tooltip helper methods extracted from OfficeGUI."""

from tkinter.constants import *

try:
    import ttkbootstrap as tb
except ModuleNotFoundError:
    from tkinter import ttk

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

    class _FallbackButton(_BootstyleMixin, ttk.Button):
        pass

    class _TBNamespace:
        Frame = _FallbackFrame
        Button = _FallbackButton

    tb = _TBNamespace()


class TooltipMixin:
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
            self._hover_tip_cls(
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
            self._hover_tip_cls(
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
            self.tr(
                "chk_markdown_strip_header_footer"
            ): "tip_toggle_markdown_strip_header_footer",
            self.tr(
                "chk_markdown_structured_headings"
            ): "tip_toggle_markdown_structured_headings",
            self.tr(
                "chk_markdown_quality_report"
            ): "tip_toggle_markdown_quality_report",
            self.tr("chk_export_records_json"): "tip_toggle_export_records_json",
            self.tr("chk_chromadb_export"): "tip_toggle_chromadb_export",
            self.tr("chk_incremental_mode"): "tip_toggle_incremental_mode",
            self.tr(
                "chk_incremental_verify_hash"
            ): "tip_toggle_incremental_verify_hash",
            self.tr(
                "chk_incremental_reprocess_renamed"
            ): "tip_toggle_incremental_reprocess_renamed",
            self.tr(
                "chk_source_priority_skip_pdf"
            ): "tip_toggle_source_priority_skip_pdf",
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
            self.tr(
                "rad_merge_convert_merge_only"
            ): "tip_option_merge_submode_merge_only",
            self.tr(
                "rad_merge_convert_pdf_to_md"
            ): "tip_option_merge_submode_pdf_to_md",
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
            self.tr(
                "lbl_office_restart_every"
            ): "tip_input_office_restart_every_n_files",
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
            (
                "ent_office_restart_every_n_files",
                "tip_input_office_restart_every_n_files",
            ),
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
            base = self.tr("tip_auto_config_item").format(
                label_text or child.winfo_class()
            )
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

# ===================== ???? =====================


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
                "程序启动异常，请关闭其他知微/OfficeGUI 进程后重试。\\n\\n错误信息：\\n"
                + str(e),
            )
        except Exception:
            pass
        raise

