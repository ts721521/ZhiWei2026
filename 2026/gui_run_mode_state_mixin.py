# -*- coding: utf-8 -*-
"""Run-mode state and toggle handlers extracted from run-tab UI mixin."""

from office_converter import (
    MODE_CONVERT_ONLY,
    MODE_MERGE_ONLY,
    MODE_CONVERT_THEN_MERGE,
    MODE_COLLECT_ONLY,
    MODE_MSHELP_ONLY,
    MERGE_MODE_ALL_IN_ONE,
)


class RunModeStateMixin:
    def _report_nonfatal_run_mode_error(self, scope, exc):
        reporter = getattr(self, "_report_nonfatal_ui_error", None)
        if callable(reporter):
            try:
                reporter(scope, exc=exc)
            except Exception:
                pass

    def _set_task_tab_highlight(self, on):
        """Highlight the task tab in task mode and restore normal style otherwise."""
        if not hasattr(self, "main_notebook") or not hasattr(self, "tab_run_tasks"):
            return
        try:
            # 1. 纭繚鏂囨湰娌℃湁绗﹀彿 (Remove previous symbol logic)
            base_text = self.tr("grp_task_runtime")
            try:
                self.main_notebook.tab(self.tab_run_tasks, text=base_text)
            except Exception:
                pass  # Ignore if tab is hidden/invalid, handled by state update

            # 2. 鍒囨崲 Notebook 鏁翠綋鏍峰紡 (Task Mode -> Success Green)
            # 娉ㄦ剰锛歵tkbootstrap Notebook bootstyle 褰卞搷閫変腑 Tab 鐨勯鑹?(underline/text)
            if on:
                self.main_notebook.configure(bootstyle="success")
            else:
                self.main_notebook.configure(bootstyle="primary")
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
                if hasattr(self, "_all_tab_text_keys") and i < len(
                    self._all_tab_text_keys
                ):
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

        # Enable only tabs relevant to the current mode閿涘牆鎮庨獮?濮婂磭鎮婇妴涓甋Help/鐎规矮缍?娑撻缚浠涢崥?tab閿?
        self._set_run_tab_state(self.tab_run_shared, "normal")
        self._set_run_tab_state(
            self.tab_run_convert, "normal" if is_convert else "disabled"
        )
        # 閵嗗苯鎮庨獮?/ 濮婂磭鎮婇妴宥忕窗娴犵粯鍓伴崥鍫濊嫙閻╃鍙ч幋鏍ㄢ拪閻炲棙膩瀵繑妞傞弰鍓с仛
        self._set_run_tab_state(
            self.tab_run_merge,
            "normal" if (is_merge_related or is_collect) else "disabled",
        )
        # MSHelp閿涙矮绮庨崷?mshelp 濡€崇础閺勫墽銇?
        self._set_run_tab_state(
            self.tab_run_mshelp, "normal" if is_mshelp else "disabled"
        )
        # 鐎规矮缍呭銉ュ徔閿涙艾顫愮紒鍫濆讲閻?
        self._set_run_tab_state(self.tab_run_locator, "normal")
        # 閹存劖鐏夐弬鍥︽閿涙艾顫愮紒鍫濆讲閻?
        self._set_run_tab_state(self.tab_run_output, "normal")
        # 浠诲姟绠＄悊椤碉細浼犵粺妯″紡涓嬮殣钘忥紝浠诲姟妯″紡涓嬪彲瑙侊紙鐢?_update_task_tab_for_app_mode 缁熶竴鍐冲畾锛?
        if hasattr(self, "_update_task_tab_for_app_mode"):
            self._update_task_tab_for_app_mode()

        # Collect options
        if is_collect:
            self._set_widget_tree_state(self.frm_collect_opts, "normal")
        else:
            self._set_widget_tree_state(self.frm_collect_opts, "disabled")

        # 閼奉亜濮╅柅澶夎厬鐎电懓绨查惃鍕€婄仦?tab閿涘牆鎮庨獮鏈电啊濮婂磭鎮婇妴涓甋Help/鐎规矮缍呴敍?
        if not bool(getattr(self, "_suppress_run_tab_autoselect", False)):
            try:
                if is_collect:
                    # 濮婂磭鎮婂Ο鈥崇础娴ｈ法鏁ら妴灞芥値楠?/ 濮婂磭鎮婇妴宥堜粵閸氬牓銆?
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
                try:
                    child.configure(state="disabled")
                except Exception as e:
                    self._report_nonfatal_run_mode_error(
                        "run_mode.disable_date_children", e
                    )
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
        for attr in (
            "lbl_strategy",
            "cb_strat",
            "chk_date_filter",
            "lbl_merge",
            "chk_enable_merge",
            "chk_markdown_strip_header_footer",
            "chk_markdown_structured_headings",
            "chk_markdown_quality_report",
            "chk_export_records_json",
            "chk_chromadb_export",
            "chk_incremental_mode",
            "chk_source_priority_skip_pdf",
            "chk_global_md5_dedup",
            "chk_enable_update_package",
            "chk_incremental_verify_hash",
            "chk_incremental_reprocess_renamed",
        ):
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
        for attr in (
            "chk_enable_sandbox",
            "entry_temp_sandbox_root",
            "btn_temp_sandbox_root",
            "entry_sandbox_min_free_gb",
            "cb_sandbox_low_space_policy",
        ):
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
                self._set_disabled_reason_in_tree(
                    self.frm_merge_opts, reason_merged_off
                )
            elif all_in_one and hasattr(self, "ent_max_merge_size_mb"):
                setattr(
                    self.ent_max_merge_size_mb,
                    "_tooltip_disabled_reason",
                    reason_all_in_one
                    if _try_state(self.ent_max_merge_size_mb) == "disabled"
                    else None,
                )
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
                if (
                    hasattr(self, "var_max_merge_size_mb")
                    and self.var_max_merge_size_mb.get()
                ):
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
            elif (
                getattr(self, "var_merge_mode", None)
                and self.var_merge_mode.get() == MERGE_MODE_ALL_IN_ONE
            ):
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
            try:
                child.configure(state=state)
            except Exception as e:
                self._report_nonfatal_run_mode_error(
                    "run_mode.toggle_date_filter.child_state", e
                )
        self.ent_date.configure(state=state)

    def _on_toggle_sandbox(self):
        mode = self.var_run_mode.get()
        is_disabled_globally = mode in (
            MODE_COLLECT_ONLY,
            MODE_MERGE_ONLY,
            MODE_MSHELP_ONLY,
        )

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

    def _on_toggle_parallel_conversion(self):
        enabled = bool(self.var_enable_parallel_conversion.get())
        state = "normal" if enabled else "disabled"
        if hasattr(self, "_frm_parallel_sub"):
            self._set_widget_tree_state(self._frm_parallel_sub, state)

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

