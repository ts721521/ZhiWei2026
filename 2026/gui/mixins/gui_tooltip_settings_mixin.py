# -*- coding: utf-8 -*-
"""Tooltip settings methods extracted from config logic mixin."""

from tkinter import colorchooser


class TooltipSettingsMixin:
    def _update_tooltip_color_preview(self):
        bg = (
            self.var_tooltip_bg.get().strip() if hasattr(self, "var_tooltip_bg") else ""
        )
        fg = (
            self.var_tooltip_fg.get().strip() if hasattr(self, "var_tooltip_fg") else ""
        )
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
        initial = (
            self.var_tooltip_bg.get().strip()
            if target == "bg"
            else self.var_tooltip_fg.get().strip()
        )
        _, hex_color = colorchooser.askcolor(
            color=initial, title=self.tr("tip_pick_color")
        )
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
            self._set_entry_valid_state(
                getattr(self, "ent_tooltip_font_size", None), ok
            )
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
                self.var_locator_result.set(
                    self.tr("msg_tooltip_invalid_color").format(invalid_label)
                )
            else:
                self.var_locator_result.set(
                    self.tr("msg_tooltip_invalid_number").format(invalid_label)
                )
        return invalid_label is None

    def reset_tooltip_settings(self):
        self.var_tooltip_delay_ms.set(str(self.TOOLTIP_DEFAULTS["tooltip_delay_ms"]))
        self.var_tooltip_font_size.set(str(self.TOOLTIP_DEFAULTS["tooltip_font_size"]))
        self.var_tooltip_bg.set(self.TOOLTIP_DEFAULTS["tooltip_bg"])
        self.var_tooltip_fg.set(self.TOOLTIP_DEFAULTS["tooltip_fg"])
        self.var_tooltip_auto_theme.set(
            1 if self.TOOLTIP_DEFAULTS["tooltip_auto_theme"] else 0
        )
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
            self.var_tooltip_delay_ms.get(),
            self.tooltip_delay_ms,
            min_value=50,
            max_value=5000,
        )
        self.var_tooltip_delay_ms.set(str(self.tooltip_delay_ms))
        self.tooltip_font_size = _to_int(
            self.var_tooltip_font_size.get(),
            self.tooltip_font_size,
            min_value=8,
            max_value=20,
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
