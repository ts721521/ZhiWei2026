# -*- coding: utf-8 -*-
"""Misc UI helper methods extracted from OfficeGUI."""

import os
import queue
import subprocess
import sys
from tkinter import filedialog
from tkinter.constants import *


class MiscUIMixin:
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
        if (
            getattr(self, "var_show_tooltip_advanced", None)
            and self.var_show_tooltip_advanced.get()
        ):
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

    def _poll_log_queue(self):
        exists_checker = getattr(self, "winfo_exists", None)
        if callable(exists_checker):
            try:
                if not exists_checker():
                    return
            except Exception:
                return
        log_queue = getattr(self, "log_queue", None)
        if log_queue is None:
            # Backward-compatible fallback for legacy callers.
            log_queue = globals().get("LOG_QUEUE")
        if log_queue is None:
            try:
                self._after_poll_log_id = self.after(200, self._poll_log_queue)
            except Exception:
                pass
            return
        try:
            while True:
                msg = log_queue.get_nowait()
                self.txt_log.insert("end", msg + "\n")
                self.txt_log.see("end")
        except queue.Empty:
            pass
        try:
            self._after_poll_log_id = self.after(200, self._poll_log_queue)
        except Exception:
            pass


# ===================== 绋嬪簭鍏ュ彛 =====================


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
                "鍚姩澶辫触",
                "绋嬪簭鍚姩寮傚父锛岃鍏抽棴鍏朵粬鐭ュ杺/OfficeGUI 杩涚▼鍚庨噸璇曘€俓n\n閿欒淇℃伅锛歕n"
                + str(e),
            )
        except Exception:
            pass
        raise

