# -*- coding: utf-8 -*-
"""Source folder dialog helpers extracted from run-tab UI mixin."""

import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
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


class SourceFolderMixin:
    def add_source_folder(self):
        """Add source folder with option to select multiple folders."""
        # Show dialog to choose single or multi select
        if not hasattr(self, "_multi_folder_dialog"):
            self._multi_folder_dialog = None

        # Ask user if they want to select multiple folders
        result = messagebox.askyesno(
            self.tr("tip_add_source_folder"),
            self.tr("msg_multi_select_folders"),
            icon="question",
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
        if (
            hasattr(self, "_multi_folder_dialog")
            and self._multi_folder_dialog is not None
        ):
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
        tb.Label(
            frame,
            text="Method 1: scan a parent folder (recommended, add many at once)",
            font=("System", 9, "bold"),
            foreground="#007bff",
        ).pack(anchor="w", pady=(0, 5))

        # Parent folder selection section
        parent_frame = tb.Frame(frame)
        parent_frame.pack(fill=X, pady=(0, 10))

        tb.Label(parent_frame, text="Parent folder:", font=("System", 9)).pack(
            side=LEFT, padx=(0, 8)
        )
        ent_parent = tb.Entry(parent_frame, width=50)
        ent_parent.pack(side=LEFT, fill=X, expand=YES)

        def pick_parent():
            p = filedialog.askdirectory(title="Select parent folder")
            if p:
                if sys.platform == "win32":
                    p = p.replace("/", "\\")
                ent_parent.delete(0, END)
                ent_parent.insert(0, p)

        def scan_subfolders():
            parent = ent_parent.get().strip()
            if not parent or not os.path.isdir(parent):
                messagebox.showwarning(
                    "Notice", "Please select a valid parent folder.", parent=dlg
                )
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
                count_label.config(text=f"Scanned {total} subfolders")

                messagebox.showinfo(
                    "Scan Complete",
                    f"Found {total} subfolders under the selected parent.",
                    parent=dlg,
                )

            except Exception as e:
                messagebox.showerror("Error", f"Scan failed: {e}", parent=dlg)

        tb.Button(
            parent_frame, text="Browse", command=pick_parent, bootstyle="info", width=8
        ).pack(side=LEFT, padx=(8, 4))
        tb.Button(
            parent_frame,
            text="Scan Subfolders",
            command=scan_subfolders,
            bootstyle="success",
            width=12,
        ).pack(side=LEFT, padx=(4, 0))

        # Manual add section
        tb.Label(
            frame,
            text="Method 2: add folders manually",
            font=("System", 9, "bold"),
            foreground="#666",
        ).pack(anchor="w", pady=(10, 5))

        # Treeview for folder selection with checkboxes
        left_frame = tb.Frame(frame)
        left_frame.pack(fill=BOTH, expand=YES)

        # Use a custom checkbox listbox approach
        cols = (
            "selected",
            "path",
        )
        folder_tree = ttk.Treeview(
            left_frame, columns=cols, show="headings", selectmode="extended"
        )
        folder_tree.heading("selected", text="", width=30)
        folder_tree.column("selected", width=30, anchor="center")
        folder_tree.heading("path", text="Folder Path")
        folder_tree.column("path", width=550, anchor="w")
        folder_tree.pack(side=LEFT, fill=BOTH, expand=YES)

        # Scrollbar for treeview
        tree_scroll = tb.Scrollbar(
            left_frame, orient=VERTICAL, command=folder_tree.yview
        )
        folder_tree.configure(yscrollcommand=tree_scroll.set)
        tree_scroll.pack(side=RIGHT, fill=Y)

        # Right: Buttons
        btn_frame = tb.Frame(frame)
        btn_frame.pack(side=RIGHT, fill=Y, padx=(10, 0))

        btn_browse = tb.Button(
            btn_frame,
            text="+ Browse",
            command=lambda: self._browse_folder_to_dialog(folder_tree),
            bootstyle="info",
            width=10,
        )
        btn_browse.pack(pady=2)

        tb.Label(btn_frame, text="", height=1).pack()  # Spacer

        btn_remove = tb.Button(
            btn_frame,
            text="- Remove",
            command=lambda: self._remove_folder_from_tree(folder_tree),
            bootstyle="danger-outline",
            width=10,
        )
        btn_remove.pack(pady=2)

        btn_clear = tb.Button(
            btn_frame,
            text="C Clear",
            command=lambda: self._clear_folder_tree(folder_tree),
            bootstyle="secondary-outline",
            width=10,
        )
        btn_clear.pack(pady=2)

        # Bottom section: Controls
        bottom_frame = tb.Frame(frame)
        bottom_frame.pack(fill=X, pady=(10, 0))

        # Count display
        self._multi_folder_count_var = tk.StringVar()
        count_label = tb.Label(
            bottom_frame, textvariable=self._multi_folder_count_var, font=("System", 9)
        )
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

        tb.Button(
            check_frame,
            text="Select All",
            command=select_all,
            bootstyle="info-outline",
            width=8,
        ).pack(side=LEFT, padx=(0, 4))
        tb.Button(
            check_frame,
            text="Deselect All",
            command=deselect_all,
            bootstyle="secondary-outline",
            width=8,
        ).pack(side=LEFT, padx=4)
        tb.Button(
            check_frame,
            text="Invert",
            command=invert_selection,
            bootstyle="secondary-outline",
            width=8,
        ).pack(side=LEFT, padx=4)

        # Initialize
        self._update_folder_count(folder_tree)

        # Action buttons
        action_btn_frame = tb.Frame(frame)
        action_btn_frame.pack(fill=X, pady=(10, 0))

        btn_add = tb.Button(
            action_btn_frame,
            text="Add Selected",
            command=lambda: self._confirm_multi_folder_selection(folder_tree),
            bootstyle="success",
            width=12,
        )
        btn_add.pack(side=LEFT)

        btn_cancel = tb.Button(
            action_btn_frame,
            text="Cancel",
            command=self._close_multi_folder_dialog,
            bootstyle="secondary-outline",
            width=10,
        )
        btn_cancel.pack(side=RIGHT)

    def _close_multi_folder_dialog(self):
        """Close the multi-folder selection dialog."""
        if (
            hasattr(self, "_multi_folder_dialog")
            and self._multi_folder_dialog is not None
        ):
            try:
                if (
                    self._multi_folder_dialog.grab_current()
                    == self._multi_folder_dialog
                ):
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
        base_dir = filedialog.askdirectory(
            title=self.tr("tip_add_source_folder") + " (select first folder)"
        )
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
                "Continue", "Add another folder?", icon="question"
            )
            if not result:
                break

            next_dir = filedialog.askdirectory(
                title=self.tr("tip_add_source_folder") + " (select next folder)"
            )
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
                self._multi_folder_count_var.set("Selected 0 folders")
            elif count == 1:
                self._multi_folder_count_var.set("Selected 1 folder")
            else:
                self._multi_folder_count_var.set(f"Selected {count} folders")

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
        if (
            hasattr(self, "_task_multi_folder_dialog")
            and self._task_multi_folder_dialog is not None
        ):
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
        dlg.protocol(
            "WM_DELETE_WINDOW", lambda: self._close_task_multi_folder_dialog(dlg)
        )
        self._task_multi_folder_dialog = dlg

        # Main frame with simple style
        frame = tk.Frame(dlg, padx=12, pady=12)
        frame.pack(fill=BOTH, expand=YES)

        # Instructions
        tk.Label(
            frame,
            text="Method 1: scan a parent folder (recommended, add many at once)",
            font=("System", 9, "bold"),
            fg="blue",
        ).pack(anchor="w", pady=(0, 5))

        # Parent folder selection section
        parent_frame = tk.Frame(frame)
        parent_frame.pack(fill=X, pady=(0, 8))

        tk.Label(parent_frame, text="Parent folder:", font=("System", 9)).pack(
            side=LEFT, padx=(0, 8)
        )
        ent_parent = tk.Entry(parent_frame, width=50)
        ent_parent.pack(side=LEFT, fill=X, expand=YES)

        def pick_parent():
            p = filedialog.askdirectory(title="Select parent folder")
            if p:
                if sys.platform == "win32":
                    p = p.replace("/", "\\")
                ent_parent.delete(0, END)
                ent_parent.insert(0, p)

        def scan_subfolders():
            parent = ent_parent.get().strip()
            if not parent or not os.path.isdir(parent):
                messagebox.showwarning(
                    "Notice", "Please select a valid parent folder.", parent=dlg
                )
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

                messagebox.showinfo(
                    "Scan Complete",
                    f"Found {len(subfolders)} subfolders under the selected parent.",
                    parent=dlg,
                )

            except Exception as e:
                messagebox.showerror("Error", f"Scan failed: {e}", parent=dlg)

        tk.Button(parent_frame, text="Browse", command=pick_parent, width=8).pack(
            side=LEFT, padx=(8, 4)
        )
        tk.Button(
            parent_frame,
            text="Scan Subfolders",
            command=scan_subfolders,
            bg="green",
            fg="white",
            width=12,
        ).pack(side=LEFT, padx=(4, 0))

        # Manual add section
        tk.Label(
            frame, text="Method 2: add folders manually", font=("System", 9, "bold"), fg="gray"
        ).pack(anchor="w", pady=(8, 5))

        # Top section with split layout
        top_frame = tk.Frame(frame)
        top_frame.pack(fill=BOTH, expand=YES, pady=(0, 10))

        # Left: Listbox for folder selection
        left_frame = tk.Frame(top_frame)
        left_frame.pack(side=LEFT, fill=BOTH, expand=YES)

        folder_listbox = tk.Listbox(
            left_frame, selectmode=EXTENDED, font=("Consolas", 9)
        )
        folder_listbox.pack(side=LEFT, fill=BOTH, expand=YES)

        # Scrollbar for listbox
        list_scroll = tk.Scrollbar(
            left_frame, orient=VERTICAL, command=folder_listbox.yview
        )
        folder_listbox.configure(yscrollcommand=list_scroll.set)
        list_scroll.pack(side=RIGHT, fill=Y)

        # Right: Buttons for add/remove
        btn_frame = tk.Frame(top_frame)
        btn_frame.pack(side=RIGHT, fill=Y, padx=(10, 0))

        tk.Button(
            btn_frame,
            text="+ Browse",
            width=8,
            command=lambda: self._task_browse_folder_to_dialog(folder_listbox),
        ).pack(pady=2)

        tk.Label(btn_frame, text="", height=1).pack()  # Spacer

        tk.Button(
            btn_frame,
            text="- Remove",
            width=8,
            command=lambda: self._task_remove_from_listbox(folder_listbox),
        ).pack(pady=2)
        tk.Button(
            btn_frame,
            text="C Clear",
            width=8,
            command=lambda: self._task_clear_listbox(folder_listbox),
        ).pack(pady=2)

        # Bottom section: Current selected folders info
        bottom_frame = tk.Frame(frame)
        bottom_frame.pack(fill=X, pady=(8, 0))

        task_count_var = tk.StringVar()
        task_count_label = tk.Label(
            bottom_frame, textvariable=task_count_var, font=("System", 9)
        )
        task_count_label.pack(anchor="w")

        # Selection controls
        check_frame = tk.Frame(bottom_frame)
        check_frame.pack(anchor="w", pady=(6, 0))

        def select_all():
            folder_listbox.selection_set(0, END)

        def deselect_all():
            folder_listbox.selection_clear(0, END)

        tk.Button(check_frame, text="Select All", command=select_all, width=8).pack(
            side=LEFT, padx=(0, 4)
        )
        tk.Button(check_frame, text="Deselect All", command=deselect_all, width=8).pack(
            side=LEFT, padx=4
        )

        # Update initial count
        def update_task_count():
            count = folder_listbox.size()
            if count == 0:
                task_count_var.set("Selected 0 folders")
            elif count == 1:
                task_count_var.set("Selected 1 folder")
            else:
                task_count_var.set(f"Selected {count} folders")

        update_task_count()

        # Action buttons
        action_btn_frame = tk.Frame(frame)
        action_btn_frame.pack(fill=X, pady=(10, 0))

        tk.Button(
            action_btn_frame,
            text="Add",
            width=10,
            command=lambda: self._task_confirm_selection(
                folder_listbox, target_listbox
            ),
        ).pack(side=LEFT)
        tk.Button(
            action_btn_frame,
            text="Cancel",
            width=10,
            command=lambda: self._close_task_multi_folder_dialog(dlg),
        ).pack(side=RIGHT)

