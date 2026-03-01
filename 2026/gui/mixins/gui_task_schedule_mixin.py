# -*- coding: utf-8 -*-
"""Task schedule (daily run at time) and scheduler thread."""

import threading
import time
from datetime import datetime

import tkinter as tk
from tkinter import messagebox


def _parse_daily_at(daily_at):
    """Return (hour, minute) or None. daily_at like '09:00' or '9:00'."""
    s = (daily_at or "").strip()
    if not s:
        return None
    parts = s.split(":")
    if len(parts) != 2:
        return None
    try:
        h, m = int(parts[0].strip()), int(parts[1].strip())
        if 0 <= h <= 23 and 0 <= m <= 59:
            return (h, m)
    except ValueError:
        pass
    return None


def _should_trigger_today(schedule, now):
    """True if schedule should run today: now >= daily_at and last_triggered is not today."""
    daily_at = (schedule.get("daily_at") or "09:00").strip()
    parsed = _parse_daily_at(daily_at)
    if not parsed:
        return False
    h, m = parsed
    now_time = now.time()
    from datetime import time as dt_time
    run_time = dt_time(h, m, 0)
    if now_time < run_time:
        return False
    last = (schedule.get("last_triggered") or "").strip()
    if not last:
        return True
    try:
        last_dt = datetime.fromisoformat(last.replace("Z", "+00:00"))
        if getattr(last_dt, "tzinfo", None):
            last_dt = last_dt.replace(tzinfo=None)
        if last_dt.date() >= now.date():
            return False
    except Exception:
        pass
    return True


class TaskScheduleMixin:
    """Schedule tasks to run daily at a given time. Scheduler runs in a daemon thread."""

    def _start_schedule_thread(self):
        """Start the schedule checker thread. Call once after GUI and task_store are ready."""
        if getattr(self, "_schedule_thread", None) is not None and self._schedule_thread.is_alive():
            return
        self._schedule_stop = False
        self._schedule_thread = threading.Thread(target=self._schedule_loop, daemon=True)
        self._schedule_thread.start()

    def _schedule_loop(self):
        while not getattr(self, "_schedule_stop", True):
            time.sleep(60)
            if getattr(self, "_schedule_stop", True):
                break
            try:
                self._schedule_tick()
            except Exception:
                pass

    def _schedule_tick(self):
        """Check schedules and trigger runs that are due."""
        task_store = getattr(self, "task_store", None)
        if not task_store:
            return
        if getattr(self, "worker_thread", None) and self.worker_thread.is_alive():
            return
        data = task_store.load_schedules()
        now = datetime.now()
        for s in data.get("schedules", []):
            if not isinstance(s, dict) or not s.get("enabled"):
                continue
            task_id = str(s.get("task_id", "")).strip()
            if not task_id:
                continue
            if not task_store.get_task(task_id):
                continue
            if not _should_trigger_today(s, now):
                continue
            now_iso = now.isoformat(timespec="seconds")
            task_store.update_schedule_last_triggered(task_id, now_iso)
            self.after(0, lambda tid=task_id: self._run_single_task(tid, False))

    def _open_task_schedule_dialog(self, task_id=None):
        """Open dialog to set daily schedule for a task. task_id defaults to selected."""
        task_id = task_id or (self._get_selected_task_id() if hasattr(self, "_get_selected_task_id") else None)
        if not task_id:
            messagebox.showinfo(
                self.tr("grp_task_runtime"), self.tr("msg_task_select_required"), parent=self
            )
            return
        task = self.task_store.get_task(task_id)
        if not task:
            return
        schedule = self.task_store.get_schedule(task_id) or {}
        win = tk.Toplevel(self)
        win.title(self.tr("win_task_schedule"))
        win.geometry("360x180")
        win.transient(self)
        win.grab_set()
        from tkinter import ttk
        frm = ttk.Frame(win, padding=12)
        frm.pack(fill="both", expand=True)
        var_enabled = tk.BooleanVar(value=bool(schedule.get("enabled")))
        var_daily_at = tk.StringVar(value=(schedule.get("daily_at") or "09:00").strip())

        ttk.Checkbutton(
            frm, text=self.tr("chk_task_schedule_enabled"), variable=var_enabled
        ).pack(anchor="w")
        row = ttk.Frame(frm)
        row.pack(fill="x", pady=(8, 0))
        ttk.Label(row, text=self.tr("lbl_task_schedule_daily_at")).pack(side="left")
        ttk.Entry(row, textvariable=var_daily_at, width=8).pack(side="left", padx=(8, 0))
        ttk.Label(row, text=" (HH:MM 24h)").pack(side="left")
        ttk.Label(frm, text=self.tr("msg_task_schedule_hint"), wraplength=320).pack(anchor="w", pady=(8, 0))

        def save():
            enabled = var_enabled.get()
            daily_at = var_daily_at.get().strip() or "09:00"
            if _parse_daily_at(daily_at) is None:
                messagebox.showwarning(
                    self.tr("win_task_schedule"),
                    self.tr("msg_task_schedule_invalid_time"),
                    parent=win,
                )
                return
            self.task_store.set_schedule(task_id, enabled, daily_at=daily_at)
            win.destroy()
            messagebox.showinfo(
                self.tr("win_task_schedule"),
                self.tr("msg_task_schedule_saved"),
                parent=self,
            )

        def cancel():
            win.destroy()

        btn_row = ttk.Frame(frm)
        btn_row.pack(fill="x", pady=(12, 0))
        ttk.Button(btn_row, text=self.tr("btn_save"), command=save).pack(side="left")
        ttk.Button(btn_row, text=self.tr("btn_cancel"), command=cancel).pack(side="left", padx=(8, 0))
        win.protocol("WM_DELETE_WINDOW", cancel)
