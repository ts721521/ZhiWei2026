# -*- coding: utf-8 -*-
"""任务定时：支持 每日 / 每周 / 间隔 / 一次性 四种频率。

调度线程每 60s 轮询 schedules.json。worker 繁忙时到期项压入 `_pending_triggers`
队列，worker 空闲后按先到先跑顺序消费，避免长任务导致整日漏触发。
"""

import threading
import time
from datetime import datetime, timedelta, time as dt_time

import tkinter as tk
from tkinter import messagebox


_VALID_KINDS = ("daily", "weekly", "interval", "once")
_SCHEDULE_TICK_SECONDS = 60


def _parse_hhmm(value):
    s = (value or "").strip()
    if not s:
        return None
    parts = s.split(":")
    if len(parts) != 2:
        return None
    try:
        h, m = int(parts[0].strip()), int(parts[1].strip())
    except ValueError:
        return None
    if 0 <= h <= 23 and 0 <= m <= 59:
        return h, m
    return None


def _parse_once_at(value):
    """'YYYY-MM-DD HH:MM' -> datetime 或 None。"""
    s = (value or "").strip()
    if not s:
        return None
    for fmt in ("%Y-%m-%d %H:%M", "%Y-%m-%d %H:%M:%S", "%Y/%m/%d %H:%M"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    return None


def _parse_iso(value):
    s = (value or "").strip()
    if not s:
        return None
    try:
        dt = datetime.fromisoformat(s.replace("Z", "+00:00"))
        if getattr(dt, "tzinfo", None):
            dt = dt.replace(tzinfo=None)
        return dt
    except Exception:
        return None


def _kind_of(schedule):
    k = str((schedule or {}).get("kind") or "daily").strip().lower()
    return k if k in _VALID_KINDS else "daily"


def _should_trigger_now(schedule, now):
    """按 kind 判断 schedule 是否到期。

    daily:    now.time >= daily_at 且 last_triggered 不是今天
    weekly:   今天的 weekday 在 weekdays 内；其余同 daily
    interval: now - last_triggered >= interval_minutes（无 last 视为立刻到期）
    once:     now >= once_at 且未触发过
    """
    if not isinstance(schedule, dict) or not schedule.get("enabled"):
        return False
    kind = _kind_of(schedule)
    last_dt = _parse_iso(schedule.get("last_triggered") or "")

    if kind == "interval":
        minutes = int(schedule.get("interval_minutes") or 60)
        if minutes < 1:
            minutes = 1
        if last_dt is None:
            return True
        return now - last_dt >= timedelta(minutes=minutes)

    if kind == "once":
        target = _parse_once_at(schedule.get("once_at") or "")
        if target is None:
            return False
        if last_dt is not None:
            return False
        return now >= target

    hm = _parse_hhmm(schedule.get("daily_at") or "09:00")
    if not hm:
        return False
    run_time = dt_time(hm[0], hm[1], 0)

    if kind == "weekly":
        weekdays = schedule.get("weekdays") or []
        if now.weekday() not in {int(x) for x in weekdays if isinstance(x, int) or str(x).isdigit()}:
            return False

    if now.time() < run_time:
        return False
    if last_dt is not None and last_dt.date() >= now.date():
        return False
    return True


def summarize_schedule(schedule, tr=None, weekday_labels=None):
    """把一条 schedule 压成一句简短文案供任务列表展示。"""
    if not isinstance(schedule, dict) or not schedule.get("enabled"):
        return "—"
    kind = _kind_of(schedule)
    tr_fn = tr if callable(tr) else (lambda k: k)
    if kind == "interval":
        return tr_fn("schedule_fmt_interval").format(int(schedule.get("interval_minutes") or 60))
    if kind == "once":
        return tr_fn("schedule_fmt_once").format(str(schedule.get("once_at") or "").strip() or "—")
    daily_at = str(schedule.get("daily_at") or "09:00").strip()[:5]
    if kind == "weekly":
        labels = weekday_labels or ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
        days = sorted({int(x) for x in (schedule.get("weekdays") or []) if 0 <= int(x) <= 6})
        if not days:
            return tr_fn("schedule_fmt_weekly_empty").format(daily_at)
        joined = "".join(labels[d] for d in days)
        return tr_fn("schedule_fmt_weekly").format(joined, daily_at)
    return tr_fn("schedule_fmt_daily").format(daily_at)


class TaskScheduleMixin:
    """定时任务：后台 daemon 线程 + 待触发队列。"""

    # ---------- lifecycle ----------
    def _start_schedule_thread(self):
        if getattr(self, "_schedule_thread", None) is not None and self._schedule_thread.is_alive():
            return
        self._schedule_stop = False
        self._pending_triggers = list(getattr(self, "_pending_triggers", []) or [])
        self._schedule_thread = threading.Thread(target=self._schedule_loop, daemon=True)
        self._schedule_thread.start()

    def _schedule_loop(self):
        while not getattr(self, "_schedule_stop", True):
            time.sleep(_SCHEDULE_TICK_SECONDS)
            if getattr(self, "_schedule_stop", True):
                break
            try:
                self._schedule_tick()
            except Exception:
                pass

    # ---------- tick + queue ----------
    def _schedule_tick(self):
        task_store = getattr(self, "task_store", None)
        if not task_store:
            return
        data = task_store.load_schedules()
        now = datetime.now()
        pending = list(getattr(self, "_pending_triggers", []) or [])
        for s in data.get("schedules", []):
            if not isinstance(s, dict):
                continue
            task_id = str(s.get("task_id", "")).strip()
            if not task_id or task_id in pending:
                continue
            if not task_store.get_task(task_id):
                continue
            if not _should_trigger_now(s, now):
                continue
            pending.append(task_id)
        self._pending_triggers = pending
        self._drain_pending_triggers()

    def _drain_pending_triggers(self):
        """从 pending 队列抽一个塞给 worker；worker 忙则下一轮再抽。"""
        if getattr(self, "worker_thread", None) and self.worker_thread.is_alive():
            return
        pending = list(getattr(self, "_pending_triggers", []) or [])
        if not pending:
            return
        task_id = pending.pop(0)
        self._pending_triggers = pending
        now_iso = datetime.now().isoformat(timespec="seconds")
        try:
            self.task_store.update_schedule_last_triggered(task_id, now_iso)
        except Exception:
            pass
        self.after(0, lambda tid=task_id: self._run_single_task(tid, False))

    # ---------- manual actions ----------
    def _trigger_task_now(self, task_id):
        """对话框「立即试触发」：立刻排入 pending，worker 空就会立刻跑。"""
        task_id = (task_id or "").strip()
        if not task_id:
            return
        pending = list(getattr(self, "_pending_triggers", []) or [])
        if task_id not in pending:
            pending.insert(0, task_id)
        self._pending_triggers = pending
        self._drain_pending_triggers()

    def _delete_task_schedule(self, task_id):
        task_id = (task_id or "").strip()
        if not task_id:
            return False
        ok = False
        try:
            ok = bool(self.task_store.delete_schedule(task_id))
        except Exception:
            ok = False
        pending = [x for x in (getattr(self, "_pending_triggers", []) or []) if x != task_id]
        self._pending_triggers = pending
        return ok

    # ---------- dialog ----------
    def _open_task_schedule_dialog(self, task_id=None):
        task_id = task_id or (
            self._get_selected_task_id() if hasattr(self, "_get_selected_task_id") else None
        )
        if not task_id:
            messagebox.showinfo(
                self.tr("grp_task_runtime"),
                self.tr("msg_task_select_required"),
                parent=self,
            )
            return
        task = self.task_store.get_task(task_id)
        if not task:
            return
        schedule = self.task_store.get_schedule(task_id) or {}

        win = tk.Toplevel(self)
        win.title(self.tr("win_task_schedule"))
        win.geometry("460x400")
        win.transient(self)
        win.grab_set()
        from tkinter import ttk

        frm = ttk.Frame(win, padding=12)
        frm.pack(fill="both", expand=True)

        var_enabled = tk.BooleanVar(value=bool(schedule.get("enabled")))
        var_kind = tk.StringVar(value=_kind_of(schedule))
        var_daily_at = tk.StringVar(value=(schedule.get("daily_at") or "09:00").strip())
        var_interval = tk.StringVar(value=str(int(schedule.get("interval_minutes") or 60)))
        var_once_at = tk.StringVar(value=(schedule.get("once_at") or "").strip())
        weekday_vars = [tk.IntVar(value=0) for _ in range(7)]
        for d in schedule.get("weekdays") or []:
            try:
                idx = int(d)
                if 0 <= idx <= 6:
                    weekday_vars[idx].set(1)
            except (TypeError, ValueError):
                continue

        ttk.Label(
            frm,
            text=str(task.get("name") or task_id),
            font=("System", 9, "bold"),
        ).pack(anchor="w")
        ttk.Checkbutton(
            frm, text=self.tr("chk_task_schedule_enabled"), variable=var_enabled
        ).pack(anchor="w", pady=(8, 4))

        kind_row = ttk.LabelFrame(frm, text=self.tr("lbl_task_schedule_kind"), padding=6)
        kind_row.pack(fill="x", pady=(4, 6))
        for value, key in (
            ("daily", "schedule_kind_daily"),
            ("weekly", "schedule_kind_weekly"),
            ("interval", "schedule_kind_interval"),
            ("once", "schedule_kind_once"),
        ):
            ttk.Radiobutton(
                kind_row, text=self.tr(key), variable=var_kind, value=value
            ).pack(side="left", padx=(0, 8))

        panels = ttk.Frame(frm)
        panels.pack(fill="x", pady=(0, 6))

        # daily / weekly 共用 HH:MM，weekly 额外挂 7 个勾
        panel_daily = ttk.Frame(panels)
        row_d = ttk.Frame(panel_daily)
        row_d.pack(fill="x")
        ttk.Label(row_d, text=self.tr("lbl_task_schedule_daily_at")).pack(side="left")
        ttk.Entry(row_d, textvariable=var_daily_at, width=8).pack(side="left", padx=(8, 0))
        ttk.Label(row_d, text=" (HH:MM 24h)").pack(side="left")

        panel_weekly = ttk.Frame(panels)
        row_w1 = ttk.Frame(panel_weekly)
        row_w1.pack(fill="x")
        ttk.Label(row_w1, text=self.tr("lbl_task_schedule_daily_at")).pack(side="left")
        ttk.Entry(row_w1, textvariable=var_daily_at, width=8).pack(side="left", padx=(8, 0))
        ttk.Label(row_w1, text=" (HH:MM 24h)").pack(side="left")
        row_w2 = ttk.Frame(panel_weekly)
        row_w2.pack(fill="x", pady=(6, 0))
        weekday_keys = (
            "weekday_mon", "weekday_tue", "weekday_wed",
            "weekday_thu", "weekday_fri", "weekday_sat", "weekday_sun",
        )
        for idx, key in enumerate(weekday_keys):
            ttk.Checkbutton(row_w2, text=self.tr(key), variable=weekday_vars[idx]).pack(side="left")

        panel_interval = ttk.Frame(panels)
        row_i = ttk.Frame(panel_interval)
        row_i.pack(fill="x")
        ttk.Label(row_i, text=self.tr("lbl_task_schedule_interval_minutes")).pack(side="left")
        ttk.Entry(row_i, textvariable=var_interval, width=8).pack(side="left", padx=(8, 0))

        panel_once = ttk.Frame(panels)
        row_o = ttk.Frame(panel_once)
        row_o.pack(fill="x")
        ttk.Label(row_o, text=self.tr("lbl_task_schedule_once_at")).pack(side="left")
        ttk.Entry(row_o, textvariable=var_once_at, width=18).pack(side="left", padx=(8, 0))
        ttk.Label(row_o, text=" (YYYY-MM-DD HH:MM)").pack(side="left")

        def _show_panel(*_):
            for p in (panel_daily, panel_weekly, panel_interval, panel_once):
                p.pack_forget()
            kind = var_kind.get()
            if kind == "weekly":
                panel_weekly.pack(fill="x")
            elif kind == "interval":
                panel_interval.pack(fill="x")
            elif kind == "once":
                panel_once.pack(fill="x")
            else:
                panel_daily.pack(fill="x")

        var_kind.trace_add("write", _show_panel)
        _show_panel()

        # 上次触发展示
        last = (schedule.get("last_triggered") or "").strip() or "—"
        ttk.Label(
            frm, text=self.tr("lbl_task_schedule_last_triggered") + ": " + last, foreground="#666"
        ).pack(anchor="w", pady=(4, 0))
        ttk.Label(
            frm, text=self.tr("msg_task_schedule_hint"), wraplength=420, foreground="#666"
        ).pack(anchor="w", pady=(6, 0))

        def _collect_and_validate():
            kind = var_kind.get()
            daily_at = var_daily_at.get().strip() or "09:00"
            interval = var_interval.get().strip() or "60"
            once_at = var_once_at.get().strip()
            weekdays = [i for i in range(7) if weekday_vars[i].get()]

            if kind in ("daily", "weekly") and _parse_hhmm(daily_at) is None:
                messagebox.showwarning(
                    self.tr("win_task_schedule"),
                    self.tr("msg_task_schedule_invalid_time"),
                    parent=win,
                )
                return None
            if kind == "weekly" and not weekdays:
                messagebox.showwarning(
                    self.tr("win_task_schedule"),
                    self.tr("msg_task_schedule_weekly_empty"),
                    parent=win,
                )
                return None
            if kind == "interval":
                try:
                    minutes = int(interval)
                    if minutes < 1:
                        raise ValueError
                except ValueError:
                    messagebox.showwarning(
                        self.tr("win_task_schedule"),
                        self.tr("msg_task_schedule_invalid_interval"),
                        parent=win,
                    )
                    return None
            else:
                minutes = 60
            if kind == "once" and _parse_once_at(once_at) is None:
                messagebox.showwarning(
                    self.tr("win_task_schedule"),
                    self.tr("msg_task_schedule_invalid_once_at"),
                    parent=win,
                )
                return None
            return {
                "kind": kind,
                "daily_at": daily_at,
                "weekdays": weekdays,
                "interval_minutes": minutes,
                "once_at": once_at,
            }

        def save():
            cfg = _collect_and_validate()
            if cfg is None:
                return
            self.task_store.set_schedule(
                task_id,
                bool(var_enabled.get()),
                daily_at=cfg["daily_at"],
                kind=cfg["kind"],
                weekdays=cfg["weekdays"],
                interval_minutes=cfg["interval_minutes"],
                once_at=cfg["once_at"],
            )
            if hasattr(self, "_refresh_task_list_ui"):
                self.after(0, self._refresh_task_list_ui)
            win.destroy()
            messagebox.showinfo(
                self.tr("win_task_schedule"),
                self.tr("msg_task_schedule_saved"),
                parent=self,
            )

        def delete():
            if not messagebox.askyesno(
                self.tr("win_task_schedule"),
                self.tr("msg_task_schedule_confirm_delete"),
                parent=win,
            ):
                return
            self._delete_task_schedule(task_id)
            if hasattr(self, "_refresh_task_list_ui"):
                self.after(0, self._refresh_task_list_ui)
            win.destroy()

        def trigger_now():
            self._trigger_task_now(task_id)
            messagebox.showinfo(
                self.tr("win_task_schedule"),
                self.tr("msg_task_schedule_triggered"),
                parent=win,
            )

        def cancel():
            win.destroy()

        btn_row = ttk.Frame(frm)
        btn_row.pack(fill="x", pady=(14, 0))
        ttk.Button(btn_row, text=self.tr("btn_save"), command=save).pack(side="left")
        ttk.Button(
            btn_row, text=self.tr("btn_task_schedule_trigger_now"), command=trigger_now
        ).pack(side="left", padx=(8, 0))
        ttk.Button(
            btn_row, text=self.tr("btn_task_schedule_delete"), command=delete
        ).pack(side="left", padx=(8, 0))
        ttk.Button(btn_row, text=self.tr("btn_cancel"), command=cancel).pack(side="right")
        win.protocol("WM_DELETE_WINDOW", cancel)
