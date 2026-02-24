# -*- coding: utf-8 -*-
"""Callback emission helpers extracted from office_converter.py."""


def emit_file_plan(callback, file_list, warn_func=None):
    if not callable(callback):
        return
    try:
        callback(list(file_list or []))
    except Exception as e:
        if callable(warn_func):
            warn_func(f"file_plan_callback failed: {e}")


def emit_file_done(callback, record, warn_func=None):
    if not callable(callback):
        return
    if not isinstance(record, dict):
        return
    try:
        callback(dict(record))
    except Exception as e:
        if callable(warn_func):
            warn_func(f"file_done_callback failed: {e}")
