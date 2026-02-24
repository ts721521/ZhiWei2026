# -*- coding: utf-8 -*-
"""Path resolution helpers extracted from office_converter.py."""

import os


def get_path_from_config(cfg, key_base, prefer_win=False, prefer_mac=False):
    cfg = cfg or {}
    val = None
    if prefer_win:
        val = cfg.get(f"{key_base}_win")
    elif prefer_mac:
        val = cfg.get(f"{key_base}_mac")
    if not val:
        val = cfg.get(key_base)
    if val:
        return os.path.abspath(val)
    return ""
