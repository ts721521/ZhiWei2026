# -*- coding: utf-8 -*-
"""Office process cycle helpers extracted from office_converter.py."""


def should_reuse_office_app(cfg, has_win32, is_mac_platform):
    if not has_win32 or is_mac_platform:
        return False
    cfg = cfg or {}
    return bool(cfg.get("office_reuse_app", True))


def get_office_restart_every(cfg):
    cfg = cfg or {}
    try:
        value = int(cfg.get("office_restart_every_n_files", 25))
    except (TypeError, ValueError, AttributeError):
        value = 25
    return value if value > 0 else 0


def get_app_type_for_ext(cfg, ext):
    cfg = cfg or {}
    exts = cfg.get("allowed_extensions", {})
    ext_lower = (ext or "").lower()
    if ext_lower in exts.get("word", []):
        return "word"
    if ext_lower in exts.get("excel", []):
        return "excel"
    if ext_lower in exts.get("powerpoint", []):
        return "ppt"
    return ""
