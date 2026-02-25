# -*- coding: utf-8 -*-
"""Process operation helpers extracted from office_converter.py."""

import subprocess

from converter.constants import ENGINE_MS, ENGINE_WPS


def process_names_for_engine(engine_type):
    apps = (
        ["wps", "et", "wpp", "wpscenter", "wpscloudsvr"]
        if engine_type == ENGINE_WPS or engine_type is None
        else []
    )
    if engine_type == ENGINE_MS or engine_type is None:
        apps.extend(["winword", "excel", "powerpnt"])
    return apps


def kill_process_by_name(app_name, has_win32, run_cmd=None):
    if not app_name or not has_win32:
        return
    run_cmd = run_cmd or subprocess.run
    cmd = f"taskkill /F /IM {app_name}.exe"
    run_cmd(cmd, shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)


def kill_process_by_name_for_converter(
    app_name,
    *,
    has_win32,
    run_cmd=None,
):
    try:
        kill_process_by_name(app_name, has_win32=has_win32, run_cmd=run_cmd)
    except (OSError, RuntimeError, TypeError, ValueError, AttributeError):
        return False
    return True
