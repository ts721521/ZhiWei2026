# -*- coding: utf-8 -*-
"""Process handling policy helpers extracted from office_converter.py."""

from converter.constants import (
    KILL_MODE_AUTO,
    KILL_MODE_KEEP,
    MODE_COLLECT_ONLY,
    MODE_MERGE_ONLY,
    MODE_MSHELP_ONLY,
)


def resolve_process_handling(run_mode, kill_process_mode, interactive):
    if run_mode in (MODE_MERGE_ONLY, MODE_COLLECT_ONLY, MODE_MSHELP_ONLY):
        return {"skip": True, "cleanup_all": False, "reuse_process": False}

    mode = kill_process_mode
    if mode == KILL_MODE_KEEP:
        return {"skip": False, "cleanup_all": False, "reuse_process": True}
    if mode == KILL_MODE_AUTO or not interactive:
        return {"skip": False, "cleanup_all": True, "reuse_process": False}
    return {"skip": False, "cleanup_all": False, "reuse_process": False}
