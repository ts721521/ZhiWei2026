# -*- coding: utf-8 -*-
"""Platform and runtime utility helpers extracted from office_converter.py."""

import os
import sys


def get_app_path():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    # Source layout:
    #   <project_root>/converter/platform_utils.py
    # Runtime config and GUI expect <project_root> as app path.
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def is_mac():
    return sys.platform == "darwin"


def is_win():
    return sys.platform == "win32"


def clear_console():
    try:
        if sys.stdout.isatty():
            os.system("cls" if os.name == "nt" else "clear")
    except (AttributeError, OSError, TypeError, ValueError):
        pass
