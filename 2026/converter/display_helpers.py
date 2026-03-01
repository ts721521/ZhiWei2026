# -*- coding: utf-8 -*-
"""Display helpers extracted from office_converter.py."""


def print_welcome(*, app_version, config_path, print_fn=print):
    print_fn("=" * 60)
    print_fn(f" 鐭ュ杺 ZhiWei 路 鐭ヨ瘑鎶曞杺宸ュ叿  v{app_version}")
    print_fn(" Supports WPS / Microsoft Office, CLI / GUI dual mode")
    print_fn("=" * 60)
    print_fn(f"Config file: {config_path}\n")


def print_step_title(text, *, print_fn=print):
    print_fn("\n" + "-" * 60)
    print_fn(text)
    print_fn("-" * 60)


def safe_console_print(text, *, end="\n", flush=False, print_fn=print):
    """Best-effort console print that never interrupts conversion flow."""
    try:
        print_fn(text, end=end, flush=flush)
        return
    except (OSError, UnicodeEncodeError, ValueError):
        pass

    fallback = str(text).encode("utf-8", "replace").decode("utf-8", "replace")
    fallback = fallback.encode("ascii", "replace").decode("ascii", "replace")
    try:
        print_fn(fallback, end=end, flush=flush)
    except (OSError, UnicodeEncodeError, ValueError):
        # Ignore console rendering failures; conversion should continue.
        return
