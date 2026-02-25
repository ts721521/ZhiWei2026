# -*- coding: utf-8 -*-
"""Interactive prompt helpers extracted from office_converter.py."""


def confirm_continue_missing_md_merge(interactive, input_fn=input, warn_func=None):
    msg = (
        "\n[WARN] Markdown merge requested but no .md files were found. "
        "Continue with remaining tasks? [Y/n]: "
    )
    if interactive:
        try:
            ans = input_fn(msg).strip().lower()
            return ans in ("", "y", "yes")
        except (EOFError, KeyboardInterrupt, OSError, RuntimeError, TypeError, ValueError):
            return False
    if callable(warn_func):
        warn_func(
            "Markdown merge requested but no .md files found. Continue by default in non-interactive mode."
        )
    return True
