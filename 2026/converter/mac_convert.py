# -*- coding: utf-8 -*-
"""macOS conversion helpers extracted from office_converter.py."""

import os
import shutil
import subprocess


def convert_on_mac(
    file_source,
    sandbox_target_pdf,
    ext,
    *,
    is_mac_fn,
    which_fn=shutil.which,
    run_cmd=subprocess.run,
    devnull=subprocess.DEVNULL,
    dirname_fn=os.path.dirname,
    splitext_fn=os.path.splitext,
    basename_fn=os.path.basename,
    exists_fn=os.path.exists,
    move_fn=shutil.move,
    log_error_fn=None,
    log_warning_fn=None,
):
    _ = ext
    if not is_mac_fn():
        return False

    soffice = which_fn("soffice")
    if soffice:
        cmd = [
            soffice,
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            dirname_fn(sandbox_target_pdf),
            file_source,
        ]
        try:
            run_cmd(cmd, check=True, stdout=devnull, stderr=devnull)
            base = splitext_fn(basename_fn(file_source))[0]
            possible_output = os.path.join(dirname_fn(sandbox_target_pdf), base + ".pdf")
            if exists_fn(possible_output) and possible_output != sandbox_target_pdf:
                move_fn(possible_output, sandbox_target_pdf)
            return True
        except (OSError, RuntimeError, TypeError, ValueError, subprocess.CalledProcessError) as exc:
            if log_error_fn:
                log_error_fn(f"LibreOffice conversion failed: {exc}")

    if log_warning_fn:
        log_warning_fn(
            "macOS Office Automation not fully implemented. Install LibreOffice for best results."
        )
    return False
