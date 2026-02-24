# -*- coding: utf-8 -*-
"""File operation helpers extracted from office_converter.py."""

import os
import shutil
from datetime import datetime


def unblock_file(file_path):
    try:
        zone_path = file_path + ":Zone.Identifier"
        try:
            os.remove(zone_path)
        except Exception:
            pass
    except Exception:
        pass


def copy_pdf_direct(source, temp_target):
    try:
        shutil.copy2(source, temp_target)
    except Exception as e:
        raise Exception(f"[PDF copy failed] {e}")


def quarantine_failed_file(source_path, failed_dir, should_copy=True, now=None):
    if not should_copy:
        return None
    try:
        fname = os.path.basename(source_path)
        target = os.path.join(failed_dir, fname)
        if os.path.exists(target):
            name, ext = os.path.splitext(fname)
            dt = now or datetime.now()
            target = os.path.join(failed_dir, f"{name}_{dt.strftime('%H%M%S')}{ext}")
        shutil.copy2(source_path, target)
        return target
    except Exception:
        return None


def handle_file_conflict(temp_pdf_path, target_pdf_path, now=None):
    if not os.path.exists(target_pdf_path):
        os.makedirs(os.path.dirname(target_pdf_path), exist_ok=True)
        shutil.move(temp_pdf_path, target_pdf_path)
        return "success", target_pdf_path

    if os.path.getsize(temp_pdf_path) == os.path.getsize(target_pdf_path):
        try:
            os.remove(target_pdf_path)
            shutil.move(temp_pdf_path, target_pdf_path)
            return "overwrite", target_pdf_path
        except Exception:
            return "overwrite_failed", target_pdf_path
    conflict_dir = os.path.join(os.path.dirname(target_pdf_path), "conflicts")
    os.makedirs(conflict_dir, exist_ok=True)
    fname = os.path.splitext(os.path.basename(target_pdf_path))[0]
    ts = (now or datetime.now()).strftime("%Y%m%d%H%M%S")
    new_path = os.path.join(conflict_dir, f"{fname}_{ts}.pdf")
    shutil.move(temp_pdf_path, new_path)
    return "conflict_saved", new_path
