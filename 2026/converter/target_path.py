# -*- coding: utf-8 -*-
"""Target PDF path helpers extracted from office_converter.py."""

import os


def get_target_path(config, source_file_path, ext, prefix_override=None):
    filename = os.path.basename(source_file_path)
    base_name = os.path.splitext(filename)[0]
    ext_lower = ext.lower()

    if prefix_override:
        prefix = prefix_override
    else:
        prefix = ""
        word_exts = config["allowed_extensions"].get("word", [])
        excel_exts = config["allowed_extensions"].get("excel", [])
        ppt_exts = config["allowed_extensions"].get("powerpoint", [])
        pdf_exts = config["allowed_extensions"].get("pdf", [])

        if ext_lower in word_exts:
            prefix = "Word_"
        elif ext_lower in excel_exts:
            prefix = "Excel_"
        elif ext_lower in ppt_exts:
            prefix = "PPT_"
        elif ext_lower in pdf_exts:
            prefix = "PDF_"

    new_filename = f"{prefix}{base_name}.pdf"
    return os.path.join(config["target_folder"], new_filename)
