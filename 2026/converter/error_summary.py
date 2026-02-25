# -*- coding: utf-8 -*-
"""Error summary presentation helper extracted from office_converter.py."""


def get_error_summary_for_display(detailed_error_records):
    summary = {}

    for record in detailed_error_records:
        et = record["error_type"]
        if et not in summary:
            summary[et] = {
                "message": record["message"],
                "suggestion": record["suggestion"],
                "is_retryable": record["is_retryable"],
                "requires_manual_action": record["requires_manual_action"],
                "files": [],
            }
        summary[et]["files"].append(record["file_name"])

    return summary
