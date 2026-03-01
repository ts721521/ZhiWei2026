# -*- coding: utf-8 -*-
"""Error classification helpers extracted from office_converter.py."""


class ConversionErrorType:
    """Normalized conversion error types used by reporting and retry logic."""

    PERMISSION_DENIED = "permission_denied"
    FILE_LOCKED = "file_locked"
    FILE_NOT_FOUND = "file_not_found"
    FILE_CORRUPTED = "file_corrupted"
    COM_ERROR = "com_error"
    TIMEOUT = "timeout"
    DISK_FULL = "disk_full"
    INVALID_FORMAT = "invalid_format"
    PASSWORD_PROTECTED = "password_protected"
    UNSUPPORTED_FORMAT = "unsupported_format"
    INVALID_PATH = "invalid_path"
    UNKNOWN = "unknown"


def classify_conversion_error(exception, context=""):
    """Classify conversion exception into a structured error descriptor."""
    err_str = str(exception).lower() if exception else ""
    ctx_str = str(context).lower() if context else ""
    signal = f"{err_str} | {ctx_str}"

    if any(
        kw in signal
        for kw in ("permission", "access denied", "unauthorized", "拒绝访问", "权限")
    ):
        return {
            "error_type": ConversionErrorType.PERMISSION_DENIED,
            "error_category": "needs_manual",
            "message": "File access permission denied",
            "suggestion": (
                "1. Run as administrator\n"
                "2. Check file/folder permissions\n"
                "3. Remove read-only/lock flags"
            ),
            "is_retryable": False,
            "requires_manual_action": True,
        }

    if any(
        kw in signal
        for kw in ("being used", "locked", "sharing violation", "占用", "另一个程序")
    ):
        return {
            "error_type": ConversionErrorType.FILE_LOCKED,
            "error_category": "retryable",
            "message": "File is locked by another process",
            "suggestion": "Close Office/File Explorer preview and retry.",
            "is_retryable": True,
            "requires_manual_action": False,
        }

    if any(kw in signal for kw in ("not found", "does not exist", "filenotfound", "no such file")):
        return {
            "error_type": ConversionErrorType.FILE_NOT_FOUND,
            "error_category": "unrecoverable",
            "message": "Source file not found",
            "suggestion": "Verify source path and file existence.",
            "is_retryable": False,
            "requires_manual_action": True,
        }

    if any(
        kw in signal
        for kw in ("corrupt", "damaged", "repair", "unreadable", "invalid data", "EOF marker not found")
    ):
        return {
            "error_type": ConversionErrorType.FILE_CORRUPTED,
            "error_category": "unrecoverable",
            "message": "File appears corrupted or unreadable",
            "suggestion": "Open and repair the file manually, then re-save and retry.",
            "is_retryable": False,
            "requires_manual_action": True,
        }

    if any(
        kw in signal
        for kw in (
            "com_error",
            "com error",
            "com object",
            "rpc",
            "0x800",
            "server busy",
            "call was rejected",
            "word.application",
            "excel.application",
            "powerpoint.application",
            "<unknown>.open",
            "conversion worker failed: attributeerror",
            "conversion worker failed: runtimeerror: com",
        )
    ):
        return {
            "error_type": ConversionErrorType.COM_ERROR,
            "error_category": "retryable",
            "message": "Office COM/automation error",
            "suggestion": (
                "1. Restart Office/WPS processes\n"
                "2. Reduce concurrency and retry\n"
                "3. Reopen Office manually once, then rerun"
            ),
            "is_retryable": True,
            "requires_manual_action": False,
        }

    if any(kw in signal for kw in ("timeout", "timed out", "超时")):
        return {
            "error_type": ConversionErrorType.TIMEOUT,
            "error_category": "retryable",
            "message": "Conversion timeout",
            "suggestion": "Increase timeout or split workload and retry.",
            "is_retryable": True,
            "requires_manual_action": False,
        }

    if any(kw in signal for kw in ("disk full", "no space", "空间不足", "storage")):
        return {
            "error_type": ConversionErrorType.DISK_FULL,
            "error_category": "needs_manual",
            "message": "Disk space is insufficient",
            "suggestion": "Free disk space or switch target/sandbox path.",
            "is_retryable": False,
            "requires_manual_action": True,
        }

    if any(kw in signal for kw in ("password", "encrypted", "protected", "加密", "密码")):
        return {
            "error_type": ConversionErrorType.PASSWORD_PROTECTED,
            "error_category": "needs_manual",
            "message": "Password-protected file",
            "suggestion": "Remove protection and retry.",
            "is_retryable": False,
            "requires_manual_action": True,
        }

    if any(
        kw in signal
        for kw in (
            "invalid format",
            "format not supported",
            "invalid pdf header",
            "unsupported",
        )
    ):
        return {
            "error_type": ConversionErrorType.INVALID_FORMAT,
            "error_category": "unrecoverable",
            "message": "Invalid or unsupported format",
            "suggestion": "Confirm extension/content match and re-save in a supported format.",
            "is_retryable": False,
            "requires_manual_action": True,
        }

    if any(
        kw in signal
        for kw in (
            "[errno 22]",
            "invalid argument",
            "filename, directory name, or volume label syntax is incorrect",
            "path too long",
        )
    ):
        return {
            "error_type": ConversionErrorType.INVALID_PATH,
            "error_category": "needs_manual",
            "message": "Invalid path or filename",
            "suggestion": "Shorten/rename path and remove invalid filename characters.",
            "is_retryable": False,
            "requires_manual_action": True,
        }

    return {
        "error_type": ConversionErrorType.UNKNOWN,
        "error_category": "unknown",
        "message": f"Unknown error: {str(exception)[:100] if exception else 'N/A'}",
        "suggestion": "Check logs and failed-file trace; then retry after cleanup.",
        "is_retryable": True,
        "requires_manual_action": True,
    }

