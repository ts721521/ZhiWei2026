# -*- coding: utf-8 -*-
"""Error classification helpers extracted from office_converter.py."""

# =============== Error Classification ===============
# 错误类型枚举 - 用于分类和生成处理建议


class ConversionErrorType:
    """转换错误类型分类"""

    PERMISSION_DENIED = "permission_denied"  # 权限不足
    FILE_LOCKED = "file_locked"  # 文件被占用
    FILE_NOT_FOUND = "file_not_found"  # 文件不存在
    FILE_CORRUPTED = "file_corrupted"  # 文件损坏
    COM_ERROR = "com_error"  # Office COM 错误
    TIMEOUT = "timeout"  # 超时
    DISK_FULL = "disk_full"  # 磁盘空间不足
    INVALID_FORMAT = "invalid_format"  # 格式无效
    PASSWORD_PROTECTED = "password_protected"  # 密码保护
    UNSUPPORTED_FORMAT = "unsupported_format"  # 不支持的格式
    UNKNOWN = "unknown"  # 未知错误


def classify_conversion_error(exception, context=""):
    """
    根据异常信息分类错误类型，返回错误类型和处理建议。

    Args:
        exception: 异常对象或异常信息字符串
        context: 额外的上下文信息（如文件路径）

    Returns:
        dict: {
            "error_type": 错误类型,
            "error_category": 错误分类（可重试/不可重试/需人工）,
            "message": 用户友好消息,
            "suggestion": 处理建议,
            "is_retryable": 是否可自动重试,
            "requires_manual_action": 是否需要人工干预
        }
    """
    err_str = str(exception).lower() if exception else ""
    err_type = type(exception).__name__ if hasattr(exception, "__name__") else ""

    # 权限错误
    if any(
        kw in err_str
        for kw in ["permission", "access denied", "拒绝访问", "权限", "unauthorized"]
    ):
        return {
            "error_type": ConversionErrorType.PERMISSION_DENIED,
            "error_category": "needs_manual",
            "message": "文件访问权限不足",
            "suggestion": "1. 以管理员身份运行程序\n2. 检查文件属性，取消「只读」\n3. 检查文件夹权限设置",
            "is_retryable": False,
            "requires_manual_action": True,
        }

    # 文件被占用
    if any(
        kw in err_str
        for kw in ["being used", "locked", "占用", "正由另一程序", "sharing violation"]
    ):
        return {
            "error_type": ConversionErrorType.FILE_LOCKED,
            "error_category": "retryable",
            "message": "文件被其他程序占用",
            "suggestion": "1. 关闭 Word/Excel/WPS 等 Office 程序\n2. 关闭文件资源管理器中该文件的预览\n3. 等待几秒后重试",
            "is_retryable": True,
            "requires_manual_action": False,
        }

    # 文件不存在
    if any(
        kw in err_str
        for kw in ["not found", "does not exist", "找不到", "不存在", "filenotfound"]
    ):
        return {
            "error_type": ConversionErrorType.FILE_NOT_FOUND,
            "error_category": "unrecoverable",
            "message": "文件不存在",
            "suggestion": "文件可能已被移动或删除，请检查源目录",
            "is_retryable": False,
            "requires_manual_action": True,
        }

    # 文件损坏
    if any(
        kw in err_str
        for kw in [
            "corrupt",
            "damaged",
            "损坏",
            "repair",
            "无法读取",
            "unreadable",
            "invalid data",
        ]
    ):
        return {
            "error_type": ConversionErrorType.FILE_CORRUPTED,
            "error_category": "unrecoverable",
            "message": "文件可能已损坏",
            "suggestion": "1. 尝试用 Office 打开文件并另存为新文件\n2. 使用 Office 的「打开并修复」功能\n3. 从备份恢复文件",
            "is_retryable": False,
            "requires_manual_action": True,
        }

    # COM 错误（Office 相关）
    if any(
        kw in err_str
        for kw in [
            "com_error",
            "com object",
            "rpc",
            "server busy",
            "call was rejected",
            "0x800",
        ]
    ):
        return {
            "error_type": ConversionErrorType.COM_ERROR,
            "error_category": "retryable",
            "message": "Office 组件通信错误",
            "suggestion": "1. 重启 Office 程序\n2. 检查 Office 安装是否完整\n3. 尝试使用其他转换引擎（WPS/MS）",
            "is_retryable": True,
            "requires_manual_action": False,
        }

    # 超时
    if any(kw in err_str for kw in ["timeout", "超时", "timed out"]):
        return {
            "error_type": ConversionErrorType.TIMEOUT,
            "error_category": "retryable",
            "message": "转换超时",
            "suggestion": "1. 文件可能过大，尝试增加超时时间\n2. 关闭其他占用资源的程序\n3. 分批处理大文件",
            "is_retryable": True,
            "requires_manual_action": False,
        }

    # 磁盘空间不足
    if any(
        kw in err_str
        for kw in ["disk full", "no space", "磁盘已满", "空间不足", "storage"]
    ):
        return {
            "error_type": ConversionErrorType.DISK_FULL,
            "error_category": "needs_manual",
            "message": "磁盘空间不足",
            "suggestion": "1. 清理磁盘空间\n2. 更换输出目录到其他磁盘\n3. 在「高级设置」中调整沙盒路径",
            "is_retryable": False,
            "requires_manual_action": True,
        }

    # 密码保护
    if any(
        kw in err_str for kw in ["password", "protected", "密码", "encrypted", "加密"]
    ):
        return {
            "error_type": ConversionErrorType.PASSWORD_PROTECTED,
            "error_category": "needs_manual",
            "message": "文件受密码保护",
            "suggestion": "1. 先用 Office 打开文件并移除密码保护\n2. 将文件另存为无密码版本后再转换",
            "is_retryable": False,
            "requires_manual_action": True,
        }

    # 格式无效
    if any(
        kw in err_str
        for kw in ["invalid format", "format not supported", "格式无效", "格式错误"]
    ):
        return {
            "error_type": ConversionErrorType.INVALID_FORMAT,
            "error_category": "unrecoverable",
            "message": "文件格式无效",
            "suggestion": "1. 确认文件扩展名与实际内容匹配\n2. 尝试用对应 Office 程序打开并另存",
            "is_retryable": False,
            "requires_manual_action": True,
        }

    # 未知错误
    return {
        "error_type": ConversionErrorType.UNKNOWN,
        "error_category": "unknown",
        "message": f"未知错误: {str(exception)[:100] if exception else 'N/A'}",
        "suggestion": "1. 查看详细日志获取更多信息\n2. 尝试手动转换该文件\n3. 如问题持续，请联系开发者",
        "is_retryable": True,
        "requires_manual_action": True,
    }
