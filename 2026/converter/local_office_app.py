# -*- coding: utf-8 -*-
"""Local Office COM app bootstrap extracted from office_converter.py."""


def get_local_app(
    *,
    app_type,
    engine_type,
    has_win32,
    engine_wps,
    engine_ms,
    pythoncom_module,
    win32_client,
):
    retryable_errors = [AttributeError, OSError, RuntimeError, TypeError, ValueError]
    com_error_cls = getattr(pythoncom_module, "com_error", None)
    if isinstance(com_error_cls, type) and issubclass(com_error_cls, BaseException):
        retryable_errors.append(com_error_cls)
    retryable_errors = tuple(retryable_errors)

    if not has_win32:
        raise RuntimeError(
            "Current system does not support Windows COM; Office conversion is unavailable."
        )

    pythoncom_module.CoInitialize()
    if engine_type == engine_wps:
        prog_id = {
            "word": "Kwps.Application",
            "excel": "Ket.Application",
            "ppt": "Kwpp.Application",
        }.get(app_type)
    else:
        prog_id = {
            "word": "Word.Application",
            "excel": "Excel.Application",
            "ppt": "PowerPoint.Application",
        }.get(app_type)
    app = None
    try:
        app = win32_client.Dispatch(prog_id)
    except retryable_errors:
        app = win32_client.DispatchEx(prog_id)

    try:
        app.Visible = False
        if app_type != "ppt":
            app.DisplayAlerts = False
    except retryable_errors:
        # Some Office/WPS COM servers reject visibility toggles; keep the app
        # instance alive and proceed with conversion calls.
        pass

    if engine_type == engine_ms and app_type == "excel":
        try:
            app.AskToUpdateLinks = False
        except (AttributeError, OSError, RuntimeError, TypeError, ValueError):
            pass

    return app
