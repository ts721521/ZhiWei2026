# -*- coding: utf-8 -*-
"""Core per-file conversion logic extracted from office_converter.py."""


def convert_logic_in_thread(
    file_source,
    sandbox_target_pdf,
    ext,
    result_context,
    *,
    is_mac_fn,
    convert_on_mac_fn,
    has_win32,
    allowed_extensions,
    get_local_app_fn,
    safe_exec_fn,
    engine_type,
    engine_wps,
    wdFormatPDF,
    xlTypePDF,
    ppSaveAsPDF,
    ppFixedFormatTypePDF,
    xlPDF_SaveAs,
    xlRepairFile,
    content_strategy,
    strategy_standard,
    strategy_price_only,
    scan_excel_content_in_thread_fn,
    setup_excel_pages_fn,
    should_reuse_office_app_fn,
    pythoncom_module,
    os_module,
):
    app = None
    doc = None

    def _ppt_candidates(primary_doc, app_obj):
        candidates = []
        if primary_doc is not None:
            candidates.append(primary_doc)
        if app_obj is not None:
            accessors = [
                lambda: app_obj.ActivePresentation,
                lambda: app_obj.ActiveWindow.Presentation,
                lambda: app_obj.Presentations.Item(1),
            ]
            for getter in accessors:
                try:
                    cand = safe_exec_fn(getter)
                except (AttributeError, OSError, RuntimeError, TypeError, ValueError):
                    cand = None
                if cand is not None and cand not in candidates:
                    candidates.append(cand)
        return candidates

    def _try_ppt_methods(primary_doc, app_obj, calls):
        for candidate in _ppt_candidates(primary_doc, app_obj):
            for method_name, args, kwargs in calls:
                try:
                    method = getattr(candidate, method_name)
                except (AttributeError, OSError, RuntimeError, TypeError, ValueError):
                    continue
                try:
                    safe_exec_fn(method, *args, **kwargs)
                    return True
                except (AttributeError, OSError, RuntimeError, TypeError, ValueError):
                    continue
        return False
    try:
        if is_mac_fn():
            if convert_on_mac_fn(file_source, sandbox_target_pdf, ext):
                return
            if not has_win32:
                raise RuntimeError(
                    "macOS conversion failed (LibreOffice not found?) and win32com not available."
                )

        if ext in allowed_extensions.get("word", []):
            app = get_local_app_fn("word")
            try:
                if engine_type == engine_wps:
                    try:
                        doc = safe_exec_fn(app.Documents.Open, file_source, ReadOnly=True)
                    except (AttributeError, OSError, RuntimeError, TypeError, ValueError):
                        doc = safe_exec_fn(app.Documents.Open, file_source)
                    try:
                        safe_exec_fn(doc.ExportAsFixedFormat, sandbox_target_pdf, wdFormatPDF)
                    except (AttributeError, OSError, RuntimeError, TypeError, ValueError):
                        if os_module.path.exists(sandbox_target_pdf):
                            os_module.remove(sandbox_target_pdf)
                        try:
                            safe_exec_fn(doc.SaveAs, sandbox_target_pdf, FileFormat=wdFormatPDF)
                        except (AttributeError, OSError, RuntimeError, TypeError, ValueError):
                            try:
                                safe_exec_fn(app.ActiveDocument.ExportAsFixedFormat, sandbox_target_pdf, wdFormatPDF)
                            except (AttributeError, OSError, RuntimeError, TypeError, ValueError):
                                if os_module.path.exists(sandbox_target_pdf):
                                    os_module.remove(sandbox_target_pdf)
                                safe_exec_fn(app.ActiveDocument.SaveAs, sandbox_target_pdf, FileFormat=wdFormatPDF)
                else:
                    doc = safe_exec_fn(
                        app.Documents.Open,
                        file_source,
                        ReadOnly=True,
                        Visible=False,
                        OpenAndRepair=True,
                    )
                    try:
                        safe_exec_fn(doc.ExportAsFixedFormat, sandbox_target_pdf, wdFormatPDF)
                    except (AttributeError, OSError, RuntimeError, TypeError, ValueError):
                        if os_module.path.exists(sandbox_target_pdf):
                            os_module.remove(sandbox_target_pdf)
                        try:
                            safe_exec_fn(doc.SaveAs, sandbox_target_pdf, FileFormat=wdFormatPDF)
                        except (AttributeError, OSError, RuntimeError, TypeError, ValueError):
                            try:
                                safe_exec_fn(app.ActiveDocument.ExportAsFixedFormat, sandbox_target_pdf, wdFormatPDF)
                            except (AttributeError, OSError, RuntimeError, TypeError, ValueError):
                                if os_module.path.exists(sandbox_target_pdf):
                                    os_module.remove(sandbox_target_pdf)
                                safe_exec_fn(app.ActiveDocument.SaveAs, sandbox_target_pdf, FileFormat=wdFormatPDF)
            finally:
                if doc:
                    try:
                        doc.Close(SaveChanges=False)
                    except (AttributeError, OSError, RuntimeError, TypeError, ValueError):
                        pass

        elif ext in allowed_extensions.get("excel", []):
            app = get_local_app_fn("excel")
            try:
                if engine_type == engine_wps:
                    try:
                        doc = safe_exec_fn(app.Workbooks.Open, file_source, ReadOnly=True)
                    except (AttributeError, OSError, RuntimeError, TypeError, ValueError):
                        doc = safe_exec_fn(app.Workbooks.Open, file_source)
                    if (
                        not result_context.get("skip_scan", False)
                        and content_strategy != strategy_standard
                    ):
                        has_kw = scan_excel_content_in_thread_fn(doc)
                        if has_kw:
                            result_context["is_price"] = True
                        elif content_strategy == strategy_price_only:
                            result_context["scan_aborted"] = True
                            return
                    setup_excel_pages_fn(doc)
                    try:
                        safe_exec_fn(doc.ExportAsFixedFormat, xlTypePDF, sandbox_target_pdf)
                    except (AttributeError, OSError, RuntimeError, TypeError, ValueError):
                        if os_module.path.exists(sandbox_target_pdf):
                            os_module.remove(sandbox_target_pdf)
                        safe_exec_fn(doc.SaveAs, sandbox_target_pdf, FileFormat=xlPDF_SaveAs)
                else:
                    try:
                        doc = safe_exec_fn(
                            app.Workbooks.Open,
                            file_source,
                            UpdateLinks=0,
                            ReadOnly=True,
                            IgnoreReadOnlyRecommended=True,
                            CorruptLoad=xlRepairFile,
                        )
                    except (AttributeError, OSError, RuntimeError, TypeError, ValueError):
                        doc = safe_exec_fn(app.Workbooks.Open, file_source)
                    if doc and (
                        not result_context.get("skip_scan", False)
                        and content_strategy != strategy_standard
                    ):
                        has_kw = scan_excel_content_in_thread_fn(doc)
                        if has_kw:
                            result_context["is_price"] = True
                        elif content_strategy == strategy_price_only:
                            result_context["scan_aborted"] = True
                            return
                    if doc:
                        setup_excel_pages_fn(doc)
                        safe_exec_fn(doc.ExportAsFixedFormat, xlTypePDF, sandbox_target_pdf)
            finally:
                if doc:
                    try:
                        doc.Close(SaveChanges=False)
                    except (AttributeError, OSError, RuntimeError, TypeError, ValueError):
                        pass

        elif ext in allowed_extensions.get("powerpoint", []):
            app = get_local_app_fn("ppt")
            try:
                if engine_type == engine_wps:
                    try:
                        doc = safe_exec_fn(app.Presentations.Open, file_source, WithWindow=False)
                    except (AttributeError, OSError, RuntimeError, TypeError, ValueError):
                        doc = safe_exec_fn(app.Presentations.Open, file_source)
                    ok = _try_ppt_methods(
                        doc,
                        app,
                        [
                            ("SaveCopyAs", (sandbox_target_pdf, ppSaveAsPDF), {}),
                            ("SaveAs", (sandbox_target_pdf, ppSaveAsPDF), {}),
                        ],
                    )
                    if not ok:
                        raise RuntimeError("powerpoint open returned no PDF-export-capable document")
                else:
                    doc = safe_exec_fn(
                        app.Presentations.Open,
                        file_source,
                        WithWindow=False,
                        ReadOnly=True,
                    )
                    ok = _try_ppt_methods(
                        doc,
                        app,
                        [
                            ("ExportAsFixedFormat", (sandbox_target_pdf, ppFixedFormatTypePDF), {}),
                            ("SaveCopyAs", (sandbox_target_pdf, ppSaveAsPDF), {}),
                            ("SaveAs", (sandbox_target_pdf, ppSaveAsPDF), {}),
                        ],
                    )
                    if not ok:
                        if os_module.path.exists(sandbox_target_pdf):
                            os_module.remove(sandbox_target_pdf)
                        raise RuntimeError("powerpoint open returned no PDF-export-capable document")
            finally:
                if doc:
                    try:
                        doc.Close()
                    except (AttributeError, OSError, RuntimeError, TypeError, ValueError):
                        pass
    finally:
        if app:
            try:
                if not should_reuse_office_app_fn():
                    safe_exec_fn(app.Quit)
            except (AttributeError, OSError, RuntimeError, TypeError, ValueError):
                pass
        pythoncom_module.CoUninitialize()
