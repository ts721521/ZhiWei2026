# -*- coding: utf-8 -*-
"""Excel page setup helpers extracted from office_converter.py."""


def setup_excel_pages(workbook):
    try:
        for sheet in workbook.Worksheets:
            try:
                _ = sheet.UsedRange
                try:
                    sheet.ResetAllPageBreaks()
                except (AttributeError, RuntimeError, TypeError, ValueError):
                    pass
                page_setup = sheet.PageSetup
                try:
                    page_setup.PrintArea = ""
                except (AttributeError, RuntimeError, TypeError, ValueError):
                    pass
                page_setup.Zoom = False
                page_setup.Orientation = 2
                page_setup.FitToPagesWide = 1
                page_setup.FitToPagesTall = False
                page_setup.CenterHorizontally = True
                try:
                    page_setup.LeftMargin = 20
                    page_setup.RightMargin = 20
                    page_setup.TopMargin = 20
                    page_setup.BottomMargin = 20
                except (AttributeError, RuntimeError, TypeError, ValueError):
                    pass
            except (AttributeError, RuntimeError, TypeError, ValueError):
                pass
    except (AttributeError, RuntimeError, TypeError, ValueError):
        pass
