import unittest
from pathlib import Path

from office_converter import OfficeConverter


class _FakePageSetup:
    def __init__(self):
        self.PrintArea = "A1:B2"
        self.Zoom = True
        self.Orientation = 1
        self.FitToPagesWide = 0
        self.FitToPagesTall = 0
        self.CenterHorizontally = False
        self.LeftMargin = 0
        self.RightMargin = 0
        self.TopMargin = 0
        self.BottomMargin = 0


class _FakeSheet:
    def __init__(self):
        self.UsedRange = object()
        self.PageSetup = _FakePageSetup()
        self.page_reset = 0

    def ResetAllPageBreaks(self):
        self.page_reset += 1


class _FakeWorkbook:
    def __init__(self):
        self.Worksheets = [_FakeSheet()]


class ConverterExcelPageSetupSplitTests(unittest.TestCase):
    def test_excel_page_setup_core_behaviors(self):
        from converter.excel_page_setup import setup_excel_pages

        wb = _FakeWorkbook()
        setup_excel_pages(wb)
        sheet = wb.Worksheets[0]
        self.assertEqual(sheet.page_reset, 1)
        self.assertEqual(sheet.PageSetup.PrintArea, "")
        self.assertFalse(sheet.PageSetup.Zoom)
        self.assertEqual(sheet.PageSetup.Orientation, 2)
        self.assertEqual(sheet.PageSetup.FitToPagesWide, 1)
        self.assertFalse(sheet.PageSetup.FitToPagesTall)
        self.assertTrue(sheet.PageSetup.CenterHorizontally)

    def test_office_converter_setup_excel_pages_delegates_to_module(self):
        import office_converter as oc

        original = oc.setup_excel_pages_impl
        try:
            seen = {}

            def _fake(workbook):
                seen["workbook"] = workbook

            oc.setup_excel_pages_impl = _fake
            dummy = OfficeConverter.__new__(OfficeConverter)
            wb = object()
            dummy._setup_excel_pages(wb)
            self.assertIs(seen["workbook"], wb)
        finally:
            oc.setup_excel_pages_impl = original

    def test_excel_page_setup_module_has_no_bare_except_exception(self):
        mod_path = Path(__file__).resolve().parents[1] / "converter" / "excel_page_setup.py"
        text = mod_path.read_text(encoding="utf-8")
        self.assertNotIn("except Exception", text)


if __name__ == "__main__":
    unittest.main()
