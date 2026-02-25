import unittest
from pathlib import Path

from office_converter import OfficeConverter


class _FakeUsedRange:
    def __init__(self, value):
        self.Value = value


class _FakeSheet:
    def __init__(self, name, value):
        self.Name = name
        self.UsedRange = _FakeUsedRange(value)


class _FakeWorkbook:
    def __init__(self, sheets):
        self.Worksheets = sheets


class ConverterExcelContentScanSplitTests(unittest.TestCase):
    def test_excel_content_scan_module_has_no_bare_except_exception(self):
        module_text = Path("converter/excel_content_scan.py").read_text(encoding="utf-8")
        self.assertNotIn("except Exception", module_text)

    def test_excel_content_scan_core_behaviors(self):
        from converter.excel_content_scan import scan_excel_content_in_thread

        wb = _FakeWorkbook([_FakeSheet("S1", (("foo",), ("price here",)))])
        hit = scan_excel_content_in_thread(
            wb,
            price_keywords=["price", "quote"],
            log_info_fn=lambda _m: None,
            log_warning_fn=lambda _m: None,
        )
        self.assertTrue(hit)

    def test_office_converter_scan_excel_content_delegates_to_module(self):
        import office_converter as oc

        original = oc.scan_excel_content_in_thread_impl
        try:
            seen = {}

            def _fake(workbook, **kwargs):
                seen["workbook"] = workbook
                seen["kwargs"] = kwargs
                return True

            oc.scan_excel_content_in_thread_impl = _fake
            dummy = OfficeConverter.__new__(OfficeConverter)
            dummy.price_keywords = ["k"]
            out = dummy.scan_excel_content_in_thread(object())
            self.assertTrue(out)
            self.assertEqual(seen["kwargs"]["price_keywords"], ["k"])
        finally:
            oc.scan_excel_content_in_thread_impl = original


if __name__ == "__main__":
    unittest.main()
