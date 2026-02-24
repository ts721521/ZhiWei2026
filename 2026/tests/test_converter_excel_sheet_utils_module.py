import unittest

from office_converter import OfficeConverter


class _FakeDim:
    def __init__(self):
        self.width = None


class _FakeCell:
    def __init__(self, value, col_letter):
        self.value = value
        self.column_letter = col_letter
        self.font = None


class _FakeWS:
    def __init__(self):
        self._row1 = [_FakeCell("h1", "A"), _FakeCell("h2", "B")]
        self.columns = [
            [_FakeCell("abc", "A"), _FakeCell("de", "A")],
            [_FakeCell("x", "B"), _FakeCell("12345", "B")],
        ]
        self.column_dimensions = {"A": _FakeDim(), "B": _FakeDim()}

    def __getitem__(self, key):
        if key == 1:
            return self._row1
        raise KeyError(key)


class _DummyFont:
    def __init__(self, bold=False):
        self.bold = bold


class ConverterExcelSheetUtilsSplitTests(unittest.TestCase):
    def test_excel_sheet_utils_core_behaviors(self):
        from converter.excel_sheet_utils import auto_fit_sheet, style_header_row

        ws = _FakeWS()
        style_header_row(ws, has_openpyxl=True, font_cls=_DummyFont)
        self.assertTrue(all(getattr(c.font, "bold", False) for c in ws[1]))

        ws2 = _FakeWS()
        style_header_row(ws2, has_openpyxl=False, font_cls=_DummyFont)
        self.assertTrue(all(c.font is None for c in ws2[1]))

        auto_fit_sheet(ws, max_width=6)
        self.assertEqual(ws.column_dimensions["A"].width, 5)  # len("abc")+2
        self.assertEqual(ws.column_dimensions["B"].width, 6)  # capped by max_width

    def test_office_converter_excel_sheet_methods_still_work(self):
        ws = _FakeWS()
        OfficeConverter._style_header_row(ws)
        self.assertEqual(len(ws[1]), 2)

        OfficeConverter._auto_fit_sheet(ws, max_width=6)
        self.assertEqual(ws.column_dimensions["A"].width, 5)
        self.assertEqual(ws.column_dimensions["B"].width, 6)


if __name__ == "__main__":
    unittest.main()
