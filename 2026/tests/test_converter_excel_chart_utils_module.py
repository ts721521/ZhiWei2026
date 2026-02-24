import unittest

from office_converter import OfficeConverter


class ConverterExcelChartUtilsSplitTests(unittest.TestCase):
    def test_excel_chart_utils_core_behaviors(self):
        from converter.excel_chart_utils import extract_chart_title_text, stringify_chart_anchor

        class _Chart:
            def __init__(self, title):
                self.title = title

        class _Run:
            def __init__(self, t):
                self.t = t

        class _Para:
            def __init__(self, runs):
                self.r = runs

        class _Rich:
            def __init__(self, ps):
                self.p = ps

        class _Tx:
            def __init__(self, rich):
                self.rich = rich

        class _Title:
            def __init__(self, tx):
                self.tx = tx

        self.assertEqual(extract_chart_title_text(_Chart(None)), "")
        self.assertEqual(extract_chart_title_text(_Chart("  Sales  ")), "Sales")
        rich_title = _Title(_Tx(_Rich([_Para([_Run("Q1"), _Run(" Report")])])))
        self.assertEqual(extract_chart_title_text(_Chart(rich_title)), "Q1 Report")

        class _Marker:
            def __init__(self, col, row):
                self.col = col
                self.row = row

        class _Anchor:
            def __init__(self, marker):
                self._from = marker

        self.assertEqual(stringify_chart_anchor(_Anchor(_Marker(1, 2))), "B3")
        self.assertEqual(stringify_chart_anchor("A1"), "A1")
        self.assertEqual(stringify_chart_anchor(None), "")

    def test_office_converter_excel_chart_methods_delegate_to_module(self):
        from converter.excel_chart_utils import extract_chart_title_text, stringify_chart_anchor

        class _Chart:
            def __init__(self, title):
                self.title = title

        self.assertEqual(
            OfficeConverter._extract_chart_title_text(_Chart("x")),
            extract_chart_title_text(_Chart("x")),
        )
        self.assertEqual(
            OfficeConverter._stringify_chart_anchor("A1"),
            stringify_chart_anchor("A1"),
        )


if __name__ == "__main__":
    unittest.main()
