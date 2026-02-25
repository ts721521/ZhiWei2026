import unittest

from office_converter import OfficeConverter


class _RefObj:
    def __init__(self, f):
        self.f = f


class _SeriesRef:
    def __init__(self, formula):
        self.numRef = _RefObj(formula)
        self.strRef = None


class _Series:
    def __init__(self, formula):
        self.val = _SeriesRef(formula)


class _Chart:
    def __init__(self, formulas):
        self.series = [_Series(f) for f in formulas]
        self.anchor = "A1"


class _Sheet:
    def __init__(self, charts, pivots=None):
        self._charts = charts
        self._pivots = pivots or []


class _PivotLocation:
    def __init__(self, ref):
        self.ref = ref


class _Pivot:
    def __init__(self, name, ref, cache_id):
        self.name = name
        self.location = _PivotLocation(ref)
        self.cacheId = cache_id


class ConverterExcelChartExtractSplitTests(unittest.TestCase):
    def test_excel_chart_extract_core_behaviors(self):
        from converter.excel_chart_extract import extract_sheet_charts

        ws = _Sheet([_Chart(["Sheet1!A1:A10", "Sheet1!A1:A10", "Sheet1!B1:B10"])])
        rows = extract_sheet_charts(
            ws,
            extract_chart_title_text_fn=lambda _c: "T",
            stringify_chart_anchor_fn=lambda a: str(a),
            series_ref_limit=2,
        )
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["title"], "T")
        self.assertEqual(rows[0]["anchor"], "A1")
        self.assertEqual(rows[0]["series_ref_count"], 2)
        self.assertFalse(rows[0]["series_refs_truncated"])

    def test_office_converter_extract_sheet_charts_delegates_to_module(self):
        import office_converter as oc

        original = oc.extract_sheet_charts_impl
        try:
            seen = {}

            def _fake(ws_formula, **kwargs):
                seen["ws_formula"] = ws_formula
                seen["kwargs"] = kwargs
                return ["ok"]

            oc.extract_sheet_charts_impl = _fake
            dummy = OfficeConverter.__new__(OfficeConverter)
            dummy._extract_chart_title_text = lambda _c: "x"
            dummy._stringify_chart_anchor = lambda _a: "A1"
            sheet = object()
            out = dummy._extract_sheet_charts(sheet, series_ref_limit=9)
            self.assertEqual(out, ["ok"])
            self.assertIs(seen["ws_formula"], sheet)
            self.assertEqual(seen["kwargs"]["series_ref_limit"], 9)
        finally:
            oc.extract_sheet_charts_impl = original

    def test_extract_sheet_pivot_tables_core_and_delegate(self):
        from converter.excel_chart_extract import extract_sheet_pivot_tables

        ws = _Sheet([], pivots=[_Pivot("P1", "A1:C8", 7)])
        rows = extract_sheet_pivot_tables(ws)
        self.assertEqual(1, len(rows))
        self.assertEqual("P1", rows[0]["name"])
        self.assertEqual("A1:C8", rows[0]["location_ref"])

        import office_converter as oc

        original = oc.extract_sheet_pivot_tables_impl
        try:
            oc.extract_sheet_pivot_tables_impl = lambda _ws: ["pivot_ok"]
            self.assertEqual(["pivot_ok"], OfficeConverter._extract_sheet_pivot_tables(object()))
        finally:
            oc.extract_sheet_pivot_tables_impl = original


if __name__ == "__main__":
    unittest.main()
