import json
import os
import tempfile
import unittest

from office_converter import OfficeConverter


class ConverterExcelJsonExportSplitTests(unittest.TestCase):
    def test_excel_json_export_core_behaviors_without_openpyxl(self):
        from converter.excel_json_export import export_single_excel_json

        root = tempfile.mkdtemp(prefix="excel_json_export_")
        source_xlsx = os.path.join(root, "book.xlsx")
        source_xls = os.path.join(root, "book.xls")
        out_xlsx = os.path.join(root, "book.xlsx.json")
        out_xls = os.path.join(root, "book.xls.json")

        with open(source_xlsx, "w", encoding="utf-8") as f:
            f.write("x")
        with open(source_xls, "w", encoding="utf-8") as f:
            f.write("x")

        path1 = export_single_excel_json(
            source_xlsx,
            config={},
            build_ai_output_path_from_source_fn=lambda *_args: out_xlsx,
            source_root_resolver=lambda _p: root,
            has_openpyxl=False,
            load_workbook_fn=None,
            extract_workbook_defined_names_fn=lambda wb: [],
            extract_sheet_charts_fn=lambda ws: [],
            extract_sheet_pivot_tables_fn=lambda ws: [],
        )
        self.assertEqual(path1, out_xlsx)
        with open(out_xlsx, "r", encoding="utf-8") as f:
            payload1 = json.load(f)
        self.assertEqual(payload1.get("parse_status"), "openpyxl_missing")

        path2 = export_single_excel_json(
            source_xls,
            config={},
            build_ai_output_path_from_source_fn=lambda *_args: out_xls,
            source_root_resolver=lambda _p: root,
            has_openpyxl=False,
            load_workbook_fn=None,
            extract_workbook_defined_names_fn=lambda wb: [],
            extract_sheet_charts_fn=lambda ws: [],
            extract_sheet_pivot_tables_fn=lambda ws: [],
        )
        self.assertEqual(path2, out_xls)
        with open(out_xls, "r", encoding="utf-8") as f:
            payload2 = json.load(f)
        self.assertEqual(payload2.get("parse_status"), "unsupported_format_xls")

        for p in [out_xlsx, out_xls, source_xlsx, source_xls]:
            try:
                os.remove(p)
            except Exception:
                pass
        try:
            os.rmdir(root)
        except Exception:
            pass

    def test_office_converter_export_single_excel_json_delegates_to_module(self):
        import office_converter as oc

        original = oc.export_single_excel_json_impl
        try:
            seen = {}

            def _fake(source_excel_path, **kwargs):
                seen["source_excel_path"] = source_excel_path
                seen["kwargs"] = kwargs
                return "ok"

            oc.export_single_excel_json_impl = _fake

            dummy = OfficeConverter.__new__(OfficeConverter)
            dummy.config = {"a": 1}
            dummy._build_ai_output_path_from_source = lambda *_args: "out.json"
            dummy._get_source_root_for_path = lambda _p: "S:/"
            dummy._extract_workbook_defined_names = lambda _wb: []
            dummy._extract_sheet_charts = lambda _ws: []
            dummy._extract_sheet_pivot_tables = lambda _ws: []

            out = dummy._export_single_excel_json("S:/a.xlsx")
            self.assertEqual(out, "ok")
            self.assertEqual(seen.get("source_excel_path"), "S:/a.xlsx")
            self.assertEqual(seen.get("kwargs", {}).get("config"), dummy.config)
        finally:
            oc.export_single_excel_json_impl = original


if __name__ == "__main__":
    unittest.main()
