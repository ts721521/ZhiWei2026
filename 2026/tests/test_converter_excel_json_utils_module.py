import unittest
from datetime import date
from datetime import datetime
from datetime import time
from pathlib import Path

from office_converter import OfficeConverter


class ConverterExcelJsonUtilsSplitTests(unittest.TestCase):
    def test_excel_json_utils_core_behaviors(self):
        from converter.excel_json_utils import (
            build_column_profiles,
            col_index_to_label,
            detect_json_value_type,
            extract_formula_sheet_refs,
            is_effectively_empty_row,
            is_empty_json_cell,
            json_safe_value,
            looks_like_header_row,
            normalize_header_row,
        )

        self.assertEqual(json_safe_value(datetime(2026, 2, 24, 12, 30, 1)), "2026-02-24T12:30:01")
        self.assertEqual(json_safe_value(date(2026, 2, 24)), "2026-02-24")
        self.assertEqual(json_safe_value(time(9, 8, 7)), "09:08:07")

        self.assertTrue(is_empty_json_cell(None))
        self.assertTrue(is_empty_json_cell("   "))
        self.assertFalse(is_empty_json_cell("x"))
        self.assertTrue(is_effectively_empty_row(["", None, "   "]))
        self.assertFalse(is_effectively_empty_row(["", 1]))

        self.assertTrue(looks_like_header_row(["name", "price", 1]))
        self.assertFalse(looks_like_header_row([1, 2, 3]))
        self.assertEqual(normalize_header_row(["A", "A", "", None], 4), ["A", "A_2", "col_3", "col_4"])

        self.assertEqual(detect_json_value_type(True), "boolean")
        self.assertEqual(detect_json_value_type(1), "integer")
        self.assertEqual(detect_json_value_type(1.5), "number")
        self.assertEqual(detect_json_value_type("x"), "string")

        profiles = build_column_profiles(
            ["c1", "c2"],
            [["v1", 2], ["", None], ["v2", 3]],
            sample_limit=10,
        )
        self.assertEqual(len(profiles), 2)
        self.assertEqual(profiles[0]["name"], "c1")
        self.assertEqual(profiles[0]["non_null_count"], 2)
        self.assertEqual(profiles[1]["inferred_type"], "integer")

        self.assertEqual(col_index_to_label(1), "A")
        self.assertEqual(col_index_to_label(27), "AA")

        refs = extract_formula_sheet_refs("=SUM('Sheet A'!A1,Sheet1!B2,[Book.xlsx]Data!C3)", "Sheet1")
        self.assertEqual(refs, {"Sheet A", "Data"})

    def test_office_converter_excel_json_methods_delegate_to_module(self):
        from converter.excel_json_utils import (
            build_column_profiles,
            col_index_to_label,
            detect_json_value_type,
            extract_formula_sheet_refs,
            is_effectively_empty_row,
            is_empty_json_cell,
            json_safe_value,
            looks_like_header_row,
            normalize_header_row,
        )

        self.assertEqual(OfficeConverter._json_safe_value(1), json_safe_value(1))
        self.assertEqual(OfficeConverter._is_empty_json_cell(""), is_empty_json_cell(""))
        self.assertEqual(
            OfficeConverter._is_effectively_empty_row(["", None]),
            is_effectively_empty_row(["", None]),
        )
        self.assertEqual(
            OfficeConverter._looks_like_header_row(["A", "B", 1]),
            looks_like_header_row(["A", "B", 1]),
        )
        self.assertEqual(
            OfficeConverter._normalize_header_row(["A", "A"], 2),
            normalize_header_row(["A", "A"], 2),
        )
        self.assertEqual(OfficeConverter._detect_json_value_type(2.0), detect_json_value_type(2.0))
        self.assertEqual(
            OfficeConverter._build_column_profiles(["h"], [["v"]], 5),
            build_column_profiles(["h"], [["v"]], 5),
        )
        self.assertEqual(OfficeConverter._col_index_to_label(52), col_index_to_label(52))
        self.assertEqual(
            OfficeConverter._extract_formula_sheet_refs("=A!A1+'B B'!C1", "A"),
            extract_formula_sheet_refs("=A!A1+'B B'!C1", "A"),
        )

    def test_excel_json_utils_module_has_no_bare_except_exception(self):
        mod_path = Path(__file__).resolve().parents[1] / "converter" / "excel_json_utils.py"
        text = mod_path.read_text(encoding="utf-8")
        self.assertNotIn("except Exception", text)


if __name__ == "__main__":
    unittest.main()
