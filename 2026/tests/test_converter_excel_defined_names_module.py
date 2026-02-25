import unittest
from pathlib import Path

from office_converter import OfficeConverter


class _DN:
    def __init__(self, name, local=None, hidden=False, attr="", comment="", destinations=None):
        self.name = name
        self.localSheetId = local
        self.hidden = hidden
        self.attr_text = attr
        self.comment = comment
        self._destinations = destinations or []

    @property
    def destinations(self):
        return self._destinations


class _DNContainer:
    def __init__(self, defined_name_list, items_data):
        self.definedName = defined_name_list
        self._items_data = items_data

    def items(self):
        return self._items_data


class _WB:
    def __init__(self, dn_container):
        self.defined_names = dn_container


class ConverterExcelDefinedNamesSplitTests(unittest.TestCase):
    def test_excel_defined_names_core_behaviors(self):
        from converter.excel_defined_names import extract_workbook_defined_names

        dn1 = _DN("N1", local=0, attr="=A1", destinations=[("S1", "A1")])
        dn2 = _DN("N1", local=0, attr="=A1", destinations=[("S1", "A1")])  # duplicate
        dn3 = _DN("N2", local=1, hidden=True, attr="value")
        wb = _WB(_DNContainer([dn1], [("k1", dn2), ("k2", [dn3])]))

        out = extract_workbook_defined_names(wb)
        self.assertEqual(len(out), 2)
        self.assertEqual(out[0]["name"], "N1")
        self.assertTrue(out[0]["is_formula"])
        self.assertEqual(out[1]["name"], "N2")
        self.assertFalse(out[1]["is_formula"])

    def test_office_converter_extract_defined_names_delegates_to_module(self):
        import office_converter as oc

        original = oc.extract_workbook_defined_names_impl
        try:
            seen = {}

            def _fake(wb):
                seen["wb"] = wb
                return ["x"]

            oc.extract_workbook_defined_names_impl = _fake
            wb = object()
            out = OfficeConverter._extract_workbook_defined_names(wb)
            self.assertEqual(out, ["x"])
            self.assertIs(seen["wb"], wb)
        finally:
            oc.extract_workbook_defined_names_impl = original

    def test_excel_defined_names_module_has_no_bare_except_exception(self):
        mod_path = Path(__file__).resolve().parents[1] / "converter" / "excel_defined_names.py"
        text = mod_path.read_text(encoding="utf-8")
        self.assertNotIn("except Exception", text)


if __name__ == "__main__":
    unittest.main()
