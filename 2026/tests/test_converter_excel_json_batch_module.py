import os
import tempfile
import unittest
from pathlib import Path

from office_converter import OfficeConverter


class ConverterExcelJsonBatchSplitTests(unittest.TestCase):
    def test_excel_json_batch_module_has_no_bare_except_exception(self):
        module_text = Path("converter/excel_json_batch.py").read_text(encoding="utf-8")
        self.assertNotIn("except Exception", module_text)

    def test_excel_json_batch_core_behaviors(self):
        from converter.excel_json_batch import write_excel_structured_json_exports

        root = tempfile.mkdtemp(prefix="excel_json_batch_")
        xlsx = os.path.join(root, "a.xlsx")
        txt = os.path.join(root, "a.txt")
        with open(xlsx, "w", encoding="utf-8") as f:
            f.write("x")
        with open(txt, "w", encoding="utf-8") as f:
            f.write("x")

        class Dummy:
            def __init__(self):
                self.config = {
                    "enable_excel_json": True,
                    "allowed_extensions": {"excel": [".xls", ".xlsx"]},
                }
                self.conversion_index_records = [
                    {"source_abspath": xlsx},
                    {"source_abspath": txt},
                    {"source_abspath": xlsx},
                ]
                self.generated_excel_json_outputs = []
                self.calls = []

            def _export_single_excel_json(self, source_excel_path):
                self.calls.append(source_excel_path)
                out = source_excel_path + ".json"
                with open(out, "w", encoding="utf-8") as f:
                    f.write("{}")
                return out

        d = Dummy()
        out = write_excel_structured_json_exports(d)
        self.assertEqual(d.calls, [os.path.abspath(xlsx)])
        self.assertEqual(len(out), 1)
        self.assertEqual(d.generated_excel_json_outputs, out)

    def test_office_converter_write_excel_json_batch_delegates_to_module(self):
        import office_converter as oc

        original = oc.write_excel_structured_json_exports_impl
        try:
            seen = {}

            def _fake(converter, **kwargs):
                seen["converter"] = converter
                seen["kwargs"] = kwargs
                return ["x.json"]

            oc.write_excel_structured_json_exports_impl = _fake
            dummy = OfficeConverter.__new__(OfficeConverter)
            out = dummy._write_excel_structured_json_exports()
            self.assertEqual(out, ["x.json"])
            self.assertIs(seen["converter"], dummy)
            self.assertTrue(callable(seen["kwargs"]["log_info_fn"]))
        finally:
            oc.write_excel_structured_json_exports_impl = original


if __name__ == "__main__":
    unittest.main()
