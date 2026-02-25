import json
import os
import tempfile
import unittest
from datetime import datetime

from office_converter import OfficeConverter


class ConverterRecordsJsonExportSplitTests(unittest.TestCase):
    def test_records_json_export_core_behaviors(self):
        from converter.records_json_export import (
            write_records_json_exports,
            write_records_json_exports_for_converter,
        )

        root = tempfile.mkdtemp(prefix="records_json_")
        outputs = write_records_json_exports(
            config={"enable_excel_json": True, "target_folder": root},
            conversion_index_records=[{"a": 1}],
            merge_index_records=[{"b": 2}],
            now_fn=lambda: datetime(2026, 2, 24, 20, 0, 0),
            log_info=None,
        )
        self.assertEqual(len(outputs), 2)
        for p in outputs:
            self.assertTrue(os.path.exists(p))
            with open(p, "r", encoding="utf-8") as f:
                payload = json.load(f)
            self.assertIn(payload.get("record_type"), ("convert_index", "merge_index"))

        dummy = OfficeConverter.__new__(OfficeConverter)
        dummy.config = {"enable_excel_json": True, "target_folder": root}
        dummy.conversion_index_records = [{"a": 1}]
        dummy.merge_index_records = [{"b": 2}]
        dummy.generated_records_json_outputs = []
        out2 = write_records_json_exports_for_converter(
            dummy,
            now_fn=lambda: datetime(2026, 2, 24, 20, 0, 0),
            log_info=None,
        )
        self.assertEqual(outputs, out2)
        self.assertEqual(outputs, dummy.generated_records_json_outputs)

    def test_office_converter_write_records_json_exports_delegates_to_module(self):
        import office_converter as oc

        original = oc.write_records_json_exports_for_converter
        try:
            seen = {}

            def _fake(converter, **kwargs):
                converter.generated_records_json_outputs = ["a.json"]
                seen["kwargs"] = kwargs
                return ["a.json"]

            oc.write_records_json_exports_for_converter = _fake
            dummy = OfficeConverter.__new__(OfficeConverter)
            dummy.config = {}
            dummy.conversion_index_records = []
            dummy.merge_index_records = []
            dummy.generated_records_json_outputs = []

            out = dummy._write_records_json_exports()
            self.assertEqual(out, ["a.json"])
            self.assertEqual(dummy.generated_records_json_outputs, ["a.json"])
            self.assertIn("now_fn", seen["kwargs"])
            self.assertIn("log_info", seen["kwargs"])
        finally:
            oc.write_records_json_exports_for_converter = original


if __name__ == "__main__":
    unittest.main()
