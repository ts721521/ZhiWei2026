import os
import tempfile
import unittest
from pathlib import Path

from office_converter import OfficeConverter


class ConverterTraceabilitySplitTests(unittest.TestCase):
    def test_traceability_prefix_helpers(self):
        from converter.traceability import (
            apply_short_id_prefix,
            normalize_short_id_for_match,
            strip_short_id_prefix,
        )

        self.assertEqual(apply_short_id_prefix("a1b2c3d4", "ZW-"), "ZW-A1B2C3D4")
        self.assertEqual(apply_short_id_prefix("ZW-A1B2C3D4", "ZW-"), "ZW-A1B2C3D4")
        self.assertEqual(strip_short_id_prefix("ZW-A1B2C3D4", "ZW-"), "A1B2C3D4")
        self.assertEqual(normalize_short_id_for_match("zw-a1b2c3d4"), "A1B2C3D4")

    def test_write_trace_map_xlsx_graceful_when_openpyxl_missing(self):
        from converter.traceability import write_trace_map_xlsx

        out = write_trace_map_xlsx(
            [{"source_short_id": "ZW-A1B2C3D4"}],
            output_path="trace_map.xlsx",
            has_openpyxl=False,
            workbook_cls=None,
            load_workbook_fn=None,
            font_cls=None,
            style_header_row_fn=None,
            auto_fit_sheet_fn=None,
            log_warning_fn=lambda _m: None,
        )
        self.assertIsNone(out)

    def test_write_trace_map_for_converter_core_behaviors(self):
        import converter.traceability as tmod

        original = tmod.write_trace_map_xlsx
        try:
            seen = {}

            def _fake(records, **kwargs):
                seen["records"] = records
                seen["kwargs"] = kwargs
                return kwargs["output_path"]

            tmod.write_trace_map_xlsx = _fake

            root = tempfile.mkdtemp(prefix="trace_map_core_")

            class Dummy:
                config = {
                    "target_folder": root,
                    "enable_traceability_anchor_and_map": True,
                    "short_id_prefix": "ZW-",
                }
                merge_index_records = [
                    {
                        "source_short_id": "aaaaaaaa",
                        "source_filename": "a.pdf",
                        "source_abspath": "C:/docs/a.pdf",
                        "source_relpath": "a.pdf",
                        "source_md5": "a" * 32,
                    }
                ]
                markdown_quality_records = [
                    {
                        "source_short_id": "BBBBBBBB",
                        "source_filename": "b.docx",
                        "source_abspath": "C:/docs/b.docx",
                        "source_md5": "b" * 32,
                    }
                ]
                trace_map_path = None
                generated_trace_map_outputs = []

            d = Dummy()
            out = tmod.write_trace_map_for_converter(
                d,
                has_openpyxl=False,
                log_warning_fn=lambda _m: None,
                log_info_fn=lambda _m: None,
            )
            self.assertTrue(out.endswith("trace_map.xlsx"))
            self.assertEqual(d.trace_map_path, out)
            self.assertEqual(d.generated_trace_map_outputs, [out])
            self.assertEqual(len(seen.get("records", [])), 2)
            self.assertEqual(seen["records"][0]["source_short_id"], "ZW-AAAAAAAA")
        finally:
            tmod.write_trace_map_xlsx = original

    def test_office_converter_write_trace_map_aggregates_and_delegates(self):
        import office_converter as oc

        original = oc.write_trace_map_for_converter_impl
        try:
            seen = {}

            def _fake(converter, **kwargs):
                seen["converter"] = converter
                seen["kwargs"] = kwargs
                return os.path.join(converter.config.get("target_folder", ""), "trace_map.xlsx")

            oc.write_trace_map_for_converter_impl = _fake
            root = tempfile.mkdtemp(prefix="trace_map_rt_")
            dummy = OfficeConverter.__new__(OfficeConverter)
            dummy.config = {
                "target_folder": root,
                "enable_traceability_anchor_and_map": True,
                "short_id_prefix": "ZW-",
            }
            dummy.merge_index_records = [
                {
                    "source_short_id": "AAAAAAAA",
                    "source_filename": "a.pdf",
                    "source_abspath": "C:/docs/a.pdf",
                    "source_relpath": "a.pdf",
                    "source_md5": "a" * 32,
                }
            ]
            dummy.markdown_quality_records = [
                {
                    "source_short_id": "ZW-BBBBBBBB",
                    "source_filename": "b.docx",
                    "source_abspath": "C:/docs/b.docx",
                    "source_md5": "b" * 32,
                }
            ]
            dummy._style_header_row = lambda _ws: None
            dummy._auto_fit_sheet = lambda _ws: None
            dummy.generated_trace_map_outputs = []
            dummy.trace_map_path = None

            out = dummy._write_trace_map()
            self.assertTrue(out.endswith("trace_map.xlsx"))
            self.assertIs(seen.get("converter"), dummy)
            self.assertTrue(seen["kwargs"]["has_openpyxl"] in (True, False))
            self.assertTrue(callable(seen["kwargs"]["style_header_row_fn"]))
        finally:
            oc.write_trace_map_for_converter_impl = original

    def test_traceability_module_has_no_bare_except_exception(self):
        mod_path = Path(__file__).resolve().parents[1] / "converter" / "traceability.py"
        text = mod_path.read_text(encoding="utf-8")
        self.assertNotIn("except Exception", text)


if __name__ == "__main__":
    unittest.main()
