import os
import tempfile
import unittest
from datetime import datetime
from pathlib import Path

from office_converter import OfficeConverter


class ConverterIndexRuntimeSplitTests(unittest.TestCase):
    def test_index_runtime_core_behaviors(self):
        from converter.index_runtime import (
            append_conversion_index_record,
            write_merge_map_for_converter,
            write_conversion_index_workbook_for_converter,
            write_conversion_index_workbook,
            write_merge_map,
        )

        root = tempfile.mkdtemp(prefix="index_rt_")
        out_pdf = os.path.join(root, "merged.pdf")
        csv_path, json_path = write_merge_map(
            out_pdf,
            [{"merge_batch_id": "b1", "merged_pdf_name": "m"}],
            csv_module=__import__("csv"),
            json_module=__import__("json"),
        )
        self.assertTrue(os.path.exists(csv_path))
        self.assertTrue(os.path.exists(json_path))
        dummy_map = OfficeConverter.__new__(OfficeConverter)
        self.assertEqual(
            (csv_path, json_path),
            write_merge_map_for_converter(
                dummy_map,
                out_pdf,
                [{"merge_batch_id": "b1", "merged_pdf_name": "m"}],
                csv_module=__import__("csv"),
                json_module=__import__("json"),
            ),
        )

        recs = []
        append_conversion_index_record(
            "src/a.docx",
            "out/a.pdf",
            status="ok",
            exists_fn=lambda _p: True,
            abspath_fn=lambda p: f"/abs/{p}",
            basename_fn=os.path.basename,
            relpath_fn=lambda p, b: p.replace(f"/abs/{b}/", ""),
            compute_md5_fn=lambda _p: "md5",
            get_source_root_for_path_fn=lambda _p: "src",
            target_folder="out",
            conversion_index_records=recs,
        )
        self.assertEqual(len(recs), 1)
        self.assertEqual(recs[0]["source_filename"], "a.docx")
        self.assertEqual(recs[0]["status"], "ok")

        class _WB:
            def __init__(self):
                self.active = type("WS", (), {"title": ""})()
                self.saved = None

            def save(self, path):
                self.saved = path

        seen = {}
        idx = write_conversion_index_workbook(
            recs,
            config={"target_folder": root},
            has_openpyxl=True,
            workbook_cls=_WB,
            write_conversion_index_sheet_fn=lambda ws, rows: seen.setdefault(
                "sheet", (ws, rows)
            ),
            now_fn=lambda: datetime(2026, 2, 24, 21, 0, 0),
            join_path_fn=os.path.join,
            print_fn=lambda *_a, **_k: None,
            log_warning_fn=lambda _m: None,
            log_info_fn=lambda _m: None,
        )
        self.assertTrue(idx.endswith("Convert_Index_20260224_210000.xlsx"))
        self.assertIn("sheet", seen)

        dummy = OfficeConverter.__new__(OfficeConverter)
        dummy.conversion_index_records = recs
        dummy.config = {"target_folder": root}
        dummy._write_conversion_index_sheet = lambda ws, rows: seen.setdefault("sheet2", (ws, rows))
        idx2 = write_conversion_index_workbook_for_converter(
            dummy,
            has_openpyxl=True,
            workbook_cls=_WB,
            now_fn=lambda: datetime(2026, 2, 24, 21, 0, 0),
            join_path_fn=os.path.join,
            print_fn=lambda *_a, **_k: None,
            log_warning_fn=lambda _m: None,
            log_info_fn=lambda _m: None,
        )
        self.assertEqual(idx, idx2)
        self.assertEqual(dummy.convert_index_path, idx2)

    def test_office_converter_index_runtime_methods_delegate(self):
        import office_converter as oc

        original_merge = oc.write_merge_map_for_converter
        original_append = oc.append_conversion_index_record_impl
        original_wb = oc.write_conversion_index_workbook_for_converter
        try:
            seen = {}

            def _fake_merge(converter, output_path, records, **kwargs):
                seen["merge"] = (converter, output_path, records, kwargs)
                return ("a.csv", "a.json")

            def _fake_append(source_path, pdf_path, **kwargs):
                seen["append"] = (source_path, pdf_path, kwargs)
                return "ok"

            def _fake_wb(converter, **kwargs):
                converter.convert_index_path = "idx.xlsx"
                seen["wb"] = (converter, kwargs)
                return "idx.xlsx"

            oc.write_merge_map_for_converter = _fake_merge
            oc.append_conversion_index_record_impl = _fake_append
            oc.write_conversion_index_workbook_for_converter = _fake_wb

            dummy = OfficeConverter.__new__(OfficeConverter)
            dummy.config = {"target_folder": "out"}
            dummy.conversion_index_records = []
            dummy._compute_md5 = lambda _p: "x"
            dummy._get_source_root_for_path = lambda _p: "src"
            dummy._write_conversion_index_sheet = lambda *_a, **_k: None

            out1 = dummy._write_merge_map("x.pdf", [{"a": 1}])
            out2 = dummy._append_conversion_index_record("a", "b.pdf", status="s")
            out3 = dummy._write_conversion_index_workbook()
            self.assertEqual(out1, ("a.csv", "a.json"))
            self.assertEqual(out2, "ok")
            self.assertEqual(out3, "idx.xlsx")
            self.assertEqual(dummy.convert_index_path, "idx.xlsx")
        finally:
            oc.write_merge_map_for_converter = original_merge
            oc.append_conversion_index_record_impl = original_append
            oc.write_conversion_index_workbook_for_converter = original_wb

    def test_index_runtime_module_has_no_bare_except_exception(self):
        mod_path = Path(__file__).resolve().parents[1] / "converter" / "index_runtime.py"
        text = mod_path.read_text(encoding="utf-8")
        self.assertNotIn("except Exception", text)


if __name__ == "__main__":
    unittest.main()
