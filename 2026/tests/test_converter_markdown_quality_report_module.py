import json
import os
import tempfile
import unittest
from datetime import datetime

from office_converter import OfficeConverter


class ConverterMarkdownQualityReportSplitTests(unittest.TestCase):
    def test_markdown_quality_report_core_behaviors(self):
        from converter.markdown_quality_report import (
            write_markdown_quality_report,
            write_markdown_quality_report_for_converter,
        )

        root = tempfile.mkdtemp(prefix="md_quality_")
        records = [
            {
                "source_pdf": "a.pdf",
                "markdown_path": "a.md",
                "page_count": 3,
                "non_empty_page_count": 2,
                "removed_header_lines": 1,
                "removed_footer_lines": 2,
                "removed_page_number_lines": 1,
                "heading_count": 4,
            },
            {
                "source_pdf": "b.pdf",
                "markdown_path": "b.md",
                "page_count": 5,
                "non_empty_page_count": 5,
                "removed_header_lines": 0,
                "removed_footer_lines": 1,
                "removed_page_number_lines": 2,
                "heading_count": 6,
            },
        ]
        fixed_now = datetime(2026, 2, 24, 15, 30, 45)
        out = write_markdown_quality_report(
            config={
                "target_folder": root,
                "enable_markdown_quality_report": True,
                "markdown_quality_sample_limit": 1,
            },
            markdown_quality_records=records,
            now_fn=lambda: fixed_now,
        )

        self.assertTrue(out)
        self.assertTrue(os.path.exists(out))
        with open(out, "r", encoding="utf-8") as f:
            payload = json.load(f)
        self.assertEqual(payload.get("record_count"), 2)
        self.assertEqual(payload.get("sample_limit"), 1)
        self.assertEqual(len(payload.get("samples", [])), 1)
        self.assertEqual(payload.get("summary", {}).get("total_pages"), 8)
        self.assertEqual(payload.get("summary", {}).get("non_empty_pages"), 7)

        dummy = OfficeConverter.__new__(OfficeConverter)
        dummy.config = {
            "target_folder": root,
            "enable_markdown_quality_report": True,
            "markdown_quality_sample_limit": 1,
        }
        dummy.markdown_quality_records = records
        dummy.generated_markdown_quality_outputs = []
        dummy.markdown_quality_report_path = None
        out2 = write_markdown_quality_report_for_converter(
            dummy,
            now_fn=lambda: fixed_now,
            log_info_fn=lambda *_a, **_k: None,
        )
        self.assertEqual(out, out2)
        self.assertEqual(dummy.markdown_quality_report_path, out2)
        self.assertEqual(dummy.generated_markdown_quality_outputs, [out2])

    def test_office_converter_markdown_quality_report_delegates_to_module(self):
        import office_converter as oc

        original = oc.write_markdown_quality_report_for_converter
        try:
            seen = {}

            def _fake(converter, **kwargs):
                converter.markdown_quality_report_path = "out.json"
                converter.generated_markdown_quality_outputs = ["out.json"]
                seen["kwargs"] = kwargs
                return "out.json"

            oc.write_markdown_quality_report_for_converter = _fake
            dummy = OfficeConverter.__new__(OfficeConverter)
            dummy.config = {"target_folder": "x"}
            dummy.markdown_quality_records = [{"k": "v"}]
            dummy.generated_markdown_quality_outputs = []
            dummy.markdown_quality_report_path = None

            out = dummy._write_markdown_quality_report()
            self.assertEqual(out, "out.json")
            self.assertEqual(dummy.markdown_quality_report_path, "out.json")
            self.assertEqual(dummy.generated_markdown_quality_outputs, ["out.json"])
            self.assertIn("now_fn", seen.get("kwargs", {}))
            self.assertIn("log_info_fn", seen.get("kwargs", {}))
        finally:
            oc.write_markdown_quality_report_for_converter = original


if __name__ == "__main__":
    unittest.main()
