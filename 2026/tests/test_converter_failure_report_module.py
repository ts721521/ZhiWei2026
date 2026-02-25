import json
import os
import tempfile
import unittest
from datetime import datetime
from pathlib import Path

from office_converter import OfficeConverter


class ConverterFailureReportSplitTests(unittest.TestCase):
    def test_failure_report_core_behaviors(self):
        from converter.failure_report import (
            export_failed_files_report,
            export_failed_files_report_for_converter,
        )

        result_empty = export_failed_files_report(
            [],
            tempfile.mkdtemp(prefix="failure_report_empty_"),
            run_mode="convert_only",
        )
        self.assertEqual(result_empty["summary"], "no_failed_records")
        self.assertIsNone(result_empty["json_path"])
        self.assertIsNone(result_empty["txt_path"])

        root = tempfile.mkdtemp(prefix="failure_report_")
        records = [
            {
                "file_name": "a.docx",
                "source_path": r"C:\x\a.docx",
                "error_type": "permission",
                "error_category": "needs_manual",
                "message": "m1",
                "suggestion": "s1",
                "is_retryable": False,
                "requires_manual_action": True,
            },
            {
                "file_name": "b.docx",
                "source_path": r"C:\x\b.docx",
                "error_type": "timeout",
                "error_category": "retryable",
                "message": "m2",
                "suggestion": "s2",
                "is_retryable": True,
                "requires_manual_action": False,
            },
        ]
        fixed_now = datetime(2026, 2, 24, 10, 11, 12)
        result = export_failed_files_report(
            records,
            root,
            run_mode="convert_only",
            now_fn=lambda: fixed_now,
        )
        self.assertTrue(result["json_path"].endswith("failed_files_report_20260224_101112.json"))
        self.assertTrue(result["txt_path"].endswith("failed_files_report_20260224_101112.txt"))
        self.assertIn("total_failed=2", result["summary"])
        self.assertTrue(os.path.exists(result["json_path"]))
        self.assertTrue(os.path.exists(result["txt_path"]))

        with open(result["json_path"], "r", encoding="utf-8") as f:
            payload = json.load(f)
        self.assertEqual(payload["statistics"]["retryable_count"], 1)
        self.assertEqual(payload["statistics"]["manual_action_count"], 1)

        dummy = OfficeConverter.__new__(OfficeConverter)
        dummy.config = {"target_folder": root}
        dummy.detailed_error_records = records
        dummy.get_readable_run_mode = lambda: "convert_only"
        dummy.failed_report_path = None
        result2 = export_failed_files_report_for_converter(
            dummy,
            output_dir=root,
            now_fn=lambda: fixed_now,
            log_error=lambda *_a, **_k: None,
        )
        self.assertEqual(result["txt_path"], result2["txt_path"])
        self.assertEqual(dummy.failed_report_path, result2["txt_path"])

        for p in (result["json_path"], result["txt_path"]):
            try:
                os.remove(p)
            except Exception:
                pass
        try:
            os.rmdir(root)
        except Exception:
            pass

    def test_office_converter_export_failed_files_report_delegates_to_module(self):
        import office_converter as oc

        orig = oc.export_failed_files_report_for_converter
        try:
            oc.export_failed_files_report_for_converter = (
                lambda converter, output_dir=None, **kwargs: {
                    "json_path": os.path.join(output_dir or converter.config["target_folder"], "a.json"),
                    "txt_path": os.path.join(output_dir or converter.config["target_folder"], "a.txt"),
                    "summary": "ok",
                }
            )
            root = tempfile.mkdtemp(prefix="failure_report_delegate_")
            dummy = OfficeConverter.__new__(OfficeConverter)
            dummy.config = {"target_folder": root}
            dummy.detailed_error_records = [{"error_type": "x"}]
            dummy.get_readable_run_mode = lambda: "convert_only"
            dummy.failed_report_path = None

            result = dummy.export_failed_files_report()
            self.assertEqual(result["summary"], "ok")
            os.rmdir(root)
        finally:
            oc.export_failed_files_report_for_converter = orig

    def test_failure_report_module_has_no_bare_except_exception(self):
        mod_path = Path(__file__).resolve().parents[1] / "converter" / "failure_report.py"
        text = mod_path.read_text(encoding="utf-8")
        self.assertNotIn("except Exception", text)


if __name__ == "__main__":
    unittest.main()
