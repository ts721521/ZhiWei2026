import json
import os
import tempfile
import unittest
from pathlib import Path

from office_converter import OfficeConverter


class ConverterFailureTraceUtilsSplitTests(unittest.TestCase):
    def test_failure_trace_utils_core_behaviors(self):
        from converter.failure_trace_utils import (
            build_failed_file_trace_payload,
            write_failed_file_trace_log,
        )

        payload = build_failed_file_trace_payload(
            source_path=r"C:\x\a.docx",
            error_detail={
                "raw_error": "boom",
                "context": {"a": 1},
                "error_type": "permission",
                "error_category": "needs_manual",
                "is_retryable": False,
                "requires_manual_action": True,
                "message": "m",
                "suggestion": "s",
            },
            status="failed",
            elapsed=1.23,
            is_retry=False,
            failed_copy_path="",
            extra_context={"b": 2},
            get_failure_output_expectation_fn=lambda: {"pdf": True},
            get_readable_run_mode_fn=lambda: "convert_only",
            get_readable_engine_type_fn=lambda: "wps",
            infer_failure_stage_fn=lambda p, raw_error="", context=None: "convert",
        )
        self.assertEqual(payload["failure_stage"], "convert")
        self.assertEqual(payload["expected_outputs"], {"pdf": True})
        self.assertEqual(payload["context"], {"a": 1, "b": 2})
        self.assertEqual(payload["status"], "failed")

        root = tempfile.mkdtemp(prefix="failure_trace_")
        log_path = write_failed_file_trace_log(
            payload,
            failed_copy_path="",
            enable_trace_log=True,
            failed_dir=root,
            target_folder=root,
            sanitize_stem_fn=lambda s: s,
        )
        self.assertTrue(log_path and os.path.exists(log_path))
        with open(log_path, "r", encoding="utf-8") as f:
            loaded = json.load(f)
        self.assertEqual(loaded["status"], "failed")

        try:
            os.remove(log_path)
            os.rmdir(root)
        except Exception:
            pass

    def test_office_converter_failure_trace_methods_delegate_to_module(self):
        from converter.failure_trace_utils import build_failed_file_trace_payload

        dummy = OfficeConverter.__new__(OfficeConverter)
        dummy.config = {
            "enable_failed_file_trace_log": True,
            "target_folder": tempfile.mkdtemp(prefix="failure_trace_delegate_"),
        }
        dummy.failed_dir = dummy.config["target_folder"]
        dummy._get_failure_output_expectation = lambda: {"pdf": True}
        dummy.get_readable_run_mode = lambda: "convert_only"
        dummy.get_readable_engine_type = lambda: "wps"
        dummy._infer_failure_stage = lambda p, raw_error="", context=None: "convert"

        expected = build_failed_file_trace_payload(
            source_path=r"C:\x\a.docx",
            error_detail={"raw_error": "boom", "context": {}, "message": "m"},
            status="failed",
            elapsed=0.5,
            is_retry=False,
            failed_copy_path="",
            extra_context=None,
            get_failure_output_expectation_fn=dummy._get_failure_output_expectation,
            get_readable_run_mode_fn=dummy.get_readable_run_mode,
            get_readable_engine_type_fn=dummy.get_readable_engine_type,
            infer_failure_stage_fn=dummy._infer_failure_stage,
        )
        actual = dummy._build_failed_file_trace_payload(
            source_path=r"C:\x\a.docx",
            error_detail={"raw_error": "boom", "context": {}, "message": "m"},
            status="failed",
            elapsed=0.5,
            is_retry=False,
            failed_copy_path="",
            extra_context=None,
        )
        self.assertEqual(actual, expected)

        out = dummy._write_failed_file_trace_log(actual)
        self.assertTrue(out and os.path.exists(out))

        try:
            os.remove(out)
            os.rmdir(dummy.config["target_folder"])
        except Exception:
            pass

    def test_failure_trace_utils_module_has_no_bare_except_exception(self):
        mod_path = Path(__file__).resolve().parents[1] / "converter" / "failure_trace_utils.py"
        text = mod_path.read_text(encoding="utf-8")
        self.assertNotIn("except Exception", text)


if __name__ == "__main__":
    unittest.main()
