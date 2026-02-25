import unittest
from datetime import datetime

from office_converter import OfficeConverter


class ConverterErrorRecordingModuleSplitTests(unittest.TestCase):
    def test_record_detailed_error_core_behaviors(self):
        from converter.error_recording import record_detailed_error

        records = []
        stats = {"permission_denied": 0}

        out = record_detailed_error(
            "src/a.docx",
            RuntimeError("boom"),
            context={"phase": "convert"},
            classify_conversion_error_fn=lambda _e, _p: {
                "error_type": "permission_denied",
                "error_category": "io",
                "message": "m",
                "suggestion": "s",
                "is_retryable": False,
                "requires_manual_action": True,
            },
            abspath_fn=lambda p: f"/abs/{p}",
            basename_fn=lambda p: p.split("/")[-1],
            now_fn=lambda: datetime(2026, 2, 24, 21, 0, 0),
            infer_failure_stage_fn=lambda *_a, **_k: "stage_x",
            get_failure_output_expectation_fn=lambda: ["pdf"],
            detailed_error_records=records,
            stats=stats,
        )

        self.assertEqual(out["source_path"], "/abs/src/a.docx")
        self.assertEqual(out["file_name"], "a.docx")
        self.assertEqual(out["failure_stage"], "stage_x")
        self.assertEqual(out["expected_outputs"], ["pdf"])
        self.assertEqual(stats["permission_denied"], 1)
        self.assertEqual(len(records), 1)

    def test_record_scan_access_skip_core_behaviors(self):
        from converter.error_recording import record_scan_access_skip

        stats = {"skipped": 0}
        seen = set()
        calls = {"trace": 0}

        out = record_scan_access_skip(
            "C:/x",
            RuntimeError("denied"),
            context={"elapsed": 1.2},
            seen_keys=seen,
            silent=True,
            is_win_fn=lambda: True,
            abspath_fn=lambda p: p,
            record_detailed_error_fn=lambda *_a, **_k: {
                "error_type": "permission_denied",
                "context": {"elapsed": 1.2},
            },
            stats=stats,
            build_failed_file_trace_payload_fn=lambda **kw: kw,
            write_failed_file_trace_log_fn=lambda *_a, **_k: calls.__setitem__(
                "trace", calls["trace"] + 1
            ),
            print_fn=lambda *_a, **_k: self.fail("should stay silent"),
            log_warning_fn=lambda _m: None,
        )
        self.assertEqual(out["error_type"], "permission_denied")
        self.assertIn("c:/x", seen)
        self.assertEqual(stats["skipped"], 1)
        self.assertEqual(calls["trace"], 1)

        out2 = record_scan_access_skip(
            "C:/X",
            RuntimeError("denied-again"),
            context={},
            seen_keys=seen,
            silent=False,
            is_win_fn=lambda: True,
            abspath_fn=lambda p: p,
            record_detailed_error_fn=lambda *_a, **_k: self.fail("duplicate should skip"),
            stats=stats,
            build_failed_file_trace_payload_fn=lambda **kw: kw,
            write_failed_file_trace_log_fn=lambda *_a, **_k: None,
            print_fn=lambda *_a, **_k: None,
            log_warning_fn=lambda _m: None,
        )
        self.assertIsNone(out2)
        self.assertEqual(stats["skipped"], 1)

    def test_office_converter_error_methods_delegate(self):
        import office_converter as oc

        original_record = oc.record_detailed_error_impl
        original_scan = oc.record_scan_access_skip_impl
        try:
            seen = {}

            def _fake_record(source_path, exception, **kwargs):
                seen["record"] = (source_path, str(exception), kwargs)
                return {"ok": 1}

            def _fake_scan(path, exception, **kwargs):
                seen["scan"] = (path, str(exception), kwargs)
                return {"ok": 2}

            oc.record_detailed_error_impl = _fake_record
            oc.record_scan_access_skip_impl = _fake_scan

            dummy = OfficeConverter.__new__(OfficeConverter)
            dummy.stats = {"skipped": 0}
            dummy.detailed_error_records = []
            dummy._infer_failure_stage = lambda *_a, **_k: "stage"
            dummy._get_failure_output_expectation = lambda: ["pdf"]
            dummy._build_failed_file_trace_payload = lambda **kw: kw
            dummy._write_failed_file_trace_log = lambda *_a, **_k: None

            out1 = dummy.record_detailed_error("a.docx", RuntimeError("e"))
            out2 = dummy._record_scan_access_skip("b", RuntimeError("e2"))
            self.assertEqual(out1, {"ok": 1})
            self.assertEqual(out2, {"ok": 2})
            self.assertEqual(seen["record"][0], "a.docx")
            self.assertEqual(seen["scan"][0], "b")
        finally:
            oc.record_detailed_error_impl = original_record
            oc.record_scan_access_skip_impl = original_scan


if __name__ == "__main__":
    unittest.main()
