import os
import tempfile
import unittest

from office_converter import OfficeConverter


class ConverterBatchSequentialSplitTests(unittest.TestCase):
    def test_batch_sequential_core_behaviors(self):
        from converter.batch_sequential import run_batch

        root = tempfile.mkdtemp(prefix="batch_seq_")
        src = os.path.join(root, "a.docx")
        out_pdf = os.path.join(root, "a.pdf")
        with open(src, "w", encoding="utf-8") as f:
            f.write("x")
        with open(out_pdf, "w", encoding="utf-8") as f:
            f.write("y")

        class Dummy:
            def __init__(self):
                self.config = {
                    "enable_checkpoint": True,
                    "parallel_checkpoint_interval": 1,
                }
                self.is_running = True
                self.stats = {"success": 0, "failed": 0, "timeout": 0}
                self.error_records = []
                self.progress_points = []
                self.progress_callback = lambda i, t: self.progress_points.append((i, t))
                self.done_records = []
                self.appended = []
                self.cleared = False
                self.saved = []

            def _init_checkpoint(self, files):
                return {"status": "running"}, list(files)

            def _save_checkpoint(self, checkpoint):
                self.saved.append(dict(checkpoint))

            def _mark_file_done_in_checkpoint(self, checkpoint, _fpath):
                checkpoint["status"] = "completed"
                return checkpoint

            def _clear_checkpoint(self):
                self.cleared = True

            def get_target_path(self, logical_source, _ext):
                return logical_source + ".pdf"

            def get_progress_prefix(self, idx, total):
                return f"[{idx}/{total}]"

            def process_single_file(self, *_args):
                return "success", out_pdf

            def _append_conversion_index_record(self, source, final_path, status):
                self.appended.append((source, final_path, status))

            def _emit_file_done(self, record):
                self.done_records.append(record)

            def record_detailed_error(self, *_args, **_kwargs):
                return {
                    "error_type": "unknown",
                    "error_category": "other",
                    "suggestion": "",
                    "is_retryable": False,
                    "requires_manual_action": False,
                    "failure_stage": "",
                }

            def get_readable_run_mode(self):
                return "x"

            def get_readable_engine_type(self):
                return "y"

            def quarantine_failed_file(self, _fpath):
                return ""

            def _build_failed_file_trace_payload(self, **kwargs):
                return kwargs

            def _write_failed_file_trace_log(self, _payload, failed_copy_path=None):
                return failed_copy_path

        dummy = Dummy()
        results = run_batch(dummy, [src], is_retry=False, source_alias_map=None)

        self.assertEqual(1, len(results))
        self.assertEqual("success", results[0]["status"])
        self.assertEqual(1, dummy.stats["success"])
        self.assertTrue(dummy.cleared)
        self.assertEqual([(1, 1)], dummy.progress_points)
        self.assertEqual(1, len(dummy.done_records))
        self.assertEqual(1, len(dummy.appended))

    def test_office_converter_run_batch_delegates_to_module(self):
        import office_converter as oc

        original = oc.run_batch_impl
        try:
            seen = {}

            def _fake(converter, file_list, **kwargs):
                seen["converter"] = converter
                seen["file_list"] = file_list
                seen["kwargs"] = kwargs
                return ["ok"]

            oc.run_batch_impl = _fake
            dummy = OfficeConverter.__new__(OfficeConverter)

            out = dummy.run_batch(["a.docx"], is_retry=True, source_alias_map={"a": "b"})
            self.assertEqual(["ok"], out)
            self.assertEqual(["a.docx"], seen.get("file_list"))
            self.assertTrue(seen.get("kwargs", {}).get("is_retry"))
        finally:
            oc.run_batch_impl = original


if __name__ == "__main__":
    unittest.main()
