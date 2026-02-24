import os
import tempfile
import unittest

from office_converter import OfficeConverter


class ConverterRunWorkflowSplitTests(unittest.TestCase):
    def test_run_workflow_core_behaviors_collect_mode(self):
        from converter.constants import MODE_COLLECT_ONLY
        from converter.run_workflow import run

        root = tempfile.mkdtemp(prefix="run_workflow_")

        class Dummy:
            def __init__(self):
                self.run_mode = MODE_COLLECT_ONLY
                self.config = {"target_folder": root}
                self.stats = {
                    "total": 0,
                    "success": 0,
                    "failed": 0,
                    "timeout": 0,
                    "skipped": 0,
                }
                self.perf_metrics = {}
                self.detailed_error_records = []
                self.merge_output_dir = ""
                self.temp_sandbox = ""
                self._incremental_context = None
                self.collect_called = False

            def setup_logging(self):
                pass

            def print_runtime_summary(self):
                pass

            def _reset_perf_metrics(self):
                self.perf_metrics = {}

            def _check_sandbox_free_space_or_raise(self):
                pass

            def collect_office_files_and_build_excel(self):
                self.collect_called = True

            def _write_excel_structured_json_exports(self):
                pass

            def _write_records_json_exports(self):
                pass

            def _write_chromadb_export(self):
                pass

            def _write_markdown_quality_report(self):
                pass

            def _write_corpus_manifest(self, merge_outputs=None):
                self.last_merge_outputs = merge_outputs

            def _add_perf_seconds(self, key, value):
                self.perf_metrics[key] = self.perf_metrics.get(key, 0.0) + float(value)

            def _build_perf_summary(self):
                return "perf"

            def export_failed_files_report(self):
                return {}

        dummy = Dummy()
        run(dummy, resume_file_list=None, app_version="test")

        self.assertTrue(dummy.collect_called)
        self.assertIn("total_seconds", dummy.perf_metrics)

    def test_office_converter_run_delegates_to_module(self):
        import office_converter as oc

        original = oc.run_workflow_impl
        try:
            seen = {}

            def _fake(converter, **kwargs):
                seen["converter"] = converter
                seen["kwargs"] = kwargs
                return "ok"

            oc.run_workflow_impl = _fake
            dummy = OfficeConverter.__new__(OfficeConverter)

            out = dummy.run(resume_file_list=["a"])
            self.assertEqual("ok", out)
            self.assertEqual(["a"], seen.get("kwargs", {}).get("resume_file_list"))
        finally:
            oc.run_workflow_impl = original


if __name__ == "__main__":
    unittest.main()
