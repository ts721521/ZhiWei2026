import os
import tempfile
import unittest
from unittest import mock

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

    def test_run_workflow_does_not_open_folder_in_unittest_context(self):
        from converter.constants import MODE_COLLECT_ONLY
        from converter.run_workflow import run

        root = tempfile.mkdtemp(prefix="run_workflow_no_open_")

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

            def setup_logging(self):
                pass

            def print_runtime_summary(self):
                pass

            def _reset_perf_metrics(self):
                self.perf_metrics = {}

            def _check_sandbox_free_space_or_raise(self):
                pass

            def collect_office_files_and_build_excel(self):
                pass

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
        with mock.patch("converter.run_workflow.os.startfile", create=True) as mocked_open:
            run(dummy, resume_file_list=None, app_version="test")
        mocked_open.assert_not_called()

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

    def test_run_workflow_uses_fast_md_pipeline_when_enabled(self):
        from converter.constants import MODE_CONVERT_ONLY
        from converter.run_workflow import run

        root = tempfile.mkdtemp(prefix="run_workflow_fast_md_")
        src = os.path.join(root, "a.docx")
        with open(src, "w", encoding="utf-8") as f:
            f.write("x")

        class Dummy:
            def __init__(self):
                self.run_mode = MODE_CONVERT_ONLY
                self.config = {
                    "target_folder": root,
                    "enable_fast_md_engine": True,
                    "enable_merge": True,
                    "output_enable_merged": True,
                }
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
                self.interactive = False
                self.enable_merge_excel = False
                self.reuse_process = True
                self.failed_dir = os.path.join(root, "_FAILED")
                self.trace_short_id_taken = set()
                self.generated_markdown_outputs = []
                self.markdown_quality_records = []
                self.fast_called = 0
                self.batch_called = 0

            def setup_logging(self):
                pass

            def print_runtime_summary(self):
                pass

            def _reset_perf_metrics(self):
                self.perf_metrics = {}

            def _check_sandbox_free_space_or_raise(self):
                pass

            def _scan_convert_candidates(self):
                return [src]

            def _apply_source_priority_filter(self, files):
                return files, []

            def _apply_incremental_filter(self, files):
                return files, {}

            def _apply_global_md5_dedup(self, files):
                return files, []

            def _emit_file_plan(self, _files):
                pass

            def _run_fast_md_pipeline(self, files):
                self.fast_called += 1
                out = os.path.join(root, "_MD_Corpus", "_Knowledge_Bundle.md")
                os.makedirs(os.path.dirname(out), exist_ok=True)
                with open(out, "w", encoding="utf-8") as f:
                    f.write("# bundle\n")
                self.generated_fast_md_outputs = [out]
                return {
                    "batch_results": [
                        {
                            "source_path": files[0],
                            "status": "success_md_only",
                            "detail": "fast_md",
                            "final_path": out,
                            "elapsed": 0.0,
                        }
                    ],
                    "bundle_path": out,
                    "markdown_files": [out],
                }

            def run_batch_parallel(self, files):
                self.batch_called += 1
                return []

            def run_batch(self, files):
                self.batch_called += 1
                return []

            def close_office_apps(self):
                pass

            def _flush_incremental_registry(self, _r):
                pass

            def _generate_update_package(self, _r):
                pass

            def _write_excel_structured_json_exports(self):
                pass

            def _write_records_json_exports(self):
                pass

            def _write_trace_map(self):
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
        self.assertEqual(dummy.fast_called, 1)
        self.assertEqual(dummy.batch_called, 0)
        self.assertIn("total_seconds", dummy.perf_metrics)


if __name__ == "__main__":
    unittest.main()

