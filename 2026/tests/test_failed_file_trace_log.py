import json
import shutil
import tempfile
import unittest
from pathlib import Path

from office_converter import MODE_CONVERT_ONLY, OfficeConverter


class _AlwaysFailConverter(OfficeConverter):
    def process_single_file(
        self, file_path, target_path_initial, ext, progress_str, is_retry=False
    ):
        raise RuntimeError("mock pdf conversion failed")


class FailedFileTraceLogTests(unittest.TestCase):
    def setUp(self):
        self.root = Path(tempfile.mkdtemp(prefix="failed_trace_"))
        self.addCleanup(lambda: shutil.rmtree(self.root, ignore_errors=True))
        self.src = self.root / "src"
        self.dst = self.root / "dst"
        self.src.mkdir(parents=True, exist_ok=True)
        self.dst.mkdir(parents=True, exist_ok=True)

    def _write_config(self, enable_trace_log=True):
        cfg = {
            "source_folder": str(self.src),
            "target_folder": str(self.dst),
            "run_mode": MODE_CONVERT_ONLY,
            "enable_merge": False,
            "enable_sandbox": False,
            "enable_failed_file_trace_log": bool(enable_trace_log),
            "output_enable_pdf": True,
            "output_enable_md": True,
            "output_enable_independent": True,
        }
        cfg_path = self.root / f"config_{'on' if enable_trace_log else 'off'}.json"
        cfg_path.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")
        return cfg_path

    def test_run_batch_failure_writes_per_file_trace_log(self):
        config_path = self._write_config(enable_trace_log=True)
        source = self.src / "a.docx"
        source.write_text("demo", encoding="utf-8")

        converter = _AlwaysFailConverter(str(config_path), interactive=False)
        results = converter.run_batch([str(source)])

        self.assertEqual(1, len(results))
        self.assertEqual("failed", results[0]["status"])
        self.assertIn("failure_trace_path", results[0])
        trace_path = Path(results[0]["failure_trace_path"])
        self.assertTrue(trace_path.exists())
        self.assertEqual(Path(converter.failed_dir), trace_path.parent)

        payload = json.loads(trace_path.read_text(encoding="utf-8"))
        self.assertEqual(str(source.resolve()), payload["source_path"])
        self.assertEqual("failed", payload["status"])
        self.assertEqual("office_to_pdf", payload["failure_stage"])
        self.assertTrue(payload["expected_outputs"]["need_final_pdf"])
        self.assertTrue(payload["expected_outputs"]["need_markdown"])

    def test_trace_log_can_be_disabled(self):
        config_path = self._write_config(enable_trace_log=False)
        source = self.src / "b.docx"
        source.write_text("demo", encoding="utf-8")

        converter = _AlwaysFailConverter(str(config_path), interactive=False)
        results = converter.run_batch([str(source)])

        self.assertEqual(1, len(results))
        self.assertNotIn("failure_trace_path", results[0])
        trace_logs = list(Path(converter.failed_dir).glob("*.failure.json"))
        self.assertEqual([], trace_logs)


if __name__ == "__main__":
    unittest.main()
