import json
import shutil
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from office_converter import MODE_CONVERT_ONLY, OfficeConverter


class _ResumeConverter(OfficeConverter):
    def setup_logging(self):
        return

    def print_runtime_summary(self):
        return

    def _check_sandbox_free_space_or_raise(self):
        return

    def _scan_convert_candidates(self):
        raise AssertionError("resume mode should not trigger full scan")

    def process_single_file(
        self, file_path, target_path_initial, ext, progress_str, is_retry=False
    ):
        return "success_no_output", ""

    def close_office_apps(self):
        return

    def _write_excel_structured_json_exports(self):
        return []

    def _write_records_json_exports(self):
        return []

    def _write_chromadb_export(self):
        return None

    def _write_markdown_quality_report(self):
        return None

    def _write_corpus_manifest(self, merge_outputs=None):
        return None

    def _build_perf_summary(self):
        return "perf"


class ConverterResumeTests(unittest.TestCase):
    def setUp(self):
        self.root = Path(tempfile.mkdtemp(prefix="resume_conv_"))
        self.addCleanup(lambda: shutil.rmtree(self.root, ignore_errors=True))
        self.src = self.root / "src"
        self.dst = self.root / "dst"
        self.src.mkdir(parents=True, exist_ok=True)
        self.dst.mkdir(parents=True, exist_ok=True)
        self.config_path = self.root / "config.json"
        cfg = {
            "source_folder": str(self.src),
            "target_folder": str(self.dst),
            "run_mode": MODE_CONVERT_ONLY,
            "enable_merge": False,
            "output_enable_pdf": True,
            "output_enable_md": False,
            "output_enable_merged": False,
            "output_enable_independent": True,
            "enable_corpus_manifest": False,
            "enable_markdown_quality_report": False,
            "enable_excel_json": False,
            "enable_chromadb_export": False,
            "enable_sandbox": False,
        }
        self.config_path.write_text(
            json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8"
        )

    def test_run_resume_file_list_emits_callbacks(self):
        file_a = self.src / "a.docx"
        file_b = self.src / "b.xlsx"
        file_a.write_text("a", encoding="utf-8")
        file_b.write_text("b", encoding="utf-8")

        converter = _ResumeConverter(str(self.config_path), interactive=False)
        planned = []
        done = []
        converter.file_plan_callback = lambda files: planned.append(list(files))
        converter.file_done_callback = lambda rec: done.append(dict(rec))

        with patch("office_converter.os.startfile", lambda _p: None, create=True):
            converter.run(resume_file_list=[str(file_a), str(file_b)])

        self.assertEqual([[str(file_a), str(file_b)]], planned)
        self.assertEqual(2, len(done))
        self.assertEqual(2, converter.stats["total"])
        self.assertEqual(2, converter.stats["success"])


if __name__ == "__main__":
    unittest.main()
