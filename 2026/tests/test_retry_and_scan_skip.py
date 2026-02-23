import json
import os
import shutil
import tempfile
import unittest
from pathlib import Path

from office_converter import MODE_CONVERT_ONLY, OfficeConverter


class RetryAndScanSkipTests(unittest.TestCase):
    def setUp(self):
        self.root = Path(tempfile.mkdtemp(prefix="retry_scan_"))
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
            "enable_sandbox": False,
        }
        self.config_path.write_text(
            json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        self.converter = OfficeConverter(str(self.config_path), interactive=False)

    def test_collect_retry_candidates_filters_non_retryable_errors(self):
        retryable_src = str(self.src / "a.docx")
        manual_src = str(self.src / "b.xlsx")
        self.converter.error_records = [retryable_src, manual_src]
        self.converter.detailed_error_records = [
            {"source_path": os.path.abspath(retryable_src), "is_retryable": True},
            {"source_path": os.path.abspath(manual_src), "is_retryable": False},
        ]

        failed_retryable = Path(self.converter.failed_dir) / "a.docx"
        failed_manual = Path(self.converter.failed_dir) / "b.xlsx"
        failed_retryable.write_text("retryable", encoding="utf-8")
        failed_manual.write_text("manual", encoding="utf-8")

        retry_files, retry_alias_map = self.converter._collect_retry_candidates()

        self.assertEqual([str(failed_retryable)], retry_files)
        self.assertEqual(retryable_src, retry_alias_map[str(failed_retryable)])

    def test_probe_source_root_access_records_skip(self):
        missing = self.root / "need_domain_login"
        ok = self.converter._probe_source_root_access(
            str(missing), context={"scan_scope": "convert"}, seen_keys=set()
        )

        self.assertFalse(ok)
        self.assertEqual(1, self.converter.stats["skipped"])
        self.assertEqual(1, len(self.converter.detailed_error_records))
        record = self.converter.detailed_error_records[0]
        self.assertEqual(os.path.abspath(str(missing)), record["source_path"])
        self.assertEqual("scan", record["context"]["phase"])
        self.assertTrue(record["context"]["skip_only"])


if __name__ == "__main__":
    unittest.main()
