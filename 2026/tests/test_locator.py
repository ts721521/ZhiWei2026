import json
import os
import tempfile
import unittest

from locate_source import locate_by_page, locate_by_short_id, EXIT_OK, EXIT_NOT_FOUND


class LocatorTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.map_dir = self.tmp.name
        payload = {
            "version": 1,
            "record_count": 2,
            "records": [
                {
                    "merge_batch_id": "20260210_010101",
                    "merged_pdf_name": "Merged_All_20260210_010101.pdf",
                    "merged_pdf_path": "C:/PDFs/_MERGED/Merged_All_20260210_010101.pdf",
                    "source_index": 1,
                    "source_filename": "A.pdf",
                    "source_abspath": "C:/docs/A.pdf",
                    "source_relpath": "A.pdf",
                    "source_md5": "a" * 32,
                    "source_short_id": "AAAAAAAA",
                    "start_page_1based": 1,
                    "end_page_1based": 3,
                    "page_count": 3,
                    "bookmark_title": "[ID:AAAAAAAA] A.pdf",
                },
                {
                    "merge_batch_id": "20260210_010101",
                    "merged_pdf_name": "Merged_All_20260210_010101.pdf",
                    "merged_pdf_path": "C:/PDFs/_MERGED/Merged_All_20260210_010101.pdf",
                    "source_index": 2,
                    "source_filename": "B.pdf",
                    "source_abspath": "C:/docs/B.pdf",
                    "source_relpath": "B.pdf",
                    "source_md5": "b" * 32,
                    "source_short_id": "BBBBBBBB",
                    "start_page_1based": 4,
                    "end_page_1based": 8,
                    "page_count": 5,
                    "bookmark_title": "[ID:BBBBBBBB] B.pdf",
                },
            ],
        }
        with open(os.path.join(self.map_dir, "Merged_All_20260210_010101.map.json"), "w", encoding="utf-8") as f:
            json.dump(payload, f)

    def tearDown(self):
        self.tmp.cleanup()

    def test_locate_by_page_hit(self):
        result = locate_by_page("Merged_All_20260210_010101.pdf", 5, self.map_dir)
        self.assertEqual(result.error_code, EXIT_OK)
        self.assertIsNotNone(result.record)
        self.assertEqual(result.record.source_filename, "B.pdf")

    def test_locate_by_page_not_found(self):
        result = locate_by_page("Merged_All_20260210_010101.pdf", 100, self.map_dir)
        self.assertEqual(result.error_code, EXIT_NOT_FOUND)
        self.assertIsNone(result.record)
        self.assertTrue(len(result.alternatives) >= 1)

    def test_locate_by_short_id_hit(self):
        result = locate_by_short_id("BBBBBBBB", self.map_dir)
        self.assertEqual(result.error_code, EXIT_OK)
        self.assertIsNotNone(result.record)
        self.assertEqual(result.record.source_filename, "B.pdf")

    def test_locate_by_short_id_hit_with_zw_prefix(self):
        result = locate_by_short_id("ZW-BBBBBBBB", self.map_dir)
        self.assertEqual(result.error_code, EXIT_OK)
        self.assertIsNotNone(result.record)
        self.assertEqual(result.record.source_filename, "B.pdf")


if __name__ == "__main__":
    unittest.main()
