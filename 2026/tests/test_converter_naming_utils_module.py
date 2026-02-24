import unittest
from datetime import datetime

from office_converter import OfficeConverter


class ConverterNamingUtilsSplitTests(unittest.TestCase):
    def test_naming_utils_core_behaviors(self):
        from converter.naming_utils import ext_bucket, format_merge_filename

        now = datetime(2026, 2, 24, 15, 4, 5)
        name = format_merge_filename(
            "{category}_{date}_{time}_{idx}",
            category="报价/目录",
            idx=2,
            now=now,
        )
        self.assertTrue(name.endswith(".pdf"))
        self.assertIn("20260224_150405_2", name)
        self.assertNotIn("/", name)

        allowed = {
            "word": [".doc", ".docx"],
            "excel": [".xls", ".xlsx"],
            "powerpoint": [".ppt", ".pptx"],
        }
        self.assertEqual(ext_bucket("a.DOCX", allowed), "word")
        self.assertEqual(ext_bucket("a.xlsx", allowed), "excel")
        self.assertEqual(ext_bucket("a.ppt", allowed), "powerpoint")
        self.assertEqual(ext_bucket("a.pdf", allowed), "pdf")
        self.assertEqual(ext_bucket("a.abc", allowed), ".abc")

    def test_office_converter_naming_methods_delegate_to_module(self):
        from converter.naming_utils import ext_bucket, format_merge_filename

        dummy = OfficeConverter.__new__(OfficeConverter)
        dummy.config = {
            "allowed_extensions": {
                "word": [".docx"],
                "excel": [".xlsx"],
                "powerpoint": [".pptx"],
            }
        }

        self.assertEqual(
            OfficeConverter._format_merge_filename("m_{idx}", idx=3, now=datetime(2026, 1, 1)),
            format_merge_filename("m_{idx}", idx=3, now=datetime(2026, 1, 1)),
        )
        self.assertEqual(
            dummy._ext_bucket("x.docx"),
            ext_bucket("x.docx", dummy.config["allowed_extensions"]),
        )


if __name__ == "__main__":
    unittest.main()
