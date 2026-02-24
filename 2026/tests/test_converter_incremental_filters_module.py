import os
import tempfile
import unittest

from office_converter import OfficeConverter


class ConverterIncrementalFiltersSplitTests(unittest.TestCase):
    def test_incremental_filters_core_behaviors(self):
        from converter.incremental_filters import (
            apply_global_md5_dedup,
            apply_source_priority_filter,
        )

        files = [
            r"C:\x\same.docx",
            r"C:\x\same.pdf",
            r"C:\x\other.pdf",
        ]
        kept, skipped = apply_source_priority_filter(
            files,
            {
                "source_priority_skip_same_name_pdf": True,
                "allowed_extensions": {
                    "word": [".docx"],
                    "excel": [],
                    "powerpoint": [],
                },
            },
            is_win_fn=lambda: True,
        )
        self.assertIn(r"C:\x\same.docx", kept)
        self.assertIn(r"C:\x\other.pdf", kept)
        self.assertEqual(len(skipped), 1)
        self.assertEqual(os.path.basename(skipped[0]["source_path"]).lower(), "same.pdf")

        root = tempfile.mkdtemp(prefix="inc_filters_")
        f1 = os.path.join(root, "a.docx")
        f2 = os.path.join(root, "b.docx")
        f3 = os.path.join(root, "c.xlsx")
        with open(f1, "wb") as f:
            f.write(b"same")
        with open(f2, "wb") as f:
            f.write(b"same")
        with open(f3, "wb") as f:
            f.write(b"same")
        try:
            kept2, skipped2 = apply_global_md5_dedup(
                [f1, f2, f3],
                True,
                ext_bucket_fn=lambda p: ".docx" if p.endswith(".docx") else ".xlsx",
                compute_md5_fn=lambda p: "md5-same",
            )
            self.assertIn(f1, kept2)
            self.assertIn(f3, kept2)
            self.assertEqual(len(skipped2), 1)
            self.assertEqual(os.path.abspath(f2), skipped2[0]["source_path"])
        finally:
            for p in (f1, f2, f3):
                try:
                    os.remove(p)
                except Exception:
                    pass
            try:
                os.rmdir(root)
            except Exception:
                pass

    def test_office_converter_incremental_filter_methods_delegate_to_module(self):
        from converter.incremental_filters import (
            apply_global_md5_dedup,
            apply_source_priority_filter,
        )

        dummy = OfficeConverter.__new__(OfficeConverter)
        dummy.config = {
            "source_priority_skip_same_name_pdf": True,
            "allowed_extensions": {
                "word": [".docx"],
                "excel": [],
                "powerpoint": [],
            },
            "global_md5_dedup": True,
        }
        dummy._ext_bucket = lambda p: ".docx"
        dummy._compute_md5 = lambda p: "x"

        files = [r"C:\x\a.docx", r"C:\x\b.pdf"]
        self.assertEqual(
            dummy._apply_source_priority_filter(files),
            apply_source_priority_filter(
                files,
                dummy.config,
                is_win_fn=lambda: os.name == "nt",
            ),
        )

        self.assertEqual(
            dummy._apply_global_md5_dedup([r"C:\x\a.docx", r"C:\x\b.docx"]),
            apply_global_md5_dedup(
                [r"C:\x\a.docx", r"C:\x\b.docx"],
                True,
                dummy._ext_bucket,
                dummy._compute_md5,
            ),
        )


if __name__ == "__main__":
    unittest.main()
