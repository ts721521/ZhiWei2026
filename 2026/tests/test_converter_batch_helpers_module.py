import os
import tempfile
import unittest

from office_converter import OfficeConverter


class ConverterBatchHelpersSplitTests(unittest.TestCase):
    def test_batch_helpers_core_behaviors(self):
        from converter.batch_helpers import collect_retry_candidates, get_progress_prefix

        self.assertEqual(
            get_progress_prefix(1, 4),
            "[ 25%]#####--------------- [1/4]",
        )
        self.assertEqual(
            get_progress_prefix(0, 0),
            "[  0%]-------------------- [0/0]",
        )

        root = tempfile.mkdtemp(prefix="retry_helpers_")
        failed_dir = os.path.join(root, "failed")
        os.makedirs(failed_dir, exist_ok=True)
        f1 = os.path.join(failed_dir, "a.docx")
        f2 = os.path.join(failed_dir, "b.xlsx")
        with open(f1, "w", encoding="utf-8") as f:
            f.write("a")
        with open(f2, "w", encoding="utf-8") as f:
            f.write("b")

        try:
            retry_files, alias_map = collect_retry_candidates(
                failed_dir,
                {"word": [".docx"], "excel": [".xlsx"]},
                [r"C:\src\a.docx", r"C:\src\b.xlsx"],
                [
                    {"source_path": r"C:\src\a.docx", "is_retryable": True},
                    {"source_path": r"C:\src\b.xlsx", "is_retryable": False},
                ],
            )
            self.assertEqual(len(retry_files), 1)
            self.assertEqual(os.path.basename(retry_files[0]), "a.docx")
            self.assertEqual(alias_map[retry_files[0]], r"C:\src\a.docx")
        finally:
            for p in (f1, f2):
                try:
                    os.remove(p)
                except Exception:
                    pass
            for d in (failed_dir, root):
                try:
                    os.rmdir(d)
                except Exception:
                    pass

    def test_office_converter_batch_helper_methods_delegate_to_module(self):
        from converter.batch_helpers import collect_retry_candidates, get_progress_prefix

        self.assertEqual(
            OfficeConverter.__new__(OfficeConverter).get_progress_prefix(2, 10),
            get_progress_prefix(2, 10),
        )

        root = tempfile.mkdtemp(prefix="retry_delegate_")
        failed_dir = os.path.join(root, "failed")
        os.makedirs(failed_dir, exist_ok=True)
        f1 = os.path.join(failed_dir, "c.docx")
        with open(f1, "w", encoding="utf-8") as f:
            f.write("c")

        dummy = OfficeConverter.__new__(OfficeConverter)
        dummy.failed_dir = failed_dir
        dummy.config = {"allowed_extensions": {"word": [".docx"]}}
        dummy.error_records = [r"C:\src\c.docx"]
        dummy.detailed_error_records = [{"source_path": r"C:\src\c.docx", "is_retryable": True}]

        try:
            self.assertEqual(
                dummy._collect_retry_candidates(),
                collect_retry_candidates(
                    failed_dir,
                    dummy.config["allowed_extensions"],
                    dummy.error_records,
                    dummy.detailed_error_records,
                ),
            )
        finally:
            try:
                os.remove(f1)
            except Exception:
                pass
            for d in (failed_dir, root):
                try:
                    os.rmdir(d)
                except Exception:
                    pass


if __name__ == "__main__":
    unittest.main()
