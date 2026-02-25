import os
import tempfile
import unittest
from datetime import datetime
from pathlib import Path

from office_converter import OfficeConverter


class ConverterFileOpsSplitTests(unittest.TestCase):
    def test_file_ops_core_behaviors(self):
        from converter.file_ops import (
            copy_pdf_direct,
            handle_file_conflict,
            quarantine_failed_file,
            unblock_file,
        )

        root = tempfile.mkdtemp(prefix="file_ops_")
        src = os.path.join(root, "a.pdf")
        with open(src, "wb") as f:
            f.write(b"123")

        try:
            # copy_pdf_direct
            copied = os.path.join(root, "copied.pdf")
            copy_pdf_direct(src, copied)
            self.assertTrue(os.path.exists(copied))

            # handle conflict: target absent -> success
            temp1 = os.path.join(root, "tmp1.pdf")
            with open(temp1, "wb") as f:
                f.write(b"abc")
            target = os.path.join(root, "out", "x.pdf")
            status, path = handle_file_conflict(temp1, target)
            self.assertEqual(status, "success")
            self.assertEqual(path, target)
            self.assertTrue(os.path.exists(target))

            # target same size -> overwrite
            temp2 = os.path.join(root, "tmp2.pdf")
            with open(temp2, "wb") as f:
                f.write(b"111")
            with open(target, "wb") as f:
                f.write(b"222")
            status, path = handle_file_conflict(temp2, target)
            self.assertEqual(status, "overwrite")
            self.assertEqual(path, target)

            # target different size -> conflict_saved
            temp3 = os.path.join(root, "tmp3.pdf")
            with open(temp3, "wb") as f:
                f.write(b"short")
            with open(target, "wb") as f:
                f.write(b"long-content")
            status, path = handle_file_conflict(temp3, target, now=datetime(2026, 2, 24, 1, 2, 3))
            self.assertEqual(status, "conflict_saved")
            self.assertIn(os.path.join("conflicts", "x_20260224010203.pdf"), path)
            self.assertTrue(os.path.exists(path))

            # quarantine
            failed_dir = os.path.join(root, "failed")
            os.makedirs(failed_dir, exist_ok=True)
            q1 = quarantine_failed_file(src, failed_dir, should_copy=True)
            self.assertTrue(os.path.exists(q1))
            q2 = quarantine_failed_file(src, failed_dir, should_copy=True, now=datetime(2026, 2, 24, 9, 9, 9))
            self.assertTrue(os.path.exists(q2))
            self.assertNotEqual(q1, q2)

            # unblock should be no-op on missing ADS
            unblock_file(src)
        finally:
            for p in [
                os.path.join(root, "copied.pdf"),
                os.path.join(root, "tmp1.pdf"),
                os.path.join(root, "tmp2.pdf"),
                os.path.join(root, "tmp3.pdf"),
                os.path.join(root, "out", "x.pdf"),
                os.path.join(root, "out", "conflicts", "x_20260224010203.pdf"),
                os.path.join(root, "failed", "a.pdf"),
                os.path.join(root, "failed", "a_090909.pdf"),
                src,
            ]:
                try:
                    os.remove(p)
                except Exception:
                    pass
            for d in [
                os.path.join(root, "out", "conflicts"),
                os.path.join(root, "out"),
                os.path.join(root, "failed"),
                root,
            ]:
                try:
                    os.rmdir(d)
                except Exception:
                    pass

    def test_office_converter_file_ops_methods_delegate_to_module(self):
        root = tempfile.mkdtemp(prefix="file_ops_delegate_")
        src = os.path.join(root, "a.pdf")
        with open(src, "wb") as f:
            f.write(b"abc")
        dummy = OfficeConverter.__new__(OfficeConverter)
        dummy.failed_dir = os.path.join(root, "failed")
        os.makedirs(dummy.failed_dir, exist_ok=True)

        try:
            copied = os.path.join(root, "b.pdf")
            dummy.copy_pdf_direct(src, copied)
            self.assertTrue(os.path.exists(copied))

            q = dummy.quarantine_failed_file(src)
            self.assertTrue(os.path.exists(q))

            temp = os.path.join(root, "tmp.pdf")
            with open(temp, "wb") as f:
                f.write(b"1")
            target = os.path.join(root, "out", "x.pdf")
            status, _ = dummy.handle_file_conflict(temp, target)
            self.assertIn(status, {"success", "overwrite", "conflict_saved"})

            dummy._unblock_file(src)
        finally:
            for p in [
                os.path.join(root, "b.pdf"),
                os.path.join(root, "failed", "a.pdf"),
                os.path.join(root, "tmp.pdf"),
                os.path.join(root, "out", "x.pdf"),
                src,
            ]:
                try:
                    os.remove(p)
                except Exception:
                    pass
            for d in [
                os.path.join(root, "out", "conflicts"),
                os.path.join(root, "out"),
                os.path.join(root, "failed"),
                root,
            ]:
                try:
                    os.rmdir(d)
                except Exception:
                    pass

    def test_file_ops_module_has_no_bare_except_exception(self):
        mod_path = Path(__file__).resolve().parents[1] / "converter" / "file_ops.py"
        text = mod_path.read_text(encoding="utf-8")
        self.assertNotIn("except Exception", text)


if __name__ == "__main__":
    unittest.main()
