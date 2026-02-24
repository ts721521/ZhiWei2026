import hashlib
import os
import tempfile
import unittest

from office_converter import OfficeConverter


class ConverterHashUtilsSplitTests(unittest.TestCase):
    def test_hash_utils_core_behaviors(self):
        from converter.hash_utils import (
            build_short_id,
            compute_file_hash,
            compute_md5,
            make_file_hyperlink,
            mask_md5,
        )

        fd, path = tempfile.mkstemp()
        os.close(fd)
        try:
            with open(path, "wb") as f:
                f.write(b"abc123")

            self.assertEqual(compute_md5(path), hashlib.md5(b"abc123").hexdigest())
            self.assertEqual(compute_file_hash(path), hashlib.sha256(b"abc123").hexdigest())

            md5_value = "0123456789abcdef0123456789abcdef"
            self.assertEqual(mask_md5(md5_value), "01234567...cdef")
            self.assertEqual(mask_md5("short"), "short")

            taken = {"01234567"}
            short_id = build_short_id(md5_value, taken)
            self.assertEqual(short_id, "0123456789")

            link = make_file_hyperlink(path)
            self.assertTrue(link.startswith("file:///"))
            self.assertIn(os.path.abspath(path).replace("\\", "/"), link)
        finally:
            try:
                os.remove(path)
            except Exception:
                pass

    def test_office_converter_hash_methods_delegate_to_module(self):
        from converter.hash_utils import (
            build_short_id,
            compute_file_hash,
            compute_md5,
            make_file_hyperlink,
            mask_md5,
        )

        fd, path = tempfile.mkstemp()
        os.close(fd)
        try:
            with open(path, "wb") as f:
                f.write(b"delegate-check")

            self.assertEqual(OfficeConverter._compute_md5(path), compute_md5(path))
            self.assertEqual(OfficeConverter._compute_file_hash(path), compute_file_hash(path))
            self.assertEqual(
                OfficeConverter._make_file_hyperlink(path),
                make_file_hyperlink(path),
            )
            self.assertEqual(OfficeConverter._mask_md5("abcdef123456"), mask_md5("abcdef123456"))

            taken_a = set()
            taken_b = set()
            md5_value = "aaaaaaaa11111111bbbbbbbb22222222"
            self.assertEqual(
                OfficeConverter._build_short_id(md5_value, taken_a),
                build_short_id(md5_value, taken_b),
            )
        finally:
            try:
                os.remove(path)
            except Exception:
                pass


if __name__ == "__main__":
    unittest.main()
