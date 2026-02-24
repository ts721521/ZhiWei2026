import os
import tempfile
import unittest
import zipfile

from office_converter import OfficeConverter


class ConverterTextHelpersSplitTests(unittest.TestCase):
    def test_text_helpers_core_behaviors(self):
        from converter.text_helpers import (
            extract_mshc_payload,
            find_files_recursive,
            meta_content_by_names,
            normalize_md_line,
            wrap_plain_text_for_pdf,
        )

        root = tempfile.mkdtemp(prefix="text_helpers_")
        nested = os.path.join(root, "a", "b")
        os.makedirs(nested, exist_ok=True)
        p1 = os.path.join(root, "x.htm")
        p2 = os.path.join(nested, "y.html")
        with open(p1, "w", encoding="utf-8") as f:
            f.write("x")
        with open(p2, "w", encoding="utf-8") as f:
            f.write("y")
        try:
            found = find_files_recursive(root, (".htm", ".html"))
            self.assertIn(p1, found)
            self.assertIn(p2, found)
        finally:
            for p in (p2, p1):
                try:
                    os.remove(p)
                except Exception:
                    pass
            for d in (nested, os.path.join(root, "a"), root):
                try:
                    os.rmdir(d)
                except Exception:
                    pass

        self.assertEqual(normalize_md_line("  a \t \n b  "), "a b")
        wrapped = wrap_plain_text_for_pdf("a bb ccc dddd", width=5)
        self.assertEqual(wrapped, ["a bb", "ccc", "dddd"])
        self.assertEqual(wrap_plain_text_for_pdf("", width=10), [""])

        class _Meta:
            def __init__(self, name, content):
                self._name = name
                self._content = content

            def get(self, key, default=""):
                if key == "name":
                    return self._name
                if key == "content":
                    return self._content
                return default

        class _Soup:
            def find_all(self, tag):
                if tag != "meta":
                    return []
                return [_Meta("Title", "Doc A"), _Meta("Other", "X")]

        self.assertEqual(meta_content_by_names(_Soup(), ["title"]), "Doc A")
        self.assertEqual(meta_content_by_names(_Soup(), ["none"]), "")

        zip_root = tempfile.mkdtemp(prefix="mshc_zip_")
        archive = os.path.join(zip_root, "payload.mshc")
        out_dir = os.path.join(zip_root, "out")
        with zipfile.ZipFile(archive, "w") as zf:
            zf.writestr("topic/index.html", "<h1>ok</h1>")
        try:
            extract_mshc_payload(archive, out_dir)
            self.assertTrue(os.path.exists(os.path.join(out_dir, "topic", "index.html")))
        finally:
            try:
                os.remove(os.path.join(out_dir, "topic", "index.html"))
            except Exception:
                pass
            for d in (os.path.join(out_dir, "topic"), out_dir):
                try:
                    os.rmdir(d)
                except Exception:
                    pass
            try:
                os.remove(archive)
            except Exception:
                pass
            try:
                os.rmdir(zip_root)
            except Exception:
                pass

    def test_office_converter_text_helper_methods_delegate_to_module(self):
        from converter.text_helpers import (
            extract_mshc_payload,
            find_files_recursive,
            meta_content_by_names,
            normalize_md_line,
            wrap_plain_text_for_pdf,
        )

        self.assertEqual(
            OfficeConverter._normalize_md_line(" a  b "),
            normalize_md_line(" a  b "),
        )
        self.assertEqual(
            OfficeConverter._wrap_plain_text_for_pdf("a b c", width=2),
            wrap_plain_text_for_pdf("a b c", width=2),
        )

        class _Soup:
            def find_all(self, tag):
                class _Meta:
                    def get(self, key, default=""):
                        return "k" if key == "name" else "v"

                return [_Meta()]

        self.assertEqual(
            OfficeConverter._meta_content_by_names(_Soup(), ["k"]),
            meta_content_by_names(_Soup(), ["k"]),
        )

        root = tempfile.mkdtemp(prefix="delegate_find_")
        file_path = os.path.join(root, "a.htm")
        with open(file_path, "w", encoding="utf-8") as f:
            f.write("x")
        try:
            self.assertEqual(
                sorted(OfficeConverter._find_files_recursive(root, (".htm",))),
                sorted(find_files_recursive(root, (".htm",))),
            )
        finally:
            try:
                os.remove(file_path)
            except Exception:
                pass
            try:
                os.rmdir(root)
            except Exception:
                pass

        zip_root = tempfile.mkdtemp(prefix="delegate_zip_")
        archive = os.path.join(zip_root, "a.mshc")
        out_dir = os.path.join(zip_root, "out")
        with zipfile.ZipFile(archive, "w") as zf:
            zf.writestr("x.txt", "1")
        try:
            OfficeConverter._extract_mshc_payload(archive, out_dir)
            self.assertTrue(os.path.exists(os.path.join(out_dir, "x.txt")))
            os.remove(os.path.join(out_dir, "x.txt"))
            os.rmdir(out_dir)
            extract_mshc_payload(archive, out_dir)
            self.assertTrue(os.path.exists(os.path.join(out_dir, "x.txt")))
        finally:
            try:
                os.remove(os.path.join(out_dir, "x.txt"))
            except Exception:
                pass
            try:
                os.rmdir(out_dir)
            except Exception:
                pass
            try:
                os.remove(archive)
            except Exception:
                pass
            try:
                os.rmdir(zip_root)
            except Exception:
                pass


if __name__ == "__main__":
    unittest.main()
