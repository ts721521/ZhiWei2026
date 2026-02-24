import os
import tempfile
import unittest

from office_converter import OfficeConverter


class ConverterProcessSingleSplitTests(unittest.TestCase):
    def test_process_single_core_empty_file_behavior(self):
        from converter.process_single import process_single_file

        root = tempfile.mkdtemp(prefix="process_single_")
        src = os.path.join(root, "a.docx")
        with open(src, "w", encoding="utf-8"):
            pass

        class Dummy:
            def __init__(self):
                self.stats = {"skipped": 0}

        dummy = Dummy()
        status, out = process_single_file(
            dummy,
            src,
            "target.pdf",
            ".docx",
            "[1/1]",
            is_retry=False,
        )
        self.assertEqual("skip_empty", status)
        self.assertEqual("target.pdf", out)
        self.assertEqual(1, dummy.stats["skipped"])

    def test_office_converter_process_single_file_delegates_to_module(self):
        import office_converter as oc

        original = oc.process_single_file_impl
        try:
            seen = {}

            def _fake(converter, file_path, target_path_initial, ext, progress_str, **kwargs):
                seen["converter"] = converter
                seen["file_path"] = file_path
                seen["kwargs"] = kwargs
                return "ok", "out.pdf"

            oc.process_single_file_impl = _fake
            dummy = OfficeConverter.__new__(OfficeConverter)

            out = dummy.process_single_file("a.docx", "x.pdf", ".docx", "[1/1]", is_retry=True)
            self.assertEqual(("ok", "out.pdf"), out)
            self.assertEqual("a.docx", seen.get("file_path"))
            self.assertTrue(seen.get("kwargs", {}).get("is_retry"))
        finally:
            oc.process_single_file_impl = original


if __name__ == "__main__":
    unittest.main()
