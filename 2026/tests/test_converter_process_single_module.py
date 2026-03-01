import os
import tempfile
import unittest

from pathlib import Path

from converter.constants import STRATEGY_STANDARD
from office_converter import OfficeConverter


class ConverterProcessSingleSplitTests(unittest.TestCase):
    def test_process_single_module_has_no_bare_except_exception(self):
        module_text = Path("converter/process_single.py").read_text(encoding="utf-8")
        self.assertNotIn("except Exception", module_text)

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

    def test_process_single_fast_fail_invalid_docx_container(self):
        from converter.process_single import process_single_file

        root = tempfile.mkdtemp(prefix="process_single_bad_docx_")
        src = os.path.join(root, "broken.docx")
        with open(src, "wb") as f:
            f.write(b"\x00\x00\x00\x00invalid")

        class Dummy:
            def __init__(self, sandbox):
                self.temp_sandbox = sandbox
                self.stats = {"skipped": 0}
                self.price_keywords = []
                self.content_strategy = STRATEGY_STANDARD
                self.run_mode = "convert_only"
                self.config = {
                    "enable_sandbox": False,
                    "allowed_extensions": {
                        "word": [".doc", ".docx"],
                        "excel": [".xls", ".xlsx"],
                        "powerpoint": [".ppt", ".pptx"],
                        "cab": [".cab"],
                    },
                }

            @staticmethod
            def compute_convert_output_plan(*_args, **_kwargs):
                return {"need_final_pdf": False, "need_markdown": False}

            @staticmethod
            def _on_office_file_processed(_ext):
                return None

        with self.assertRaisesRegex(ValueError, "invalid .docx package signature"):
            process_single_file(
                Dummy(root),
                src,
                "target.pdf",
                ".docx",
                "[1/1]",
                is_retry=False,
            )


if __name__ == "__main__":
    unittest.main()
