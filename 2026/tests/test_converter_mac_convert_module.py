import os
import tempfile
import unittest
from pathlib import Path

from office_converter import OfficeConverter


class ConverterMacConvertSplitTests(unittest.TestCase):
    def test_mac_convert_module_has_no_bare_except_exception(self):
        module_text = Path("converter/mac_convert.py").read_text(encoding="utf-8")
        self.assertNotIn("except Exception", module_text)

    def test_mac_convert_core_behaviors(self):
        from converter.mac_convert import convert_on_mac

        self.assertFalse(
            convert_on_mac(
                "a.docx",
                "a.pdf",
                ".docx",
                is_mac_fn=lambda: False,
            )
        )

        root = tempfile.mkdtemp(prefix="mac_convert_")
        src = os.path.join(root, "demo.docx")
        out = os.path.join(root, "out.pdf")
        with open(src, "w", encoding="utf-8") as f:
            f.write("x")

        def _run(cmd, check=True, stdout=None, stderr=None):
            self.assertTrue(check)
            self.assertIsNotNone(stdout)
            self.assertIsNotNone(stderr)
            generated = os.path.join(root, "demo.pdf")
            with open(generated, "w", encoding="utf-8") as f:
                f.write("pdf")

        ok = convert_on_mac(
            src,
            out,
            ".docx",
            is_mac_fn=lambda: True,
            which_fn=lambda _name: "soffice",
            run_cmd=_run,
            log_error_fn=lambda _m: None,
            log_warning_fn=lambda _m: None,
        )
        self.assertTrue(ok)
        self.assertTrue(os.path.exists(out))

    def test_office_converter_convert_on_mac_delegates_to_module(self):
        import office_converter as oc

        original = oc.convert_on_mac_impl
        try:
            seen = {}

            def _fake(file_source, sandbox_target_pdf, ext, **kwargs):
                seen["args"] = (file_source, sandbox_target_pdf, ext)
                seen["kwargs"] = kwargs
                return True

            oc.convert_on_mac_impl = _fake
            dummy = OfficeConverter.__new__(OfficeConverter)
            out = dummy._convert_on_mac("a.docx", "a.pdf", ".docx")
            self.assertTrue(out)
            self.assertEqual(seen["args"], ("a.docx", "a.pdf", ".docx"))
            self.assertTrue(callable(seen["kwargs"]["is_mac_fn"]))
        finally:
            oc.convert_on_mac_impl = original


if __name__ == "__main__":
    unittest.main()
