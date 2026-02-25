import unittest

from pathlib import Path

from office_converter import OfficeConverter


class ConverterMergePdfsSplitTests(unittest.TestCase):
    def test_merge_pdfs_module_has_no_bare_except_exception(self):
        module_text = Path("converter/merge_pdfs.py").read_text(encoding="utf-8")
        self.assertNotIn("except Exception", module_text)

    def test_merge_pdfs_core_behaviors_when_pypdf_missing(self):
        from converter.merge_pdfs import merge_pdfs

        class Dummy:
            def __init__(self):
                self.config = {"enable_merge": True}

        dummy = Dummy()
        out = merge_pdfs(
            dummy,
            has_pypdf=False,
            has_openpyxl=False,
            pdf_writer_cls=None,
            pdf_reader_cls=None,
            workbook_cls=None,
            pythoncom_mod=None,
            win32_client=None,
        )
        self.assertEqual([], out)

    def test_office_converter_merge_pdfs_delegates_to_module(self):
        import office_converter as oc

        original = oc.merge_pdfs_impl
        try:
            seen = {}

            def _fake(converter, **kwargs):
                seen["converter"] = converter
                seen["kwargs"] = kwargs
                return ["ok"]

            oc.merge_pdfs_impl = _fake
            dummy = OfficeConverter.__new__(OfficeConverter)

            out = dummy.merge_pdfs()
            self.assertEqual(["ok"], out)
            self.assertIn("has_pypdf", seen.get("kwargs", {}))
        finally:
            oc.merge_pdfs_impl = original


if __name__ == "__main__":
    unittest.main()
