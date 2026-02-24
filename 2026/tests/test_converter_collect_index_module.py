import unittest

from office_converter import OfficeConverter


class ConverterCollectIndexSplitTests(unittest.TestCase):
    def test_collect_index_core_behaviors_when_openpyxl_missing(self):
        from converter.collect_index import collect_office_files_and_build_excel

        class Dummy:
            def __init__(self):
                self.config = {
                    "target_folder": "T",
                    "allowed_extensions": {
                        "word": [".docx"],
                        "excel": [".xlsx"],
                        "powerpoint": [".pptx"],
                    },
                }

        dummy = Dummy()
        out = collect_office_files_and_build_excel(
            dummy,
            has_openpyxl=False,
            workbook_cls=None,
            font_cls=None,
        )
        self.assertIsNone(out)

    def test_office_converter_collect_method_delegates_to_module(self):
        import office_converter as oc

        original = oc.collect_office_files_and_build_excel_impl
        try:
            seen = {}

            def _fake(converter, **kwargs):
                seen["converter"] = converter
                seen["kwargs"] = kwargs
                return "ok"

            oc.collect_office_files_and_build_excel_impl = _fake
            dummy = OfficeConverter.__new__(OfficeConverter)

            out = dummy.collect_office_files_and_build_excel()
            self.assertEqual("ok", out)
            self.assertIn("has_openpyxl", seen.get("kwargs", {}))
        finally:
            oc.collect_office_files_and_build_excel_impl = original


if __name__ == "__main__":
    unittest.main()
