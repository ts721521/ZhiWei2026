import unittest

from office_converter import OfficeConverter


class ConverterConvertThreadRuntimeSplitTests(unittest.TestCase):
    def test_convert_thread_runtime_core_behaviors(self):
        from converter.convert_thread_runtime import convert_logic_in_thread_for_converter

        class Dummy:
            def __init__(self):
                self.config = {"allowed_extensions": {"word": [".docx"]}}
                self.engine_type = "wps"
                self.content_strategy = "standard"
                self._convert_on_mac = lambda *_a, **_k: False
                self._get_local_app = lambda *_a, **_k: None
                self._safe_exec = lambda *_a, **_k: None
                self.scan_excel_content_in_thread = lambda *_a, **_k: False
                self._setup_excel_pages = lambda *_a, **_k: None
                self._should_reuse_office_app = lambda: False

        seen = {}

        def _fake(*args, **kwargs):
            seen["args"] = args
            seen["kwargs"] = kwargs
            return "ok"

        d = Dummy()
        out = convert_logic_in_thread_for_converter(
            d,
            "a.docx",
            "a.pdf",
            ".docx",
            {"x": 1},
            convert_logic_in_thread_fn=_fake,
            is_mac_fn=lambda: False,
            has_win32=True,
            engine_wps="wps",
            wd_format_pdf=17,
            xl_type_pdf=0,
            pp_save_as_pdf=0,
            pp_fixed_format_type_pdf=0,
            xl_pdf_save_as=0,
            xl_repair_file=0,
            strategy_standard="standard",
            strategy_price_only="price_only",
            pythoncom_module=object(),
            os_module=__import__("os"),
        )
        self.assertEqual(out, "ok")
        self.assertEqual(seen["args"][:3], ("a.docx", "a.pdf", ".docx"))
        self.assertEqual(seen["kwargs"]["engine_type"], "wps")

    def test_office_converter_convert_logic_delegates_to_runtime_module(self):
        import office_converter as oc

        original = oc.convert_logic_in_thread_for_converter_impl
        try:
            seen = {}

            def _fake(converter, *args, **kwargs):
                seen["converter"] = converter
                seen["args"] = args
                seen["kwargs"] = kwargs
                return "ok2"

            oc.convert_logic_in_thread_for_converter_impl = _fake
            dummy = OfficeConverter.__new__(OfficeConverter)
            out = dummy.convert_logic_in_thread("a.docx", "a.pdf", ".docx", {})
            self.assertEqual(out, "ok2")
            self.assertIs(seen["converter"], dummy)
            self.assertEqual(seen["args"][:3], ("a.docx", "a.pdf", ".docx"))
        finally:
            oc.convert_logic_in_thread_for_converter_impl = original


if __name__ == "__main__":
    unittest.main()
