import unittest
from pathlib import Path

from office_converter import OfficeConverter


class _FakeDoc:
    def __init__(self):
        self.exported = None
        self.closed = False

    def ExportAsFixedFormat(self, out, fmt):
        self.exported = (out, fmt)

    def Close(self, SaveChanges=False):
        self.closed = True


class _FakeDocuments:
    def __init__(self, doc):
        self._doc = doc

    def Open(self, *_args, **_kwargs):
        return self._doc


class _FakeApp:
    def __init__(self, doc):
        self.Documents = _FakeDocuments(doc)
        self.quit_called = False

    def Quit(self):
        self.quit_called = True


class _FakePythonCom:
    def __init__(self):
        self.uninit_count = 0

    def CoUninitialize(self):
        self.uninit_count += 1


class ConverterConvertThreadSplitTests(unittest.TestCase):
    def test_convert_thread_module_has_no_bare_except_exception(self):
        module_text = Path("converter/convert_thread.py").read_text(encoding="utf-8")
        self.assertNotIn("except Exception", module_text)

    def test_convert_thread_core_behaviors_word(self):
        from converter.convert_thread import convert_logic_in_thread

        doc = _FakeDoc()
        app = _FakeApp(doc)
        pycom = _FakePythonCom()

        convert_logic_in_thread(
            "a.docx",
            "a.pdf",
            ".docx",
            {},
            is_mac_fn=lambda: False,
            convert_on_mac_fn=lambda *_a, **_k: False,
            has_win32=True,
            allowed_extensions={"word": [".docx"], "excel": [], "powerpoint": []},
            get_local_app_fn=lambda _k: app,
            safe_exec_fn=lambda fn, *a, **k: fn(*a, **k),
            engine_type="wps",
            engine_wps="wps",
            wdFormatPDF=17,
            xlTypePDF=0,
            ppSaveAsPDF=0,
            ppFixedFormatTypePDF=0,
            xlPDF_SaveAs=0,
            xlRepairFile=0,
            content_strategy="standard",
            strategy_standard="standard",
            strategy_price_only="price_only",
            scan_excel_content_in_thread_fn=lambda _d: False,
            setup_excel_pages_fn=lambda _d: None,
            should_reuse_office_app_fn=lambda: False,
            pythoncom_module=pycom,
            os_module=__import__("os"),
        )

        self.assertEqual(doc.exported, ("a.pdf", 17))
        self.assertTrue(doc.closed)
        self.assertTrue(app.quit_called)
        self.assertEqual(pycom.uninit_count, 1)

    def test_office_converter_convert_logic_in_thread_delegates_to_module(self):
        import office_converter as oc

        original = oc.convert_logic_in_thread_impl
        try:
            seen = {}

            def _fake(file_source, sandbox_target_pdf, ext, result_context, **kwargs):
                seen["args"] = (file_source, sandbox_target_pdf, ext, result_context)
                seen["kwargs"] = kwargs
                return "ok"

            oc.convert_logic_in_thread_impl = _fake
            dummy = OfficeConverter.__new__(OfficeConverter)
            dummy._convert_on_mac = lambda *_a, **_k: False
            dummy.config = {"allowed_extensions": {}}
            dummy._get_local_app = lambda *_a, **_k: None
            dummy._safe_exec = lambda *_a, **_k: None
            dummy.engine_type = "wps"
            dummy.content_strategy = "standard"
            dummy.scan_excel_content_in_thread = lambda *_a, **_k: False
            dummy._setup_excel_pages = lambda *_a, **_k: None
            dummy._should_reuse_office_app = lambda: False

            out = dummy.convert_logic_in_thread("a.docx", "a.pdf", ".docx", {"k": "v"})
            self.assertEqual(out, "ok")
            self.assertEqual(seen["args"], ("a.docx", "a.pdf", ".docx", {"k": "v"}))
        finally:
            oc.convert_logic_in_thread_impl = original


if __name__ == "__main__":
    unittest.main()
