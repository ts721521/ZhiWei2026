import unittest
from pathlib import Path

from office_converter import OfficeConverter


class _FakeSelection:
    def __init__(self):
        class _F:
            Name = ""
            Size = 0
            Bold = False

        class _P:
            Alignment = 0
            LineSpacingRule = 0
            LineSpacing = 0

        self.Font = _F()
        self.ParagraphFormat = _P()
        self.texts = []
        self.breaks = []

    def TypeText(self, text):
        self.texts.append(text)

    def InsertBreak(self, t):
        self.breaks.append(t)


class _FakeDoc:
    def __init__(self):
        class _PageSetup:
            PaperSize = None
            TopMargin = None
            BottomMargin = None
            LeftMargin = None
            RightMargin = None

        self.PageSetup = _PageSetup()
        self.exported = None
        self.closed = False

    def ExportAsFixedFormat(self, **kwargs):
        self.exported = kwargs

    def Close(self, SaveChanges=0):
        self.closed = True


class _FakeDocuments:
    def __init__(self, doc):
        self._doc = doc

    def Add(self):
        return self._doc


class _FakeWordApp:
    def __init__(self):
        self.doc = _FakeDoc()
        self.Selection = _FakeSelection()
        self.Documents = _FakeDocuments(self.doc)
        self.Visible = True


class ConverterMergeIndexDocSplitTests(unittest.TestCase):
    def test_merge_index_doc_core_behaviors(self):
        from converter.merge_index_doc import create_index_doc_and_convert

        app = _FakeWordApp()
        out = create_index_doc_and_convert(
            app,
            ["a.docx", "b.docx", "c" * 60 + ".docx"],
            "My Title",
            temp_sandbox="C:\\tmp",
            uuid4_hex_fn=lambda: "u1",
            log_error_fn=lambda _msg: None,
        )
        self.assertEqual(out, "C:\\tmp\\index_u1.pdf")
        self.assertFalse(app.Visible)
        self.assertTrue(app.doc.closed)
        self.assertEqual(app.doc.exported.get("OutputFileName"), "C:\\tmp\\index_u1.pdf")
        merged_text = "".join(app.Selection.texts)
        self.assertIn("My Title", merged_text)
        self.assertIn("...", merged_text)

    def test_office_converter_create_index_doc_delegates_to_module(self):
        import office_converter as oc

        original = oc.create_index_doc_and_convert_impl
        try:
            seen = {}

            def _fake(word_app, file_list, title, **kwargs):
                seen["args"] = (word_app, file_list, title)
                seen["kwargs"] = kwargs
                return "x.pdf"

            oc.create_index_doc_and_convert_impl = _fake
            dummy = OfficeConverter.__new__(OfficeConverter)
            dummy.temp_sandbox = "tmp"

            out = dummy._create_index_doc_and_convert("app", ["a"], "title")
            self.assertEqual(out, "x.pdf")
            self.assertEqual(seen["args"], ("app", ["a"], "title"))
            self.assertEqual(seen["kwargs"]["temp_sandbox"], "tmp")
        finally:
            oc.create_index_doc_and_convert_impl = original

    def test_merge_index_doc_module_has_no_bare_except_exception(self):
        mod_path = Path(__file__).resolve().parents[1] / "converter" / "merge_index_doc.py"
        text = mod_path.read_text(encoding="utf-8")
        self.assertNotIn("except Exception", text)


if __name__ == "__main__":
    unittest.main()
