import os
import tempfile
import unittest

from office_converter import OfficeConverter


class _FakeCanvas:
    def __init__(self, out_path, pagesize=None):
        self.out_path = out_path
        self.pagesize = pagesize
        self.calls = []

    def showPage(self):
        self.calls.append(("showPage",))

    def drawString(self, x, y, text):
        self.calls.append(("drawString", x, y, text))

    def save(self):
        with open(self.out_path, "w", encoding="utf-8") as f:
            f.write("ok")


class ConverterMarkdownPdfExportSplitTests(unittest.TestCase):
    def test_markdown_pdf_export_core_behaviors(self):
        from converter.markdown_pdf_export import export_markdown_to_pdf

        root = tempfile.mkdtemp(prefix="md_pdf_")
        md = os.path.join(root, "a.md")
        out_pdf = os.path.join(root, "a.pdf")
        with open(md, "w", encoding="utf-8") as f:
            f.write("line1\nline2")

        self.assertRaises(
            RuntimeError,
            export_markdown_to_pdf,
            md,
            out_pdf,
            has_reportlab=False,
            canvas_cls=None,
            page_size=None,
        )

        export_markdown_to_pdf(
            md,
            out_pdf,
            has_reportlab=True,
            canvas_cls=_FakeCanvas,
            page_size=(595, 842),
            wrap_plain_text_for_pdf_fn=lambda text, width=100: [text],
        )
        self.assertTrue(os.path.exists(out_pdf))

        try:
            os.remove(md)
            os.remove(out_pdf)
            os.rmdir(root)
        except Exception:
            pass

    def test_office_converter_export_markdown_to_pdf_delegates_to_module(self):
        import office_converter as oc

        root = tempfile.mkdtemp(prefix="md_pdf_delegate_")
        md = os.path.join(root, "a.md")
        out_pdf = os.path.join(root, "a.pdf")
        with open(md, "w", encoding="utf-8") as f:
            f.write("x")

        original = oc.export_markdown_to_pdf_impl
        try:
            seen = {}

            def _fake(*args, **kwargs):
                seen["args"] = args
                seen["kwargs"] = kwargs
                return "ok"

            oc.export_markdown_to_pdf_impl = _fake
            dummy = OfficeConverter.__new__(OfficeConverter)
            dummy._wrap_plain_text_for_pdf = lambda text, width=100: [text]
            result = dummy._export_markdown_to_pdf(md, out_pdf)
            self.assertEqual(result, "ok")
            self.assertEqual(seen["args"][0], md)
            self.assertEqual(seen["args"][1], out_pdf)
        finally:
            oc.export_markdown_to_pdf_impl = original
            try:
                os.remove(md)
                os.rmdir(root)
            except Exception:
                pass


if __name__ == "__main__":
    unittest.main()
