import os
import tempfile
import unittest

from office_converter import OfficeConverter


class _FakeDocument:
    def __init__(self):
        self.ops = []
        self.saved = None

    def add_paragraph(self, text, style=None):
        self.ops.append(("p", text, style))

    def add_heading(self, text, level=1):
        self.ops.append(("h", text, level))

    def save(self, path):
        self.saved = path
        with open(path, "w", encoding="utf-8") as f:
            f.write("ok")


class ConverterMarkdownDocxExportSplitTests(unittest.TestCase):
    def test_markdown_docx_export_core_behaviors(self):
        from converter.markdown_docx_export import export_markdown_to_docx

        root = tempfile.mkdtemp(prefix="md_docx_")
        md = os.path.join(root, "a.md")
        out_docx = os.path.join(root, "a.docx")
        with open(md, "w", encoding="utf-8") as f:
            f.write(
                "# H1\n## H2\n### H3\n- item\n1. step\n\n```\ncode\n```\nplain\n"
            )

        self.assertRaises(
            RuntimeError,
            export_markdown_to_docx,
            md,
            out_docx,
            has_pydocx=False,
            document_cls=None,
        )

        export_markdown_to_docx(
            md,
            out_docx,
            has_pydocx=True,
            document_cls=_FakeDocument,
        )
        self.assertTrue(os.path.exists(out_docx))

        try:
            os.remove(md)
            os.remove(out_docx)
            os.rmdir(root)
        except Exception:
            pass

    def test_office_converter_export_markdown_to_docx_delegates_to_module(self):
        import office_converter as oc

        root = tempfile.mkdtemp(prefix="md_docx_delegate_")
        md = os.path.join(root, "a.md")
        out_docx = os.path.join(root, "a.docx")
        with open(md, "w", encoding="utf-8") as f:
            f.write("x")

        original = oc.export_markdown_to_docx_impl
        try:
            seen = {}

            def _fake(*args, **kwargs):
                seen["args"] = args
                seen["kwargs"] = kwargs
                return "ok"

            oc.export_markdown_to_docx_impl = _fake
            dummy = OfficeConverter.__new__(OfficeConverter)
            result = dummy._export_markdown_to_docx(md, out_docx)
            self.assertEqual(result, "ok")
            self.assertEqual(seen["args"][0], md)
            self.assertEqual(seen["args"][1], out_docx)
        finally:
            oc.export_markdown_to_docx_impl = original
            try:
                os.remove(md)
                os.rmdir(root)
            except Exception:
                pass


if __name__ == "__main__":
    unittest.main()
