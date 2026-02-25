import os
import tempfile
import unittest
from datetime import datetime
from pathlib import Path

from office_converter import OfficeConverter


class _FakePage:
    def __init__(self, text):
        self._text = text

    def extract_text(self):
        return self._text


class _FakeReader:
    def __init__(self, _path):
        self.pages = [_FakePage("header\nbody a\n1"), _FakePage("header\nbody b\n2")]


class ConverterPdfMarkdownExportSplitTests(unittest.TestCase):
    def test_pdf_markdown_export_core_behaviors(self):
        from converter.pdf_markdown_export import export_pdf_markdown

        root = tempfile.mkdtemp(prefix="pdf_md_export_")
        pdf = os.path.join(root, "a.pdf")
        with open(pdf, "wb") as f:
            f.write(b"pdf")

        out_md_path = os.path.join(root, "out.md")
        md_path, rec = export_pdf_markdown(
            pdf,
            config={
                "output_enable_md": True,
                "markdown_strip_header_footer": True,
                "markdown_structured_headings": True,
            },
            has_pypdf=True,
            pdf_reader_cls=_FakeReader,
            source_path_hint=None,
            build_ai_output_path_fn=lambda _p, _folder, _ext: out_md_path,
            build_ai_output_path_from_source_fn=lambda _p, _folder, _ext: out_md_path,
            collect_margin_candidates_fn=lambda _pages: ({"header"}, set()),
            clean_markdown_page_lines_fn=lambda text, _h, _f: (
                [ln for ln in text.splitlines() if ln != "header"],
                {
                    "removed_header_lines": 1,
                    "removed_footer_lines": 0,
                    "removed_page_number_lines": 0,
                },
            ),
            render_markdown_blocks_fn=lambda lines, structured_headings=True: (
                [" ".join(lines)],
                1 if structured_headings else 0,
            ),
            compute_md5_fn=lambda _p: "a" * 32,
            build_short_id_fn=lambda _md5, _taken: "AAAAAAAA",
            short_id_taken_ids=set(),
            short_id_prefix="ZW-",
            now_fn=lambda: datetime(2026, 2, 24, 18, 0, 0),
        )

        self.assertEqual(md_path, out_md_path)
        self.assertTrue(os.path.exists(out_md_path))
        self.assertEqual(rec.get("page_count"), 2)
        self.assertEqual(rec.get("non_empty_page_count"), 2)
        self.assertEqual(rec.get("removed_header_lines"), 2)
        self.assertEqual(rec.get("heading_count"), 2)
        self.assertTrue(rec.get("source_short_id", "").startswith("ZW-"))
        with open(out_md_path, "r", encoding="utf-8") as f:
            text = f.read()
        self.assertIn("source_short_id: ZW-", text)

    def test_office_converter_export_pdf_markdown_delegates_to_module(self):
        import office_converter as oc

        root = tempfile.mkdtemp(prefix="pdf_md_delegate_")
        pdf = os.path.join(root, "a.pdf")
        with open(pdf, "wb") as f:
            f.write(b"pdf")

        original = oc.export_pdf_markdown_impl
        try:
            seen = {}

            def _fake(pdf_path, **kwargs):
                seen["pdf_path"] = pdf_path
                seen["kwargs"] = kwargs
                return "x.md", {"k": "v"}

            oc.export_pdf_markdown_impl = _fake
            dummy = OfficeConverter.__new__(OfficeConverter)
            dummy.config = {}
            dummy.generated_markdown_outputs = []
            dummy.markdown_quality_records = []
            dummy._build_ai_output_path = lambda *_a, **_k: "out.md"
            dummy._build_ai_output_path_from_source = lambda *_a, **_k: "out.md"
            dummy._collect_margin_candidates = lambda _p: (set(), set())
            dummy._clean_markdown_page_lines = lambda _t, _h, _f: ([], {})
            dummy._render_markdown_blocks = lambda _l, structured_headings=True: ([], 0)
            dummy._compute_md5 = lambda _p: "a" * 32
            dummy._build_short_id = lambda _md5, _taken: "AAAAAAAA"
            dummy.trace_short_id_taken = set()

            out = dummy._export_pdf_markdown(pdf, source_path_hint=pdf)
            self.assertEqual(out, "x.md")
            self.assertEqual(dummy.generated_markdown_outputs, ["x.md"])
            self.assertEqual(dummy.markdown_quality_records, [{"k": "v"}])
            self.assertEqual(seen.get("pdf_path"), pdf)
        finally:
            oc.export_pdf_markdown_impl = original

    def test_pdf_markdown_export_module_has_no_bare_except_exception(self):
        mod_path = Path(__file__).resolve().parents[1] / "converter" / "pdf_markdown_export.py"
        text = mod_path.read_text(encoding="utf-8")
        self.assertNotIn("except Exception", text)


if __name__ == "__main__":
    unittest.main()
