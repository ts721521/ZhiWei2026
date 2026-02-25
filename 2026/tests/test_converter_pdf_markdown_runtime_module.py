import unittest
from pathlib import Path

from office_converter import OfficeConverter


class ConverterPdfMarkdownRuntimeSplitTests(unittest.TestCase):
    def test_pdf_markdown_runtime_core_behaviors(self):
        from converter.pdf_markdown_runtime import export_pdf_markdown_for_converter

        class Dummy:
            def __init__(self):
                self.config = {"short_id_prefix": "ZW-"}
                self.trace_short_id_taken = set()
                self.generated_markdown_outputs = []
                self.markdown_quality_records = []
                self._build_ai_output_path = lambda *_a, **_k: "x.md"
                self._build_ai_output_path_from_source = lambda *_a, **_k: "x.md"
                self._collect_margin_candidates = lambda _p: (set(), set())
                self._clean_markdown_page_lines = lambda _t, _h, _f: ([], {})
                self._render_markdown_blocks = lambda _l, **_k: ([], 0)
                self._compute_md5 = lambda _p: "a" * 32
                self._build_short_id = lambda _m, _t: "AAAAAAAA"

        seen = {}

        def _fake_export(pdf_path, **kwargs):
            seen["pdf_path"] = pdf_path
            seen["kwargs"] = kwargs
            return "out.md", {"ok": 1}

        d = Dummy()
        out = export_pdf_markdown_for_converter(
            d,
            "a.pdf",
            source_path_hint="a.pdf",
            export_pdf_markdown_fn=_fake_export,
            has_pypdf=True,
            pdf_reader_cls=object,
            now_fn=lambda: None,
            log_info_fn=lambda _m: None,
            log_error_fn=lambda _m: None,
        )
        self.assertEqual(out, "out.md")
        self.assertEqual(d.generated_markdown_outputs, ["out.md"])
        self.assertEqual(d.markdown_quality_records, [{"ok": 1}])
        self.assertEqual(seen["pdf_path"], "a.pdf")

    def test_office_converter_export_pdf_markdown_delegates_to_runtime_module(self):
        import office_converter as oc

        original = oc.export_pdf_markdown_for_converter_impl
        try:
            seen = {}

            def _fake(converter, pdf_path, **kwargs):
                seen["converter"] = converter
                seen["pdf_path"] = pdf_path
                seen["kwargs"] = kwargs
                return "ok.md"

            oc.export_pdf_markdown_for_converter_impl = _fake
            dummy = OfficeConverter.__new__(OfficeConverter)
            out = dummy._export_pdf_markdown("a.pdf", source_path_hint="s.pdf")
            self.assertEqual(out, "ok.md")
            self.assertIs(seen["converter"], dummy)
            self.assertEqual(seen["pdf_path"], "a.pdf")
            self.assertEqual(seen["kwargs"]["source_path_hint"], "s.pdf")
        finally:
            oc.export_pdf_markdown_for_converter_impl = original

    def test_pdf_markdown_runtime_module_has_no_bare_except_exception(self):
        mod_path = Path(__file__).resolve().parents[1] / "converter" / "pdf_markdown_runtime.py"
        text = mod_path.read_text(encoding="utf-8")
        self.assertNotIn("except Exception", text)


if __name__ == "__main__":
    unittest.main()
