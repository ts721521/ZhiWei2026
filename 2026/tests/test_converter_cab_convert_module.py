import os
import tempfile
import unittest
from datetime import datetime
from pathlib import Path

from office_converter import OfficeConverter


class ConverterCabConvertSplitTests(unittest.TestCase):
    def test_cab_convert_core_behaviors(self):
        from converter.cab_convert import convert_cab_to_markdown

        root = tempfile.mkdtemp(prefix="cab_convert_")
        cab = os.path.join(root, "a.cab")
        html = os.path.join(root, "topic.html")
        out_md = os.path.join(root, "out.md")
        with open(cab, "wb") as f:
            f.write(b"cab")
        with open(html, "w", encoding="utf-8") as f:
            f.write("<h1>x</h1>")

        generated = []
        seen = {"append": None}

        md_path, rendered_count = convert_cab_to_markdown(
            cab,
            cab,
            has_bs4=True,
            temp_sandbox=root,
            uuid4_hex_fn=lambda: "u1",
            extract_cab_with_fallback_fn=lambda _c, d: os.makedirs(d, exist_ok=True),
            find_files_recursive_fn=lambda _r, _exts: [],
            extract_mshc_payload_fn=lambda _m, _c: None,
            parse_mshelp_topics_fn=lambda _root: [{"title": "T1", "file": html}],
            build_ai_output_path_from_source_fn=lambda _s, _k, _e: out_md,
            normalize_md_line_fn=lambda s: str(s).strip(),
            render_html_to_markdown_fn=lambda _f: "body",
            append_mshelp_record_fn=lambda s, m, c: seen.update({"append": (s, m, c)}),
            now_fn=lambda: datetime(2026, 2, 24, 19, 0, 0),
            generated_markdown_outputs=generated,
            log_warning_fn=lambda _msg: None,
        )

        self.assertEqual(md_path, out_md)
        self.assertEqual(rendered_count, 1)
        self.assertTrue(os.path.exists(out_md))
        self.assertEqual(generated, [out_md])
        self.assertEqual(seen["append"], (cab, out_md, 1))

    def test_office_converter_convert_cab_to_markdown_delegates_to_module(self):
        import office_converter as oc

        original = oc.convert_cab_to_markdown_impl
        try:
            seen = {}

            def _fake(cab_path, source_path_for_output, **kwargs):
                seen["args"] = (cab_path, source_path_for_output)
                seen["kwargs"] = kwargs
                return "x.md", 2

            oc.convert_cab_to_markdown_impl = _fake
            dummy = OfficeConverter.__new__(OfficeConverter)
            dummy.temp_sandbox = "tmp"
            dummy.generated_markdown_outputs = []
            dummy._extract_cab_with_fallback = lambda *_a, **_k: None
            dummy._find_files_recursive = lambda *_a, **_k: []
            dummy._extract_mshc_payload = lambda *_a, **_k: None
            dummy._parse_mshelp_topics = lambda *_a, **_k: []
            dummy._build_ai_output_path_from_source = lambda *_a, **_k: ""
            dummy._normalize_md_line = lambda s: s
            dummy._render_html_to_markdown = lambda *_a, **_k: ""
            dummy._append_mshelp_record = lambda *_a, **_k: None

            out = dummy._convert_cab_to_markdown("a.cab", "a.cab")
            self.assertEqual(out, ("x.md", 2))
            self.assertEqual(seen["args"], ("a.cab", "a.cab"))
        finally:
            oc.convert_cab_to_markdown_impl = original

    def test_cab_convert_module_has_no_bare_except_exception(self):
        mod_path = Path(__file__).resolve().parents[1] / "converter" / "cab_convert.py"
        text = mod_path.read_text(encoding="utf-8")
        self.assertNotIn("except Exception", text)


if __name__ == "__main__":
    unittest.main()
