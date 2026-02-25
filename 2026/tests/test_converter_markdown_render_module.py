import os
import tempfile
import unittest

from office_converter import OfficeConverter


class ConverterMarkdownRenderSplitTests(unittest.TestCase):
    def test_markdown_render_core_behaviors_without_bs4(self):
        from converter.markdown_render import render_html_to_markdown

        root = tempfile.mkdtemp(prefix="md_render_")
        html = os.path.join(root, "a.html")
        with open(html, "w", encoding="utf-8") as f:
            f.write("<html><body><h1>Title</h1><p>Hello</p></body></html>")

        out = render_html_to_markdown(
            html,
            has_bs4=False,
            beautifulsoup_cls=None,
            normalize_md_line_fn=lambda s: str(s or "").strip(),
            table_to_markdown_lines_fn=lambda _n: [],
        )
        self.assertEqual(out, "Title Hello")

    def test_office_converter_render_html_to_markdown_delegates_to_module(self):
        import office_converter as oc

        root = tempfile.mkdtemp(prefix="md_render_delegate_")
        html = os.path.join(root, "a.html")
        with open(html, "w", encoding="utf-8") as f:
            f.write("<p>x</p>")

        original = oc.render_html_to_markdown_impl
        try:
            seen = {}

            def _fake(html_path, **kwargs):
                seen["html_path"] = html_path
                seen["kwargs"] = kwargs
                return "md"

            oc.render_html_to_markdown_impl = _fake
            dummy = OfficeConverter.__new__(OfficeConverter)
            dummy._normalize_md_line = lambda s: s
            dummy._table_to_markdown_lines = lambda _n: []

            out = dummy._render_html_to_markdown(html)
            self.assertEqual(out, "md")
            self.assertEqual(seen.get("html_path"), html)
        finally:
            oc.render_html_to_markdown_impl = original

    def test_markdown_table_lines_and_converter_delegate(self):
        from converter.markdown_render import table_to_markdown_lines
        import office_converter as oc

        class _Cell:
            def __init__(self, text):
                self._text = text

            def get_text(self, *_a, **_k):
                return self._text

        class _Row:
            def __init__(self, cells):
                self._cells = cells

            def find_all(self, names):
                if names == ["th", "td"]:
                    return self._cells
                return []

        class _Table:
            def find_all(self, name):
                if name != "tr":
                    return []
                return [
                    _Row([_Cell("A|B"), _Cell("C")]),
                    _Row([_Cell("1"), _Cell("2")]),
                ]

        lines = table_to_markdown_lines(
            _Table(),
            normalize_md_line_fn=lambda s: str(s).strip(),
        )
        self.assertEqual(lines[0], "| A\\|B | C |")
        self.assertEqual(lines[1], "| --- | --- |")

        original = oc.table_to_markdown_lines_impl
        try:
            seen = {}

            def _fake(table_tag, **kwargs):
                seen["table_tag"] = table_tag
                seen["kwargs"] = kwargs
                return ["ok"]

            oc.table_to_markdown_lines_impl = _fake
            dummy = OfficeConverter.__new__(OfficeConverter)
            dummy._normalize_md_line = lambda s: s
            out = dummy._table_to_markdown_lines(object())
            self.assertEqual(out, ["ok"])
            self.assertIn("table_tag", seen)
        finally:
            oc.table_to_markdown_lines_impl = original


if __name__ == "__main__":
    unittest.main()
