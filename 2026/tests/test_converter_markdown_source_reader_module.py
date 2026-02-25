import tempfile
import unittest

from office_converter import OfficeConverter


class _MarkResult:
    def __init__(self, text_content=None, markdown=None, text=None):
        self.text_content = text_content
        self.markdown = markdown
        self.text = text


class _MarkItDownFake:
    def __init__(self, result):
        self._result = result

    def convert(self, _path):
        return self._result


class ConverterMarkdownSourceReaderSplitTests(unittest.TestCase):
    def test_markdown_source_reader_core_behaviors(self):
        from converter.markdown_source_reader import convert_source_to_markdown_text

        with tempfile.NamedTemporaryFile("w", suffix=".txt", delete=False, encoding="utf-8") as fh:
            fh.write("fallback text")
            src = fh.name

        class _MDCls:
            def __init__(self):
                self._impl = _MarkItDownFake(_MarkResult(text_content="from_md"))

            def convert(self, path):
                return self._impl.convert(path)

        self.assertEqual(
            "from_md",
            convert_source_to_markdown_text(
                src,
                has_markitdown=True,
                markitdown_cls=_MDCls,
            ),
        )
        self.assertEqual(
            "fallback text",
            convert_source_to_markdown_text(
                src,
                has_markitdown=False,
                markitdown_cls=None,
            ),
        )

    def test_office_converter_method_delegates_to_module(self):
        import office_converter as oc

        dummy = OfficeConverter.__new__(OfficeConverter)
        original = oc.convert_source_to_markdown_text_impl
        try:
            oc.convert_source_to_markdown_text_impl = lambda *args, **kwargs: "ok"
            self.assertEqual("ok", dummy._convert_source_to_markdown_text("x.txt"))
        finally:
            oc.convert_source_to_markdown_text_impl = original

    def test_module_has_no_bare_except_exception(self):
        with open("converter/markdown_source_reader.py", "r", encoding="utf-8") as f:
            src = f.read()
        self.assertNotIn("except Exception", src)


if __name__ == "__main__":
    unittest.main()
