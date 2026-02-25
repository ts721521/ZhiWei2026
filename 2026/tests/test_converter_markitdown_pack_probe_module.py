import os
import tempfile
import unittest
from pathlib import Path


class _FakeResult:
    def __init__(self, text_content):
        self.text_content = text_content


class _FakeMarkItDown:
    def convert(self, source_path):
        return _FakeResult(f"# converted\n\nsource={os.path.basename(source_path)}")


class ConverterMarkitdownPackProbeTests(unittest.TestCase):
    def test_markitdown_probe_core_behaviors(self):
        from converter.markitdown_pack_probe import run_markitdown_probe

        root = tempfile.mkdtemp(prefix="md_probe_")
        src = os.path.join(root, "a.docx")
        out = os.path.join(root, "out.md")
        with open(src, "w", encoding="utf-8") as f:
            f.write("x")

        result = run_markitdown_probe(src, out, markitdown_cls=_FakeMarkItDown)
        self.assertEqual(result.get("status"), "ok")
        self.assertTrue(os.path.exists(out))
        with open(out, "r", encoding="utf-8") as f:
            text = f.read()
        self.assertIn("converted", text)

    def test_markitdown_probe_reports_missing_input(self):
        from converter.markitdown_pack_probe import run_markitdown_probe

        result = run_markitdown_probe(
            "Z:/nope/a.docx",
            "out.md",
            markitdown_cls=_FakeMarkItDown,
        )
        self.assertEqual(result.get("status"), "missing_input")

    def test_markitdown_pack_probe_module_has_no_bare_except_exception(self):
        mod_path = Path(__file__).resolve().parents[1] / "converter" / "markitdown_pack_probe.py"
        text = mod_path.read_text(encoding="utf-8")
        self.assertNotIn("except Exception", text)


if __name__ == "__main__":
    unittest.main()
