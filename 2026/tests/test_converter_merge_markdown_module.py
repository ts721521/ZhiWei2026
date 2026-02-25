import os
import tempfile
import unittest
from pathlib import Path

from office_converter import OfficeConverter


class ConverterMergeMarkdownSplitTests(unittest.TestCase):
    def test_merge_markdown_module_has_no_bare_except_exception(self):
        module_text = Path("converter/merge_markdown.py").read_text(encoding="utf-8")
        self.assertNotIn("except Exception", module_text)

    def test_merge_markdowns_module_core_behaviors(self):
        from converter.merge_markdown import merge_markdowns

        root = tempfile.mkdtemp(prefix="merge_md_core_")
        out_dir = os.path.join(root, "out")
        os.makedirs(out_dir, exist_ok=True)
        md_a = os.path.join(root, "a.md")
        md_b = os.path.join(root, "b.md")
        with open(md_a, "w", encoding="utf-8") as f:
            f.write("# A\n")
        with open(md_b, "w", encoding="utf-8") as f:
            f.write("# B\n")

        class Dummy:
            merge_output_dir = out_dir
            generated_merge_markdown_outputs = []

            def _scan_merge_candidates_by_ext(self, ext):
                self.ext = ext
                return [md_a, md_b]

            def _build_markdown_merge_tasks(self, md_files):
                return [("Merged_All.md", md_files)]

        dummy = Dummy()
        generated = merge_markdowns(dummy)
        self.assertEqual(len(generated), 1)
        self.assertTrue(generated[0].endswith("Merged_All.md"))
        self.assertEqual(dummy.generated_merge_markdown_outputs, generated)
        with open(generated[0], "r", encoding="utf-8") as f:
            content = f.read()
        self.assertIn("## Source Map", content)
        self.assertIn("### [1] a.md", content)

    def test_office_converter_merge_markdowns_delegates_to_module(self):
        import office_converter as oc

        original = oc.merge_markdowns_impl
        try:
            seen = {}

            def _fake(converter, **kwargs):
                seen["converter"] = converter
                seen["kwargs"] = kwargs
                return ["ok.md"]

            oc.merge_markdowns_impl = _fake
            dummy = OfficeConverter.__new__(OfficeConverter)
            out = dummy.merge_markdowns(candidates=["x.md"])
            self.assertEqual(out, ["ok.md"])
            self.assertIs(seen["converter"], dummy)
            self.assertEqual(seen["kwargs"]["candidates"], ["x.md"])
        finally:
            oc.merge_markdowns_impl = original


if __name__ == "__main__":
    unittest.main()
