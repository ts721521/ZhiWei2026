import os
import tempfile
import unittest
from datetime import datetime

from office_converter import OfficeConverter


class ConverterPromptWrapperSplitTests(unittest.TestCase):
    def test_prompt_wrapper_core_behaviors(self):
        from converter.prompt_wrapper import (
            collect_prompt_ready_candidates,
            write_prompt_ready,
            write_prompt_ready_for_converter,
        )

        root = tempfile.mkdtemp(prefix="prompt_ready_")
        md = os.path.join(root, "a.md")
        with open(md, "w", encoding="utf-8") as f:
            f.write("# A\ncontent\n")

        out = write_prompt_ready(
            config={"enable_prompt_wrapper": True, "prompt_template_type": "new_solution"},
            target_folder=root,
            candidate_markdown_paths=[md],
            now_fn=lambda: datetime(2026, 2, 24, 22, 0, 0),
        )
        self.assertTrue(out)
        self.assertTrue(os.path.exists(out))
        with open(out, "r", encoding="utf-8") as f:
            text = f.read()
        self.assertIn("generated_at:", text)
        self.assertIn("template_type:", text)
        self.assertIn("a.md", text)

        self.assertEqual(
            ["a.md", "b.md", "c.md"],
            collect_prompt_ready_candidates(["a.md"], ["b.md"], ["c.md"]),
        )

        dummy = OfficeConverter.__new__(OfficeConverter)
        dummy.config = {"target_folder": root, "enable_prompt_wrapper": True}
        dummy.generated_fast_md_outputs = [md]
        dummy.generated_merge_markdown_outputs = []
        dummy.generated_markdown_outputs = []
        dummy.generated_prompt_outputs = []
        dummy.prompt_ready_path = None
        out2 = write_prompt_ready_for_converter(
            dummy,
            now_fn=lambda: datetime(2026, 2, 24, 22, 0, 0),
            log_info_fn=lambda *_a, **_k: None,
        )
        self.assertEqual(out, out2)
        self.assertEqual(dummy.prompt_ready_path, out2)
        self.assertEqual(dummy.generated_prompt_outputs, [out2])

    def test_office_converter_prompt_ready_delegate(self):
        import office_converter as oc

        original = oc.write_prompt_ready_for_converter
        try:
            seen = {}

            def _fake(converter, **kwargs):
                seen["converter"] = converter
                seen["kwargs"] = kwargs
                converter.prompt_ready_path = "Prompt_Ready.txt"
                converter.generated_prompt_outputs = ["Prompt_Ready.txt"]
                return "Prompt_Ready.txt"

            oc.write_prompt_ready_for_converter = _fake
            dummy = OfficeConverter.__new__(OfficeConverter)
            dummy.config = {"target_folder": "T", "enable_prompt_wrapper": True}
            dummy.generated_fast_md_outputs = ["a.md"]
            dummy.generated_merge_markdown_outputs = []
            dummy.generated_markdown_outputs = []
            dummy.generated_prompt_outputs = []
            dummy.prompt_ready_path = None

            out = dummy._write_prompt_ready()
            self.assertEqual(out, "Prompt_Ready.txt")
            self.assertEqual(dummy.prompt_ready_path, "Prompt_Ready.txt")
            self.assertEqual(dummy.generated_prompt_outputs, ["Prompt_Ready.txt"])
            self.assertIs(seen["converter"], dummy)
        finally:
            oc.write_prompt_ready_for_converter = original


if __name__ == "__main__":
    unittest.main()
