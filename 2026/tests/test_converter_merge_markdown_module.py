import os
import json
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
            config = {"markdown_max_size_mb": 80}

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

    def test_merge_markdowns_splits_when_size_limit_exceeded(self):
        from converter.merge_markdown import merge_markdowns

        root = tempfile.mkdtemp(prefix="merge_md_split_")
        out_dir = os.path.join(root, "out")
        os.makedirs(out_dir, exist_ok=True)
        md_a = os.path.join(root, "a.md")
        md_b = os.path.join(root, "b.md")
        with open(md_a, "w", encoding="utf-8") as f:
            f.write("A" * (700 * 1024))
        with open(md_b, "w", encoding="utf-8") as f:
            f.write("B" * (700 * 1024))

        class Dummy:
            merge_output_dir = out_dir
            generated_merge_markdown_outputs = []
            config = {"markdown_max_size_mb": 1}

            def _scan_merge_candidates_by_ext(self, ext):
                self.ext = ext
                return [md_a, md_b]

            def _build_markdown_merge_tasks(self, md_files):
                return [("Merged_All.md", md_files)]

        dummy = Dummy()
        generated = merge_markdowns(dummy)
        self.assertEqual(len(generated), 2)
        self.assertTrue(any(path.endswith("_part_001.md") for path in generated))
        self.assertTrue(any(path.endswith("_part_002.md") for path in generated))

    def test_merge_markdowns_generates_image_manifest_and_unique_assets(self):
        from converter.merge_markdown import merge_markdowns

        root = tempfile.mkdtemp(prefix="merge_md_img_")
        out_dir = os.path.join(root, "out")
        os.makedirs(out_dir, exist_ok=True)

        src1 = os.path.join(root, "src1")
        src2 = os.path.join(root, "src2")
        os.makedirs(src1, exist_ok=True)
        os.makedirs(src2, exist_ok=True)

        with open(os.path.join(src1, "image1.png"), "wb") as f:
            f.write(b"img-one")
        with open(os.path.join(src2, "image1.png"), "wb") as f:
            f.write(b"img-two")

        doc1 = os.path.join(src1, "doc1.docx")
        doc2 = os.path.join(src2, "doc2.docx")
        with open(doc1, "wb") as f:
            f.write(b"x")
        with open(doc2, "wb") as f:
            f.write(b"y")

        md_a = os.path.join(src1, "a.md")
        md_b = os.path.join(src2, "b.md")
        with open(md_a, "w", encoding="utf-8") as f:
            f.write("---\nsource_file: " + doc1 + "\n---\n\n![a](image1.png)\n")
        with open(md_b, "w", encoding="utf-8") as f:
            f.write("---\nsource_file: " + doc2 + "\n---\n\n![b](image1.png)\n")

        class Dummy:
            merge_output_dir = out_dir
            generated_merge_markdown_outputs = []
            generated_markdown_manifest_outputs = []
            config = {"markdown_max_size_mb": 80, "enable_markdown_image_manifest": True}
            merge_index_records = [
                {
                    "source_abspath": doc1,
                    "source_filename": "doc1.docx",
                    "merged_pdf_name": "Merged_All.pdf",
                    "merged_pdf_path": os.path.join(out_dir, "Merged_All.pdf"),
                    "start_page_1based": 5,
                    "end_page_1based": 9,
                },
                {
                    "source_abspath": doc2,
                    "source_filename": "doc2.docx",
                    "merged_pdf_name": "Merged_All.pdf",
                    "merged_pdf_path": os.path.join(out_dir, "Merged_All.pdf"),
                    "start_page_1based": 10,
                    "end_page_1based": 12,
                },
            ]

            def _scan_merge_candidates_by_ext(self, ext):
                self.ext = ext
                return [md_a, md_b]

            def _build_markdown_merge_tasks(self, md_files):
                return [("Merged_All.md", md_files)]

        dummy = Dummy()
        generated = merge_markdowns(dummy)
        self.assertEqual(1, len(generated))

        merged_md = generated[0]
        with open(merged_md, "r", encoding="utf-8") as f:
            text = f.read()
        self.assertIn("IMG_REF:doc_001_img_001", text)
        self.assertIn("IMG_REF:doc_002_img_001", text)

        self.assertEqual(1, len(dummy.generated_markdown_manifest_outputs))
        manifest_path = dummy.generated_markdown_manifest_outputs[0]
        self.assertTrue(os.path.exists(manifest_path))
        with open(manifest_path, "r", encoding="utf-8") as f:
            payload = json.load(f)
        self.assertEqual(2, payload.get("record_count"))
        rels = [rec.get("copied_asset_relpath", "") for rec in payload.get("records", [])]
        self.assertEqual(2, len([x for x in rels if x]))
        self.assertEqual(2, len(set(rels)))
        hints = [rec.get("merged_pdf_page_hint_1based") for rec in payload.get("records", [])]
        self.assertIn(5, hints)
        self.assertIn(10, hints)

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
