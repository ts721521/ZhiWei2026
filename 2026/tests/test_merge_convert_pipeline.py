import json
import unittest
import uuid
from pathlib import Path

from office_converter import (
    MERGE_CONVERT_SUBMODE_MERGE_ONLY,
    MERGE_CONVERT_SUBMODE_PDF_TO_MD,
    MODE_MERGE_ONLY,
    OfficeConverter,
)


class _FakeConverter(OfficeConverter):
    def __init__(self, config_path, scan_map=None):
        super().__init__(config_path, interactive=False)
        self.run_mode = MODE_MERGE_ONLY
        self.scan_map = scan_map or {}
        self.calls = []
        self.generated_markdown_outputs = []

    def _scan_merge_candidates_by_ext(self, ext):
        return list(self.scan_map.get(str(ext).lower(), []))

    def merge_pdfs(self):
        self.calls.append(("merge_pdfs", None))
        return ["merged_pdf_out.pdf"]

    def merge_markdowns(self, candidates=None):
        self.calls.append(("merge_markdowns", list(candidates or [])))
        return ["merged_md_out.md"]

    def _export_pdf_markdown(self, pdf_path, source_path_hint=None):
        out_path = Path(f"{pdf_path}.md")
        if not out_path.is_absolute():
            out_path = Path(".tmp_test_work") / out_path
        out_path.parent.mkdir(parents=True, exist_ok=True)
        out_path.write_text("md", encoding="utf-8")
        out = str(out_path)
        self.generated_markdown_outputs.append(out)
        self.calls.append(("pdf_to_md", pdf_path))
        return out


class MergeConvertPipelineTests(unittest.TestCase):
    def _make_converter(self, overrides=None, scan_map=None):
        overrides = overrides or {}
        local_tmp_root = Path(".tmp_test_work")
        local_tmp_root.mkdir(parents=True, exist_ok=True)
        root = local_tmp_root / f"case_{uuid.uuid4().hex}"
        root.mkdir(parents=True, exist_ok=True)
        self.addCleanup(lambda: __import__("shutil").rmtree(root, ignore_errors=True))
        source = root / "src"
        target = root / "out"
        source.mkdir(parents=True, exist_ok=True)
        target.mkdir(parents=True, exist_ok=True)
        cfg_path = root / "config.json"
        cfg = {
            "source_folder": str(source),
            "target_folder": str(target),
            "run_mode": MODE_MERGE_ONLY,
            "output_enable_pdf": True,
            "output_enable_md": True,
            "output_enable_merged": True,
            "output_enable_independent": False,
            "merge_convert_submode": MERGE_CONVERT_SUBMODE_MERGE_ONLY,
            "enable_merge": True,
        }
        cfg.update(overrides)
        cfg_path.write_text(
            json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        return _FakeConverter(str(cfg_path), scan_map=scan_map)

    def test_merge_only_calls_pdf_and_md_merge(self):
        conv = self._make_converter(scan_map={".md": ["a.md"]})
        out = conv._run_merge_mode_pipeline(batch_results=[])
        self.assertIn("merged_pdf_out.pdf", out)
        self.assertIn("merged_md_out.md", out)
        self.assertIn(("merge_pdfs", None), conv.calls)
        self.assertIn(("merge_markdowns", ["a.md"]), conv.calls)

    def test_merge_only_missing_md_can_abort(self):
        conv = self._make_converter(scan_map={".md": []})
        conv._confirm_continue_missing_md_merge = lambda: False
        with self.assertRaises(RuntimeError):
            conv._run_merge_mode_pipeline(batch_results=[])

    def test_pdf_to_md_without_merged_only_converts(self):
        conv = self._make_converter(
            overrides={
                "merge_convert_submode": MERGE_CONVERT_SUBMODE_PDF_TO_MD,
                "output_enable_merged": False,
                "output_enable_independent": True,
                "output_enable_pdf": False,
                "output_enable_md": True,
            },
            scan_map={".pdf": ["x.pdf", "y.pdf"]},
        )
        out = conv._run_merge_mode_pipeline(batch_results=[])
        self.assertEqual([], out)
        self.assertIn(("pdf_to_md", "x.pdf"), conv.calls)
        self.assertIn(("pdf_to_md", "y.pdf"), conv.calls)
        self.assertNotIn(("merge_pdfs", None), conv.calls)
        self.assertFalse(any(c[0] == "merge_markdowns" for c in conv.calls))

    def test_pdf_to_md_with_merged_md_merges_generated_md(self):
        conv = self._make_converter(
            overrides={
                "merge_convert_submode": MERGE_CONVERT_SUBMODE_PDF_TO_MD,
                "output_enable_merged": True,
                "output_enable_pdf": False,
                "output_enable_md": True,
            },
            scan_map={".pdf": ["x.pdf"]},
        )
        out = conv._run_merge_mode_pipeline(batch_results=[])
        self.assertIn("merged_md_out.md", out)
        md_calls = [c for c in conv.calls if c[0] == "merge_markdowns"]
        self.assertEqual(1, len(md_calls))
        self.assertEqual(1, len(md_calls[0][1]))
        self.assertTrue(str(md_calls[0][1][0]).endswith("x.pdf.md"))
        self.assertNotIn(("merge_pdfs", None), conv.calls)


if __name__ == "__main__":
    unittest.main()
