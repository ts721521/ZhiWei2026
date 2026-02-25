import os
import tempfile
import unittest
from datetime import datetime

from office_converter import MERGE_MODE_ALL_IN_ONE, MODE_MERGE_ONLY, OfficeConverter


class ConverterMergeCandidatesSplitTests(unittest.TestCase):
    def test_merge_candidates_module_core_behaviors(self):
        from converter.merge_candidates import (
            build_markdown_merge_tasks,
            resolve_merge_scan_context,
            scan_candidates_by_ext,
        )

        root = tempfile.mkdtemp(prefix="merge_scan_")
        inc = os.path.join(root, "in")
        exc = os.path.join(root, "exclude")
        os.makedirs(inc, exist_ok=True)
        os.makedirs(exc, exist_ok=True)
        keep_pdf = os.path.join(inc, "a.pdf")
        skip_pdf = os.path.join(exc, "b.pdf")
        with open(keep_pdf, "w", encoding="utf-8") as f:
            f.write("a")
        with open(skip_pdf, "w", encoding="utf-8") as f:
            f.write("b")
        try:
            files = scan_candidates_by_ext(".pdf", [root], [exc])
            self.assertIn(keep_pdf, files)
            self.assertNotIn(skip_pdf, files)
        finally:
            for p in (keep_pdf, skip_pdf):
                try:
                    os.remove(p)
                except Exception:
                    pass
            for d in (inc, exc, root):
                try:
                    os.rmdir(d)
                except Exception:
                    pass

        md_files = [
            r"C:\x\Price_A.md",
            r"C:\x\Word_B.md",
            r"C:\x\misc.md",
        ]
        all_in_one = build_markdown_merge_tasks(
            md_files, MERGE_MODE_ALL_IN_ONE, now=datetime(2026, 2, 24, 1, 2, 3)
        )
        self.assertEqual(len(all_in_one), 1)
        self.assertTrue(all_in_one[0][0].startswith("Merged_All_20260224_010203"))

        by_cat = build_markdown_merge_tasks(md_files, "by_category", now=datetime(2026, 2, 24, 1, 2, 3))
        names = [n for n, _ in by_cat]
        self.assertIn("Merged_Price_20260224_010203_001.md", names)
        self.assertIn("Merged_Word_20260224_010203_001.md", names)
        self.assertIn("Merged_Markdown_20260224_010203_001.md", names)

        cfg = {"target_folder": r"C:\T", "merge_source": "source"}
        roots, excludes = resolve_merge_scan_context(
            run_mode=MODE_MERGE_ONLY,
            config=cfg,
            get_source_roots_fn=lambda: [r"C:\S"],
            failed_dir=r"C:\F",
            merge_output_dir=r"C:\M",
            mode_merge_only=MODE_MERGE_ONLY,
        )
        self.assertEqual([r"C:\S"], roots)
        self.assertIn(r"C:\T", excludes)

    def test_office_converter_merge_candidate_methods_delegate_to_module(self):
        from converter.merge_candidates import (
            build_markdown_merge_tasks,
            scan_candidates_by_ext,
        )

        root = tempfile.mkdtemp(prefix="merge_delegate_")
        src = os.path.join(root, "src")
        os.makedirs(src, exist_ok=True)
        pdf_path = os.path.join(src, "x.pdf")
        with open(pdf_path, "w", encoding="utf-8") as f:
            f.write("x")

        dummy = OfficeConverter.__new__(OfficeConverter)
        dummy.run_mode = MODE_MERGE_ONLY
        dummy.merge_mode = "by_category"
        dummy.config = {"merge_source": "source", "target_folder": os.path.join(root, "target")}
        dummy.failed_dir = os.path.join(root, "failed")
        dummy.merge_output_dir = os.path.join(root, "merged")
        dummy._get_source_roots = lambda: [src]

        try:
            self.assertEqual(
                dummy._scan_merge_candidates_by_ext(".pdf"),
                scan_candidates_by_ext(
                    ".pdf",
                    [src],
                    [dummy.failed_dir, dummy.merge_output_dir, dummy.config["target_folder"]],
                ),
            )
            md_files = [r"C:\x\Price_A.md", r"C:\x\free.md"]
            self.assertEqual(
                dummy._build_markdown_merge_tasks(md_files),
                build_markdown_merge_tasks(md_files, dummy.merge_mode),
            )
        finally:
            try:
                os.remove(pdf_path)
            except Exception:
                pass
            for d in (src, root):
                try:
                    os.rmdir(d)
                except Exception:
                    pass

    def test_office_converter_scan_merge_candidates_delegates_to_wrapper(self):
        import office_converter as oc

        original = oc.scan_merge_candidates_by_ext_for_converter
        try:
            seen = {}

            def _fake(converter, ext, **kwargs):
                seen["converter"] = converter
                seen["ext"] = ext
                seen["kwargs"] = kwargs
                return ["a.pdf"]

            oc.scan_merge_candidates_by_ext_for_converter = _fake
            dummy = OfficeConverter.__new__(OfficeConverter)
            out = dummy._scan_merge_candidates_by_ext(".pdf")
            self.assertEqual(["a.pdf"], out)
            self.assertIs(seen["converter"], dummy)
            self.assertEqual(".pdf", seen["ext"])
            self.assertEqual(oc.MODE_MERGE_ONLY, seen["kwargs"]["mode_merge_only"])
        finally:
            oc.scan_merge_candidates_by_ext_for_converter = original


if __name__ == "__main__":
    unittest.main()
