import os
import tempfile
import unittest
from datetime import datetime
from pathlib import Path

from office_converter import MERGE_MODE_ALL_IN_ONE, MODE_MERGE_ONLY, OfficeConverter


class ConverterMergeTasksSplitTests(unittest.TestCase):
    def test_merge_tasks_core_behaviors(self):
        from converter.merge_tasks import get_merge_tasks

        root = tempfile.mkdtemp(prefix="merge_tasks_")
        src = os.path.join(root, "src")
        target = os.path.join(root, "target")
        failed = os.path.join(target, "_FAILED_FILES")
        merged = os.path.join(target, "_MERGED")
        os.makedirs(src, exist_ok=True)
        os.makedirs(target, exist_ok=True)
        os.makedirs(failed, exist_ok=True)
        os.makedirs(merged, exist_ok=True)

        p1 = os.path.join(src, "Price_A.pdf")
        p2 = os.path.join(src, "Price_B.pdf")
        p3 = os.path.join(src, "Word_A.pdf")
        skip = os.path.join(failed, "Price_SKIP.pdf")
        for p in (p1, p2, p3, skip):
            with open(p, "wb") as f:
                f.write(b"x")

        now = datetime(2026, 2, 24, 10, 20, 30)
        tasks = get_merge_tasks(
            run_mode=MODE_MERGE_ONLY,
            merge_source="source",
            target_folder=target,
            get_source_roots_fn=lambda: [src, failed],
            failed_dir=failed,
            merge_output_dir=merged,
            merge_mode="by_category",
            merge_mode_all_in_one=MERGE_MODE_ALL_IN_ONE,
            mode_merge_only=MODE_MERGE_ONLY,
            merge_filename_pattern="Merged_{category}_{timestamp}_{idx}",
            max_merge_size_mb=0.000001,
            now_fn=lambda: now,
            format_merge_filename_fn=lambda pattern, category, idx, now: (
                pattern.replace("{category}", category)
                .replace("{timestamp}", now.strftime("%Y%m%d_%H%M%S"))
                .replace("{idx}", f"{idx:03d}")
                + ".pdf"
            ),
            print_fn=lambda *_args, **_kwargs: None,
        )
        self.assertGreaterEqual(len(tasks), 3)
        out_names = [n for n, _ in tasks]
        self.assertTrue(any("Merged_Price_" in n for n in out_names))
        self.assertTrue(any("Merged_Word_" in n for n in out_names))
        flattened = [p for _, group in tasks for p in group]
        self.assertIn(p1, flattened)
        self.assertIn(p2, flattened)
        self.assertIn(p3, flattened)
        self.assertNotIn(skip, flattened)

        all_in_one = get_merge_tasks(
            run_mode="convert_then_merge",
            merge_source="source",
            target_folder=src,
            get_source_roots_fn=lambda: [root],
            failed_dir=failed,
            merge_output_dir=merged,
            merge_mode=MERGE_MODE_ALL_IN_ONE,
            merge_mode_all_in_one=MERGE_MODE_ALL_IN_ONE,
            mode_merge_only=MODE_MERGE_ONLY,
            merge_filename_pattern="Merged_{category}_{timestamp}_{idx}",
            max_merge_size_mb=80,
            now_fn=lambda: now,
            format_merge_filename_fn=lambda pattern, category, idx, now: f"{category}_{idx}.pdf",
            print_fn=lambda *_args, **_kwargs: None,
        )
        self.assertEqual(len(all_in_one), 1)
        self.assertEqual(all_in_one[0][0], "All_1.pdf")

    def test_office_converter_get_merge_tasks_delegates_to_module(self):
        import office_converter as oc

        original = oc.get_merge_tasks_impl
        try:
            seen = {}

            def _fake(**kwargs):
                seen["kwargs"] = kwargs
                return [("out.pdf", ["a.pdf"])]

            oc.get_merge_tasks_impl = _fake
            dummy = OfficeConverter.__new__(OfficeConverter)
            dummy.run_mode = "merge_only"
            dummy.config = {
                "merge_source": "source",
                "target_folder": "target",
                "merge_filename_pattern": "",
                "max_merge_size_mb": 80,
            }
            dummy.failed_dir = "failed"
            dummy.merge_output_dir = "merged"
            dummy.merge_mode = "by_category"
            dummy._get_source_roots = lambda: ["source"]
            dummy._format_merge_filename = lambda *args, **kwargs: "x"

            out = dummy._get_merge_tasks()
            self.assertEqual(out, [("out.pdf", ["a.pdf"])])
            self.assertEqual(seen["kwargs"]["target_folder"], "target")
            self.assertEqual(seen["kwargs"]["merge_mode"], "by_category")
        finally:
            oc.get_merge_tasks_impl = original

    def test_merge_tasks_module_has_no_bare_except_exception(self):
        mod_path = Path(__file__).resolve().parents[1] / "converter" / "merge_tasks.py"
        text = mod_path.read_text(encoding="utf-8")
        self.assertNotIn("except Exception", text)


if __name__ == "__main__":
    unittest.main()
