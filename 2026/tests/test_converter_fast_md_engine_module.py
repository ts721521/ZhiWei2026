import os
import tempfile
import unittest
from pathlib import Path

from office_converter import OfficeConverter


class ConverterFastMdEngineSplitTests(unittest.TestCase):
    def test_fast_md_engine_core_behaviors(self):
        from converter.fast_md_engine import run_fast_md_pipeline

        root = tempfile.mkdtemp(prefix="fast_md_")
        target = os.path.join(root, "target")
        os.makedirs(target, exist_ok=True)
        src = os.path.join(root, "a.docx")
        with open(src, "w", encoding="utf-8") as f:
            f.write("hello")

        generated = []
        quality = []
        stats = {"success": 0, "failed": 0}
        progress_calls = []
        done_records = []
        out = run_fast_md_pipeline(
            [src],
            config={"short_id_prefix": "ZW-"},
            target_folder=target,
            generated_markdown_outputs=generated,
            markdown_quality_records=quality,
            trace_short_id_taken=set(),
            stats=stats,
            source_root_resolver_fn=lambda _p: root,
            compute_md5_fn=lambda _p: "a" * 32,
            build_short_id_fn=lambda _md5, _taken: "AAAAAAAA",
            convert_source_to_markdown_text_fn=lambda _p: "md body",
            progress_callback=lambda c, t: progress_calls.append((c, t)),
            emit_file_done_fn=lambda rec: done_records.append(rec),
            log_info_fn=lambda _m: None,
            log_warning_fn=lambda _m: None,
        )
        self.assertEqual(stats["success"], 1)
        self.assertEqual(stats["failed"], 0)
        self.assertEqual(len(out.get("batch_results", [])), 1)
        self.assertTrue(out.get("bundle_path", "").endswith("_Knowledge_Bundle.md"))
        self.assertTrue(os.path.exists(out["bundle_path"]))
        self.assertGreaterEqual(len(generated), 2)
        self.assertEqual(quality[0]["source_short_id"], "ZW-AAAAAAAA")
        self.assertEqual(progress_calls, [(1, 1)])
        self.assertEqual(len(done_records), 1)
        self.assertEqual(done_records[0]["status"], "success_md_only")

    def test_office_converter_fast_md_delegate(self):
        import office_converter as oc

        original = oc.run_fast_md_pipeline_impl
        try:
            seen = {}

            def _fake(files, **kwargs):
                seen["files"] = files
                seen["kwargs"] = kwargs
                return {"batch_results": [], "bundle_path": None}

            oc.run_fast_md_pipeline_impl = _fake
            dummy = OfficeConverter.__new__(OfficeConverter)
            dummy.config = {"target_folder": "T"}
            dummy.generated_markdown_outputs = []
            dummy.markdown_quality_records = []
            dummy.trace_short_id_taken = set()
            dummy.stats = {"success": 0, "failed": 0}
            dummy._get_source_root_for_path = lambda _p: "S"
            dummy._compute_md5 = lambda _p: "x"
            dummy._build_short_id = lambda _m, _t: "Y"
            dummy._convert_source_to_markdown_text = lambda _p: "z"

            out = dummy._run_fast_md_pipeline(["a.docx"])
            self.assertEqual(out.get("bundle_path"), None)
            self.assertEqual(seen["files"], ["a.docx"])
            self.assertIn("progress_callback", seen["kwargs"])
            self.assertIn("emit_file_done_fn", seen["kwargs"])
        finally:
            oc.run_fast_md_pipeline_impl = original

    def test_fast_md_engine_module_has_no_bare_except_exception(self):
        mod_path = Path(__file__).resolve().parents[1] / "converter" / "fast_md_engine.py"
        text = mod_path.read_text(encoding="utf-8")
        self.assertNotIn("except Exception", text)

    def test_fast_md_engine_marks_oversize_markdown_as_failed(self):
        from converter.fast_md_engine import run_fast_md_pipeline

        root = tempfile.mkdtemp(prefix="fast_md_oversize_")
        target = os.path.join(root, "target")
        os.makedirs(target, exist_ok=True)
        src = os.path.join(root, "big.docx")
        with open(src, "w", encoding="utf-8") as f:
            f.write("x")

        generated = []
        quality = []
        stats = {"success": 0, "failed": 0}
        out = run_fast_md_pipeline(
            [src],
            config={"short_id_prefix": "ZW-", "markdown_max_size_mb": 1},
            target_folder=target,
            generated_markdown_outputs=generated,
            markdown_quality_records=quality,
            trace_short_id_taken=set(),
            stats=stats,
            source_root_resolver_fn=lambda _p: root,
            compute_md5_fn=lambda _p: "b" * 32,
            build_short_id_fn=lambda _md5, _taken: "BBBBBBBB",
            convert_source_to_markdown_text_fn=lambda _p: "x" * (2 * 1024 * 1024),
            log_info_fn=lambda _m: None,
            log_warning_fn=lambda _m: None,
        )
        self.assertEqual(stats["success"], 0)
        self.assertEqual(stats["failed"], 1)
        self.assertEqual(len(generated), 0)
        self.assertEqual(len(quality), 0)
        self.assertIsNone(out.get("bundle_path"))
        self.assertIn("oversize", out["batch_results"][0]["detail"])

    def test_fast_md_engine_splits_knowledge_bundle_by_size_limit(self):
        from converter.fast_md_engine import run_fast_md_pipeline

        root = tempfile.mkdtemp(prefix="fast_md_bundle_split_")
        target = os.path.join(root, "target")
        os.makedirs(target, exist_ok=True)
        src_a = os.path.join(root, "a.docx")
        src_b = os.path.join(root, "b.docx")
        with open(src_a, "w", encoding="utf-8") as f:
            f.write("a")
        with open(src_b, "w", encoding="utf-8") as f:
            f.write("b")

        body_map = {
            src_a: "A" * (700 * 1024),
            src_b: "B" * (700 * 1024),
        }
        generated = []
        quality = []
        stats = {"success": 0, "failed": 0}
        out = run_fast_md_pipeline(
            [src_a, src_b],
            config={"short_id_prefix": "ZW-", "markdown_max_size_mb": 1},
            target_folder=target,
            generated_markdown_outputs=generated,
            markdown_quality_records=quality,
            trace_short_id_taken=set(),
            stats=stats,
            source_root_resolver_fn=lambda _p: root,
            compute_md5_fn=lambda p: ("a" * 31 + "1") if p == src_a else ("b" * 31 + "2"),
            build_short_id_fn=lambda _md5, _taken: "ID",
            convert_source_to_markdown_text_fn=lambda p: body_map[p],
            log_info_fn=lambda _m: None,
            log_warning_fn=lambda _m: None,
        )
        self.assertEqual(stats["success"], 2)
        self.assertEqual(stats["failed"], 0)
        bundles = [p for p in generated if "_Knowledge_Bundle_" in os.path.basename(p)]
        self.assertEqual(len(bundles), 2)
        self.assertTrue(out.get("bundle_path", "").endswith("_Knowledge_Bundle_001.md"))


if __name__ == "__main__":
    unittest.main()
