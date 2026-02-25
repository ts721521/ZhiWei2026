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
        finally:
            oc.run_fast_md_pipeline_impl = original

    def test_fast_md_engine_module_has_no_bare_except_exception(self):
        mod_path = Path(__file__).resolve().parents[1] / "converter" / "fast_md_engine.py"
        text = mod_path.read_text(encoding="utf-8")
        self.assertNotIn("except Exception", text)


if __name__ == "__main__":
    unittest.main()
