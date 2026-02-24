import json
import os
import tempfile
import unittest

from office_converter import OfficeConverter


class ConverterCorpusManifestSplitTests(unittest.TestCase):
    def test_llm_delivery_hub_core_behaviors(self):
        from converter.corpus_manifest import maybe_build_llm_delivery_hub

        root = tempfile.mkdtemp(prefix="llm_hub_")
        target = os.path.join(root, "target")
        hub = os.path.join(root, "hub")
        os.makedirs(target, exist_ok=True)
        source_md = os.path.join(target, "_AI", "Markdown", "a.md")
        os.makedirs(os.path.dirname(source_md), exist_ok=True)
        with open(source_md, "w", encoding="utf-8") as f:
            f.write("# test")

        class Dummy:
            def __init__(self):
                self.config = {
                    "enable_llm_delivery_hub": True,
                    "llm_delivery_root": hub,
                    "llm_delivery_include_pdf": False,
                    "llm_delivery_flatten": False,
                    "upload_dedup_merged": True,
                    "enable_upload_json_manifest": True,
                    "enable_upload_readme": True,
                }
                self.run_mode = "convert_then_merge"
                self.llm_hub_root = ""
                self.mshelp_records = []

        dummy = Dummy()
        artifacts = [
            {
                "kind": "markdown_export",
                "path_abs": source_md,
                "path_rel_to_target": os.path.relpath(source_md, target),
                "md5": "",
                "sha256": "",
            }
        ]
        out = maybe_build_llm_delivery_hub(dummy, target, artifacts)
        self.assertIsNotNone(out)
        self.assertEqual(dummy.llm_hub_root, hub)
        self.assertTrue(os.path.exists(os.path.join(hub, "llm_upload_manifest.json")))
        self.assertTrue(os.path.exists(os.path.join(hub, "README_UPLOAD_LIST.txt")))

    def test_write_corpus_manifest_core_behaviors(self):
        from converter.artifact_meta import add_artifact
        from converter.corpus_manifest import write_corpus_manifest

        root = tempfile.mkdtemp(prefix="corpus_manifest_")
        target = os.path.join(root, "target")
        os.makedirs(target, exist_ok=True)
        pdf = os.path.join(target, "a.pdf")
        with open(pdf, "w", encoding="utf-8") as f:
            f.write("x")

        class Dummy:
            def __init__(self):
                self.config = {
                    "enable_corpus_manifest": True,
                    "target_folder": target,
                    "source_folder": root,
                    "enable_llm_delivery_hub": False,
                }
                self.run_mode = "convert_only"
                self.collect_mode = "collect"
                self.merge_mode = "category"
                self.content_strategy = "standard"
                self.generated_pdfs = [pdf]
                self.generated_merge_outputs = []
                self.generated_merge_markdown_outputs = []
                self.generated_map_outputs = []
                self.generated_markdown_outputs = []
                self.generated_markdown_quality_outputs = []
                self.generated_excel_json_outputs = []
                self.generated_records_json_outputs = []
                self.generated_chromadb_outputs = []
                self.generated_update_package_outputs = []
                self.generated_mshelp_outputs = []
                self.conversion_index_records = []
                self.merge_index_records = []
                self.convert_index_path = None
                self.collect_index_path = None
                self.merge_excel_path = None
                self.corpus_manifest_path = None

            def _add_artifact(self, artifacts, kind, path):
                add_artifact(artifacts, kind, path, self.config.get("target_folder", ""))

            def _maybe_build_llm_delivery_hub(self, target_folder, artifacts):
                return None

        dummy = Dummy()
        out = write_corpus_manifest(dummy)
        self.assertTrue(out)
        self.assertTrue(os.path.exists(out))
        self.assertEqual(dummy.corpus_manifest_path, out)
        with open(out, "r", encoding="utf-8") as f:
            payload = json.load(f)
        self.assertEqual(payload.get("summary", {}).get("converted_pdf_count"), 1)
        self.assertGreaterEqual(payload.get("summary", {}).get("artifact_count", 0), 1)

    def test_office_converter_corpus_methods_delegate_to_module(self):
        import office_converter as oc

        original_hub = oc.maybe_build_llm_delivery_hub_impl
        original_manifest = oc.write_corpus_manifest_impl
        try:
            seen = {}

            def _fake_hub(converter, target_folder, artifacts):
                seen["hub"] = (converter, target_folder, artifacts)
                return {"kind": "llm_delivery_hub"}

            def _fake_manifest(converter, merge_outputs=None):
                seen["manifest"] = (converter, merge_outputs)
                return "corpus.json"

            oc.maybe_build_llm_delivery_hub_impl = _fake_hub
            oc.write_corpus_manifest_impl = _fake_manifest

            dummy = OfficeConverter.__new__(OfficeConverter)
            hub_out = dummy._maybe_build_llm_delivery_hub("T", [])
            manifest_out = dummy._write_corpus_manifest(["a.pdf"])

            self.assertEqual(hub_out, {"kind": "llm_delivery_hub"})
            self.assertEqual(manifest_out, "corpus.json")
            self.assertEqual(seen.get("manifest", (None, None))[1], ["a.pdf"])
        finally:
            oc.maybe_build_llm_delivery_hub_impl = original_hub
            oc.write_corpus_manifest_impl = original_manifest


if __name__ == "__main__":
    unittest.main()
