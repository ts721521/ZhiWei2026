import json
import os
import tempfile
import unittest
from datetime import datetime
from pathlib import Path

from office_converter import OfficeConverter


class ConverterChromaDbExportSplitTests(unittest.TestCase):
    def test_chromadb_export_module_has_no_bare_except_exception(self):
        module_text = Path("converter/chromadb_export.py").read_text(encoding="utf-8")
        self.assertNotIn("except Exception", module_text)

    def test_chromadb_export_core_behaviors(self):
        from converter.chromadb_export import write_chromadb_export

        root = tempfile.mkdtemp(prefix="chromadb_export_")
        docs = [
            {"id": "1", "document": "abc", "metadata": {"a": 1, "b": None, "c": {"x": 1}}}
        ]
        fixed_now = datetime(2026, 2, 24, 12, 0, 0)
        manifest, outputs = write_chromadb_export(
            docs,
            config={"chromadb_write_jsonl_fallback": True, "chromadb_collection_name": "x"},
            target_root=root,
            has_chromadb=False,
            chromadb_module=None,
            sanitize_collection_name_fn=lambda s: "x",
            resolve_persist_dir_fn=lambda: os.path.join(root, "_AI", "ChromaDB", "db"),
            now_fn=lambda: fixed_now,
        )
        self.assertTrue(manifest.endswith("chroma_export_20260224_120000.json"))
        self.assertEqual(2, len(outputs))
        self.assertTrue(all(os.path.exists(p) for p in outputs))

        with open(manifest, "r", encoding="utf-8") as f:
            payload = json.load(f)
        self.assertEqual(payload["status"], "chromadb_missing")
        self.assertEqual(payload["record_count"], 1)

        for p in outputs:
            try:
                os.remove(p)
            except Exception:
                pass
        for d in (
            os.path.join(root, "_AI", "ChromaDB"),
            os.path.join(root, "_AI"),
            root,
        ):
            try:
                os.rmdir(d)
            except Exception:
                pass

    def test_office_converter_write_chromadb_export_delegates_to_module(self):
        import office_converter as oc

        root = tempfile.mkdtemp(prefix="chromadb_export_delegate_")
        orig = oc.write_chromadb_export_impl
        try:
            oc.write_chromadb_export_impl = (
                lambda docs, **kwargs: (os.path.join(root, "m.json"), [os.path.join(root, "m.json")])
            )
            dummy = OfficeConverter.__new__(OfficeConverter)
            dummy.config = {"enable_chromadb_export": True, "target_folder": root}
            dummy.generated_chromadb_outputs = []
            dummy.chromadb_export_manifest_path = None
            dummy._collect_chromadb_documents = lambda: [{"id": "1", "document": "x", "metadata": {}}]
            dummy._sanitize_chromadb_collection_name = lambda s: s
            dummy._resolve_chromadb_persist_dir = lambda: os.path.join(root, "db")

            out = dummy._write_chromadb_export()
            self.assertEqual(out, os.path.join(root, "m.json"))
            self.assertEqual(dummy.chromadb_export_manifest_path, out)
            self.assertEqual(dummy.generated_chromadb_outputs, [out])
        finally:
            oc.write_chromadb_export_impl = orig
            try:
                os.rmdir(root)
            except Exception:
                pass

    def test_office_converter_write_chromadb_export_clears_state_when_no_docs(self):
        dummy = OfficeConverter.__new__(OfficeConverter)
        dummy.config = {"enable_chromadb_export": True, "target_folder": tempfile.mkdtemp(prefix="chromadb_no_docs_")}
        dummy.generated_chromadb_outputs = ["x"]
        dummy.chromadb_export_manifest_path = "y"
        dummy._collect_chromadb_documents = lambda: []
        out = dummy._write_chromadb_export()
        self.assertIsNone(out)
        self.assertEqual(dummy.generated_chromadb_outputs, [])
        self.assertIsNone(dummy.chromadb_export_manifest_path)


if __name__ == "__main__":
    unittest.main()
