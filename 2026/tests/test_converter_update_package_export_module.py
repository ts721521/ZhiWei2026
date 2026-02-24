import hashlib
import json
import os
import tempfile
import unittest
from datetime import datetime

from office_converter import OfficeConverter


class ConverterUpdatePackageExportSplitTests(unittest.TestCase):
    def test_update_package_export_core_behaviors(self):
        from converter.update_package_export import generate_update_package

        root = tempfile.mkdtemp(prefix="update_pkg_")
        source_root = os.path.join(root, "src")
        target_root = os.path.join(root, "target")
        update_root = os.path.join(root, "update")
        os.makedirs(source_root, exist_ok=True)
        os.makedirs(target_root, exist_ok=True)
        os.makedirs(update_root, exist_ok=True)

        src = os.path.join(source_root, "a.docx")
        final_pdf = os.path.join(target_root, "a.pdf")
        with open(src, "w", encoding="utf-8") as f:
            f.write("source")
        with open(final_pdf, "w", encoding="utf-8") as f:
            f.write("pdf")

        def _md5(path):
            with open(path, "rb") as f:
                return hashlib.md5(f.read()).hexdigest()

        def _write_xlsx(path, records):
            with open(path, "w", encoding="utf-8") as f:
                f.write(str(len(records)))
            return path

        fixed_now = datetime(2026, 2, 24, 13, 20, 30)
        manifest_path, outputs = generate_update_package(
            [
                {
                    "source_path": src,
                    "status": "success",
                    "detail": "",
                    "final_path": final_pdf,
                }
            ],
            incremental_context={
                "enabled": True,
                "registry_path": os.path.join(root, "registry.json"),
                "scan_meta": {
                    src: {
                        "change_state": "modified",
                        "source_hash_sha256": "sha256-x",
                        "renamed_from": "",
                        "rename_match_type": "",
                    }
                },
                "scanned_count": 1,
                "added_count": 0,
                "modified_count": 1,
                "renamed_count": 0,
                "unchanged_count": 0,
                "deleted_count": 0,
                "deleted_paths": [],
                "renamed_pairs": [],
            },
            config={
                "enable_update_package": True,
                "source_folder": source_root,
                "target_folder": target_root,
            },
            run_mode="convert_then_merge",
            resolve_update_package_root_fn=lambda: update_root,
            compute_md5_fn=_md5,
            write_update_package_index_xlsx_fn=_write_xlsx,
            now_fn=lambda: fixed_now,
        )

        self.assertTrue(manifest_path)
        self.assertTrue(os.path.exists(manifest_path))
        self.assertTrue(outputs)
        self.assertIn(manifest_path, outputs)
        copied = [p for p in outputs if p.lower().endswith(".pdf")]
        self.assertEqual(len(copied), 1)
        self.assertTrue(os.path.exists(copied[0]))

        with open(manifest_path, "r", encoding="utf-8") as f:
            manifest = json.load(f)
        self.assertEqual(manifest.get("record_count"), 1)
        self.assertEqual(manifest.get("packaged_pdf_count"), 1)
        self.assertEqual(manifest.get("status_counts", {}).get("success"), 1)

    def test_office_converter_generate_update_package_delegates_to_module(self):
        import office_converter as oc

        original = oc.generate_update_package_impl
        try:
            seen = {}

            def _fake(process_results, **kwargs):
                seen["process_results"] = process_results
                seen["kwargs"] = kwargs
                return "manifest.json", ["manifest.json", "index.json"]

            oc.generate_update_package_impl = _fake

            dummy = OfficeConverter.__new__(OfficeConverter)
            dummy._incremental_context = {"enabled": True}
            dummy.config = {}
            dummy.run_mode = "x"
            dummy._resolve_update_package_root = lambda: "out"
            dummy._compute_md5 = lambda _p: "md5"
            dummy._write_update_package_index_xlsx = lambda _p, _r: None

            out = dummy._generate_update_package([{"source_path": "a"}])
            self.assertEqual(out, "manifest.json")
            self.assertEqual(dummy.generated_update_package_outputs, ["manifest.json", "index.json"])
            self.assertEqual(dummy.update_package_manifest_path, "manifest.json")
            self.assertEqual(seen.get("process_results"), [{"source_path": "a"}])
        finally:
            oc.generate_update_package_impl = original


if __name__ == "__main__":
    unittest.main()
