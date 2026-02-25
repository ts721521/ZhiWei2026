import json
import os
import tempfile
import unittest
from pathlib import Path

from converter.file_registry import FileRegistry
from office_converter import OfficeConverter


class ConverterIncrementalRegistryOpsSplitTests(unittest.TestCase):
    def test_build_source_meta_core_behaviors(self):
        from converter.incremental_registry_ops import build_source_meta

        root = tempfile.mkdtemp(prefix="inc_meta_")
        path = os.path.join(root, "a.docx")
        with open(path, "wb") as f:
            f.write(b"hello")

        warnings = []
        meta = build_source_meta(path, include_hash=False)
        self.assertEqual(meta["ext"], ".docx")
        self.assertEqual(meta["source_size"], 5)
        self.assertEqual(meta["source_hash_sha256"], "")

        meta2 = build_source_meta(
            path,
            include_hash=True,
            compute_file_hash_fn=lambda p: "sha",
        )
        self.assertEqual(meta2["source_hash_sha256"], "sha")

        meta3 = build_source_meta(
            path,
            include_hash=True,
            compute_file_hash_fn=lambda p: (_ for _ in ()).throw(RuntimeError("boom")),
            log_warning=warnings.append,
        )
        self.assertEqual(meta3["source_hash_sha256"], "")
        self.assertTrue(any("failed to compute source hash" in w for w in warnings))

        try:
            os.remove(path)
            os.rmdir(root)
        except Exception:
            pass

    def test_flush_incremental_registry_core_behaviors(self):
        from converter.incremental_registry_ops import flush_incremental_registry

        root = tempfile.mkdtemp(prefix="inc_flush_")
        reg_path = os.path.join(root, "registry.json")
        src = os.path.join(root, "a.docx")
        out_pdf = os.path.join(root, "a.pdf")
        with open(src, "wb") as f:
            f.write(b"a")
        with open(out_pdf, "wb") as f:
            f.write(b"pdf")

        reg = FileRegistry(reg_path, base_root=root)
        reg.entries = {}
        context = {
            "enabled": True,
            "registry": reg,
            "registry_path": reg_path,
            "scan_meta": {
                src: {
                    "ext": ".docx",
                    "source_size": 1,
                    "source_mtime": "2026-01-01T00:00:00",
                    "source_mtime_ns": 1,
                    "source_hash_sha256": "h",
                    "change_state": "modified",
                }
            },
            "scanned_count": 1,
            "added_count": 0,
            "modified_count": 1,
            "renamed_count": 0,
            "unchanged_count": 0,
            "deleted_count": 0,
        }
        process_results = [
            {
                "source_path": src,
                "status": "success",
                "error": "",
                "final_path": out_pdf,
            }
        ]

        info_logs = []
        flush_incremental_registry(
            context,
            process_results,
            run_mode="convert_then_merge",
            compute_md5_fn=lambda p: "md5",
            log_info=info_logs.append,
        )

        key = reg.normalize_path(src)
        self.assertIn(key, reg.entries)
        self.assertEqual(reg.entries[key]["last_status"], "success")
        self.assertEqual(reg.entries[key]["last_output_pdf_md5"], "md5")
        self.assertTrue(os.path.exists(reg_path))

        with open(reg_path, "r", encoding="utf-8") as f:
            payload = json.load(f)
        self.assertIn("last_run", payload)
        self.assertTrue(any("registry updated" in msg for msg in info_logs))

        for p in (out_pdf, src, reg_path):
            try:
                os.remove(p)
            except Exception:
                pass
        try:
            os.rmdir(root)
        except Exception:
            pass

    def test_office_converter_incremental_registry_methods_delegate_to_module(self):
        from converter.incremental_registry_ops import build_source_meta

        root = tempfile.mkdtemp(prefix="inc_delegate_")
        path = os.path.join(root, "x.docx")
        with open(path, "wb") as f:
            f.write(b"x")

        dummy = OfficeConverter.__new__(OfficeConverter)
        dummy._compute_file_hash = lambda p: "sha-x"

        expected = build_source_meta(
            path,
            include_hash=True,
            compute_file_hash_fn=dummy._compute_file_hash,
        )
        actual = dummy._build_source_meta(path, include_hash=True)
        self.assertEqual(actual, expected)

        try:
            os.remove(path)
            os.rmdir(root)
        except Exception:
            pass

    def test_incremental_registry_ops_module_has_no_bare_except_exception(self):
        mod_path = (
            Path(__file__).resolve().parents[1] / "converter" / "incremental_registry_ops.py"
        )
        text = mod_path.read_text(encoding="utf-8")
        self.assertNotIn("except Exception", text)


if __name__ == "__main__":
    unittest.main()
