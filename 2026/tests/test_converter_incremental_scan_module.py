import os
import tempfile
import unittest
from pathlib import Path

from converter.file_registry import FileRegistry
from office_converter import OfficeConverter


class ConverterIncrementalScanSplitTests(unittest.TestCase):
    def test_incremental_scan_module_has_no_bare_except_exception(self):
        module_text = Path("converter/incremental_scan.py").read_text(encoding="utf-8")
        self.assertNotIn("except Exception", module_text)

    def test_incremental_scan_disabled_mode_passthrough(self):
        from converter.incremental_scan import apply_incremental_filter

        files = [r"C:\x\a.docx", r"C:\x\b.pdf"]
        process_files, context = apply_incremental_filter(
            files,
            {"enable_incremental_mode": False},
            resolve_registry_path_fn=lambda: "",
            build_source_meta_fn=lambda p, include_hash=False: None,
            compute_file_hash_fn=lambda p: "",
        )
        self.assertEqual(process_files, files)
        self.assertFalse(context["enabled"])
        self.assertEqual(context["scanned_count"], len(files))

    def test_incremental_scan_core_behaviors(self):
        from converter.incremental_registry_ops import build_source_meta
        from converter.incremental_scan import apply_incremental_filter

        root = tempfile.mkdtemp(prefix="inc_scan_")
        reg_path = os.path.join(root, "reg.json")
        a = os.path.join(root, "a.docx")
        b = os.path.join(root, "b.docx")
        c_deleted = os.path.join(root, "c.docx")

        for path, data in ((a, b"a"), (b, b"bb"), (c_deleted, b"ccc")):
            with open(path, "wb") as f:
                f.write(data)

        reg = FileRegistry(reg_path, base_root=root)
        a_meta = build_source_meta(a, include_hash=False)
        b_meta = build_source_meta(b, include_hash=False)
        c_meta = build_source_meta(c_deleted, include_hash=False)
        reg.entries = {
            reg.normalize_path(a): dict(a_meta),
            reg.normalize_path(b): dict(b_meta, source_mtime_ns=0),
            reg.normalize_path(c_deleted): dict(c_meta),
        }
        reg.save()
        os.remove(c_deleted)

        process_files, context = apply_incremental_filter(
            [a, b],
            {
                "enable_incremental_mode": True,
                "incremental_verify_hash": False,
                "incremental_reprocess_renamed": False,
                "source_folder": root,
            },
            resolve_registry_path_fn=lambda: reg_path,
            build_source_meta_fn=lambda p, include_hash=False: build_source_meta(
                p,
                include_hash=include_hash,
                compute_file_hash_fn=lambda x: "sha",
            ),
            compute_file_hash_fn=lambda p: "sha",
        )

        self.assertTrue(context["enabled"])
        self.assertEqual(context["scanned_count"], 2)
        self.assertEqual(context["unchanged_count"], 1)
        self.assertEqual(context["modified_count"], 1)
        self.assertEqual(context["added_count"], 0)
        self.assertEqual(context["deleted_count"], 1)
        self.assertEqual(process_files, [b])

        for p in (a, b, reg_path):
            try:
                os.remove(p)
            except Exception:
                pass
        try:
            os.rmdir(root)
        except Exception:
            pass

    def test_office_converter_incremental_scan_method_delegates_to_module(self):
        from converter.incremental_scan import (
            apply_incremental_filter,
            apply_incremental_filter_for_converter,
        )

        dummy = OfficeConverter.__new__(OfficeConverter)
        dummy.config = {"enable_incremental_mode": False}
        dummy._incremental_context = {"enabled": True}
        dummy.incremental_registry_path = "x"
        dummy._resolve_incremental_registry_path = lambda: ""
        dummy._build_source_meta = lambda p, include_hash=False: None
        dummy._compute_file_hash = lambda p: ""

        files = [r"C:\x\a.docx"]
        expected = apply_incremental_filter(
            files,
            dummy.config,
            resolve_registry_path_fn=dummy._resolve_incremental_registry_path,
            build_source_meta_fn=dummy._build_source_meta,
            compute_file_hash_fn=dummy._compute_file_hash,
        )
        actual = dummy._apply_incremental_filter(files)
        self.assertEqual(actual, expected)
        self.assertIsNone(dummy._incremental_context)
        self.assertEqual(dummy.incremental_registry_path, "")
        self.assertEqual(actual, apply_incremental_filter_for_converter(dummy, files))


if __name__ == "__main__":
    unittest.main()
