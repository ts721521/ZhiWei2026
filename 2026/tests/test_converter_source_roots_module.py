import os
import tempfile
import unittest

from office_converter import OfficeConverter


class ConverterSourceRootsSplitTests(unittest.TestCase):
    def test_source_roots_core_behaviors(self):
        from converter.source_roots import (
            get_configured_source_roots,
            get_source_root_for_path,
            get_source_roots,
            probe_source_root_access,
        )

        root = tempfile.mkdtemp(prefix="src_roots_")
        a = os.path.join(root, "a")
        b = os.path.join(root, "b")
        os.makedirs(a, exist_ok=True)
        os.makedirs(b, exist_ok=True)
        try:
            cfg = {"source_folders": [a, " ", b]}
            configured = get_configured_source_roots(cfg)
            self.assertEqual(configured, [os.path.abspath(a), os.path.abspath(b)])

            accessible = get_source_roots(cfg)
            self.assertEqual(accessible, [os.path.abspath(a), os.path.abspath(b)])

            cfg2 = {"source_folder": a}
            self.assertEqual(get_configured_source_roots(cfg2), [os.path.abspath(a)])

            picked = get_source_root_for_path(
                os.path.join(a, "child", "f.docx"),
                [root, a],
                fallback="",
            )
            self.assertEqual(picked, os.path.abspath(a))

            skips = []
            ok = probe_source_root_access(
                a,
                record_skip_fn=lambda *args, **kwargs: skips.append((args, kwargs)),
            )
            self.assertTrue(ok)
            self.assertEqual(skips, [])

            bad = probe_source_root_access(
                os.path.join(root, "missing"),
                record_skip_fn=lambda *args, **kwargs: skips.append((args, kwargs)),
            )
            self.assertFalse(bad)
            self.assertGreaterEqual(len(skips), 1)
        finally:
            for d in (a, b, root):
                try:
                    os.rmdir(d)
                except Exception:
                    pass

    def test_office_converter_source_root_methods_delegate_to_module(self):
        from converter.source_roots import (
            get_configured_source_roots,
            get_source_root_for_path,
            get_source_roots,
            probe_source_root_access,
        )

        root = tempfile.mkdtemp(prefix="src_roots_delegate_")
        src = os.path.join(root, "src")
        os.makedirs(src, exist_ok=True)
        dummy = OfficeConverter.__new__(OfficeConverter)
        dummy.config = {"source_folders": [src]}
        dummy._record_scan_access_skip = lambda *args, **kwargs: ("skipped", args, kwargs)
        try:
            self.assertEqual(
                dummy._get_configured_source_roots(),
                get_configured_source_roots(dummy.config),
            )
            self.assertEqual(
                dummy._get_source_roots(),
                get_source_roots(dummy.config, is_dir_func=os.path.isdir),
            )
            self.assertEqual(
                dummy._get_source_root_for_path(os.path.join(src, "a.docx")),
                get_source_root_for_path(
                    os.path.join(src, "a.docx"),
                    dummy._get_source_roots(),
                    fallback=dummy.config.get("source_folder", "") or "",
                ),
            )
            self.assertEqual(
                dummy._probe_source_root_access(src),
                probe_source_root_access(
                    src,
                    dummy._record_scan_access_skip,
                    is_dir_func=os.path.isdir,
                    listdir_fn=os.listdir,
                ),
            )
        finally:
            for d in (src, root):
                try:
                    os.rmdir(d)
                except Exception:
                    pass


if __name__ == "__main__":
    unittest.main()
