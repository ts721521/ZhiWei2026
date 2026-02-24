import os
import tempfile
import unittest

from office_converter import OfficeConverter


class ConverterMshelpScanSplitTests(unittest.TestCase):
    def test_mshelp_scan_core_behaviors(self):
        from converter.mshelp_scan import (
            find_mshelpviewer_dirs,
            scan_mshelp_cab_candidates,
        )

        root = tempfile.mkdtemp(prefix="mshelp_scan_")
        d1 = os.path.join(root, "A", "MSHelpViewer")
        d2 = os.path.join(root, "B", "Nested", "MSHelpViewer")
        os.makedirs(d1, exist_ok=True)
        os.makedirs(d2, exist_ok=True)

        dirs = find_mshelpviewer_dirs(
            root,
            folder_name="MSHelpViewer",
            is_win_fn=lambda: True,
        )
        self.assertIn(os.path.abspath(d1), dirs)
        self.assertIn(os.path.abspath(d2), dirs)

        cab1 = os.path.join(d1, "a.cab")
        cab2 = os.path.join(d1, "A.CAB")
        with open(cab1, "wb") as f:
            f.write(b"x")

        dirs2, files2 = scan_mshelp_cab_candidates(
            {"allowed_extensions": {"cab": [".cab"]}},
            [root],
            find_mshelpviewer_dirs_fn=lambda src: [d1],
            find_files_recursive_fn=lambda d, exts: [cab1, cab2],
            is_win_fn=lambda: True,
        )
        self.assertEqual(dirs2, [os.path.abspath(d1)])
        self.assertEqual(files2, [os.path.abspath(cab1)])

        try:
            os.remove(cab1)
        except Exception:
            pass
        for d in (
            d2,
            os.path.dirname(d2),
            os.path.join(root, "B"),
            d1,
            os.path.join(root, "A"),
            root,
        ):
            try:
                os.rmdir(d)
            except Exception:
                pass

    def test_office_converter_mshelp_scan_methods_delegate_to_module(self):
        from converter.mshelp_scan import (
            find_mshelpviewer_dirs,
            scan_mshelp_cab_candidates,
        )

        root = tempfile.mkdtemp(prefix="mshelp_scan_delegate_")
        mshelp_dir = os.path.join(root, "MSHelpViewer")
        os.makedirs(mshelp_dir, exist_ok=True)
        cab = os.path.join(mshelp_dir, "a.cab")
        with open(cab, "wb") as f:
            f.write(b"x")

        dummy = OfficeConverter.__new__(OfficeConverter)
        dummy.config = {
            "mshelpviewer_folder_name": "MSHelpViewer",
            "allowed_extensions": {"cab": [".cab"]},
        }
        dummy._get_source_roots = lambda: [root]
        dummy._find_files_recursive = lambda d, exts: [cab]

        self.assertEqual(
            dummy._find_mshelpviewer_dirs(root),
            find_mshelpviewer_dirs(
                root,
                folder_name=dummy.config.get("mshelpviewer_folder_name", "MSHelpViewer"),
            ),
        )

        expected = scan_mshelp_cab_candidates(
            dummy.config,
            dummy._get_source_roots(),
            find_mshelpviewer_dirs_fn=dummy._find_mshelpviewer_dirs,
            find_files_recursive_fn=dummy._find_files_recursive,
        )
        actual = dummy._scan_mshelp_cab_candidates()
        self.assertEqual(actual, expected)

        try:
            os.remove(cab)
            os.rmdir(mshelp_dir)
            os.rmdir(root)
        except Exception:
            pass


if __name__ == "__main__":
    unittest.main()
