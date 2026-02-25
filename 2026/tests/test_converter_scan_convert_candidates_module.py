import os
import tempfile
import unittest

from office_converter import OfficeConverter


class ConverterScanConvertCandidatesSplitTests(unittest.TestCase):
    def test_scan_convert_candidates_core_behaviors(self):
        from converter.scan_convert_candidates import (
            iter_convert_candidates,
            scan_convert_candidates,
        )

        root = tempfile.mkdtemp(prefix="scan_convert_")
        keep = os.path.join(root, "a.docx")
        skip_ext = os.path.join(root, "a.txt")
        excluded = os.path.join(root, "excluded")
        os.makedirs(excluded, exist_ok=True)
        skip_excluded = os.path.join(excluded, "b.docx")
        for p in (keep, skip_ext, skip_excluded):
            with open(p, "wb") as f:
                f.write(b"x")

        cfg = {
            "source_folder": root,
            "excluded_folders": ["excluded"],
            "allowed_extensions": {"word": [".docx"], "excel": [], "powerpoint": []},
        }
        files = scan_convert_candidates(
            cfg,
            [root],
            probe_source_root_access_fn=lambda path, context=None, seen_keys=None: True,
            record_scan_access_skip_fn=lambda *args, **kwargs: None,
            filter_date=None,
            filter_mode="after",
        )
        self.assertEqual(files, [keep])
        iter_files = list(
            iter_convert_candidates(
                cfg,
                [root],
                probe_source_root_access_fn=lambda path, context=None, seen_keys=None: True,
                record_scan_access_skip_fn=lambda *args, **kwargs: None,
                filter_date=None,
                filter_mode="after",
            )
        )
        self.assertEqual(iter_files, [keep])

        warns = []
        errs = []
        files2 = scan_convert_candidates(
            cfg,
            [root],
            probe_source_root_access_fn=lambda path, context=None, seen_keys=None: False,
            record_scan_access_skip_fn=lambda *args, **kwargs: None,
            filter_date=None,
            filter_mode="after",
            print_warn_fn=warns.append,
            log_error_fn=errs.append,
        )
        self.assertEqual(files2, [])
        self.assertTrue(warns)
        self.assertTrue(errs)

        for p in (skip_excluded, excluded, keep, skip_ext):
            try:
                if os.path.isdir(p):
                    os.rmdir(p)
                else:
                    os.remove(p)
            except Exception:
                pass
        try:
            os.rmdir(root)
        except Exception:
            pass

    def test_office_converter_scan_convert_candidates_delegates_to_module(self):
        from converter.scan_convert_candidates import scan_convert_candidates

        root = tempfile.mkdtemp(prefix="scan_convert_delegate_")
        keep = os.path.join(root, "a.docx")
        with open(keep, "wb") as f:
            f.write(b"x")

        dummy = OfficeConverter.__new__(OfficeConverter)
        dummy.config = {
            "source_folder": root,
            "excluded_folders": [],
            "allowed_extensions": {"word": [".docx"], "excel": [], "powerpoint": []},
        }
        dummy.filter_date = None
        dummy.filter_mode = "after"
        dummy._get_configured_source_roots = lambda: [root]
        dummy._probe_source_root_access = lambda path, context=None, seen_keys=None: True
        dummy._record_scan_access_skip = lambda *args, **kwargs: None

        expected = scan_convert_candidates(
            dummy.config,
            dummy._get_configured_source_roots(),
            probe_source_root_access_fn=dummy._probe_source_root_access,
            record_scan_access_skip_fn=dummy._record_scan_access_skip,
            filter_date=dummy.filter_date,
            filter_mode=dummy.filter_mode,
        )
        actual = dummy._scan_convert_candidates()
        self.assertEqual(actual, expected)

        try:
            os.remove(keep)
            os.rmdir(root)
        except Exception:
            pass


if __name__ == "__main__":
    unittest.main()
