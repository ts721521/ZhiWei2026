import unittest
from pathlib import Path

from office_converter import OfficeConverter


class ConverterRuntimeExtractsSplitTests(unittest.TestCase):
    def test_target_path_core_and_delegate(self):
        from converter.target_path import get_target_path

        cfg = {
            "target_folder": r"C:\out",
            "allowed_extensions": {
                "word": [".docx"],
                "excel": [".xlsx"],
                "powerpoint": [".pptx"],
                "pdf": [".pdf"],
            },
        }
        out = get_target_path(cfg, r"C:\src\a.docx", ".docx")
        self.assertTrue(out.endswith(r"Word_a.pdf"))

        import office_converter as oc

        original = oc.get_target_path_impl
        try:
            oc.get_target_path_impl = lambda *_a, **_k: "x.pdf"
            dummy = OfficeConverter.__new__(OfficeConverter)
            dummy.config = cfg
            self.assertEqual(dummy.get_target_path("a.docx", ".docx"), "x.pdf")
        finally:
            oc.get_target_path_impl = original

    def test_pdf_content_scan_core_and_delegate(self):
        from converter.pdf_content_scan import scan_pdf_content

        class _Page:
            def __init__(self, text):
                self._text = text

            def extract_text(self):
                return self._text

        class _Reader:
            def __init__(self, _path):
                self.pages = [_Page("abc"), _Page("price hit")]

        self.assertTrue(
            scan_pdf_content(
                "x.pdf",
                price_keywords=["price"],
                has_pypdf=True,
                pdf_reader_cls=_Reader,
            )
        )
        self.assertFalse(
            scan_pdf_content(
                "x.pdf",
                price_keywords=["price"],
                has_pypdf=False,
                pdf_reader_cls=None,
            )
        )

        import office_converter as oc

        original = oc.scan_pdf_content_impl
        try:
            oc.scan_pdf_content_impl = lambda *_a, **_k: True
            dummy = OfficeConverter.__new__(OfficeConverter)
            dummy.price_keywords = ["x"]
            self.assertTrue(dummy.scan_pdf_content("x.pdf"))
        finally:
            oc.scan_pdf_content_impl = original

    def test_error_summary_core_and_delegate(self):
        from converter.error_summary import get_error_summary_for_display

        records = [
            {
                "error_type": "timeout",
                "message": "m1",
                "suggestion": "s1",
                "is_retryable": True,
                "requires_manual_action": False,
                "file_name": "a.docx",
            },
            {
                "error_type": "timeout",
                "message": "m1",
                "suggestion": "s1",
                "is_retryable": True,
                "requires_manual_action": False,
                "file_name": "b.docx",
            },
        ]
        out = get_error_summary_for_display(records)
        self.assertEqual(out["timeout"]["files"], ["a.docx", "b.docx"])

        import office_converter as oc

        original = oc.get_error_summary_for_display_impl
        try:
            oc.get_error_summary_for_display_impl = lambda _r: {"x": 1}
            dummy = OfficeConverter.__new__(OfficeConverter)
            dummy.detailed_error_records = records
            self.assertEqual(dummy.get_error_summary_for_display(), {"x": 1})
        finally:
            oc.get_error_summary_for_display_impl = original

    def test_mshelp_workflow_core_and_delegate(self):
        from converter.mshelp_workflow import run_mshelp_only

        stats = {}
        result = run_mshelp_only(
            stats=stats,
            scan_mshelp_cab_candidates_fn=lambda: (["d"], ["a.cab"]),
            run_batch_fn=lambda files: [{"f": files[0]}],
            write_mshelp_index_files_fn=lambda: ["idx.json"],
            merge_mshelp_markdowns_fn=lambda: ["m.md"],
            add_perf_seconds_fn=lambda _k, _v: None,
            perf_counter_fn=lambda: 1.0,
            log_info_fn=lambda *_a, **_k: None,
            print_fn=lambda *_a, **_k: None,
        )
        self.assertEqual(stats["total"], 1)
        self.assertEqual(result[0][0]["f"], "a.cab")

        import office_converter as oc

        original = oc.run_mshelp_only_impl
        try:
            oc.run_mshelp_only_impl = lambda **_k: ("r", "d", "i", "m")
            dummy = OfficeConverter.__new__(OfficeConverter)
            dummy.stats = {}
            dummy._scan_mshelp_cab_candidates = lambda: ([], [])
            dummy.run_batch = lambda _f: []
            dummy._write_mshelp_index_files = lambda: []
            dummy._merge_mshelp_markdowns = lambda: []
            dummy._add_perf_seconds = lambda _k, _v: None
            self.assertEqual(dummy._run_mshelp_only(), ("r", "d", "i", "m"))
        finally:
            oc.run_mshelp_only_impl = original

    def test_runtime_summary_core_and_delegate(self):
        from converter.runtime_summary import print_runtime_summary

        logs = []
        print_runtime_summary(
            config={"source_folder": "S", "target_folder": "T"},
            run_mode="convert_only",
            merge_mode="category",
            content_strategy="standard",
            mode_merge_only="merge_only",
            get_output_pref_fn=lambda: {
                "pdf": True,
                "md": False,
                "merged": True,
                "independent": True,
            },
            get_merge_convert_submode_fn=lambda: "x",
            should_reuse_office_app_fn=lambda: False,
            get_office_restart_every_fn=lambda: 10,
            print_fn=lambda m: logs.append(m),
        )
        self.assertTrue(any("Runtime Summary" in line for line in logs))

        import office_converter as oc

        original = oc.print_runtime_summary_impl
        try:
            seen = {}

            def _fake(**kwargs):
                seen["kwargs"] = kwargs

            oc.print_runtime_summary_impl = _fake
            dummy = OfficeConverter.__new__(OfficeConverter)
            dummy.config = {}
            dummy.run_mode = "r"
            dummy.merge_mode = "m"
            dummy.content_strategy = "s"
            dummy._get_output_pref = lambda: {}
            dummy._get_merge_convert_submode = lambda: "sub"
            dummy._should_reuse_office_app = lambda: False
            dummy._get_office_restart_every = lambda: 0
            dummy.print_runtime_summary()
            self.assertIn("config", seen.get("kwargs", {}))
        finally:
            oc.print_runtime_summary_impl = original

    def test_new_runtime_modules_have_no_bare_except_exception(self):
        for rel in (
            "converter/target_path.py",
            "converter/pdf_content_scan.py",
            "converter/error_summary.py",
            "converter/mshelp_workflow.py",
            "converter/runtime_summary.py",
        ):
            text = Path(rel).read_text(encoding="utf-8")
            self.assertNotIn("except Exception", text)


if __name__ == "__main__":
    unittest.main()
