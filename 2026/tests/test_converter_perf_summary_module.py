import unittest
from unittest.mock import patch

from office_converter import OfficeConverter


class ConverterPerfSummarySplitTests(unittest.TestCase):
    def test_add_perf_seconds_core_behaviors(self):
        from converter.perf_summary import add_perf_seconds

        perf = {"total_seconds": 1.0}
        add_perf_seconds(perf, "total_seconds", "2.5")
        add_perf_seconds(perf, "missing", 1)
        add_perf_seconds(perf, "total_seconds", -1)
        add_perf_seconds(perf, "total_seconds", "bad")
        self.assertEqual(3.5, perf["total_seconds"])

    def test_converter_perf_summary_module_exists(self):
        from converter.perf_summary import build_perf_summary

        text = build_perf_summary(
            {"total_seconds": 10.0, "scan_seconds": 1.0},
            {"success": 2},
        )
        self.assertIsInstance(text, str)
        self.assertIn("10.00", text)

    def test_office_converter_build_perf_summary_matches_module(self):
        from converter.perf_summary import build_perf_summary

        dummy = OfficeConverter.__new__(OfficeConverter)
        dummy.perf_metrics = {
            "scan_seconds": 1.5,
            "batch_seconds": 2.0,
            "convert_core_seconds": 1.0,
            "pdf_wait_seconds": 0.1,
            "markdown_seconds": 0.2,
            "mshelp_merge_seconds": 0.3,
            "merge_seconds": 0.4,
            "postprocess_seconds": 0.5,
            "total_seconds": 4.0,
        }
        dummy.stats = {"success": 2}

        self.assertEqual(
            build_perf_summary(dummy.perf_metrics, dummy.stats),
            OfficeConverter._build_perf_summary(dummy),
        )

    def test_office_converter_add_perf_seconds_delegates_to_module(self):
        import office_converter as oc

        dummy = OfficeConverter.__new__(OfficeConverter)
        dummy.perf_metrics = {"total_seconds": 1.0}

        with patch.object(oc, "add_perf_seconds_impl") as impl:
            OfficeConverter._add_perf_seconds(dummy, "total_seconds", 3)
            impl.assert_called_once_with(dummy.perf_metrics, "total_seconds", 3)


if __name__ == "__main__":
    unittest.main()
