import unittest

from office_converter import OfficeConverter


class ConverterPerfSummarySplitTests(unittest.TestCase):
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


if __name__ == "__main__":
    unittest.main()
