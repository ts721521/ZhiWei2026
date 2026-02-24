import unittest
from unittest.mock import Mock, patch

from office_converter import ENGINE_MS, ENGINE_WPS, OfficeConverter


class ConverterProcessOpsSplitTests(unittest.TestCase):
    def test_converter_process_ops_module_exists(self):
        from converter.process_ops import process_names_for_engine

        both = process_names_for_engine(None)
        self.assertIn("winword", both)
        self.assertIn("wps", both)

        ms_only = process_names_for_engine(ENGINE_MS)
        self.assertIn("excel", ms_only)
        self.assertNotIn("wps", ms_only)

        wps_only = process_names_for_engine(ENGINE_WPS)
        self.assertIn("wps", wps_only)
        self.assertNotIn("winword", wps_only)

    def test_office_converter_cleanup_uses_process_ops_behavior(self):
        from converter.process_ops import process_names_for_engine

        dummy = OfficeConverter.__new__(OfficeConverter)
        dummy.engine_type = ENGINE_MS
        dummy._kill_process_by_name = Mock()

        OfficeConverter.cleanup_all_processes(dummy)
        expected = process_names_for_engine(ENGINE_MS)
        self.assertEqual(len(expected), dummy._kill_process_by_name.call_count)

    def test_office_converter_kill_process_delegates_to_process_ops(self):
        import office_converter as oc
        from converter.process_ops import kill_process_by_name

        with patch.object(oc, "HAS_WIN32", True), patch.object(
            oc.subprocess, "run"
        ) as run_mock:
            dummy = OfficeConverter.__new__(OfficeConverter)
            OfficeConverter._kill_process_by_name(dummy, "winword")
            kill_process_by_name("winword", has_win32=True, run_cmd=run_mock)
            self.assertGreaterEqual(run_mock.call_count, 2)


if __name__ == "__main__":
    unittest.main()
