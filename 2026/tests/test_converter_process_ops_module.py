import unittest
from unittest.mock import Mock, patch

from office_converter import ENGINE_MS, ENGINE_WPS, OfficeConverter


class ConverterProcessOpsSplitTests(unittest.TestCase):
    def test_converter_process_ops_module_exists(self):
        from converter.process_ops import (
            kill_process_by_name_for_converter,
            process_names_for_engine,
        )

        both = process_names_for_engine(None)
        self.assertIn("winword", both)
        self.assertIn("wps", both)

        ms_only = process_names_for_engine(ENGINE_MS)
        self.assertIn("excel", ms_only)
        self.assertNotIn("wps", ms_only)

        wps_only = process_names_for_engine(ENGINE_WPS)
        self.assertIn("wps", wps_only)
        self.assertNotIn("winword", wps_only)

        self.assertTrue(
            kill_process_by_name_for_converter(
                "winword",
                has_win32=True,
                run_cmd=lambda *_a, **_k: None,
            )
        )
        self.assertFalse(
            kill_process_by_name_for_converter(
                "winword",
                has_win32=True,
                run_cmd=lambda *_a, **_k: (_ for _ in ()).throw(OSError("x")),
            )
        )

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
        with patch.object(oc, "kill_process_by_name_for_converter", return_value=True) as kill_impl:
            dummy = OfficeConverter.__new__(OfficeConverter)
            out = OfficeConverter._kill_process_by_name(dummy, "winword")
            self.assertTrue(out)
            kill_impl.assert_called_once_with(
                "winword",
                has_win32=oc.HAS_WIN32,
                run_cmd=oc.subprocess.run,
            )


if __name__ == "__main__":
    unittest.main()
