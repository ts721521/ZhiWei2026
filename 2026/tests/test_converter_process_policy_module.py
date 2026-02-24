import unittest
from unittest.mock import Mock

from office_converter import (
    KILL_MODE_AUTO,
    KILL_MODE_KEEP,
    MODE_CONVERT_ONLY,
    MODE_MERGE_ONLY,
    OfficeConverter,
)


class ConverterProcessPolicySplitTests(unittest.TestCase):
    def test_converter_process_policy_module_exists(self):
        from converter.process_policy import resolve_process_handling

        keep = resolve_process_handling(
            run_mode=MODE_CONVERT_ONLY,
            kill_process_mode=KILL_MODE_KEEP,
            interactive=True,
        )
        self.assertTrue(keep["reuse_process"])
        self.assertFalse(keep["cleanup_all"])

        skip = resolve_process_handling(
            run_mode=MODE_MERGE_ONLY,
            kill_process_mode=KILL_MODE_AUTO,
            interactive=True,
        )
        self.assertTrue(skip["skip"])

    def test_office_converter_check_and_handle_running_processes_matches_policy(self):
        from converter.process_policy import resolve_process_handling

        dummy = OfficeConverter.__new__(OfficeConverter)
        dummy.run_mode = MODE_CONVERT_ONLY
        dummy.config = {"kill_process_mode": KILL_MODE_AUTO}
        dummy.interactive = True
        dummy.reuse_process = True
        dummy.cleanup_all_processes = Mock()

        expected = resolve_process_handling(
            run_mode=dummy.run_mode,
            kill_process_mode=dummy.config.get("kill_process_mode"),
            interactive=dummy.interactive,
        )
        OfficeConverter.check_and_handle_running_processes(dummy)

        self.assertEqual(expected["reuse_process"], dummy.reuse_process)
        if expected["cleanup_all"]:
            dummy.cleanup_all_processes.assert_called_once()
        else:
            dummy.cleanup_all_processes.assert_not_called()


if __name__ == "__main__":
    unittest.main()
