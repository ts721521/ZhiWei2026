import unittest
from unittest.mock import patch

from office_converter import OfficeConverter


class SaveConfigBehaviorTests(unittest.TestCase):
    def test_save_config_logs_error_on_unserializable_value(self):
        dummy = OfficeConverter.__new__(OfficeConverter)
        dummy.config_path = "unused.json"
        dummy.config = {"bad": object()}

        with patch("office_converter.logging.error") as log_error:
            dummy.save_config()

        self.assertTrue(log_error.called)
        self.assertIn("failed to save config", str(log_error.call_args[0][0]))


if __name__ == "__main__":
    unittest.main()
