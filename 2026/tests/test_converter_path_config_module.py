import os
import unittest
from unittest.mock import patch

from office_converter import OfficeConverter


class ConverterPathConfigSplitTests(unittest.TestCase):
    def test_converter_path_config_module_exists(self):
        from converter.path_config import get_path_from_config

        cfg = {
            "source_folder": "C:\\Base",
            "source_folder_win": "D:\\WinPath",
        }
        path = get_path_from_config(cfg, "source_folder", prefer_win=True, prefer_mac=False)
        self.assertEqual(os.path.abspath("D:\\WinPath"), path)

    def test_office_converter_method_matches_path_config_module(self):
        from converter.path_config import get_path_from_config

        dummy = OfficeConverter.__new__(OfficeConverter)
        dummy.config = {
            "target_folder": "C:\\TargetBase",
            "target_folder_win": "E:\\TargetWin",
        }
        expected = get_path_from_config(
            dummy.config, "target_folder", prefer_win=True, prefer_mac=False
        )

        import office_converter as oc

        with patch.object(oc, "is_win", return_value=True), patch.object(
            oc, "is_mac", return_value=False
        ):
            self.assertEqual(
                expected, OfficeConverter._get_path_from_config(dummy, "target_folder")
            )


if __name__ == "__main__":
    unittest.main()
