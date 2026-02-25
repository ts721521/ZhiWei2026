import unittest
from pathlib import Path
from unittest.mock import patch

import office_converter as oc
from office_converter import OfficeConverter


class ConverterOfficeCycleSplitTests(unittest.TestCase):
    def test_converter_office_cycle_module_exists(self):
        from converter.office_cycle import (
            get_app_type_for_ext,
            get_office_restart_every,
            should_reuse_office_app,
        )

        cfg = {
            "office_reuse_app": True,
            "office_restart_every_n_files": "3",
            "allowed_extensions": {
                "word": [".doc", ".docx"],
                "excel": [".xls", ".xlsx"],
                "powerpoint": [".ppt", ".pptx"],
            },
        }
        self.assertTrue(
            should_reuse_office_app(cfg, has_win32=True, is_mac_platform=False)
        )
        self.assertEqual(3, get_office_restart_every(cfg))
        self.assertEqual("word", get_app_type_for_ext(cfg, ".docx"))

    def test_office_converter_methods_match_office_cycle_module(self):
        from converter.office_cycle import (
            get_app_type_for_ext,
            get_office_restart_every,
            should_reuse_office_app,
        )

        dummy = OfficeConverter.__new__(OfficeConverter)
        dummy.config = {
            "office_reuse_app": True,
            "office_restart_every_n_files": "5",
            "allowed_extensions": {
                "word": [".doc", ".docx"],
                "excel": [".xls", ".xlsx"],
                "powerpoint": [".ppt", ".pptx"],
            },
        }

        with patch.object(oc, "HAS_WIN32", True), patch.object(
            oc, "is_mac", return_value=False
        ):
            self.assertEqual(
                should_reuse_office_app(
                    dummy.config, has_win32=True, is_mac_platform=False
                ),
                dummy._should_reuse_office_app(),
            )

        self.assertEqual(get_office_restart_every(dummy.config), dummy._get_office_restart_every())
        self.assertEqual(
            get_app_type_for_ext(dummy.config, ".xlsx"),
            dummy._get_app_type_for_ext(".xlsx"),
        )

    def test_office_cycle_module_has_no_bare_except_exception(self):
        mod_path = Path(__file__).resolve().parents[1] / "converter" / "office_cycle.py"
        text = mod_path.read_text(encoding="utf-8")
        self.assertNotIn("except Exception", text)


if __name__ == "__main__":
    unittest.main()
