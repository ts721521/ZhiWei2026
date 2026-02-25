import unittest
from pathlib import Path


class ConverterPlatformUtilsSplitTests(unittest.TestCase):
    def test_converter_platform_utils_module_exists(self):
        from converter.platform_utils import get_app_path, is_mac, is_win

        self.assertIsInstance(get_app_path(), str)
        self.assertIsInstance(is_win(), bool)
        self.assertIsInstance(is_mac(), bool)

    def test_office_converter_re_exports_platform_utils(self):
        from converter.platform_utils import get_app_path as split_get_app_path
        from office_converter import get_app_path as office_get_app_path

        self.assertIs(split_get_app_path, office_get_app_path)

    def test_platform_utils_module_has_no_bare_except_exception(self):
        mod_path = Path(__file__).resolve().parents[1] / "converter" / "platform_utils.py"
        text = mod_path.read_text(encoding="utf-8")
        self.assertNotIn("except Exception", text)


if __name__ == "__main__":
    unittest.main()
