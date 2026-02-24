import unittest


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


if __name__ == "__main__":
    unittest.main()
