import unittest


class ConverterConstantsModuleSplitTests(unittest.TestCase):
    def test_converter_constants_module_exists_with_expected_modes(self):
        from converter.constants import MODE_CONVERT_THEN_MERGE, MODE_MSHELP_ONLY

        self.assertEqual("convert_then_merge", MODE_CONVERT_THEN_MERGE)
        self.assertEqual("mshelp_only", MODE_MSHELP_ONLY)

    def test_office_converter_re_exports_mode_constants(self):
        from converter.constants import MODE_CONVERT_ONLY as split_mode
        from office_converter import MODE_CONVERT_ONLY as office_mode

        self.assertEqual(split_mode, office_mode)


if __name__ == "__main__":
    unittest.main()
