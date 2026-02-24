import unittest

from office_converter import OfficeConverter


class ConverterRuntimePrefsSplitTests(unittest.TestCase):
    def test_converter_runtime_prefs_module_exists(self):
        from converter.runtime_prefs import get_merge_convert_submode, get_output_pref

        cfg = {
            "output_enable_pdf": True,
            "output_enable_md": False,
            "output_enable_merged": True,
            "output_enable_independent": False,
            "merge_convert_submode": "pdf_to_md",
        }
        pref = get_output_pref(cfg)
        self.assertTrue(pref["pdf"])
        self.assertFalse(pref["md"])
        self.assertEqual("pdf_to_md", get_merge_convert_submode(cfg))

    def test_office_converter_methods_match_runtime_prefs_module(self):
        from converter.runtime_prefs import get_merge_convert_submode, get_output_pref

        dummy = OfficeConverter.__new__(OfficeConverter)
        dummy.config = {
            "output_enable_pdf": False,
            "output_enable_md": True,
            "output_enable_merged": False,
            "output_enable_independent": True,
            "merge_convert_submode": "invalid_submode",
        }
        self.assertEqual(get_output_pref(dummy.config), dummy._get_output_pref())
        self.assertEqual(
            get_merge_convert_submode(dummy.config), dummy._get_merge_convert_submode()
        )


if __name__ == "__main__":
    unittest.main()
