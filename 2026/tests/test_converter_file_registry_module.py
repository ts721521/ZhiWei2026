import unittest


class ConverterFileRegistrySplitTests(unittest.TestCase):
    def test_converter_file_registry_module_exists(self):
        from converter.file_registry import FileRegistry

        reg = FileRegistry(path="", base_root="")
        self.assertEqual("", reg.normalize_path(""))

    def test_office_converter_re_exports_file_registry(self):
        from converter.file_registry import FileRegistry as split_registry
        from office_converter import FileRegistry as office_registry

        self.assertIs(split_registry, office_registry)


if __name__ == "__main__":
    unittest.main()
