import unittest
from pathlib import Path


class ConverterFileRegistrySplitTests(unittest.TestCase):
    def test_converter_file_registry_module_exists(self):
        from converter.file_registry import FileRegistry

        reg = FileRegistry(path="", base_root="")
        self.assertEqual("", reg.normalize_path(""))

    def test_office_converter_re_exports_file_registry(self):
        from converter.file_registry import FileRegistry as split_registry
        from office_converter import FileRegistry as office_registry

        self.assertIs(split_registry, office_registry)

    def test_file_registry_module_has_no_bare_except_exception(self):
        mod_path = Path(__file__).resolve().parents[1] / "converter" / "file_registry.py"
        text = mod_path.read_text(encoding="utf-8")
        self.assertNotIn("except Exception", text)


if __name__ == "__main__":
    unittest.main()
