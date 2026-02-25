import unittest
from pathlib import Path


class OfficeConverterModuleTests(unittest.TestCase):
    def test_office_converter_module_has_no_bare_except_exception(self):
        module_text = Path("office_converter.py").read_text(encoding="utf-8")
        self.assertNotIn("except Exception", module_text)


if __name__ == "__main__":
    unittest.main()
