import json
import os
import tempfile
import unittest
from pathlib import Path


class ConverterDefaultConfigModuleSplitTests(unittest.TestCase):
    def test_converter_default_config_module_exists_and_writes_file(self):
        from converter.default_config import create_default_config

        with tempfile.TemporaryDirectory(prefix="cfg_split_") as tmp:
            path = os.path.join(tmp, "config.json")
            self.assertTrue(create_default_config(path))
            with open(path, "r", encoding="utf-8") as f:
                cfg = json.load(f)

        self.assertEqual("classic", cfg.get("app_mode"))
        self.assertIn("ui", cfg)

    def test_office_converter_re_exports_create_default_config(self):
        from converter.default_config import create_default_config as split_create
        from office_converter import create_default_config as office_create

        self.assertIs(split_create, office_create)

    def test_default_config_module_has_no_bare_except_exception(self):
        mod_path = Path(__file__).resolve().parents[1] / "converter" / "default_config.py"
        text = mod_path.read_text(encoding="utf-8")
        self.assertNotIn("except Exception", text)


if __name__ == "__main__":
    unittest.main()
