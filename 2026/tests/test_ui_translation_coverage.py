import glob
import os
import re
import unittest

from ui_translations import TRANSLATIONS


_TEST_DIR = os.path.dirname(os.path.abspath(__file__))
_ROOT = os.path.dirname(_TEST_DIR)


class UiTranslationCoverageTests(unittest.TestCase):
    def test_all_tr_keys_exist_in_zh_and_en_maps(self):
        pattern = re.compile(r"\btr\(\s*['\"]([^'\"]+)['\"]\s*\)")
        used_keys = set()

        for path in glob.glob(os.path.join(_ROOT, "*.py")):
            with open(path, "r", encoding="utf-8") as f:
                used_keys.update(pattern.findall(f.read()))

        zh_keys = set((TRANSLATIONS.get("zh") or {}).keys())
        en_keys = set((TRANSLATIONS.get("en") or {}).keys())

        missing_zh = sorted(k for k in used_keys if k not in zh_keys)
        missing_en = sorted(k for k in used_keys if k not in en_keys)

        self.assertEqual([], missing_zh, f"Missing zh translation keys: {missing_zh}")
        self.assertEqual([], missing_en, f"Missing en translation keys: {missing_en}")


if __name__ == "__main__":
    unittest.main()
