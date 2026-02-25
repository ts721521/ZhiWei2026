import unittest

from gui.mixins.gui_config_logic_mixin import ConfigLogicMixin


class _DummyConfigLogic(ConfigLogicMixin):
    pass


class ConfigLogicMixinTests(unittest.TestCase):
    def test_safe_positive_int_works_on_instance(self):
        dummy = _DummyConfigLogic()
        self.assertEqual(12, dummy._safe_positive_int("12", 5))
        self.assertEqual(5, dummy._safe_positive_int("-1", 5))

    def test_is_valid_hex_color_works_on_instance(self):
        dummy = _DummyConfigLogic()
        self.assertTrue(dummy._is_valid_hex_color("#A0B1C2"))
        self.assertFalse(dummy._is_valid_hex_color("A0B1C2"))


if __name__ == "__main__":
    unittest.main()

