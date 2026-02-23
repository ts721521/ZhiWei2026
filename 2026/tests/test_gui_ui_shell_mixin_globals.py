import unittest

from gui_ui_shell_mixin import UIShellMixin, tb


class UIShellMixinGlobalBindingsTests(unittest.TestCase):
    def test_finish_init_dependencies_are_imported(self):
        globals_map = UIShellMixin._finish_init.__globals__
        self.assertIn("os", globals_map)
        self.assertIn("create_default_config", globals_map)

    def test_build_ui_dependencies_are_imported(self):
        globals_map = UIShellMixin._build_ui.__globals__
        self.assertIn("sys", globals_map)
        self.assertIn("HAS_TTKBOOTSTRAP", globals_map)

    def test_tb_namespace_exposes_radiobutton(self):
        self.assertTrue(
            hasattr(tb, "Radiobutton"),
            "Fallback tb namespace must provide Radiobutton for app mode selector",
        )


if __name__ == "__main__":
    unittest.main()
