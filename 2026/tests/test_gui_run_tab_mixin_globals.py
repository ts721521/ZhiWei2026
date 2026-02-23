import unittest

from gui_run_tab_mixin import RunTabUIMixin


class RunTabUIMixinGlobalBindingsTests(unittest.TestCase):
    def test_build_run_tab_content_imports_datetime(self):
        globals_map = RunTabUIMixin._build_run_tab_content.__globals__
        self.assertIn(
            "datetime",
            globals_map,
            "Run tab builder uses datetime.now() and must import datetime",
        )


if __name__ == "__main__":
    unittest.main()
