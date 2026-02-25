import unittest

from gui.mixins.gui_locator_mixin import LocatorMixin


class LocatorMixinGlobalBindingsTests(unittest.TestCase):
    def test_refresh_locator_maps_imports_glob(self):
        globals_map = LocatorMixin.refresh_locator_maps.__globals__
        self.assertIn(
            "glob",
            globals_map,
            "Locator refresh uses glob.glob() and must import glob",
        )


if __name__ == "__main__":
    unittest.main()

