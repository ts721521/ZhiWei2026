import unittest

from gui.mixins.gui_runtime_status_mixin import RuntimeStatusMixin
from office_converter import MODE_CONVERT_ONLY


class _DummyRuntimeStatus(RuntimeStatusMixin):
    def tr(self, key):
        return key


class RuntimeStatusMixinTests(unittest.TestCase):
    def test_sanitize_runtime_config_for_fast_md(self):
        dummy = _DummyRuntimeStatus()
        cfg = {
            "enable_fast_md_engine": True,
            "output_enable_pdf": True,
            "output_enable_md": False,
            "output_enable_merged": True,
            "output_enable_independent": False,
            "enable_parallel_conversion": True,
            "merge_mode": "category",
            "max_merge_size_mb": 80,
        }
        msgs = dummy._sanitize_runtime_config_for_mode(cfg, MODE_CONVERT_ONLY)
        self.assertFalse(cfg["output_enable_pdf"])
        self.assertTrue(cfg["output_enable_md"])
        self.assertFalse(cfg["output_enable_merged"])
        self.assertTrue(cfg["output_enable_independent"])
        self.assertFalse(cfg["enable_parallel_conversion"])
        self.assertTrue(any("Fast MD" in m for m in msgs))


if __name__ == "__main__":
    unittest.main()

