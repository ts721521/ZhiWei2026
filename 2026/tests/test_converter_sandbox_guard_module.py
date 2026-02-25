import types
import unittest
from pathlib import Path

from office_converter import OfficeConverter


class ConverterSandboxGuardSplitTests(unittest.TestCase):
    def test_sandbox_guard_core_behaviors(self):
        from converter.sandbox_guard import check_sandbox_free_space_or_raise

        logs = {"info": [], "warn": [], "print": []}

        # enough free space
        check_sandbox_free_space_or_raise(
            {"enable_sandbox": True, "sandbox_min_free_gb": 1, "temp_sandbox_root": "C:\\x"},
            exists_fn=lambda _p: True,
            splitdrive_fn=lambda p: ("C:", p),
            getcwd_fn=lambda: "C:\\cwd",
            disk_usage_fn=lambda _p: types.SimpleNamespace(free=5 * 1024 * 1024 * 1024),
            log_info_fn=lambda m: logs["info"].append(m),
            log_warning_fn=lambda m: logs["warn"].append(m),
            print_fn=lambda m: logs["print"].append(m),
        )
        self.assertTrue(logs["info"])

        # low space + block
        with self.assertRaises(RuntimeError):
            check_sandbox_free_space_or_raise(
                {
                    "enable_sandbox": True,
                    "sandbox_min_free_gb": 10,
                    "temp_sandbox_root": "D:\\x",
                    "sandbox_low_space_policy": "block",
                },
                exists_fn=lambda _p: True,
                splitdrive_fn=lambda p: ("D:", p),
                getcwd_fn=lambda: "D:\\cwd",
                disk_usage_fn=lambda _p: types.SimpleNamespace(free=1 * 1024 * 1024 * 1024),
                log_info_fn=lambda m: logs["info"].append(m),
                log_warning_fn=lambda m: logs["warn"].append(m),
                print_fn=lambda m: logs["print"].append(m),
            )

    def test_office_converter_sandbox_guard_delegates_to_module(self):
        import office_converter as oc

        original = oc.check_sandbox_free_space_or_raise_impl
        try:
            seen = {}

            def _fake(config, **kwargs):
                seen["config"] = config
                seen["kwargs"] = kwargs
                return "ok"

            oc.check_sandbox_free_space_or_raise_impl = _fake
            dummy = OfficeConverter.__new__(OfficeConverter)
            dummy.config = {"a": 1}

            out = dummy._check_sandbox_free_space_or_raise()
            self.assertEqual(out, "ok")
            self.assertEqual(seen["config"], {"a": 1})
        finally:
            oc.check_sandbox_free_space_or_raise_impl = original

    def test_sandbox_guard_module_has_no_bare_except_exception(self):
        mod_path = Path(__file__).resolve().parents[1] / "converter" / "sandbox_guard.py"
        text = mod_path.read_text(encoding="utf-8")
        self.assertNotIn("except Exception", text)


if __name__ == "__main__":
    unittest.main()
