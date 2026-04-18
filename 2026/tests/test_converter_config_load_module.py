import unittest
from pathlib import Path

from office_converter import OfficeConverter


class ConverterConfigLoadSplitTests(unittest.TestCase):
    def test_config_load_core_behaviors(self):
        from converter.config_load import load_config

        class Dummy:
            def __init__(self):
                self.config = {}
                self.defaults_called = 0

            def _get_path_from_config(self, key):
                mapping = {
                    "source_folder": "C:/src",
                    "target_folder": "C:/target",
                    "obsidian_root": "C:/obs",
                }
                return mapping[key]

            def _apply_config_defaults(self):
                self.defaults_called += 1
                self.config.update(
                    {
                        "run_mode": "convert_then_merge",
                        "collect_mode": "copy_and_index",
                        "collect_copy_layout": "preserve_tree",
                        "content_strategy": "standard",
                        "default_engine": "ask",
                        "kill_process_mode": "ask",
                        "merge_mode": "category_split",
                        "merge_convert_submode": "merge_only",
                        "timeout_seconds": 60,
                        "pdf_wait_seconds": 15,
                        "ppt_timeout_seconds": 60,
                        "ppt_pdf_wait_seconds": 15,
                        "parallel_workers": 4,
                        "parallel_checkpoint_interval": 10,
                        "enable_parallel_conversion": False,
                        "enable_merge": True,
                        "output_enable_pdf": True,
                        "output_enable_md": True,
                        "output_enable_merged": True,
                        "output_enable_independent": False,
                        "enable_fast_md_engine": False,
                        "excluded_folders": [],
                        "price_keywords": [],
                        "allowed_extensions": {},
                    }
                )

        d = Dummy()
        load_config(
            d,
            "cfg.json",
            open_fn=lambda *_a, **_k: __import__("io").StringIO(
                '{"source_folders":[" C:/s1 ",""],"x":1}'
            ),
            json_loads_fn=__import__("json").loads,
            abspath_fn=lambda p: f"ABS::{p}",
            print_fn=lambda _m: None,
            exit_fn=lambda _c: None,
        )
        self.assertEqual(d.config["source_folder"], "ABS::C:/s1")
        self.assertEqual(d.config["target_folder"], "C:/target")
        self.assertEqual(d.defaults_called, 1)

    def test_config_load_exits_on_invalid_schema(self):
        from converter.config_load import load_config

        class Dummy:
            def __init__(self):
                self.config = {}

            def _get_path_from_config(self, key):
                mapping = {
                    "source_folder": "C:/src",
                    "target_folder": "C:/target",
                    "obsidian_root": "C:/obs",
                }
                return mapping[key]

            def _apply_config_defaults(self):
                return None

        d = Dummy()
        seen = {"code": None, "msg": ""}

        def _exit(code):
            seen["code"] = code

        def _print(msg):
            seen["msg"] = str(msg)

        load_config(
            d,
            "cfg.json",
            open_fn=lambda *_a, **_k: __import__("io").StringIO(
                '{"source_folders":["C:/s1"],"run_mode":"bad_mode"}'
            ),
            json_loads_fn=__import__("json").loads,
            abspath_fn=lambda p: p,
            print_fn=_print,
            exit_fn=_exit,
        )
        self.assertEqual(1, seen["code"])
        self.assertIn("Invalid config schema", seen["msg"])

    def test_config_load_exits_on_invalid_json(self):
        from converter.config_load import load_config

        class Dummy:
            def __init__(self):
                self.config = {}

            def _get_path_from_config(self, _key):
                return ""

            def _apply_config_defaults(self):
                return None

        d = Dummy()
        seen = {"code": None, "msg": ""}

        load_config(
            d,
            "cfg.json",
            open_fn=lambda *_a, **_k: __import__("io").StringIO("{bad_json"),
            json_loads_fn=__import__("json").loads,
            abspath_fn=lambda p: p,
            print_fn=lambda msg: seen.__setitem__("msg", str(msg)),
            exit_fn=lambda code: seen.__setitem__("code", code),
        )
        self.assertEqual(1, seen["code"])
        self.assertIn("Invalid JSON", seen["msg"])

    def test_office_converter_load_config_delegates_to_module(self):
        import office_converter as oc

        original = oc.load_config_impl
        try:
            seen = {}

            def _fake(converter, path, **kwargs):
                seen["converter"] = converter
                seen["path"] = path
                seen["kwargs"] = kwargs
                return "ok"

            oc.load_config_impl = _fake
            dummy = OfficeConverter.__new__(OfficeConverter)
            out = dummy.load_config("x.json")
            self.assertEqual(out, "ok")
            self.assertIs(seen["converter"], dummy)
            self.assertEqual(seen["path"], "x.json")
        finally:
            oc.load_config_impl = original

    def test_config_load_module_has_no_bare_except_exception(self):
        mod_path = Path(__file__).resolve().parents[1] / "converter" / "config_load.py"
        text = mod_path.read_text(encoding="utf-8")
        self.assertNotIn("except Exception", text)


if __name__ == "__main__":
    unittest.main()
