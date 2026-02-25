import unittest
from pathlib import Path

from office_converter import OfficeConverter


class ConverterInteractiveChoicesSplitTests(unittest.TestCase):
    def test_interactive_choices_core_behaviors(self):
        from converter.interactive_choices import (
            ask_for_subfolder,
            select_collect_mode,
            select_content_strategy,
            select_engine_mode,
            select_merge_mode,
            select_run_mode,
        )

        cfg = {"target_folder": r"C:\out"}
        ask_for_subfolder(
            cfg,
            print_step_title_fn=lambda _m: None,
            print_fn=lambda _m: None,
            input_fn=lambda _p: "a:b",
            abspath_fn=lambda p: p,
        )
        self.assertEqual(cfg["target_folder"], r"C:\out\ab")

        run_mode = select_run_mode(
            print_step_title_fn=lambda _m: None,
            get_readable_run_mode_fn=lambda m: m,
            mode_convert_only="convert_only",
            mode_merge_only="merge_only",
            mode_convert_then_merge="convert_then_merge",
            mode_collect_only="collect_only",
            mode_mshelp_only="mshelp_only",
            print_fn=lambda _m: None,
            input_fn=lambda _p: "5",
        )
        self.assertEqual(run_mode, "mshelp_only")

        collect_mode = select_collect_mode(
            print_step_title_fn=lambda _m: None,
            get_readable_collect_mode_fn=lambda m: m,
            collect_mode_copy_and_index="copy_and_index",
            collect_mode_index_only="index_only",
            print_fn=lambda _m: None,
            input_fn=lambda _p: "2",
        )
        self.assertEqual(collect_mode, "index_only")

        merge_mode = select_merge_mode(
            {"enable_merge": True, "merge_mode": "unknown"},
            print_step_title_fn=lambda _m: None,
            get_readable_merge_mode_fn=lambda m: m,
            merge_mode_category="category",
            merge_mode_all_in_one="all_in_one",
            print_fn=lambda _m: None,
            input_fn=lambda _p: "2",
        )
        self.assertEqual(merge_mode, "all_in_one")

        strategy = select_content_strategy(
            ["price"],
            print_step_title_fn=lambda _m: None,
            get_readable_content_strategy_fn=lambda m: m,
            strategy_standard="standard",
            strategy_smart_tag="smart_tag",
            strategy_price_only="price_only",
            print_fn=lambda _m: None,
            input_fn=lambda _p: "3",
        )
        self.assertEqual(strategy, "price_only")

        engine = select_engine_mode(
            {"default_engine": "ask"},
            print_step_title_fn=lambda _m: None,
            get_readable_engine_type_fn=lambda m: m,
            engine_ask="ask",
            engine_wps="wps",
            engine_ms="ms",
            print_fn=lambda _m: None,
            input_fn=lambda _p: "2",
        )
        self.assertEqual(engine, "ms")

    def test_office_converter_interactive_choice_methods_delegate_to_module(self):
        import office_converter as oc

        originals = (
            oc.ask_for_subfolder_impl,
            oc.select_run_mode_impl,
            oc.select_collect_mode_impl,
            oc.select_merge_mode_impl,
            oc.select_content_strategy_impl,
            oc.select_engine_mode_impl,
        )
        try:
            seen = {}

            oc.ask_for_subfolder_impl = lambda *a, **k: seen.setdefault("sub", True)
            oc.select_run_mode_impl = lambda *a, **k: "run"
            oc.select_collect_mode_impl = lambda *a, **k: "collect"
            oc.select_merge_mode_impl = lambda *a, **k: "merge"
            oc.select_content_strategy_impl = lambda *a, **k: "strategy"
            oc.select_engine_mode_impl = lambda *a, **k: "engine"

            dummy = OfficeConverter.__new__(OfficeConverter)
            dummy.config = {"default_engine": "ask"}
            dummy.price_keywords = []
            dummy.run_mode = ""
            dummy.collect_mode = ""
            dummy.merge_mode = ""
            dummy.content_strategy = ""
            dummy.engine_type = ""
            dummy.print_step_title = lambda _m: None

            dummy.ask_for_subfolder()
            dummy.select_run_mode()
            dummy.select_collect_mode()
            dummy.select_merge_mode()
            dummy.select_content_strategy()
            dummy.select_engine_mode()

            self.assertTrue(seen.get("sub"))
            self.assertEqual(dummy.run_mode, "run")
            self.assertEqual(dummy.collect_mode, "collect")
            self.assertEqual(dummy.merge_mode, "merge")
            self.assertEqual(dummy.content_strategy, "strategy")
            self.assertEqual(dummy.engine_type, "engine")
        finally:
            (
                oc.ask_for_subfolder_impl,
                oc.select_run_mode_impl,
                oc.select_collect_mode_impl,
                oc.select_merge_mode_impl,
                oc.select_content_strategy_impl,
                oc.select_engine_mode_impl,
            ) = originals

    def test_interactive_choices_module_has_no_bare_except_exception(self):
        module_text = Path("converter/interactive_choices.py").read_text(
            encoding="utf-8"
        )
        self.assertNotIn("except Exception", module_text)


if __name__ == "__main__":
    unittest.main()
