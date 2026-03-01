import unittest
from pathlib import Path

from office_converter import OfficeConverter


class ConverterCliDisplaySplitTests(unittest.TestCase):
    def test_cli_wizard_flow_core_behaviors(self):
        from converter.cli_wizard_flow import run_cli_wizard

        calls = []
        state = {"run_mode": "convert_then_merge"}
        run_cli_wizard(
            interactive=True,
            print_welcome_fn=lambda: calls.append("welcome"),
            confirm_config_in_terminal_fn=lambda: calls.append("confirm"),
            ask_for_subfolder_fn=lambda: calls.append("subfolder"),
            select_run_mode_fn=lambda: calls.append("run_mode"),
            get_run_mode_fn=lambda: state["run_mode"],
            select_collect_mode_fn=lambda: calls.append("collect_mode"),
            select_content_strategy_fn=lambda: calls.append("content_strategy"),
            select_merge_mode_fn=lambda: calls.append("merge_mode"),
            select_engine_mode_fn=lambda: calls.append("engine_mode"),
            check_and_handle_running_processes_fn=lambda: calls.append("proc_check"),
            init_paths_from_config_fn=lambda: calls.append("init_paths"),
            config={"enable_merge": True},
            mode_collect_only="collect_only",
            mode_convert_only="convert_only",
            mode_convert_then_merge="convert_then_merge",
            mode_merge_only="merge_only",
            mode_mshelp_only="mshelp_only",
        )
        self.assertIn("content_strategy", calls)
        self.assertIn("merge_mode", calls)
        self.assertIn("engine_mode", calls)
        self.assertEqual("init_paths", calls[-1])

    def test_display_helpers_core_behaviors(self):
        from converter.display_helpers import print_step_title, print_welcome, safe_console_print

        lines = []
        print_welcome(app_version="x", config_path="cfg.json", print_fn=lambda m: lines.append(m))
        print_step_title("STEP", print_fn=lambda m: lines.append(m))
        safe_console_print("ok", print_fn=lambda m, **_k: lines.append(m))
        self.assertTrue(any("Config file:" in str(v) for v in lines))
        self.assertTrue(any("STEP" == str(v) for v in lines))

        captured = []
        state = {"n": 0}

        def flaky_print(msg, **_kwargs):
            state["n"] += 1
            if state["n"] == 1:
                raise OSError(22, "Invalid argument")
            captured.append(msg)

        safe_console_print("中文🙂", print_fn=flaky_print)
        self.assertEqual(2, state["n"])
        self.assertTrue(captured)

    def test_office_converter_methods_delegate(self):
        import office_converter as oc

        originals = (
            oc.run_cli_wizard_impl,
            oc.print_welcome_impl,
            oc.print_step_title_impl,
        )
        try:
            oc.run_cli_wizard_impl = lambda **_k: "wiz"
            oc.print_welcome_impl = lambda **_k: "wel"
            oc.print_step_title_impl = lambda *_a, **_k: "step"

            dummy = OfficeConverter.__new__(OfficeConverter)
            dummy.interactive = True
            dummy.config = {"enable_merge": True}
            dummy.run_mode = "convert_only"
            dummy.config_path = "cfg.json"
            dummy.confirm_config_in_terminal = lambda: None
            dummy.ask_for_subfolder = lambda: None
            dummy.select_run_mode = lambda: None
            dummy.select_collect_mode = lambda: None
            dummy.select_content_strategy = lambda: None
            dummy.select_merge_mode = lambda: None
            dummy.select_engine_mode = lambda: None
            dummy.check_and_handle_running_processes = lambda: None
            dummy._init_paths_from_config = lambda: None

            self.assertEqual("wel", dummy.print_welcome())
            self.assertEqual("step", dummy.print_step_title("X"))
            self.assertEqual("wiz", dummy.cli_wizard())
        finally:
            (
                oc.run_cli_wizard_impl,
                oc.print_welcome_impl,
                oc.print_step_title_impl,
            ) = originals

    def test_new_modules_have_no_bare_except_exception(self):
        for rel in (
            "converter/cli_wizard_flow.py",
            "converter/display_helpers.py",
        ):
            text = Path(rel).read_text(encoding="utf-8")
            self.assertNotIn("except Exception", text)


if __name__ == "__main__":
    unittest.main()
