import unittest

from office_converter import OfficeConverter


class ConverterConfigTerminalSplitTests(unittest.TestCase):
    def test_config_terminal_core_behaviors(self):
        from converter.config_terminal import confirm_config_in_terminal

        class Dummy:
            def __init__(self):
                self.config = {
                    "source_folder": "C:/old/source",
                    "target_folder": "C:/old/target",
                }
                self.step_titles = []
                self.saved = 0

            def print_step_title(self, text):
                self.step_titles.append(text)

            def save_config(self):
                self.saved += 1

        dummy = Dummy()
        answers = iter(["y", "new_source", "new_target", "n"])
        logs = []
        confirm_config_in_terminal(
            dummy,
            input_fn=lambda _p: next(answers),
            print_fn=lambda msg: logs.append(str(msg)),
            abspath_fn=lambda p: f"ABS::{p}",
        )
        self.assertIn("Step 1/4: Confirm Source and Target", dummy.step_titles)
        self.assertEqual(dummy.config["source_folder"], "ABS::new_source")
        self.assertEqual(dummy.config["target_folder"], "ABS::new_target")
        self.assertEqual(dummy.saved, 1)
        self.assertTrue(any("Config saved." in line for line in logs))

    def test_office_converter_confirm_config_delegates_to_module(self):
        import office_converter as oc

        original = oc.confirm_config_in_terminal_impl
        try:
            seen = {}

            def _fake(converter, **kwargs):
                seen["converter"] = converter
                seen["kwargs"] = kwargs
                return "ok"

            oc.confirm_config_in_terminal_impl = _fake
            dummy = OfficeConverter.__new__(OfficeConverter)
            out = dummy.confirm_config_in_terminal()
            self.assertEqual(out, "ok")
            self.assertIs(seen["converter"], dummy)
        finally:
            oc.confirm_config_in_terminal_impl = original


if __name__ == "__main__":
    unittest.main()
