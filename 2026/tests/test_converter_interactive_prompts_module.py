import unittest

from office_converter import OfficeConverter


class ConverterInteractivePromptsSplitTests(unittest.TestCase):
    def test_interactive_prompts_core_behaviors(self):
        from converter.interactive_prompts import confirm_continue_missing_md_merge

        self.assertTrue(
            confirm_continue_missing_md_merge(True, input_fn=lambda _: "y")
        )
        self.assertTrue(
            confirm_continue_missing_md_merge(True, input_fn=lambda _: "")
        )
        self.assertFalse(
            confirm_continue_missing_md_merge(True, input_fn=lambda _: "n")
        )
        self.assertFalse(
            confirm_continue_missing_md_merge(True, input_fn=lambda _: (_ for _ in ()).throw(RuntimeError("x")))
        )

        warnings = []
        self.assertTrue(
            confirm_continue_missing_md_merge(
                False,
                input_fn=lambda _: "n",
                warn_func=lambda m: warnings.append(m),
            )
        )
        self.assertEqual(len(warnings), 1)

    def test_office_converter_prompt_method_delegates_to_module(self):
        dummy = OfficeConverter.__new__(OfficeConverter)
        dummy.interactive = False
        self.assertTrue(dummy._confirm_continue_missing_md_merge())


if __name__ == "__main__":
    unittest.main()
