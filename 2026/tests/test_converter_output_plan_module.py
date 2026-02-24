import unittest

from office_converter import MODE_CONVERT_THEN_MERGE, OfficeConverter


class ConverterOutputPlanSplitTests(unittest.TestCase):
    def test_converter_output_plan_module_exists(self):
        from converter.output_plan import compute_convert_output_plan

        plan = compute_convert_output_plan(
            MODE_CONVERT_THEN_MERGE,
            {
                "output_enable_pdf": True,
                "output_enable_md": True,
                "output_enable_merged": True,
                "output_enable_independent": False,
                "enable_merge": True,
            },
        )
        self.assertTrue(plan["need_final_pdf"])
        self.assertTrue(plan["need_markdown"])

    def test_office_converter_uses_same_output_plan_logic(self):
        from converter.output_plan import compute_convert_output_plan

        cfg = {
            "output_enable_pdf": False,
            "output_enable_md": True,
            "output_enable_merged": True,
            "output_enable_independent": False,
            "enable_merge": True,
        }
        self.assertEqual(
            compute_convert_output_plan(MODE_CONVERT_THEN_MERGE, cfg),
            OfficeConverter.compute_convert_output_plan(MODE_CONVERT_THEN_MERGE, cfg),
        )


if __name__ == "__main__":
    unittest.main()
