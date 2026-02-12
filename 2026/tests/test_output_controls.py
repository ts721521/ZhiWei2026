import unittest

from office_converter import MODE_CONVERT_ONLY, MODE_CONVERT_THEN_MERGE, OfficeConverter


class OutputPlanTests(unittest.TestCase):
    def test_convert_only_pdf_independent(self):
        cfg = {
            "output_enable_pdf": True,
            "output_enable_md": False,
            "output_enable_merged": False,
            "output_enable_independent": True,
            "enable_merge": True,
        }
        plan = OfficeConverter.compute_convert_output_plan(MODE_CONVERT_ONLY, cfg)
        self.assertTrue(plan["need_final_pdf"])
        self.assertFalse(plan["need_markdown"])

    def test_convert_only_md_independent(self):
        cfg = {
            "output_enable_pdf": False,
            "output_enable_md": True,
            "output_enable_merged": False,
            "output_enable_independent": True,
            "enable_merge": True,
        }
        plan = OfficeConverter.compute_convert_output_plan(MODE_CONVERT_ONLY, cfg)
        self.assertFalse(plan["need_final_pdf"])
        self.assertTrue(plan["need_markdown"])

    def test_convert_then_merge_pdf_merged_only(self):
        cfg = {
            "output_enable_pdf": True,
            "output_enable_md": False,
            "output_enable_merged": True,
            "output_enable_independent": False,
            "enable_merge": True,
        }
        plan = OfficeConverter.compute_convert_output_plan(
            MODE_CONVERT_THEN_MERGE, cfg
        )
        self.assertTrue(plan["need_final_pdf"])
        self.assertFalse(plan["need_markdown"])

    def test_convert_then_merge_md_merged_only(self):
        cfg = {
            "output_enable_pdf": False,
            "output_enable_md": True,
            "output_enable_merged": True,
            "output_enable_independent": False,
            "enable_merge": True,
        }
        plan = OfficeConverter.compute_convert_output_plan(
            MODE_CONVERT_THEN_MERGE, cfg
        )
        self.assertFalse(plan["need_final_pdf"])
        self.assertTrue(plan["need_markdown"])

    def test_all_formats_disabled(self):
        cfg = {
            "output_enable_pdf": False,
            "output_enable_md": False,
            "output_enable_merged": False,
            "output_enable_independent": False,
            "enable_merge": True,
        }
        plan = OfficeConverter.compute_convert_output_plan(MODE_CONVERT_ONLY, cfg)
        self.assertFalse(plan["need_final_pdf"])
        self.assertFalse(plan["need_markdown"])


if __name__ == "__main__":
    unittest.main()
