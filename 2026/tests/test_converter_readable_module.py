import unittest

from office_converter import (
    MODE_CONVERT_THEN_MERGE,
    MERGE_MODE_CATEGORY,
    OfficeConverter,
    STRATEGY_STANDARD,
)


class ConverterReadableSplitTests(unittest.TestCase):
    def test_converter_readable_module_exists(self):
        from converter.readable import (
            readable_content_strategy,
            readable_merge_mode,
            readable_run_mode,
        )

        self.assertEqual("convert_then_merge", readable_run_mode(MODE_CONVERT_THEN_MERGE))
        self.assertEqual("standard", readable_content_strategy(STRATEGY_STANDARD))
        self.assertEqual("category_split", readable_merge_mode(MERGE_MODE_CATEGORY))

    def test_office_converter_methods_match_readable_module(self):
        from converter.readable import (
            readable_collect_mode,
            readable_content_strategy,
            readable_engine_type,
            readable_merge_mode,
            readable_run_mode,
        )

        dummy = OfficeConverter.__new__(OfficeConverter)
        dummy.run_mode = MODE_CONVERT_THEN_MERGE
        dummy.collect_mode = "copy_and_index"
        dummy.content_strategy = STRATEGY_STANDARD
        dummy.engine_type = None
        dummy.merge_mode = MERGE_MODE_CATEGORY

        self.assertEqual(readable_run_mode(dummy.run_mode), dummy.get_readable_run_mode())
        self.assertEqual(
            readable_collect_mode(dummy.collect_mode), dummy.get_readable_collect_mode()
        )
        self.assertEqual(
            readable_content_strategy(dummy.content_strategy),
            dummy.get_readable_content_strategy(),
        )
        self.assertEqual(
            readable_engine_type(dummy.engine_type), dummy.get_readable_engine_type()
        )
        self.assertEqual(
            readable_merge_mode(dummy.merge_mode), dummy.get_readable_merge_mode()
        )


if __name__ == "__main__":
    unittest.main()
