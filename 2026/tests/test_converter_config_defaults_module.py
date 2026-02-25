import unittest
from unittest.mock import patch

from office_converter import (
    COLLECT_MODE_COPY_AND_INDEX,
    MODE_CONVERT_THEN_MERGE,
    OfficeConverter,
    STRATEGY_STANDARD,
)


class ConverterConfigDefaultsSplitTests(unittest.TestCase):
    def test_converter_config_defaults_module_exists(self):
        from converter.config_defaults import apply_config_defaults

        cfg = {}
        runtime = apply_config_defaults(
            cfg,
            run_mode_default=MODE_CONVERT_THEN_MERGE,
            collect_mode_default=COLLECT_MODE_COPY_AND_INDEX,
            content_strategy_default=STRATEGY_STANDARD,
            enable_merge_index_default=False,
            enable_merge_excel_default=False,
        )
        self.assertEqual("ask", cfg.get("default_engine"))
        self.assertEqual("target", cfg.get("merge_source"))
        self.assertEqual(MODE_CONVERT_THEN_MERGE, runtime.get("run_mode"))

    def test_office_converter_apply_config_defaults_matches_module(self):
        from converter.config_defaults import (
            apply_config_defaults,
            apply_config_defaults_for_converter,
        )

        cfg_module = {}
        runtime = apply_config_defaults(
            cfg_module,
            run_mode_default=MODE_CONVERT_THEN_MERGE,
            collect_mode_default=COLLECT_MODE_COPY_AND_INDEX,
            content_strategy_default=STRATEGY_STANDARD,
            enable_merge_index_default=False,
            enable_merge_excel_default=False,
        )

        dummy = OfficeConverter.__new__(OfficeConverter)
        dummy.config = {}
        dummy.run_mode = MODE_CONVERT_THEN_MERGE
        dummy.collect_mode = COLLECT_MODE_COPY_AND_INDEX
        dummy.content_strategy = STRATEGY_STANDARD
        dummy.enable_merge_index = False
        dummy.enable_merge_excel = False

        OfficeConverter._apply_config_defaults(dummy)

        self.assertEqual(cfg_module, dummy.config)
        self.assertEqual(runtime.get("run_mode"), dummy.run_mode)
        self.assertEqual(runtime.get("collect_mode"), dummy.collect_mode)
        self.assertEqual(runtime.get("content_strategy"), dummy.content_strategy)
        self.assertEqual(runtime.get("merge_mode"), dummy.merge_mode)
        self.assertEqual(runtime.get("enable_merge_index"), dummy.enable_merge_index)
        self.assertEqual(runtime.get("enable_merge_excel"), dummy.enable_merge_excel)
        self.assertEqual(runtime.get("price_keywords"), dummy.price_keywords)
        self.assertEqual(runtime.get("excluded_folders"), dummy.excluded_folders)
        self.assertEqual(runtime, apply_config_defaults_for_converter(dummy))

    def test_office_converter_apply_config_defaults_delegates_to_module(self):
        import office_converter as oc

        dummy = OfficeConverter.__new__(OfficeConverter)
        with patch.object(oc, "apply_config_defaults_for_converter", return_value={"ok": 1}) as impl:
            self.assertEqual({"ok": 1}, OfficeConverter._apply_config_defaults(dummy))
            impl.assert_called_once_with(dummy)


if __name__ == "__main__":
    unittest.main()
