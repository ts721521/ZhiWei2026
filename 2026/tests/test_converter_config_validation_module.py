import unittest


class ConverterConfigValidationSplitTests(unittest.TestCase):
    def test_validate_runtime_config_or_raise_accepts_valid_minimal_config(self):
        from converter.config_validation import validate_runtime_config_or_raise

        cfg = {
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
            "max_merge_size_mb": 80,
            "markdown_max_size_mb": 80,
            "parallel_workers": 4,
            "parallel_checkpoint_interval": 10,
            "enable_parallel_conversion": False,
            "enable_merge": True,
            "output_enable_pdf": True,
            "output_enable_md": True,
            "output_enable_merged": True,
            "output_enable_independent": False,
            "enable_fast_md_engine": False,
            "source_folders": [],
            "excluded_folders": [],
            "price_keywords": [],
            "allowed_extensions": {},
        }
        validate_runtime_config_or_raise(cfg)

    def test_validate_runtime_config_or_raise_rejects_invalid_values(self):
        from converter.config_validation import validate_runtime_config_or_raise

        cfg = {
            "run_mode": "invalid_mode",
            "collect_mode": "copy_and_index",
            "collect_copy_layout": "preserve_tree",
            "content_strategy": "standard",
            "default_engine": "ask",
            "kill_process_mode": "ask",
            "merge_mode": "category_split",
            "merge_convert_submode": "merge_only",
            "timeout_seconds": 0,
            "pdf_wait_seconds": 15,
            "ppt_timeout_seconds": 60,
            "ppt_pdf_wait_seconds": 15,
            "max_merge_size_mb": 80,
            "markdown_max_size_mb": 80,
            "parallel_workers": "4",
            "parallel_checkpoint_interval": 10,
            "enable_parallel_conversion": "false",
            "enable_merge": True,
            "output_enable_pdf": True,
            "output_enable_md": True,
            "output_enable_merged": True,
            "output_enable_independent": False,
            "enable_fast_md_engine": False,
            "source_folders": [],
            "excluded_folders": [],
            "price_keywords": [],
            "allowed_extensions": {},
        }
        with self.assertRaises(ValueError) as cm:
            validate_runtime_config_or_raise(cfg)
        text = str(cm.exception)
        self.assertIn("run_mode", text)
        self.assertIn("timeout_seconds", text)
        self.assertIn("parallel_workers", text)
        self.assertIn("enable_parallel_conversion", text)


if __name__ == "__main__":
    unittest.main()
