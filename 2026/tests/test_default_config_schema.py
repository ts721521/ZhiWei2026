import json
import os
import tempfile
import unittest

from office_converter import (
    COLLECT_MODE_COPY_AND_INDEX,
    MODE_CONVERT_THEN_MERGE,
    STRATEGY_STANDARD,
    create_default_config,
)


class DefaultConfigSchemaTests(unittest.TestCase):
    def test_default_config_includes_gui_runtime_schema_keys(self):
        with tempfile.TemporaryDirectory(prefix="cfg_schema_") as tmp:
            config_path = os.path.join(tmp, "config.json")
            self.assertTrue(create_default_config(config_path))
            with open(config_path, "r", encoding="utf-8") as f:
                cfg = json.load(f)

        self.assertEqual("classic", cfg.get("app_mode"))
        self.assertEqual(MODE_CONVERT_THEN_MERGE, cfg.get("run_mode"))
        self.assertEqual(COLLECT_MODE_COPY_AND_INDEX, cfg.get("collect_mode"))
        self.assertEqual(STRATEGY_STANDARD, cfg.get("content_strategy"))

        self.assertIn("enable_parallel_conversion", cfg)
        self.assertIn("parallel_workers", cfg)
        self.assertIn("enable_checkpoint", cfg)
        self.assertIn("checkpoint_auto_resume", cfg)
        self.assertIn("enable_fast_md_engine", cfg)
        self.assertIn("enable_traceability_anchor_and_map", cfg)
        self.assertIn("enable_prompt_wrapper", cfg)
        self.assertIn("prompt_template_type", cfg)
        self.assertIn("short_id_prefix", cfg)

        self.assertIn("source_folders", cfg)
        self.assertIn("source_folder_win", cfg)
        self.assertIn("source_folders_win", cfg)
        self.assertIn("target_folder_win", cfg)
        self.assertIn("source_folder_mac", cfg)
        self.assertIn("source_folders_mac", cfg)
        self.assertIn("target_folder_mac", cfg)

        ui_cfg = cfg.get("ui", {})
        self.assertIsInstance(ui_cfg, dict)
        self.assertIn("confirm_revert_dirty", ui_cfg)
        self.assertIn("task_current_config_only", ui_cfg)
        self.assertTrue(ui_cfg.get("task_current_config_only"))


if __name__ == "__main__":
    unittest.main()
