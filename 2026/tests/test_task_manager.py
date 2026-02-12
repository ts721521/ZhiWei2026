import json
import os
import shutil
import tempfile
import unittest

from task_manager import (
    TaskStore,
    build_task_runtime_config,
    create_checkpoint,
    mark_checkpoint_file_done,
)


class TaskManagerTests(unittest.TestCase):
    def setUp(self):
        self.root_dir = tempfile.mkdtemp(prefix="taskmgr_")
        self.addCleanup(lambda: shutil.rmtree(self.root_dir, ignore_errors=True))
        self.store = TaskStore(self.root_dir)

    def test_save_task_updates_index_and_task_file(self):
        task = self.store.save_task(
            {
                "id": "alpha",
                "name": "Alpha Task",
                "source_folder": r"C:\src",
                "target_folder": r"C:\out",
                "run_incremental": True,
                "config_overrides": {"run_mode": "convert_then_merge"},
            }
        )
        self.assertEqual("alpha", task["id"])
        self.assertTrue(os.path.isfile(self.store.task_path("alpha")))

        index = self.store.load_index()
        self.assertEqual(1, len(index["tasks"]))
        self.assertEqual("alpha", index["tasks"][0]["id"])
        self.assertEqual("Alpha Task", index["tasks"][0]["name"])

    def test_checkpoint_lifecycle(self):
        cp = create_checkpoint("task1", ["a.docx", "b.xlsx"], run_id="run-1")
        saved = self.store.save_checkpoint("task1", cp)
        self.assertEqual("task1", saved["task_id"])
        loaded = self.store.load_checkpoint("task1")
        self.assertEqual(["a.docx", "b.xlsx"], loaded["planned_files"])

        updated = mark_checkpoint_file_done(loaded, "a.docx")
        self.assertIn("a.docx", updated["completed_files"])
        self.store.save_checkpoint("task1", updated)
        loaded2 = self.store.load_checkpoint("task1")
        self.assertEqual(1, len(loaded2["completed_files"]))

        self.store.clear_checkpoint("task1")
        self.assertIsNone(self.store.load_checkpoint("task1"))

    def test_build_task_runtime_config_full_vs_incremental(self):
        project_cfg = {
            "run_mode": "convert_then_merge",
            "enable_incremental_mode": True,
            "enable_markdown": True,
            "output_enable_md": True,
            "target_folder": r"C:\default_out",
        }
        task = {
            "id": "t001",
            "source_folder": r"C:\src_task",
            "target_folder": r"C:\out_task",
            "run_incremental": True,
            "config_overrides": {"merge_source": "source"},
        }

        cfg_inc = build_task_runtime_config(project_cfg, task, force_full_rebuild=False)
        self.assertTrue(cfg_inc["enable_incremental_mode"])
        self.assertEqual("target", cfg_inc["merge_source"])
        self.assertTrue(
            cfg_inc["incremental_registry_path"].endswith(
                os.path.join("_AI", "registry", "task_t001_incremental_registry.json")
            )
        )

        cfg_full = build_task_runtime_config(project_cfg, task, force_full_rebuild=True)
        self.assertFalse(cfg_full["enable_incremental_mode"])
        self.assertEqual("target", cfg_full["merge_source"])

    def test_delete_task_removes_files_and_index_entry(self):
        self.store.save_task(
            {
                "id": "beta",
                "name": "Beta Task",
                "source_folder": r"C:\src2",
                "target_folder": r"C:\out2",
            }
        )
        self.store.save_checkpoint("beta", create_checkpoint("beta", ["x.docx"]))
        self.store.delete_task("beta")
        self.assertIsNone(self.store.get_task("beta"))
        self.assertIsNone(self.store.load_checkpoint("beta"))
        self.assertEqual([], self.store.list_tasks())

    def test_build_task_runtime_config_resolves_conflicting_keys(self):
        project_cfg = {
            "run_mode": "convert_then_merge",
            "merge_source": "source",
            "enable_incremental_mode": False,
            "output_enable_md": True,
            "enable_markdown": True,
        }
        task = {
            "id": "gamma",
            "source_folder": r"C:\src3",
            "target_folder": r"C:\out3",
            "run_incremental": True,
            "config_overrides": {
                "enable_incremental_mode": False,
                "merge_source": "source",
                "output_enable_md": False,
                "enable_markdown": True,
            },
        }

        cfg = build_task_runtime_config(project_cfg, task, force_full_rebuild=False)
        self.assertTrue(cfg["enable_incremental_mode"])
        self.assertEqual("target", cfg["merge_source"])
        self.assertFalse(cfg["output_enable_md"])
        self.assertFalse(cfg["enable_markdown"])


if __name__ == "__main__":
    unittest.main()
