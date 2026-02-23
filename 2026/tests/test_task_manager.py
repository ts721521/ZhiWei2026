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

    def test_save_task_persists_config_binding_summary_fields(self):
        self.store.save_task(
            {
                "id": "cfg_binding",
                "name": "Config Binding Task",
                "source_folder": r"C:\src",
                "target_folder": r"C:\out",
                "config_snapshot_path": r"C:\profiles\prod.json",
                "config_snapshot_profile_name": "生产配置",
                "config_snapshot_profile_file": "prod.json",
            }
        )
        index = self.store.load_index()
        row = index["tasks"][0]
        self.assertEqual(r"C:\profiles\prod.json", row.get("config_snapshot_path"))
        self.assertEqual("生产配置", row.get("config_snapshot_profile_name"))
        self.assertEqual("prod.json", row.get("config_snapshot_profile_file"))

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

    def test_build_task_runtime_config_prefers_task_snapshot(self):
        project_cfg = {
            "run_mode": "convert_then_merge",
            "output_enable_md": False,
            "default_engine": "wps",
            "merge_source": "source",
        }
        task = {
            "id": "snapshot_1",
            "source_folder": r"C:\src_snap",
            "target_folder": r"C:\out_snap",
            "run_incremental": True,
            "project_config_snapshot": {
                "run_mode": "convert_only",
                "output_enable_md": True,
                "default_engine": "ms",
                "merge_source": "source",
            },
            "config_binding_mode": "snapshot",
            "config_overrides": {},
        }

        cfg = build_task_runtime_config(project_cfg, task, force_full_rebuild=False)
        self.assertEqual("convert_only", cfg["run_mode"])
        self.assertTrue(cfg["output_enable_md"])
        self.assertEqual("ms", cfg["default_engine"])

    def test_build_task_runtime_config_prefers_project_snapshot_over_runtime_snapshot(self):
        project_cfg = {
            "run_mode": "convert_then_merge",
            "output_enable_md": False,
            "default_engine": "wps",
        }
        task = {
            "id": "snapshot_2",
            "source_folder": r"C:\src_snap2",
            "target_folder": r"C:\out_snap2",
            "run_incremental": True,
            "project_config_snapshot": {
                "run_mode": "merge_only",
                "output_enable_md": True,
                "default_engine": "ms",
            },
            "runtime_config_snapshot": {
                "run_mode": "collect_only",
                "output_enable_md": False,
                "default_engine": "wps",
            },
            "config_binding_mode": "snapshot",
            "config_overrides": {},
        }

        cfg = build_task_runtime_config(project_cfg, task, force_full_rebuild=False)
        self.assertEqual("merge_only", cfg["run_mode"])
        self.assertTrue(cfg["output_enable_md"])
        self.assertEqual("ms", cfg["default_engine"])
        self.assertEqual("task.project_config_snapshot", cfg.get("_task_config_source"))

    def test_build_task_runtime_config_uses_runtime_snapshot_as_fallback_only(self):
        project_cfg = {
            "run_mode": "convert_then_merge",
            "output_enable_md": False,
            "default_engine": "ms",
        }
        task = {
            "id": "snapshot_3",
            "source_folder": r"C:\src_snap3",
            "target_folder": r"C:\out_snap3",
            "run_incremental": True,
            "runtime_config_snapshot": {
                "run_mode": "collect_only",
                "output_enable_md": True,
                "default_engine": "wps",
            },
            "config_binding_mode": "snapshot",
            "config_overrides": {},
        }

        cfg = build_task_runtime_config(project_cfg, task, force_full_rebuild=False)
        self.assertEqual("collect_only", cfg["run_mode"])
        self.assertTrue(cfg["output_enable_md"])
        self.assertEqual("wps", cfg["default_engine"])
        self.assertEqual("task.runtime_config_snapshot", cfg.get("_task_config_source"))

    def test_build_task_runtime_config_defaults_legacy_task_to_active(self):
        project_cfg = {
            "run_mode": "convert_then_merge",
            "output_enable_md": False,
            "default_engine": "ms",
        }
        task = {
            "id": "legacy_active",
            "source_folder": r"C:\src_legacy",
            "target_folder": r"C:\out_legacy",
            "project_config_snapshot": {
                "run_mode": "merge_only",
                "output_enable_md": True,
                "default_engine": "wps",
            },
            "config_overrides": {},
        }

        cfg = build_task_runtime_config(project_cfg, task, force_full_rebuild=False)
        self.assertEqual("convert_then_merge", cfg["run_mode"])
        self.assertFalse(cfg["output_enable_md"])
        self.assertEqual("ms", cfg["default_engine"])
        self.assertEqual("active", cfg.get("_task_config_binding_mode"))
        self.assertEqual("project_config(active)", cfg.get("_task_config_source"))

    def test_build_task_runtime_config_profile_binding_uses_bound_profile_file(self):
        profile_cfg = {
            "run_mode": "merge_only",
            "output_enable_md": True,
            "default_engine": "wps",
            "merge_source": "source",
        }
        profile_path = os.path.join(self.root_dir, "profile_bound.json")
        with open(profile_path, "w", encoding="utf-8") as f:
            json.dump(profile_cfg, f, ensure_ascii=False, indent=2)

        project_cfg = {"run_mode": "convert_then_merge", "default_engine": "ms"}
        task = {
            "id": "profile_task",
            "source_folder": r"C:\src_profile",
            "target_folder": r"C:\out_profile",
            "config_binding_mode": "profile",
            "config_snapshot_path": profile_path,
            "config_overrides": {},
        }

        cfg = build_task_runtime_config(project_cfg, task, force_full_rebuild=False)
        self.assertEqual("merge_only", cfg["run_mode"])
        self.assertTrue(cfg["output_enable_md"])
        self.assertEqual("wps", cfg["default_engine"])
        self.assertEqual("profile", cfg.get("_task_config_binding_mode"))
        self.assertEqual("task.profile_config", cfg.get("_task_config_source"))

    def test_migrate_legacy_tasks_sets_binding_mode_active(self):
        task_id = "legacy_migrate"
        legacy_task = {
            "id": task_id,
            "name": "Legacy",
            "source_folder": r"C:\src_legacy2",
            "target_folder": r"C:\out_legacy2",
            "run_incremental": True,
        }
        with open(self.store.task_path(task_id), "w", encoding="utf-8") as f:
            json.dump(legacy_task, f, ensure_ascii=False, indent=2)
        self.store.save_index(
            {
                "version": 1,
                "tasks": [
                    {
                        "id": task_id,
                        "name": "Legacy",
                        "source_folder": r"C:\src_legacy2",
                        "target_folder": r"C:\out_legacy2",
                    }
                ],
            }
        )

        migrated = self.store.migrate_legacy_tasks()
        self.assertEqual(1, migrated)
        task = self.store.get_task(task_id)
        self.assertEqual("active", task.get("config_binding_mode"))


if __name__ == "__main__":
    unittest.main()
