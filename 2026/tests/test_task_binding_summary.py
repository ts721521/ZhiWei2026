import unittest

from gui.mixins.gui_task_workflow_mixin import TaskWorkflowMixin


class _DummyTaskWorkflow(TaskWorkflowMixin):
    def __init__(self, config_path, resolved):
        self.config_path = config_path
        self._resolved = resolved

    def _resolve_task_bound_profile(self, task):
        return dict(self._resolved)

    def tr(self, key):
        return key


class TaskBindingSummaryTests(unittest.TestCase):
    def test_bound_profile_same_as_active_is_current(self):
        wf = _DummyTaskWorkflow(
            config_path=r"C:\profiles\prod.json",
            resolved={
                "config_path": r"C:\profiles\prod.json",
                "profile_name": "生产配置",
                "profile_file": "prod.json",
                "match_mode": "task.config_snapshot_path",
            },
        )
        summary = wf._summarize_task_config_binding(
            {"config_binding_mode": "active"}, runtime_preview={}
        )
        self.assertEqual("prod.json", summary["display_name"])
        self.assertEqual("跟随当前活动配置", summary["relation_label"])

    def test_bound_profile_different_from_active_is_fixed(self):
        wf = _DummyTaskWorkflow(
            config_path=r"C:\profiles\dev.json",
            resolved={
                "config_path": r"C:\profiles\prod.json",
                "profile_name": "生产配置",
                "profile_file": "prod.json",
                "match_mode": "task.config_snapshot_path",
            },
        )
        summary = wf._summarize_task_config_binding(
            {"config_binding_mode": "profile"}, runtime_preview={}
        )
        self.assertEqual("生产配置", summary["display_name"])
        self.assertEqual("绑定指定配置", summary["relation_label"])

    def test_snapshot_only_task_uses_snapshot_relation(self):
        wf = _DummyTaskWorkflow(
            config_path=r"C:\profiles\dev.json",
            resolved={"config_path": "", "profile_name": "", "profile_file": "", "match_mode": "unknown"},
        )
        task = {
            "config_binding_mode": "snapshot",
            "project_config_snapshot": {"run_mode": "convert_only"},
        }
        summary = wf._summarize_task_config_binding(task, runtime_preview={})
        self.assertEqual("使用任务快照", summary["relation_label"])


if __name__ == "__main__":
    unittest.main()

