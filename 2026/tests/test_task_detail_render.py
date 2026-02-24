import unittest

from gui_task_workflow_mixin import TaskWorkflowMixin


class _FakeVar:
    def __init__(self, value):
        self._value = value

    def get(self):
        return self._value


class _FakeTextCore:
    def __init__(self, initial=""):
        self.value = initial
        self.state = "normal"

    def configure(self, cnf=None, **kwargs):
        if isinstance(cnf, dict):
            kwargs.update(cnf)
        if "state" in kwargs:
            self.state = kwargs["state"]

    def delete(self, _start, _end):
        self.value = ""

    def insert(self, _end, text):
        self.value += str(text)

    def get(self, _start, _end):
        return self.value


class _FakeScrolledTextWrapper:
    """Simulate ttkbootstrap ScrolledText: outer wrapper rejects 'state'."""

    def __init__(self, initial=""):
        self.text = _FakeTextCore(initial=initial)

    def configure(self, cnf=None, **kwargs):
        if isinstance(cnf, dict):
            kwargs.update(cnf)
        if "state" in kwargs:
            raise RuntimeError('unknown option "-state"')
        return None

    def delete(self, _start, _end):
        self.text.delete(_start, _end)

    def insert(self, _end, text):
        self.text.insert(_end, text)

    def get(self, _start, _end):
        return self.text.get(_start, _end)


class _FakeButton:
    def __init__(self):
        self.last_state = None

    def configure(self, **kwargs):
        if "state" in kwargs:
            self.last_state = kwargs["state"]


class _FakeTaskStore:
    def get_task(self, task_id):
        if not task_id:
            return None
        return {
            "id": task_id,
            "name": "DemoTask",
            "description": "",
            "source_folder": r"C:\src",
            "source_folders": [r"C:\src"],
            "target_folder": r"C:\out",
            "run_incremental": True,
            "status": "idle",
            "config_overrides": {},
            "project_config_snapshot": {},
            "runtime_config_snapshot": {},
        }

    def load_checkpoint(self, _task_id):
        return {"planned_files": [], "completed_files": []}

    def task_path(self, task_id):
        return f"/tmp/{task_id}.json"


class _DummyTaskWorkflow(TaskWorkflowMixin):
    def __init__(self):
        self.task_store = _FakeTaskStore()
        self.config_path = r"C:\config.json"
        self.var_app_mode = _FakeVar("task")
        self.txt_task_detail = _FakeScrolledTextWrapper(initial="当前未选择任务。\n")
        self.btn_task_resume = _FakeButton()

    def _get_selected_task_id(self):
        return "task_demo"

    def _load_config_for_write(self):
        return {}

    def _ensure_task_config_snapshots(self, task, project_cfg=None, persist=True):
        runtime = {
            "run_mode": "convert_then_merge",
            "source_folders": [task.get("source_folder", "")],
            "target_folder": task.get("target_folder", ""),
            "_task_config_source": "project_config(active)",
        }
        return task, runtime

    def _summarize_task_config_binding(self, task, runtime_preview=None):
        return {
            "binding_mode": "active",
            "display_name": "config.json",
            "config_path": self.config_path,
            "profile_name": "",
            "profile_file": "config.json",
            "relation_label": "跟随当前活动配置",
            "runtime_source_desc": "当前活动配置",
            "match_mode": "active_config",
        }

    def _safe_abs_path(self, path):
        return str(path or "")

    def _task_binding_mode_text(self, mode):
        return str(mode or "")

    def _apply_task_runtime_to_ui(self, cfg, preserve_current_run_tab=False):
        return None

    def _update_task_tab_for_app_mode(self):
        return None

    def tr(self, key):
        mapping = {
            "msg_task_none_selected": "当前未选择任务。",
            "msg_task_detail": "任务：{}\n源目录：{}\n目标目录：{}\n运行模式：{}\n增量：{}\n状态：{}\n断点进度：{}/{}",
            "lbl_task_full_config": "完整配置",
        }
        return mapping.get(key, key)


class TaskDetailRenderTests(unittest.TestCase):
    def test_on_task_select_writes_detail_even_when_outer_scrolledtext_rejects_state(self):
        wf = _DummyTaskWorkflow()
        wf._on_task_select()
        detail = wf.txt_task_detail.text.get("1.0", "end")
        self.assertIn("任务：DemoTask", detail)
        self.assertIn("完整配置", detail)


if __name__ == "__main__":
    unittest.main()
