import unittest

from gui.mixins.gui_task_workflow_mixin import TaskWorkflowMixin


class _FakeVar:
    def __init__(self, value=""):
        self._value = value

    def get(self):
        return self._value

    def set(self, value):
        self._value = value


class _FakeTree:
    def __init__(self):
        self._rows = {}
        self._selection = []
        self._focus = None

    def selection(self):
        return tuple(self._selection)

    def get_children(self):
        return list(self._rows.keys())

    def delete(self, iid):
        self._rows.pop(str(iid), None)

    def insert(self, _parent, _index, iid, values):
        self._rows[str(iid)] = tuple(values)

    def exists(self, iid):
        return str(iid) in self._rows

    def selection_set(self, iid):
        self._selection = [str(iid)]

    def focus(self, iid):
        self._focus = str(iid)


class _FakeTaskStore:
    def __init__(self, tasks):
        self._tasks = [dict(t) for t in tasks]
        self._by_id = {str(t.get("id")): dict(t) for t in tasks}

    def list_tasks(self):
        return [dict(t) for t in self._tasks]

    def get_task(self, task_id):
        return dict(self._by_id.get(str(task_id), {})) or None


class _DummyTaskWorkflow(TaskWorkflowMixin):
    def __init__(self, tasks):
        self.task_store = _FakeTaskStore(tasks)
        self.tree_tasks = _FakeTree()
        self.var_task_filter_text = _FakeVar("")
        self.var_task_status_filter = _FakeVar("all")
        self.var_task_sort_by = _FakeVar("updated_desc")
        self.var_task_scope_current_config_only = _FakeVar(0)
        self.config_path = r"C:\cfg\current.json"
        self._selected_called = 0

    def _summarize_task_config_binding(self, _task, runtime_preview=None):
        return {"display_name": "config.json", "relation_label": "active"}

    def _on_task_select(self):
        self._selected_called += 1


class TaskListFilterSortTests(unittest.TestCase):
    def setUp(self):
        self.tasks = [
            {
                "id": "t3",
                "name": "Gamma",
                "source_folder": r"C:\src\gamma",
                "target_folder": r"C:\out",
                "status": "idle",
                "last_run_at": "2026-02-20T10:00:00",
                "updated_at": "2026-02-20T10:00:00",
            },
            {
                "id": "t1",
                "name": "Alpha",
                "source_folder": r"C:\src\alpha",
                "target_folder": r"C:\out",
                "status": "running",
                "last_run_at": "2026-02-24T09:00:00",
                "updated_at": "2026-02-24T09:00:00",
            },
            {
                "id": "t2",
                "name": "Beta",
                "source_folder": r"C:\src\beta",
                "target_folder": r"C:\out",
                "status": "failed",
                "last_run_at": "2026-02-23T08:00:00",
                "updated_at": "2026-02-23T08:00:00",
            },
        ]

    def _row_ids(self, wf):
        return list(wf.tree_tasks._rows.keys())

    def test_filter_keyword_and_status_then_sort_by_name(self):
        wf = _DummyTaskWorkflow(self.tasks)
        wf.var_task_filter_text.set("a")
        wf.var_task_status_filter.set("running")
        wf.var_task_sort_by.set("name_asc")

        wf._refresh_task_list_ui()

        self.assertEqual(["t1"], self._row_ids(wf))
        self.assertEqual(1, wf._selected_called)

    def test_sort_by_last_run_desc(self):
        wf = _DummyTaskWorkflow(self.tasks)
        wf.var_task_sort_by.set("last_run_desc")

        wf._refresh_task_list_ui()

        self.assertEqual(["t1", "t2", "t3"], self._row_ids(wf))

    def test_scope_filter_only_current_config(self):
        scoped_tasks = [
            {
                "id": "cur",
                "name": "CurrentCfgTask",
                "source_folder": r"C:\src\cur",
                "target_folder": r"C:\out",
                "status": "idle",
                "last_run_at": "2026-02-24T10:00:00",
                "updated_at": "2026-02-24T10:00:00",
                "config_snapshot_path": r"C:\cfg\current.json",
            },
            {
                "id": "other",
                "name": "OtherCfgTask",
                "source_folder": r"C:\src\other",
                "target_folder": r"C:\out",
                "status": "idle",
                "last_run_at": "2026-02-24T09:00:00",
                "updated_at": "2026-02-24T09:00:00",
                "config_snapshot_path": r"C:\cfg\other.json",
            },
        ]
        wf = _DummyTaskWorkflow(scoped_tasks)
        wf.var_task_scope_current_config_only.set(1)

        wf._refresh_task_list_ui()

        self.assertEqual(["cur"], self._row_ids(wf))


if __name__ == "__main__":
    unittest.main()

