import unittest
from unittest import mock
from pathlib import Path

from gui.mixins.gui_task_workflow_mixin import TaskWorkflowMixin


class _BadVar:
    def get(self):
        raise TypeError("bad var")

    def set(self, _value):
        raise ValueError("bad set")


class _FakeVar:
    def __init__(self, value=""):
        self._value = value

    def get(self):
        return self._value

    def set(self, value):
        self._value = value


class _FallbackCombo:
    def __init__(self):
        self.values = ()

    def cget(self, _key):
        raise AttributeError("no cget")

    def configure(self, **_kwargs):
        raise TypeError("configure unsupported")

    def __setitem__(self, key, value):
        if key != "values":
            raise KeyError(key)
        self.values = tuple(value)


class _BadQueue:
    def put(self, _message):
        raise RuntimeError("queue down")


class _Dummy(TaskWorkflowMixin):
    pass


class TaskWorkflowExceptionNarrowingTests(unittest.TestCase):
    def test_task_workflow_mixin_has_no_bare_except_exception(self):
        mixin_path = (
            Path(__file__).resolve().parents[1]
            / "gui"
            / "mixins"
            / "gui_task_workflow_mixin.py"
        )
        text = mixin_path.read_text(encoding="utf-8")
        self.assertNotIn("except Exception", text)

    def test_task_list_filter_helpers_fallback_on_type_error(self):
        wf = _Dummy()
        wf.var_task_filter_text = _BadVar()
        wf.var_task_status_filter = _BadVar()
        wf.var_task_sort_by = _BadVar()

        self.assertEqual("", wf._task_list_filter_text())
        self.assertEqual("all", wf._task_list_status_filter())
        self.assertEqual("updated_desc", wf._task_list_sort_by())

    def test_refresh_status_filter_values_uses_setitem_fallback(self):
        wf = _Dummy()
        wf.cb_task_status_filter = _FallbackCombo()
        wf.var_task_status_filter = _FakeVar("unknown")

        wf._refresh_task_status_filter_values(
            [{"status": "running"}, {"status": "idle"}, {"status": "running"}]
        )

        self.assertEqual(("all", "idle", "running"), wf.cb_task_status_filter.values)
        self.assertEqual("all", wf.var_task_status_filter.get())

    def test_report_nonfatal_ui_error_ignores_queue_runtime_error(self):
        wf = _Dummy()
        wf.log_queue = _BadQueue()

        rec = wf._report_nonfatal_ui_error("task.queue", detail="warn")

        self.assertEqual("task.queue", rec["scope"])
        self.assertEqual("warn", rec["message"])
        self.assertEqual(1, len(getattr(wf, "_ui_nonfatal_errors", [])))

    def test_normalize_config_for_compare_unserializable_returns_empty(self):
        wf = _Dummy()
        self.assertEqual("", wf._normalize_config_for_compare({"bad": {1, 2, 3}}))

    def test_safe_abs_path_value_error_returns_original(self):
        wf = _Dummy()
        with mock.patch(
            "gui.mixins.gui_task_workflow_mixin.os.path.abspath",
            side_effect=ValueError("bad path"),
        ):
            self.assertEqual("abc", wf._safe_abs_path("abc"))


if __name__ == "__main__":
    unittest.main()
