import queue
import unittest

from gui_task_workflow_mixin import TaskWorkflowMixin


class _DummyTaskWorkflow(TaskWorkflowMixin):
    pass


class NonFatalUiErrorReportingTests(unittest.TestCase):
    def test_report_nonfatal_ui_error_records_and_logs(self):
        wf = _DummyTaskWorkflow()
        wf.log_queue = queue.Queue()

        record = wf._report_nonfatal_ui_error("task.refresh", exc=ValueError("boom"))

        self.assertEqual("task.refresh", record.get("scope"))
        self.assertIn("boom", record.get("message", ""))
        self.assertEqual(1, len(getattr(wf, "_ui_nonfatal_errors", [])))
        logged = wf.log_queue.get_nowait()
        self.assertIn("task.refresh", logged)
        self.assertIn("boom", logged)

    def test_report_nonfatal_ui_error_keeps_recent_window(self):
        wf = _DummyTaskWorkflow()
        wf.log_queue = queue.Queue()

        for i in range(80):
            wf._report_nonfatal_ui_error("bulk", detail=f"err-{i}")

        errors = getattr(wf, "_ui_nonfatal_errors", [])
        self.assertEqual(50, len(errors))
        self.assertTrue(errors[0]["message"].endswith("err-30"))
        self.assertTrue(errors[-1]["message"].endswith("err-79"))


if __name__ == "__main__":
    unittest.main()
