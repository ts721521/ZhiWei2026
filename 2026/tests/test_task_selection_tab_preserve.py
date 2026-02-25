import os
import re
import unittest


_TEST_DIR = os.path.dirname(os.path.abspath(__file__))
_ROOT = os.path.dirname(_TEST_DIR)


class TaskSelectionTabPreserveTests(unittest.TestCase):
    def test_task_select_preserves_current_run_tab_when_syncing_runtime(self):
        path = os.path.join(_ROOT, "gui", "mixins", "gui_task_workflow_mixin.py")
        with open(path, "r", encoding="utf-8") as f:
            source = f.read()

        pattern = re.compile(
            r"_apply_task_runtime_to_ui\(\s*runtime_preview,\s*preserve_current_run_tab=True\s*\)"
        )
        self.assertRegex(
            source,
            pattern,
            "Selecting a task should not auto-jump away from the task tab.",
        )


if __name__ == "__main__":
    unittest.main()
