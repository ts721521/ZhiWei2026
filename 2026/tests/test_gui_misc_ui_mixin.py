import queue
import unittest
from pathlib import Path
from unittest import mock

from gui.mixins.gui_misc_ui_mixin import MiscUIMixin


class _DummyLogWidget:
    def __init__(self):
        self.lines = []

    def insert(self, _where, text):
        self.lines.append(text)

    def see(self, _where):
        pass


class _DummyMiscUI(MiscUIMixin):
    def __init__(self):
        self.log_queue = queue.Queue()
        self.txt_log = _DummyLogWidget()
        self.after_calls = []

    def after(self, ms, callback):
        self.after_calls.append((ms, callback))


class MiscUIMixinTests(unittest.TestCase):
    def test_poll_log_queue_uses_instance_queue_without_global_name(self):
        ui = _DummyMiscUI()
        ui.log_queue.put("line-1")
        ui.log_queue.put("line-2")
        ui._poll_log_queue()
        joined = "".join(ui.txt_log.lines)
        self.assertIn("line-1", joined)
        self.assertIn("line-2", joined)
        self.assertTrue(ui.after_calls)

    def test_open_path_skips_open_in_unittest_context(self):
        ui = _DummyMiscUI()
        with mock.patch(
            "gui.mixins.gui_misc_ui_mixin.sys.argv",
            ["python", "-m", "unittest", "discover"],
        ), mock.patch(
            "gui.mixins.gui_misc_ui_mixin.sys.platform", "win32"
        ), mock.patch(
            "gui.mixins.gui_misc_ui_mixin.os.path.exists", return_value=True
        ), mock.patch(
            "gui.mixins.gui_misc_ui_mixin.os.startfile", create=True
        ) as mocked_start:
            ui._open_path(r"C:\\tmp")
        mocked_start.assert_not_called()

    def test_misc_ui_mixin_has_no_bare_except_exception(self):
        mixin_path = (
            Path(__file__).resolve().parents[1]
            / "gui"
            / "mixins"
            / "gui_misc_ui_mixin.py"
        )
        text = mixin_path.read_text(encoding="utf-8")
        self.assertNotIn("except Exception", text)


if __name__ == "__main__":
    unittest.main()
