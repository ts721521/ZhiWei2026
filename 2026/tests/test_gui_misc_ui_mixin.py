import queue
import unittest

from gui_misc_ui_mixin import MiscUIMixin


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


if __name__ == "__main__":
    unittest.main()
