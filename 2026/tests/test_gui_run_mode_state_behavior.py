import os
import unittest

from gui_run_mode_state_mixin import RunModeStateMixin


class _FakeVar:
    def __init__(self, value):
        self._value = value

    def get(self):
        return self._value


class _BrokenWidget:
    def configure(self, **kwargs):
        raise RuntimeError("boom")


class _GoodWidget:
    def __init__(self):
        self.state = None

    def configure(self, **kwargs):
        self.state = kwargs.get("state")


class _FakeFrame:
    def __init__(self, children):
        self._children = list(children)

    def winfo_children(self):
        return list(self._children)


class _DummyRunModeState(RunModeStateMixin):
    def __init__(self):
        self.var_enable_date_filter = _FakeVar(0)
        self.good_child = _GoodWidget()
        self.frm_date = _FakeFrame([_BrokenWidget(), self.good_child])
        self.ent_date = _GoodWidget()
        self.errors = []

    def _report_nonfatal_ui_error(self, scope, exc=None, detail=""):
        self.errors.append((scope, str(exc or detail)))


class RunModeStateBehaviorTests(unittest.TestCase):
    def test_on_toggle_date_filter_reports_nonfatal_child_error(self):
        dummy = _DummyRunModeState()

        dummy._on_toggle_date_filter()

        self.assertEqual("disabled", dummy.good_child.state)
        self.assertEqual("disabled", dummy.ent_date.state)
        self.assertEqual(1, len(dummy.errors))
        self.assertEqual(
            "run_mode.toggle_date_filter.child_state", dummy.errors[0][0]
        )
        self.assertIn("boom", dummy.errors[0][1])

    def test_run_mode_mixin_has_no_bare_except(self):
        test_dir = os.path.dirname(os.path.abspath(__file__))
        root = os.path.dirname(test_dir)
        target = os.path.join(root, "gui_run_mode_state_mixin.py")
        with open(target, "r", encoding="utf-8") as f:
            lines = f.readlines()
        bare = [i + 1 for i, line in enumerate(lines) if line.strip() == "except:"]
        self.assertEqual([], bare, f"bare except found in run mode mixin: {bare}")


if __name__ == "__main__":
    unittest.main()
