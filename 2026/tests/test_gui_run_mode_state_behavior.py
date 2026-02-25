import os
import unittest

from gui.mixins.gui_run_mode_state_mixin import RunModeStateMixin


class _FakeVar:
    def __init__(self, value):
        self._value = value

    def get(self):
        return self._value

    def set(self, value):
        self._value = value


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
        target = os.path.join(root, "gui", "mixins", "gui_run_mode_state_mixin.py")
        with open(target, "r", encoding="utf-8") as f:
            lines = f.readlines()
        bare = [i + 1 for i, line in enumerate(lines) if line.strip() == "except:"]
        self.assertEqual([], bare, f"bare except found in run mode mixin: {bare}")

    def test_on_toggle_fast_md_engine_forces_md_only_and_disables_conflicts(self):
        class _W:
            def __init__(self):
                self.state = "normal"

            def configure(self, **kwargs):
                if "state" in kwargs:
                    self.state = kwargs["state"]

        class _D(RunModeStateMixin):
            def __init__(self):
                from office_converter import MODE_CONVERT_ONLY

                self.var_run_mode = _FakeVar(MODE_CONVERT_ONLY)
                self.var_enable_fast_md_engine = _FakeVar(1)
                self.var_output_enable_md = _FakeVar(0)
                self.var_output_enable_pdf = _FakeVar(1)
                self.var_output_enable_merged = _FakeVar(1)
                self.var_output_enable_independent = _FakeVar(0)
                self.var_enable_parallel_conversion = _FakeVar(1)
                self.chk_output_enable_pdf = _W()
                self.chk_output_enable_merged = _W()
                self.chk_output_enable_independent = _W()
                self.chk_enable_merge = _W()
                self.rb_merge_submode_merge = _W()
                self.rb_merge_submode_pdf_to_md = _W()
                self.chk_enable_parallel_conversion = _W()
                self._frm_parallel_sub = _W()
                self._tooltip_updated = 0
                self._summary_updated = 0

            def _set_widget_tree_state(self, widget, state):
                widget.configure(state=state)

            def _on_toggle_parallel_conversion(self):
                self._frm_parallel_sub.configure(
                    state="normal" if self.var_enable_parallel_conversion.get() else "disabled"
                )

            def _apply_disabled_reason_tooltips(self):
                self._tooltip_updated += 1

            def _update_output_summary_label(self):
                self._summary_updated += 1

        d = _D()
        d._on_toggle_fast_md_engine()

        self.assertEqual(1, d.var_output_enable_md.get())
        self.assertEqual(0, d.var_output_enable_pdf.get())
        self.assertEqual(0, d.var_output_enable_merged.get())
        self.assertEqual(1, d.var_output_enable_independent.get())
        self.assertEqual(0, d.var_enable_parallel_conversion.get())
        self.assertEqual("disabled", d.chk_output_enable_pdf.state)
        self.assertEqual("disabled", d.chk_output_enable_merged.state)
        self.assertEqual("disabled", d.chk_output_enable_independent.state)
        self.assertEqual("disabled", d.chk_enable_parallel_conversion.state)
        self.assertEqual(1, d._tooltip_updated)
        self.assertEqual(1, d._summary_updated)


if __name__ == "__main__":
    unittest.main()

