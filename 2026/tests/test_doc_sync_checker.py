import subprocess
import sys
import unittest
from pathlib import Path


class DocSyncCheckerTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.root = Path(__file__).resolve().parent.parent
        cls.script = cls.root / "scripts" / "check_doc_sync.py"

    def run_checker(self, *changed):
        cmd = [sys.executable, str(self.script), "--changed", *changed]
        return subprocess.run(cmd, cwd=self.root, capture_output=True, text=True)

    def test_pass_when_code_and_required_docs_and_agents_all_present(self):
        res = self.run_checker(
            "office_gui.py",
            "docs/plans/2026-02-24-office-converter-split-handover.md",
            "docs/test-reports/TEST_REPORT_SUMMARY.md",
            "AGENTS.md",
        )
        self.assertEqual(0, res.returncode, res.stdout + res.stderr)
        self.assertIn("passed", res.stdout)

    def test_fail_when_code_changes_without_required_docs(self):
        res = self.run_checker("office_gui.py")
        self.assertEqual(1, res.returncode)
        self.assertIn("missing update", res.stdout)
        self.assertIn("TEST_REPORT_SUMMARY.md", res.stdout)

    def test_fail_when_arch_change_missing_agents(self):
        res = self.run_checker(
            "gui/mixins/gui_run_tab_mixin.py",
            "docs/plans/2026-02-24-office-converter-split-handover.md",
            "docs/test-reports/TEST_REPORT_SUMMARY.md",
        )
        self.assertEqual(1, res.returncode)
        self.assertIn("AGENTS.md", res.stdout)


if __name__ == "__main__":
    unittest.main()
