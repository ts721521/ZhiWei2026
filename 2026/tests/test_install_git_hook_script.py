import tempfile
import unittest
from pathlib import Path

from scripts.install_git_hook import _find_git_dir


class InstallGitHookScriptTests(unittest.TestCase):
    def test_find_git_dir_walks_up(self):
        with tempfile.TemporaryDirectory(prefix="hook_find_") as tmp:
            root = Path(tmp)
            (root / ".git").mkdir()
            deep = root / "a" / "b" / "c"
            deep.mkdir(parents=True)
            found = _find_git_dir(deep)
            self.assertEqual(root / ".git", found)

    def test_find_git_dir_returns_none_when_absent(self):
        with tempfile.TemporaryDirectory(prefix="hook_find_none_") as tmp:
            root = Path(tmp)
            deep = root / "x" / "y"
            deep.mkdir(parents=True)
            self.assertIsNone(_find_git_dir(deep))


if __name__ == "__main__":
    unittest.main()
