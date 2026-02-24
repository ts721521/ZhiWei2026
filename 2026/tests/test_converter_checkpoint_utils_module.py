import json
import os
import tempfile
import unittest

from office_converter import OfficeConverter


class ConverterCheckpointUtilsSplitTests(unittest.TestCase):
    def test_checkpoint_utils_core_behaviors(self):
        from converter.checkpoint_utils import (
            clear_checkpoint_file,
            get_checkpoint_path,
            mark_file_done_in_checkpoint,
            save_checkpoint,
        )

        root = tempfile.mkdtemp(prefix="ckpt_utils_")
        cfg = {"target_folder": root, "source_folder": r"C:\src\a"}
        checkpoint_path = get_checkpoint_path(cfg)
        self.assertTrue(checkpoint_path.endswith(".json"))
        self.assertTrue(os.path.isdir(os.path.dirname(checkpoint_path)))

        payload = {
            "planned_files": ["a", "b"],
            "completed_files": [],
            "status": "running",
        }
        save_checkpoint(payload, checkpoint_path)
        self.assertTrue(os.path.exists(checkpoint_path))
        with open(checkpoint_path, "r", encoding="utf-8") as f:
            loaded = json.load(f)
        self.assertIn("updated_at", loaded)

        updated = mark_file_done_in_checkpoint(
            {"planned_files": ["a"], "completed_files": [], "status": "running"},
            "a",
        )
        self.assertEqual(updated["status"], "completed")
        self.assertEqual(len(updated["completed_files"]), 1)

        clear_checkpoint_file(checkpoint_path)
        self.assertFalse(os.path.exists(checkpoint_path))

        for d in (
            os.path.join(root, "_AI", "checkpoints"),
            os.path.join(root, "_AI"),
            root,
        ):
            try:
                os.rmdir(d)
            except Exception:
                pass

    def test_office_converter_checkpoint_methods_delegate_to_module(self):
        from converter.checkpoint_utils import (
            get_checkpoint_path,
            mark_file_done_in_checkpoint,
        )

        root = tempfile.mkdtemp(prefix="ckpt_delegate_")
        dummy = OfficeConverter.__new__(OfficeConverter)
        dummy.config = {"target_folder": root, "source_folder": r"C:\src\b"}

        try:
            self.assertEqual(dummy._get_checkpoint_path(), get_checkpoint_path(dummy.config))

            checkpoint = {"planned_files": ["x"], "completed_files": [], "status": "running"}
            expected = mark_file_done_in_checkpoint(
                {"planned_files": ["x"], "completed_files": [], "status": "running"},
                "x",
            )
            self.assertEqual(dummy._mark_file_done_in_checkpoint(checkpoint, "x"), expected)

            dummy._save_checkpoint(checkpoint)
            self.assertTrue(os.path.exists(dummy._get_checkpoint_path()))
            dummy._clear_checkpoint()
            self.assertFalse(os.path.exists(dummy._get_checkpoint_path()))
        finally:
            for d in (
                os.path.join(root, "_AI", "checkpoints"),
                os.path.join(root, "_AI"),
                root,
            ):
                try:
                    os.rmdir(d)
                except Exception:
                    pass


if __name__ == "__main__":
    unittest.main()
