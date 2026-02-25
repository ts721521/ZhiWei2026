import json
import os
import tempfile
import unittest
from datetime import datetime
from pathlib import Path

from office_converter import OfficeConverter


class ConverterCheckpointRuntimeSplitTests(unittest.TestCase):
    def test_checkpoint_runtime_module_has_no_bare_except_exception(self):
        module_text = Path("converter/checkpoint_runtime.py").read_text(encoding="utf-8")
        self.assertNotIn("except Exception", module_text)

    def test_checkpoint_runtime_core_behaviors(self):
        from converter.checkpoint_runtime import init_checkpoint

        root = tempfile.mkdtemp(prefix="ckpt_rt_")
        ckpt = os.path.join(root, "checkpoint.json")
        files = ["a", "b", "c"]

        saved = {}
        out_ckpt, pending = init_checkpoint(
            files,
            config={"enable_checkpoint": True, "checkpoint_auto_resume": True},
            get_checkpoint_path_fn=lambda: ckpt,
            checkpoint_resume_callback=None,
            save_checkpoint_fn=lambda c: saved.setdefault("value", c),
            now_fn=lambda: datetime(2026, 2, 24, 21, 0, 0),
            exists_fn=lambda _p: False,
            remove_fn=lambda _p: None,
            open_fn=open,
            print_fn=lambda *_a, **_k: None,
            log_warning_fn=lambda _m: None,
        )
        self.assertEqual(pending, files)
        self.assertEqual(out_ckpt.get("status"), "running")
        self.assertIn("value", saved)

        with open(ckpt, "w", encoding="utf-8") as f:
            json.dump(
                {
                    "status": "running",
                    "completed_files": ["a"],
                },
                f,
            )
        out_ckpt2, pending2 = init_checkpoint(
            files,
            config={"enable_checkpoint": True, "checkpoint_auto_resume": True},
            get_checkpoint_path_fn=lambda: ckpt,
            checkpoint_resume_callback=lambda c, t: True,
            save_checkpoint_fn=lambda _c: None,
            now_fn=lambda: datetime(2026, 2, 24, 21, 0, 0),
            exists_fn=lambda p: os.path.exists(p),
            remove_fn=lambda p: os.remove(p),
            open_fn=open,
            print_fn=lambda *_a, **_k: None,
            log_warning_fn=lambda _m: None,
        )
        self.assertEqual(out_ckpt2.get("status"), "running")
        self.assertEqual(pending2, ["b", "c"])

    def test_office_converter_init_checkpoint_delegates_to_module(self):
        import office_converter as oc

        original = oc.init_checkpoint_impl
        try:
            seen = {}

            def _fake(file_list, **kwargs):
                seen["file_list"] = file_list
                seen["kwargs"] = kwargs
                return "ckpt", ["x"]

            oc.init_checkpoint_impl = _fake
            dummy = OfficeConverter.__new__(OfficeConverter)
            dummy.config = {}
            dummy._get_checkpoint_path = lambda: "cp.json"
            dummy._save_checkpoint = lambda _c: None

            out = dummy._init_checkpoint(["a"])
            self.assertEqual(out, ("ckpt", ["x"]))
            self.assertEqual(seen["file_list"], ["a"])
        finally:
            oc.init_checkpoint_impl = original


if __name__ == "__main__":
    unittest.main()
