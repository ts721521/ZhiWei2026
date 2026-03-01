import tempfile
import unittest
from pathlib import Path
import json

from gui.mixins.gui_profile_mixin import ProfileManagementMixin


class _DummyProfileMixin(ProfileManagementMixin):
    def __init__(self, script_dir=""):
        self.script_dir = script_dir


class ProfileManagementMixinTests(unittest.TestCase):
    def test_sanitize_profile_stem_replaces_invalid_chars(self):
        mixin = _DummyProfileMixin()
        out = mixin._sanitize_profile_stem('prod:/\\*?"<>| config.')
        self.assertEqual("prod_ config", out)

    def test_load_builtin_config_records_discovers_presets_and_scenarios(self):
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            preset_dir = root / "configs" / "presets" / "notebooklm"
            scenario_dir = root / "configs" / "scenarios" / "notebooklm"
            preset_dir.mkdir(parents=True, exist_ok=True)
            scenario_dir.mkdir(parents=True, exist_ok=True)
            (preset_dir / "config.a.json").write_text("{}", encoding="utf-8")
            (scenario_dir / "config.b.json").write_text("{}", encoding="utf-8")

            mixin = _DummyProfileMixin(script_dir=str(root))
            records = mixin._load_builtin_config_records()
            files = {rec.get("file") for rec in records}
            self.assertIn("configs/presets/notebooklm/config.a.json", files)
            self.assertIn("configs/scenarios/notebooklm/config.b.json", files)

    def test_load_builtin_profile_keeps_active_config_path(self):
        class _Var:
            def __init__(self, value=""):
                self._v = value

            def get(self):
                return self._v

            def set(self, v):
                self._v = v

        class _DummyLoader(_DummyProfileMixin):
            def __init__(self, script_dir, active_config):
                super().__init__(script_dir=script_dir)
                self.config_path = active_config
                self.cfg_dirty = False
                self.var_config_path = _Var(active_config)
                self.var_profile_active_path = _Var(active_config)
                self.profile_manager_win = None
                self.load_profile_dialog = None
                self.loaded = False

            def _load_config_to_ui(self):
                self.loaded = True

            def tr(self, key):
                if key == "msg_profile_load_ok":
                    return "{}"
                return key

        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            active = root / "active.json"
            builtin = root / "builtin.json"
            active.write_text(json.dumps({"k": 1}), encoding="utf-8")
            builtin.write_text(json.dumps({"k": 2, "app_mode": "classic"}), encoding="utf-8")
            mixin = _DummyLoader(str(root), str(active))
            rec = {
                "name": "builtin",
                "file": "configs/presets/notebooklm/config.x.json",
                "abs_path": str(builtin),
                "is_builtin": True,
            }
            ok = mixin._load_profile_record(rec, confirm_dirty=False, show_msg=False)
            self.assertTrue(ok)
            self.assertEqual(str(active), mixin.config_path)
            self.assertTrue(mixin.loaded)
            loaded_on_disk = json.loads(active.read_text(encoding="utf-8"))
            self.assertEqual(2, loaded_on_disk.get("k"))


if __name__ == "__main__":
    unittest.main()
