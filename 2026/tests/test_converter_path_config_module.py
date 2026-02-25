import os
import tempfile
import unittest
from unittest.mock import patch

from office_converter import OfficeConverter


class ConverterPathConfigSplitTests(unittest.TestCase):
    def test_converter_path_config_module_exists(self):
        from converter.path_config import get_path_from_config

        cfg = {
            "source_folder": "C:\\Base",
            "source_folder_win": "D:\\WinPath",
        }
        path = get_path_from_config(cfg, "source_folder", prefer_win=True, prefer_mac=False)
        self.assertEqual(os.path.abspath("D:\\WinPath"), path)

    def test_office_converter_method_matches_path_config_module(self):
        from converter.path_config import get_path_from_config

        dummy = OfficeConverter.__new__(OfficeConverter)
        dummy.config = {
            "target_folder": "C:\\TargetBase",
            "target_folder_win": "E:\\TargetWin",
        }
        expected = get_path_from_config(
            dummy.config, "target_folder", prefer_win=True, prefer_mac=False
        )

        import office_converter as oc

        with patch.object(oc, "is_win", return_value=True), patch.object(
            oc, "is_mac", return_value=False
        ):
            self.assertEqual(
                expected, OfficeConverter._get_path_from_config(dummy, "target_folder")
            )

    def test_init_paths_from_config_core_behaviors(self):
        from converter.path_config import (
            init_paths_from_config,
            init_paths_from_config_for_converter,
        )

        with tempfile.TemporaryDirectory() as td:
            cfg = {"target_folder": td, "temp_sandbox_root": "tmp_sandbox"}
            paths = init_paths_from_config(
                cfg,
                get_app_path_fn=lambda: td,
                gettempdir_fn=lambda: td,
            )
            self.assertTrue(paths["temp_sandbox"].endswith("OfficeToPDF_Sandbox"))
            self.assertTrue(os.path.isdir(paths["temp_sandbox"]))
            self.assertTrue(os.path.isdir(paths["failed_dir"]))
            self.assertTrue(os.path.isdir(paths["merge_output_dir"]))
            dummy = OfficeConverter.__new__(OfficeConverter)
            dummy.config = cfg
            out = init_paths_from_config_for_converter(
                dummy,
                get_app_path_fn=lambda: td,
                gettempdir_fn=lambda: td,
            )
            self.assertTrue(out["temp_sandbox"].endswith("OfficeToPDF_Sandbox"))
            self.assertEqual(out["temp_sandbox"], dummy.temp_sandbox)

    def test_save_config_core_behaviors(self):
        from converter.path_config import save_config

        with tempfile.TemporaryDirectory() as td:
            cfg_path = os.path.join(td, "cfg.json")
            ok = save_config(cfg_path, {"a": 1})
            self.assertTrue(ok)
            self.assertTrue(os.path.exists(cfg_path))

    def test_office_converter_init_and_save_delegate_to_path_module(self):
        import office_converter as oc

        dummy = OfficeConverter.__new__(OfficeConverter)
        dummy.config = {"target_folder": "C:\\TargetBase"}
        dummy.config_path = "cfg.json"

        with patch.object(
            oc,
            "init_paths_from_config_for_converter",
            return_value={"temp_sandbox": "tsandbox"},
        ) as init_impl:
            OfficeConverter._init_paths_from_config(dummy)
            init_impl.assert_called_once_with(
                dummy,
                get_app_path_fn=oc.get_app_path,
                gettempdir_fn=oc.tempfile.gettempdir,
                isabs_fn=oc.os.path.isabs,
                abspath_fn=oc.os.path.abspath,
                join_fn=oc.os.path.join,
                makedirs_fn=oc.os.makedirs,
            )

        with patch.object(oc, "save_config_impl", return_value=True) as save_impl:
            self.assertTrue(OfficeConverter.save_config(dummy))
            save_impl.assert_called_once_with(
                "cfg.json",
                dummy.config,
                open_fn=open,
                dump_fn=oc.json.dump,
                log_error_fn=oc.logging.error,
            )


if __name__ == "__main__":
    unittest.main()
