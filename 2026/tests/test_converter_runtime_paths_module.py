import os
import unittest

from office_converter import OfficeConverter


class ConverterRuntimePathsSplitTests(unittest.TestCase):
    def test_runtime_paths_core_behaviors(self):
        from converter.runtime_paths import (
            resolve_incremental_registry_path,
            resolve_update_package_root,
        )

        cfg = {"target_folder": r"C:\tmp\target"}
        self.assertEqual(
            resolve_incremental_registry_path(cfg),
            os.path.join(r"C:\tmp\target", "_AI", "registry", "incremental_registry.json"),
        )
        self.assertEqual(
            resolve_update_package_root(cfg),
            os.path.join(r"C:\tmp\target", "_AI", "Update_Package"),
        )

        cfg2 = {
            "target_folder": r"C:\tmp\target",
            "incremental_registry_path": "reg.json",
            "update_package_root": "pkg",
        }
        self.assertEqual(
            resolve_incremental_registry_path(cfg2),
            os.path.abspath(os.path.join(r"C:\tmp\target", "reg.json")),
        )
        self.assertEqual(
            resolve_update_package_root(cfg2),
            os.path.abspath(os.path.join(r"C:\tmp\target", "pkg")),
        )

    def test_office_converter_runtime_path_methods_delegate_to_module(self):
        from converter.runtime_paths import (
            resolve_incremental_registry_path,
            resolve_update_package_root,
        )

        dummy = OfficeConverter.__new__(OfficeConverter)
        dummy.config = {
            "target_folder": r"C:\tmp\target",
            "incremental_registry_path": "registry.json",
            "update_package_root": "upd",
        }
        self.assertEqual(
            dummy._resolve_incremental_registry_path(),
            resolve_incremental_registry_path(dummy.config),
        )
        self.assertEqual(
            dummy._resolve_update_package_root(),
            resolve_update_package_root(dummy.config),
        )


if __name__ == "__main__":
    unittest.main()
