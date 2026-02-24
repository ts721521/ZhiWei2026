# -*- coding: utf-8 -*-
"""Runtime path resolver helpers extracted from office_converter.py."""

import os


def resolve_incremental_registry_path(config):
    configured = str(config.get("incremental_registry_path", "") or "").strip()
    if configured:
        if not os.path.isabs(configured):
            return os.path.abspath(os.path.join(config.get("target_folder", ""), configured))
        return configured
    return os.path.join(
        config.get("target_folder", ""),
        "_AI",
        "registry",
        "incremental_registry.json",
    )


def resolve_update_package_root(config):
    configured = str(config.get("update_package_root", "") or "").strip()
    if configured:
        if not os.path.isabs(configured):
            return os.path.abspath(os.path.join(config.get("target_folder", ""), configured))
        return configured
    return os.path.join(config.get("target_folder", ""), "_AI", "Update_Package")
