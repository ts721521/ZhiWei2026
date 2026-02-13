# -*- coding: utf-8 -*-
"""
GUI 任务模式 / 传统模式 功能测试（Windows 为主）.
验证：app_mode 持久化、传统模式隐藏任务 Tab、任务模式显示任务 Tab、向导可打开.
"""
import json
import os
import sys
import tempfile
import unittest
from unittest.mock import patch

# 项目根目录加入 path
_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
_ROOT = os.path.dirname(_SCRIPT_DIR)
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)

# 有 ttkbootstrap 时才能验证「任务模式显示 Tab」（无时用 tab state，恢复逻辑与 ttkbootstrap hide/restore 不同）
def _has_ttkbootstrap():
    try:
        import office_gui
        return getattr(office_gui, "HAS_TTKBOOTSTRAP", False)
    except Exception:
        return False


def _task_tab_visible(app):
    """任务 Tab 当前是否在 Notebook 中显示（未被 hide 且 state 非 hidden）。"""
    try:
        app.main_notebook.index(app.tab_run_tasks)
        try:
            s = app.main_notebook.tab(app.tab_run_tasks, "state")
            return str(s) != "hidden"
        except Exception:
            return True
    except Exception:
        return False


class TestAppModeConfig(unittest.TestCase):
    """app_mode 配置读写（不启动 GUI）."""

    def test_app_mode_roundtrip_classic(self):
        with tempfile.TemporaryDirectory(prefix="gui_cfg_") as tmp:
            path = os.path.join(tmp, "config.json")
            cfg = {"app_mode": "classic", "run_mode": "convert_then_merge"}
            with open(path, "w", encoding="utf-8") as f:
                json.dump(cfg, f, ensure_ascii=False)
            with open(path, "r", encoding="utf-8") as f:
                loaded = json.load(f)
            self.assertEqual(loaded.get("app_mode"), "classic")

    def test_app_mode_roundtrip_task(self):
        with tempfile.TemporaryDirectory(prefix="gui_cfg_") as tmp:
            path = os.path.join(tmp, "config.json")
            cfg = {"app_mode": "task", "run_mode": "convert_then_merge"}
            with open(path, "w", encoding="utf-8") as f:
                json.dump(cfg, f, ensure_ascii=False)
            with open(path, "r", encoding="utf-8") as f:
                loaded = json.load(f)
            self.assertEqual(loaded.get("app_mode"), "task")


@unittest.skipIf(sys.platform != "win32", "GUI 实例化测试仅在 Windows 下运行")
class TestGuiTaskTabVisibility(unittest.TestCase):
    """GUI 实例化后任务 Tab 显隐（需 Windows + 可显示窗口）."""

    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="gui_test_")
        self.config_path = os.path.join(self.tmp, "config.json")
        # 使用 office_converter 的默认配置并写入 app_mode
        from office_converter import create_default_config
        ok = create_default_config(self.config_path)
        self.assertTrue(ok, "default config should be created")
        with open(self.config_path, "r", encoding="utf-8") as f:
            cfg = json.load(f)
        cfg["app_mode"] = "classic"
        with open(self.config_path, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=4, ensure_ascii=False)

    def tearDown(self):
        try:
            if hasattr(self, "app") and self.app.winfo_exists():
                self.app.destroy()
        except Exception:
            pass
        try:
            import shutil
            shutil.rmtree(self.tmp, ignore_errors=True)
        except Exception:
            pass

    @patch("office_gui.messagebox.showinfo")
    def test_classic_mode_hides_task_tab(self, mock_showinfo):
        import office_gui
        self.app = office_gui.OfficeGUI(config_path=self.config_path)
        self.app.update_idletasks()
        self.assertEqual(self.app.var_app_mode.get(), "classic")
        self.assertFalse(
            _task_tab_visible(self.app),
            "传统模式下任务 Tab 应被隐藏",
        )
        self.app.destroy()
        self.app = None

    @unittest.skipIf(not _has_ttkbootstrap(), "任务 Tab 恢复显示需 ttkbootstrap；Windows 下推荐安装以完整验证")
    @patch("office_gui.messagebox.showinfo")
    def test_switch_to_task_mode_shows_task_tab(self, mock_showinfo):
        import office_gui
        self.app = office_gui.OfficeGUI(config_path=self.config_path)
        self.app.update_idletasks()
        self.app.var_app_mode.set("task")
        self.app._update_task_tab_for_app_mode()
        self.app.update_idletasks()
        self.assertTrue(
            _task_tab_visible(self.app),
            "切换到任务模式后任务 Tab 应显示",
        )
        self.app.destroy()
        self.app = None

    @unittest.skipIf(not _has_ttkbootstrap(), "任务 Tab 恢复显示需 ttkbootstrap；Windows 下推荐安装以完整验证")
    @patch("office_gui.messagebox.showinfo")
    def test_task_mode_config_shows_task_tab(self, mock_showinfo):
        with open(self.config_path, "r", encoding="utf-8") as f:
            cfg = json.load(f)
        cfg["app_mode"] = "task"
        with open(self.config_path, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=4, ensure_ascii=False)
        import office_gui
        self.app = office_gui.OfficeGUI(config_path=self.config_path)
        self.app.update_idletasks()
        self.assertEqual(self.app.var_app_mode.get(), "task")
        self.assertTrue(
            _task_tab_visible(self.app),
            "配置为 task 时启动后任务 Tab 应显示",
        )
        self.app.destroy()
        self.app = None


class TestWizardExists(unittest.TestCase):
    """新建任务向导入口与步骤存在（不弹窗，只检查方法存在）."""

    def test_open_task_wizard_method_exists(self):
        import office_gui
        self.assertTrue(
            hasattr(office_gui.OfficeGUI, "_open_task_wizard"),
            "_open_task_wizard 应存在",
        )

    def test_update_task_tab_for_app_mode_exists(self):
        import office_gui
        self.assertTrue(
            hasattr(office_gui.OfficeGUI, "_update_task_tab_for_app_mode"),
            "_update_task_tab_for_app_mode 应存在",
        )

    def test_wizard_step_keys_defined(self):
        """向导 4 步的 tr key 存在（文档 5.3）。"""
        gui_path = os.path.join(_ROOT, "office_gui.py")
        with open(gui_path, encoding="utf-8") as f:
            source = f.read()
        for i in range(1, 5):
            key = f"wizard_step{i}"
            self.assertIn(
                key,
                source,
                f"向导第 {i} 步 key {key!r} 应在 office_gui.py 中存在",
            )

    def test_start_button_branches_on_app_mode(self):
        """开始按钮按 app_mode 分支（文档 6.3 / 7.1）。"""
        gui_path = os.path.join(_ROOT, "office_gui.py")
        with open(gui_path, encoding="utf-8") as f:
            source = f.read()
        self.assertIn("_on_click_start", source)
        self.assertIn("var_app_mode", source)
        self.assertIn('"task"', source, "应有 task 模式分支")
        self.assertTrue(
            "var_app_mode" in source and "task" in source and "_on_click_start" in source,
            "开始按钮应根据 app_mode 分支到任务或传统逻辑",
        )


@unittest.skipIf(sys.platform != "win32", "仅 Windows 运行")
class TestWizardStepLabelsOnWindows(unittest.TestCase):
    """Windows 下向导 4 步标签可解析（需 GUI 实例）。"""

    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="gui_wizard_")
        self.config_path = os.path.join(self.tmp, "config.json")
        from office_converter import create_default_config
        create_default_config(self.config_path)
        with open(self.config_path, "r", encoding="utf-8") as f:
            cfg = json.load(f)
        cfg["app_mode"] = "classic"
        with open(self.config_path, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=4, ensure_ascii=False)

    def tearDown(self):
        try:
            if hasattr(self, "app") and self.app.winfo_exists():
                self.app.destroy()
        except Exception:
            pass
        try:
            import shutil
            shutil.rmtree(self.tmp, ignore_errors=True)
        except Exception:
            pass

    @patch("office_gui.messagebox.showinfo")
    def test_wizard_four_step_labels_return_strings(self, mock_showinfo):
        import office_gui
        self.app = office_gui.OfficeGUI(config_path=self.config_path)
        for i in range(1, 5):
            label = self.app.tr(f"wizard_step{i}")
            self.assertIsInstance(label, str, f"wizard_step{i} 应为字符串")
            self.assertTrue(len(label) > 0, f"wizard_step{i} 不应为空")
        self.app.destroy()
        self.app = None


if __name__ == "__main__":
    unittest.main()
