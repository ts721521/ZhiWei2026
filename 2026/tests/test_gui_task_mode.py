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


def _drain_events(app, rounds=5):
    for _ in range(rounds):
        try:
            app.update_idletasks()
            app.update()
        except Exception:
            break


def _close_app(app):
    if app is None:
        return
    try:
        _drain_events(app, rounds=2)
    except Exception:
        pass
    try:
        # Suppress noisy async Tcl callbacks after window teardown in GUI tests.
        app.tk.call("proc", "bgerror", "msg", "return")
    except Exception:
        pass
    try:
        if hasattr(app, "_on_close_main_window"):
            app._on_close_main_window()
        elif app.winfo_exists():
            app.destroy()
    except Exception:
        try:
            if app.winfo_exists():
                app.destroy()
        except Exception:
            pass


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
    """GUI 实例化后任务 Tab 显隐（仅任务模式可运行）."""

    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="gui_test_")
        self.config_path = os.path.join(self.tmp, "config.json")
        # 使用 office_converter 的默认配置并写入 app_mode（运行时会被强制归一为 task）
        from office_converter import create_default_config
        ok = create_default_config(self.config_path)
        self.assertTrue(ok, "default config should be created")
        with open(self.config_path, "r", encoding="utf-8") as f:
            cfg = json.load(f)
        cfg["app_mode"] = "classic"
        with open(self.config_path, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=4, ensure_ascii=False)

    def tearDown(self):
        _close_app(getattr(self, "app", None))
        self.app = None
        try:
            import shutil
            shutil.rmtree(self.tmp, ignore_errors=True)
        except Exception:
            pass

    @patch("office_gui.messagebox.showinfo")
    def test_task_mode_always_shows_task_tab(self, mock_showinfo):
        with open(self.config_path, "r", encoding="utf-8") as f:
            cfg = json.load(f)
        cfg["app_mode"] = "classic"
        with open(self.config_path, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=4, ensure_ascii=False)
        import office_gui
        self.app = office_gui.OfficeGUI(config_path=self.config_path)
        _drain_events(self.app)
        self.assertEqual(self.app.var_app_mode.get(), "task")
        self.assertTrue(
            _task_tab_visible(self.app),
            "统一任务模式下任务 Tab 应始终显示",
        )
        _close_app(self.app)
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
        source = ""
        paths = [
            os.path.join(_ROOT, "office_gui.py"),
            os.path.join(_ROOT, "gui", "mixins", "gui_task_workflow_mixin.py"),
            os.path.join(_ROOT, "ui_translations.py"),
        ]
        for p in paths:
            with open(p, encoding="utf-8") as f:
                source += "\n" + f.read()
        for i in range(1, 5):
            key = f"wizard_step{i}"
            self.assertIn(
                key,
                source,
                f"向导第 {i} 步 key {key!r} 应在 GUI 功能文件中存在",
            )

    def test_start_button_branches_on_app_mode(self):
        """开始按钮始终通过任务路径运行（统一任务模式）。"""
        source = ""
        paths = [
            os.path.join(_ROOT, "office_gui.py"),
            os.path.join(_ROOT, "gui", "mixins", "gui_execution_mixin.py"),
        ]
        for p in paths:
            with open(p, encoding="utf-8") as f:
                source += "\n" + f.read()
        self.assertIn("_on_click_start", source)
        # 只要存在 _on_click_start 且任务运行入口存在即可；不再要求按 classic/task 分支
        self.assertIn("_on_click_task_run", source)


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
        _close_app(getattr(self, "app", None))
        self.app = None
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
        _close_app(self.app)
        self.app = None


if __name__ == "__main__":
    unittest.main()
