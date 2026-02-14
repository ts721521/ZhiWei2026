# -*- coding: utf-8 -*-
"""
一键打包：将 office_gui 打成 exe（单文件模式）。
使用前安装: pip install pyinstaller
在本目录执行: python build_exe.py
"""
import os
import shutil
import subprocess
import sys
import importlib.util

# Import version from office_converter
APP_NAME = "ZhiWei"
ENTRY = "office_gui.py"
ENV_KEYS_TO_CLEAR = ("PYTHONPATH", "PYTHONHOME")

# Get version from office_converter
def get_version():
    try:
        spec = importlib.util.spec_from_file_location("office_converter", "office_converter.py")
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        return getattr(module, "__version__", "unknown")
    except Exception:
        return "unknown"

APP_VERSION = get_version()
APP_NAME_WITH_VERSION = f"{APP_NAME}_v{APP_VERSION}"


def ensure_pyinstaller_installed():
    if importlib.util.find_spec("PyInstaller") is not None:
        return True
    exe = sys.executable
    print("[ERROR] 当前解释器未安装 PyInstaller。")
    print(f"[HINT] 请先执行: \"{exe}\" -m pip install -U pyinstaller")
    print(f"[HINT] 可检查是否安装成功: \"{exe}\" -m PyInstaller --version")
    return False


def build_clean_env():
    env = os.environ.copy()
    cleared = []
    for key in ENV_KEYS_TO_CLEAR:
        if env.get(key):
            cleared.append((key, env.get(key)))
            env.pop(key, None)

    if cleared:
        print("[WARN] detected polluted Python environment variables; cleared for build:")
        for key, value in cleared:
            print(f"       {key}={value}")
    return env


def main():
    root = os.path.dirname(os.path.abspath(__file__))
    os.chdir(root)
    if not ensure_pyinstaller_installed():
        return 1

    dist_root = os.path.join(root, "dist")
    build_root = os.path.join(root, "build")

    # 打包前清空目标目录，避免残留旧文件
    for target_dir in (dist_root, build_root):
        if os.path.isdir(target_dir):
            try:
                shutil.rmtree(target_dir, ignore_errors=True)
                print(f"已清空: {os.path.basename(target_dir)}\\")
            except Exception as e:
                print(f"[WARN] 清空 {os.path.basename(target_dir)} 时出错: {e}")

    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--name", APP_NAME_WITH_VERSION,  # Use versioned name
        "--onefile",
        "--windowed",
        "--noconfirm",
        "--clean",
        "--hidden-import", "ttkbootstrap",
        "--hidden-import", "tkinter",
        "--hidden-import", "win32com.client",
        "--hidden-import", "pythoncom",
        "--hidden-import", "pypdf",
        "--hidden-import", "openpyxl",
        "--hidden-import", "docx",
        "--hidden-import", "bs4",
        "--hidden-import", "chromadb",
        "--collect-all", "ttkbootstrap",
        ENTRY,
    ]
    print(f"正在打包 {APP_NAME_WITH_VERSION}，请稍候...")
    r = subprocess.call(cmd, env=build_clean_env())
    if r != 0:
        return r

    dist_dir = dist_root
    exe_path = os.path.join(dist_dir, APP_NAME_WITH_VERSION + ".exe")  # Use versioned exe path
    config_src = os.path.join(root, "config.json")
    config_dst = os.path.join(dist_dir, "config.json")
    if os.path.isfile(config_src) and os.path.isfile(exe_path):
        shutil.copy2(config_src, config_dst)
        print("已复制 config.json 到 dist（与 exe 同目录）。")

    print("")
    print(f"打包完成！版本：{APP_VERSION}")
    print(f"  可执行文件: dist\\{APP_NAME_WITH_VERSION}.exe")
    print("  重要: 请从 dist 目录运行 exe，不要从 build 目录运行（build 为临时目录，会报错）。")
    print(f"  分发时至少发送该 exe；如需预置配置，请同时发送同目录 config.json。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
