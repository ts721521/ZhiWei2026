# -*- coding: utf-8 -*-
"""
一键打包：将 office_gui 打成 exe（单目录模式）。
使用前安装: pip install pyinstaller
在本目录执行: python build_exe.py
"""
import os
import shutil
import subprocess
import sys
import importlib.util

APP_NAME = "OfficeBatchConverter"
ENTRY = "office_gui.py"
ENV_KEYS_TO_CLEAR = ("PYTHONPATH", "PYTHONHOME")


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

    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--name", APP_NAME,
        "--onedir",
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
    print("正在打包，请稍候...")
    r = subprocess.call(cmd, env=build_clean_env())
    if r != 0:
        return r

    dist_dir = os.path.join(root, "dist", APP_NAME)
    config_src = os.path.join(root, "config.json")
    config_dst = os.path.join(dist_dir, "config.json")
    if os.path.isfile(config_src) and os.path.isdir(dist_dir):
        shutil.copy2(config_src, config_dst)
        print("已复制 config.json 到 exe 同目录。")

    print("")
    print("打包完成。")
    print("  可执行文件: dist\\" + APP_NAME + "\\" + APP_NAME + ".exe")
    print("  重要: 请从 dist 目录运行 exe，不要从 build 目录运行（build 为临时目录，会报错）。")
    print("  分发时请压缩整个 dist\\" + APP_NAME + " 文件夹。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
