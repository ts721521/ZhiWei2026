#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""检查 Google Drive 上传相关配置与依赖（不打印敏感内容）。"""
from __future__ import print_function

import os
import sys

_REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_CONFIG_PATH = os.path.join(_REPO_ROOT, "config.json")


def _main():
    print("=== Google Drive 上传配置检查 ===\n")
    if not os.path.isfile(_CONFIG_PATH):
        print("未找到 config.json，路径:", _CONFIG_PATH)
        print("请先运行一次 GUI 或从 configs/templates/config.example.json 复制并填写。")
        return 1

    try:
        with open(_CONFIG_PATH, "r", encoding="utf-8") as f:
            cfg = __import__("json").load(f)
    except Exception as e:
        print("读取 config.json 失败:", e)
        return 1

    enable = cfg.get("enable_gdrive_upload", False)
    secrets_path = (cfg.get("gdrive_client_secrets_path") or "").strip()
    folder_id = (cfg.get("gdrive_folder_id") or "").strip()
    token_path = (cfg.get("gdrive_token_path") or "").strip()

    print("1. 启用 Google Drive 上传:", "是" if enable else "否")
    print("2. 客户端密钥路径:", "已设置" if secrets_path else "未设置")
    if secrets_path:
        exists = os.path.isfile(secrets_path)
        print("   - 文件存在:", "是" if exists else "否")
        if not exists:
            print("   - 请确认路径正确且 client_secrets.json 已从 Google Cloud 控制台下载到该位置。")
    print("3. 目标文件夹 ID:", "已设置" if folder_id else "未设置（留空则上传到 Drive 根下「知喂上传/Run_*」）")
    if token_path:
        print("4. Token 缓存路径: 已设置")
    else:
        print("4. Token 缓存路径: 使用默认（用户目录/知喂/gdrive_token.json）")

    try:
        import gdrive_upload as gd
        has_dep = gd.HAS_GDEPEND
    except Exception:
        has_dep = False
    print("5. Google 依赖 (google-auth-oauthlib, google-api-python-client):", "已安装" if has_dep else "未安装")
    if not has_dep:
        print("   - 安装: pip install google-auth-oauthlib google-api-python-client")
        print("   - 或在 GUI 中点击「一键安装 Google Drive 依赖」")

    ok = enable and bool(secrets_path) and os.path.isfile(secrets_path) and has_dep
    print()
    if ok:
        print("结论: 配置完整，可以进行上传。")
        print("  - 在知喂 GUI「成果文件」页勾选「启用 Google Drive 上传」，填写密钥路径后点击「上传 _LLM_UPLOAD 到 Google Drive」。")
        print("  - 首次上传会打开浏览器完成 OAuth 授权。")
    else:
        print("结论: 请补全上述未通过项后再使用上传功能。")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(_main())
