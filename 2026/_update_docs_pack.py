# -*- coding: utf-8 -*-
import os
_root = os.path.dirname(os.path.abspath(__file__))
path = os.path.join(_root, "打包说明.md")
with open(path, "r", encoding="utf-8") as f:
    s = f.read()
old = "1. **缺少 DLL 或模块**"
new = "1. **报错「Failed to load Python DLL」或「找不到指定的模块」**：说明运行的是 build 下的 exe。请改为运行 dist\\OfficeBatchConverter\\OfficeBatchConverter.exe。\n2. **缺少 DLL 或模块**"
if old in s:
    s = s.replace(old, new)
    with open(path, "w", encoding="utf-8") as f:
        f.write(s)
    print("OK: 打包说明.md updated")
else:
    print("SKIP: pattern not found")

path2 = os.path.join(_root, "使用说明书.md")
with open(path2, "r", encoding="utf-8") as f:
    s2 = f.read()
if "打包" in s2 and "打包说明" not in s2:
    # append a short section or find 附录/其他
    pass
# Add a line in 使用说明书 about 打包成 exe
marker = "## 2. 界面结构"
if marker in s2 and "打包" not in s2[:s2.find(marker) + 200]:
    intro = "\n\n## 打包为 exe\n\n如需将程序封装为单机 exe，请在本目录运行 `python build_exe.py`（需先安装 PyInstaller）。详细说明见 **打包说明.md**。\n\n"
    # insert before first ## 
    idx = s2.find("\n## ")
    if idx != -1:
        s2 = s2[:idx] + intro + s2[idx:]
        with open(path2, "w", encoding="utf-8") as f:
            f.write(s2)
        print("OK: 使用说明书.md added 打包 section")
else:
    print("SKIP: 使用说明书 already has 打包 or no marker")
