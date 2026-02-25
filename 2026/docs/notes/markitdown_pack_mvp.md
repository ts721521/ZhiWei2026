# MarkItDown 打包验证（V6.0 MVP）

用于落实 V6.0 计划中的风险项 8.1：先验证 `markitdown` 在 PyInstaller 场景下可运行，再决定是否作为正式分发路径。

## 1. 最小脚本

脚本路径：

- `scripts/markitdown_smoke_mvp.py`

运行方式：

```bash
python scripts/markitdown_smoke_mvp.py --input "D:\sample\a.docx" --output "D:\sample\probe.md"
```

返回码：

- `0`：转换成功
- `1`：缺依赖、输入不存在或转换异常

## 2. 本地打包

```bash
python -m pip install -U pyinstaller markitdown
pyinstaller --onefile --console --name MarkItDownProbe scripts/markitdown_smoke_mvp.py
```

产物示例：

- `dist/MarkItDownProbe.exe`

## 3. 干净机验证清单

在无开发环境的 Windows 虚拟机验证：

1. 拷贝 `MarkItDownProbe.exe` 与一个 Office 样本文件。
2. 运行：
   `MarkItDownProbe.exe --input ".\a.docx" --output ".\probe.md"`
3. 检查返回码是否为 `0`。
4. 检查输出 `probe.md` 是否存在且非空。

## 4. 失败处理

若打包后失败，记录以下信息：

- 错误堆栈（stdout/stderr）
- 样本文件类型（docx/xlsx/pptx）
- 操作系统版本

并按 V6.0 计划切换到备选方案：

1. 评估 `docling` 方案；
2. 或将“极速 MD”标注为本地开发环境能力，不随 exe 分发。
