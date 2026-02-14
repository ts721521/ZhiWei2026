# Office 文档批量转换 & 梳理工具

将 Office（Word / Excel / PowerPoint）批量转为 PDF，支持合并、梳理去重，以及面向 NotebookLM 的来源反查与本机检索增强。

**当前版本**: v5.17.0

---

## 功能概览

- **批量转换**：Word / Excel / PPT → PDF（支持 WPS 或 MS Office 引擎）
- **PDF 合并**：分类拆卷或全部合一，可配置合并输出文件名规则
- **多源目录**：支持同时选择多个源文件夹，一次运行处理所有目录下的文件
- **文件梳理**：归集模式下去重、索引、Excel 清单
- **NotebookLM 溯源**：合并 PDF + 页码或短 ID 定位回源文件，配合 Everything / Listary 检索
- **任务模式**：保存多组“源目录 + 目标目录 + 运行参数”，一键切换与执行
- **增量同步**：大体积语料只转换新增/修改文件，并生成增量包
- **MSHelp 模式**：将 MS Help Viewer 的 CAB 帮助包转为 Markdown，便于 RAG / NotebookLM 使用
- **LLM 交付目录**：可集中输出供上传的 Markdown / PDF 等，带清单与去重选项

---

## 环境要求

- **操作系统**：Windows（依赖 COM 调用 Office/WPS）
- **Python**：3.9+
- **Office 环境**：已安装 WPS 或 Microsoft Office（用于转换）

---

## 快速开始

### 1. 克隆或下载本仓库

```bash
git clone <你的仓库地址>
cd <仓库内 2026 所在目录>   # 例如 2026 或 GPTVersion/2026
```

### 2. 安装依赖

```bash
pip install -r requirements.txt
```

若使用 GUI，建议安装 ttkbootstrap（界面与深色主题）：

```bash
pip install ttkbootstrap
```

### 3. 运行

**图形界面（推荐）**

```bash
python office_gui.py
```

首次运行会在当前目录生成默认 `config.json`，可在界面内修改路径与选项。

**命令行**

```bash
python office_converter.py --source "C:\文档" --target "D:\PDF" --run-mode convert_then_merge
```

更多参数见 `python office_converter.py --help`。

---

## 打包为 exe

在项目目录下执行：

```bash
python build_exe.py
```

- 脚本会**先清空** `dist` 与 `build` 目录，再执行 PyInstaller 打包。
- 完成后在 `dist/` 下得到 `OfficeBatchConverter_v5.17.0.exe`（版本号随 `office_converter.py` 中的 `__version__`）。
- 请从 `dist` 目录运行 exe，勿运行 `build` 下的文件。

详细说明见 [打包说明.md](打包说明.md)。

---

## 文档

| 文档 | 说明 |
|------|------|
| [使用说明书.md](使用说明书.md) | 界面说明、数据流程、增量与定位、配置字段、常见问题 |
| [打包说明.md](打包说明.md) | 使用 PyInstaller / cx_Freeze 打包与分发 |
| [CHANGELOG.md](CHANGELOG.md) | 版本更新记录 |

---

## 项目结构（简要）

```
2026/
├── office_gui.py        # GUI 入口（Tk/ttkbootstrap）
├── office_converter.py  # 核心：转换、合并、梳理、MSHelp、增量等
├── ui_translations.py   # 界面文案（当前仅中文）
├── build_exe.py         # 一键打包脚本
├── config.json          # 运行时配置（首次运行自动生成）
├── 使用说明书.md
├── 打包说明.md
├── CHANGELOG.md
└── README.md
```

---

## 许可证与贡献

本仓库为个人/团队使用的工具项目。如需二次开发或参与改进，欢迎提 Issue 或 Pull Request。

---

*README 随版本更新，当前对应 v5.17.0。*
