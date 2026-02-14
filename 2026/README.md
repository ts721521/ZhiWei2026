# 知喂 (ZhiWei)

面向知识库与 AI 语料准备的一体化工具：将 Office 文档批量转换为 PDF、合并成卷、梳理去重，并支持与 NotebookLM 等场景的溯源与检索集成。

---

## 概述

知喂用于在本地将 Word、Excel、PowerPoint 等 Office 文档批量转为 PDF，并按分类或单卷合并；同时提供归集梳理、增量同步、映射文件生成与页码/短 ID 溯源能力，便于将产出物接入 NotebookLM、RAG 或其它知识库管线。

**当前版本**：v5.17.0

---

## 功能特性

| 能力 | 说明 |
|------|------|
| **批量转换** | Word / Excel / PPT → PDF，支持 WPS 或 Microsoft Office 引擎 |
| **PDF 合并** | 按分类拆卷或全部合一，合并输出文件名规则可配置 |
| **多源目录** | 支持多文件夹同时作为源，单次运行统一处理 |
| **文件梳理** | 归集模式下支持去重、索引与 Excel 清单输出 |
| **NotebookLM 溯源** | 通过合并 PDF 页码或短 ID 反查源文件，可配合 Everything / Listary |
| **任务模式** | 多组「源目录 + 目标目录 + 运行参数」配置，一键切换与执行 |
| **增量同步** | 仅处理新增/变更文件，支持增量包与账本式记录 |
| **MSHelp 模式** | 将 MS Help Viewer CAB 帮助包转为 Markdown，便于 RAG 使用 |
| **LLM 交付** | 集中输出可上传的 Markdown/PDF 等，支持清单与去重策略 |

---

## 环境要求

- **操作系统**：Windows（依赖 COM 调用本地 Office/WPS）
- **Python**：3.9 或更高
- **Office**：已安装 WPS 或 Microsoft Office（用于文档转换）

---

## 安装与运行

### 获取代码

```bash
git clone <仓库地址>
cd <仓库内 2026 所在目录>
```

### 安装依赖

```bash
pip install -r requirements.txt
```

如需完整 GUI 体验（含深色主题），建议安装：

```bash
pip install ttkbootstrap
```

### 启动方式

**图形界面（推荐）**

```bash
python office_gui.py
```

首次运行将在当前目录生成默认 `config.json`；也可复制 [config.example.json](config.example.json) 为 `config.json` 后修改路径与选项。

**命令行**

```bash
python office_converter.py --source "C:\文档" --target "D:\PDF" --run-mode convert_then_merge
```

更多参数请执行 `python office_converter.py --help` 查看。

---

## 构建与分发

在项目根目录执行：

```bash
python build_exe.py
```

- 构建前会自动清空 `dist` 与 `build` 目录。
- 产出物位于 `dist/`，文件名为 `ZhiWei_v<版本号>.exe`（版本号来自 `office_converter.py` 中的 `__version__`）。
- 请仅从 `dist` 目录运行生成的 exe，不要直接运行 `build` 目录下的文件。

详细步骤与常见问题见 [打包说明.md](打包说明.md)。

---

## 文档

| 文档 | 内容 |
|------|------|
| [使用说明书.md](使用说明书.md) | 界面说明、数据流程、增量与溯源、配置说明、常见问题 |
| [打包说明.md](打包说明.md) | PyInstaller / cx_Freeze 打包与分发说明 |
| [CHANGELOG.md](CHANGELOG.md) | 版本更新记录 |

---

## 项目结构

```
├── office_gui.py        # GUI 入口（Tk / ttkbootstrap）
├── office_converter.py  # 核心逻辑：转换、合并、梳理、MSHelp、增量等
├── ui_translations.py   # 界面文案
├── build_exe.py         # 一键打包脚本
├── config.example.json   # 配置示例（复制为 config.json 后修改）
├── 使用说明书.md
├── 打包说明.md
├── CHANGELOG.md
└── README.md
```

---

## 许可证与贡献

本项目为工具型仓库，供个人或团队自用及二次开发。欢迎通过 Issue 反馈问题或通过 Pull Request 参与改进。

---

*文档与版本同步维护，当前对应 v5.17.0。*
