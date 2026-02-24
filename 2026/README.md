# 知喂 (ZhiWei)

知识投喂工具：将 Office 文档批量转换为 PDF，支持合并/归集、任务模式、增量处理与 NotebookLM 溯源。

当前代码版本：`v5.19.1`（来自 `office_converter.py` 的 `__version__`）

## 核心能力

- 批量转换：Word / Excel / PowerPoint 转 PDF（WPS 或 Microsoft Office 引擎）。
- PDF 合并：按分类或全量合并，支持索引页、映射文件、书签短 ID。
- 任务模式：保存多组任务配置，一键运行；支持任务列表筛选、排序、状态过滤。
- 配置隔离：任务列表支持“仅当前配置”视图，避免跨配置任务干扰。
- 增量同步：只处理新增/变更文件，支持更新包输出。
- NotebookLM 溯源：通过 `map.json/csv` 按页码或短 ID 反查源文件。
- MSHelp 模式：将 CAB 帮助包转换为可检索文档。

## 运行环境

- Windows（依赖本地 Office/WPS COM）
- Python 3.9+
- WPS 或 Microsoft Office（用于转换）

## 快速开始

```bash
git clone <仓库地址>
cd 2026
pip install -r requirements.txt
python office_gui.py
```

可选（更完整 GUI 主题）：

```bash
pip install ttkbootstrap
```

命令行模式示例：

```bash
python office_converter.py --source "C:\输入目录" --target "D:\输出目录" --run-mode convert_then_merge
```

更多参数：

```bash
python office_converter.py --help
```

## 打包

```bash
python build_exe.py
```

- 产物位于 `dist/`。
- EXE 命名规则：`ZhiWei_v<版本号>.exe`。
- 版本号读取自 `office_converter.py` 的 `__version__`。

## 测试状态

2026-02-24 全量验证：

```bash
python -m unittest discover -s tests -p "test_*.py" -v
```

结果：`Ran 71 tests ... OK`

详见：

- [docs/test-reports/README.md](docs/test-reports/README.md)
- [docs/test-reports/TEST_REPORT_SUMMARY.md](docs/test-reports/TEST_REPORT_SUMMARY.md)

## 文档导航

- [CHANGELOG.md](CHANGELOG.md)
- [AGENTS.md](AGENTS.md)（项目规范与命名，供 AI/开发者）
- [docs/notes/使用说明书.md](docs/notes/使用说明书.md)
- [docs/notes/打包说明.md](docs/notes/打包说明.md)
- [docs/plans/V6.0_需求书与项目现状对照评估.md](docs/plans/V6.0_需求书与项目现状对照评估.md)（V6.0 规划与风险建议）
- [docs/plans/2026-02-24-office-converter-split-plan.md](docs/plans/2026-02-24-office-converter-split-plan.md)
- [docs/AI_交接文档_下一阶段开发.md](docs/AI_交接文档_下一阶段开发.md)

## 项目结构（核心）

```text
office_gui.py                   GUI 入口
office_converter.py             核心转换与流水线
task_manager.py                 任务存储与运行态合成
ui_translations.py              UI 文案
docs/
  notes/
  plans/
  test-reports/
tests/                          自动化测试
scripts/                        辅助脚本
```

