# AGENTS.md（仓库协作规范）

适用范围：`2026/` 仓库内所有 AI 代理与自动化助手。
目标：保证“代码、测试、交接文档、规范文档”始终同步。

下文规定的「必须记录」「必须同步更新」等规则适用于本项目，不因任何一方偏好而改变。

**本项目就是 AI 主导的**：**人提需求，AI 开发**（写死）。用户只提需求、不参与代码；开发与记录均由 AI 完成。AI 须明确：用户提需求，AI 做好的一切都要有记录。具体包括但不限于：按第 3、4 节更新交接文档与测试汇总；凡涉及「接手必读」中的文档、路径、结构迁移时，当轮 AI 必须同步更新本文件（AGENTS.md）。不能依赖用户提醒，也不能省略或简化任何应记录项。

## 0. 接手必读（新 AI / 新开发者入仓后请按顺序执行）

1. **读本文档**（AGENTS.md），了解强制规则与目录约定。
2. **读交接与现状**  
   - [docs/AI_交接文档_下一阶段开发.md](docs/AI_交接文档_下一阶段开发.md) — 项目概览、入口、已实现能力、建议下一阶段。  
   - [docs/plans/2026-02-24-office-converter-split-handover.md](docs/plans/2026-02-24-office-converter-split-handover.md) — 最后一节的「当前状态」为代码/拆分现状。  
   - [docs/test-reports/TEST_REPORT_SUMMARY.md](docs/test-reports/TEST_REPORT_SUMMARY.md) — 顶部为最新全量回归结果。
3. **读待办与计划**  
   - [docs/TASK_LIST.md](docs/TASK_LIST.md) — 任务清单与 Phase，未勾选为待办。  
   - [docs/plans/2026-02-24-code-review-optimization-suggestions.md](docs/plans/2026-02-24-code-review-optimization-suggestions.md) — 代码审查与优化建议。  
   其他计划见 `docs/plans/` 目录。
4. **开始工作前**（建议）：若你对本轮改动负责，请在「本轮变更摘要」中注明**执行者/会话标识**（如 Agent 名或会话 ID + 日期），便于多 AI 追溯。

## 1. 基本原则

- 代码改动必须有可追溯记录。
- 文档不是补充材料，而是交付物的一部分。
- 任何**导致本文档所列路径或目录失效的变更**（如模块搬迁、文件归档、导入路径调整）都必须同步更新 `AGENTS.md`。

## 2. 命名与目录

- Python 模块统一 `snake_case`。
- 测试文件统一 `tests/test_*.py`。
- 计划/交接文档在 `docs/plans/`。
- 测试汇总在 `docs/test-reports/`。
- 临时或历史兼容文件归档到 `docs/archive/`，不放在运行主路径。

## 3. 强制文档同步规则（必须遵守）

当以下任一代码路径发生变更时，必须同步更新文档：

- `office_converter.py`
- `office_gui.py`
- `task_manager.py`
- `converter/**`
- `gui/**`
- `tests/**`

必须同时更新：

- `docs/plans/2026-02-24-office-converter-split-handover.md`
- `docs/test-reports/TEST_REPORT_SUMMARY.md`

若变更涉及架构/目录迁移（例如模块搬迁、导入规则变化、兼容层新增/移除），还必须更新：

- `AGENTS.md`

## 4. 强制测试记录规则（必须遵守）

完成一轮开发后必须执行：

```bash
python -m unittest discover -s tests -p "test_*.py" -v
```

并在上述两个文档中记录：

- 执行命令
- 用例总数（Ran N tests）
- 结果（OK/FAILED）
- 本轮变更摘要（做了什么）

补充：配置加载链路已接入 schema 校验（`converter/config_validation.py`），涉及配置字段调整时必须同步更新校验逻辑与对应测试。

## 5. 自动校验（建议纳入提交前流程）

新增脚本：`scripts/check_doc_sync.py`

用途：校验“代码变更是否同步更新文档 + 架构迁移是否同步更新 AGENTS”。

示例：

```bash
python scripts/check_doc_sync.py --changed office_gui.py gui/mixins/gui_run_tab_mixin.py docs/plans/2026-02-24-office-converter-split-handover.md docs/test-reports/TEST_REPORT_SUMMARY.md AGENTS.md
```

在有 git 的环境可直接使用：

```bash
python scripts/check_doc_sync.py --staged
```

## 5.1 本地强制执行（推荐）

- 安装仓库 pre-commit 钩子：

```bash
python scripts/install_git_hook.py
```

- 钩子模板路径：`.githooks/pre-commit`
- 钩子行为：提交前自动执行 `python scripts/check_doc_sync.py --staged`

## 5.2 CI 强制执行（推荐）

- 仓库已提供 GitHub Actions 工作流：
  - `.github/workflows/quality-gate.yml`
- 门禁内容：
  - 变更文件文档同步校验（`check_doc_sync.py`）
  - 全量单元测试（`python -m unittest discover -s tests -p "test_*.py" -v`）

## 6. 禁止项

- 只改代码不记文档。
- 只跑测试不写回归结果。
- 做了**导致本文档所列路径或目录失效的**变更但不更新 `AGENTS.md`。
- 在根目录散落临时脚本/临时文档。

## 7. 执行口径

若规则与临时口头要求冲突，以“可追溯、可回放、可验收”为优先。
任何例外必须在交接文档注明“例外原因 + 影响范围 + 后续补齐计划”。

---

**其他项目复用**：本仓库提供通用模板 [docs/AGENTS_TEMPLATE.md](docs/AGENTS_TEMPLATE.md)，可复制到其他项目根目录并重命名为 `AGENTS.md`；若该项目无对应文档目录，默认在当前项目下创建，或由用户指定路径后 AI 按模板「初始化清单」执行。
