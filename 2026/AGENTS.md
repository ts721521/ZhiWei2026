# AGENTS.md（仓库协作规范）

## 核心目标（必须牢记）

**本项目为 NotebookLM 服务**：知喂 (ZhiWei) 是为 NotebookLM 准备语料的工具。

**核心原则：程序是转换+合并，不是复制！**

| NotebookLM 限制 | 值 |
|-----------------|-----|
| 单文件大小 | <= 100MB |
| 每 Notebook 来源数 | <= 300（免费版） |
| 支持格式 | PDF, DOCX, TXT, Markdown, Google Docs/Sheets/Slides |

**关键规则**：
1. `_LLM_UPLOAD` 目录每次运行前必须清空，只包含本次产物
2. 输出文件必须是转换或合并后的结果，不能是源文件的简单复制
3. 合并后的文件大小和数量必须符合 NotebookLM 限制

---

适用范围：`2026/` 仓库内所有 AI 代理与自动化助手。
目标：保证“代码、测试、交接文档、规范文档”始终同步。

下文规定的「必须记录」「必须同步更新」等规则适用于本项目，不因任何一方偏好而改变。

**本项目就是 AI 主导的**：**人提需求，AI 开发**（写死）。用户只提需求、不参与代码；开发与记录均由 AI 完成。AI 须明确：用户提需求，AI 做好的一切都要有记录。具体包括但不限于：按第 3、4 节更新交接文档与测试汇总；凡涉及「接手必读」中的文档、路径、结构迁移时，当轮 AI 必须同步更新本文件（AGENTS.md）。不能依赖用户提醒，也不能省略或简化任何应记录项。

## 0. 接手必读（新 AI / 新开发者入仓后请按顺序执行）

1. **读本文档**（AGENTS.md），了解强制规则与目录约定。
2. **读交接与现状**  
   - [docs/dev/AI_交接文档_下一阶段开发.md](docs/dev/AI_交接文档_下一阶段开发.md) — 项目概览、入口、已实现能力、建议下一阶段。  
   - [docs/archive/plans-2026-02-landed/2026-02-24-office-converter-split-handover.md](docs/archive/plans-2026-02-landed/2026-02-24-office-converter-split-handover.md) — 最后一节的「当前状态」为代码/拆分现状。  
   - [docs/test-reports/TEST_REPORT_SUMMARY.md](docs/test-reports/TEST_REPORT_SUMMARY.md) — 顶部为最新全量回归结果。  
   - [docs/dev/程序与脚本清单.md](docs/dev/程序与脚本清单.md) — 列出主程序、E2E 脚本与辅助脚本的职责，明确 NotebookLM E2E 测试的是真实程序（OfficeConverter + run_workflow），而非替代实现。
3. **读待办与计划**  
   - [docs/design/TASK_LIST.md](docs/design/TASK_LIST.md) — 任务清单与 Phase，未勾选为待办。  
   - [docs/plans/2026-02-24-code-review-optimization-suggestions.md](docs/plans/2026-02-24-code-review-optimization-suggestions.md) — 代码审查与优化建议。  
   其他计划见 `docs/plans/` 目录。
4. **开始工作前**（建议）：若你对本轮改动负责，请在「本轮变更摘要」中注明**执行者/会话标识**（如 Agent 名或会话 ID + 日期），便于多 AI 追溯。
5. **本轮结束前**（必须）：回填**变更摘要、测试结果、未完成事项与风险**到交接文档或 [docs/test-reports/TEST_REPORT_SUMMARY.md](docs/test-reports/TEST_REPORT_SUMMARY.md)，确保下一轮可接续。

## 1. 基本原则

- **可追溯**：代码、测试、决策、例外必须有记录。
- **可回放**：他人能根据文档复现关键步骤与结论。
- **可验收**：每轮应给出可验证证据（命令、输出摘要、结果状态）。
- **单一事实源**：规范以本文档（AGENTS.md）为准；跨轮事实以交接文档与测试汇总为准。
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

## 3.1 版本与变更记录（必须遵守）

凡交付**新功能**或**重要修复**，必须同时完成：

1. **版本号**：在 [office_converter.py](office_converter.py) 中更新 `__version__`（语义化或递进，如 5.19.1 → 5.20.0）。
2. **CHANGELOG**：在 [CHANGELOG.md](CHANGELOG.md) 顶部追加对应版本条目，注明日期、Added/Fixed/Changed/Documentation。
3. **AGENTS 变更摘要**：若涉及路径/目录/规范变更，在本文档末尾添加 dated note（如「YYYY-MM-DD note: …」）；若仅功能迭代，可在当轮交接/测试汇总中记录，或视需要在 AGENTS 末尾补一条摘要。

可参考外部模板（如模板库中的 00_Template_Index 等）中的版本与文档更新习惯，按本项目结构适配；不得只改代码不升版本、不记变更。

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

**NotebookLM 知识库 E2E**：执行该场景测试时须**循环运行直至通过或需人工介入**。见 [docs/plans/NotebookLM_知识库测试计划_5_投标_ZWPDFTSEST.md](docs/plans/NotebookLM_知识库测试计划_5_投标_ZWPDFTSEST.md) 第六节；可执行命令（在 2026 目录下）：`python scripts/run_notebooklm_e2e.py`；失败时修复提示文件：`docs/test-reports/notebooklm_e2e_repair_prompt.txt`。

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

## 5.3 版本发布门禁（可选）

发布新版本前，须在交接文档或 CHANGELOG 中附**本轮测试证据**（命令、用例数、OK/FAILED）与**风险说明**（已知限制、后续计划），便于验收与回滚决策。

## 6. AI 共享记忆规则（可选）

若启用多 AI 接手的共享记忆，建议使用 [docs/AI_AGENT_MEMORY.md](docs/AI_AGENT_MEMORY.md)（随仓库发布，可被其他 AI 读取）：

- 采用**追加写**（append-only），除非事实错误，不改历史。
- 每条建议格式：`UTC时间 | 范围 | 变更 | 原因 | TODO/风险`。
- 只写仓库相关事实，不写密钥、令牌、账号、内网地址、本机私有路径等敏感信息。
- 若本轮修改了协作规范（目录、命令、流程、约束），必须同时更新 `AGENTS.md` 与共享记忆。

## 7. 禁止项

- 只改代码不记文档。
- 只跑测试不写回归结果。
- 做了**导致本文档所列路径或目录失效的**变更但不更新 `AGENTS.md`。
- 在根目录散落临时脚本/临时文档。

## 8. 执行口径与例外管理

若规则与临时口头要求冲突，以「可追溯、可回放、可验收」为优先。
任何**例外**必须在交接文档中注明：**例外原因、影响范围、临时措施、补齐计划、目标日期**。

---

**其他项目复用**：本仓库提供通用模板 [docs/dev/AGENTS_TEMPLATE.md](docs/dev/AGENTS_TEMPLATE.md)，可复制到其他项目根目录并重命名为 `AGENTS.md`；若该项目无对应文档目录，默认在当前项目下创建，或由用户指定路径后 AI 按模板「初始化清单」执行。

## 2026-02-27 note (GUI config key)
- Added runtime/config key `enable_markdown_image_manifest` for merged markdown image mapping manifest.
- When changing GUI/runtime config keys, keep these files in sync: `converter/config_defaults.py`, `converter/default_config.py`, GUI config mixins, and docs test summary/handover.

## 2026-02-28 note (任务保存/批量/定时 + 版本与变更习惯)
- **功能**：实现「保存到任务」一键按钮、任务多选与「批量运行」排队执行、定时运行（每日 HH:MM，`tasks/schedules.json` + 进程内调度）。涉及：`gui_ui_shell_mixin.py`、`gui_task_workflow_mixin.py`、`gui_execution_mixin.py`、`gui_task_schedule_mixin.py`、`task_manager.py`（schedules 读写）、`ui_translations.py`、`gui/mixins/__init__.py`、`office_gui.py`。
- **版本与变更习惯**：新增 §3.1「版本与变更记录」：交付新功能或重要修复须更新 `office_converter.py` 的 `__version__`、`CHANGELOG.md` 及视情况在 AGENTS 末尾添加 dated note。可参考外部模板库（如 00_Template_Index）的版本与文档更新习惯，按本项目结构适配。
- 版本已升至 v5.20.0，CHANGELOG 已追加 [v5.20.0] - 2026-02-28。

## 2026-04-18 note (任务向导复制布局步骤)
- v5.21.1：任务向导第 3 步（路径）显示「归集子模式 + 复制布局」，与第 2 步选「归集/索引」后的操作路径一致；见 `CHANGELOG.md` [v5.21.1]。

## 2026-04-18 note (Collect 复制布局)
- 新增 `collect_copy_layout`（`preserve_tree` | `flat`），仅在 `copy_and_index` 时影响拷贝目标路径；默认 `preserve_tree` 与旧行为一致。
- 门禁用交接摘要：`docs/plans/2026-02-24-office-converter-split-handover.md` §66；完整历史仍见 `docs/archive/plans-2026-02-landed/2026-02-24-office-converter-split-handover.md`。

## 2026-02-28 note (借鉴 03_AGENTS_TEMPLATE)
- **§0**：增加第 5 条「本轮结束前必须回填变更摘要、测试结果、未完成事项与风险」。
- **§1**：原则细化为可追溯、可回放、可验收、单一事实源。
- **§5.3**：新增可选「版本发布门禁」：发布前须附测试证据与风险说明。
- **§6**：新增可选「AI 共享记忆规则」：`docs/AI_AGENT_MEMORY.md`、追加写、行格式、改规范时同步更新。
- **§8**（原 §7 执行口径）：例外管理补充为「例外原因、影响范围、临时措施、补齐计划、目标日期」。
- 参考模板：`\\192.168.3.2\...\03_AGENTS_TEMPLATE.md`（v1.2 通用方法论增强版）。
