# 针对本项目的 Skill / MCP / Hooks 安装建议

适用于：知喂 (ZhiWei) 2026 仓库 — Python + 文档/测试强约束 + Git/CI。

---

## 一、建议优先安装的 **Skills（插件技能）**

| 插件名 | 来源 | 理由 |
|--------|------|------|
| **Superpowers** | Cursor Marketplace | TDD、调试、代码审查等结构化流程，与 AGENTS.md「先测试、再记录回归」一致；可用 `/plugin-add superpowers` 或从市场安装。 |
| **Cursor Team Kit** | Cursor Marketplace | 官方插件：CI 监控与修复、PR、合并冲突、smoke test、编译检查、work summary，与 `.github/workflows/quality-gate.yml` 和交接文档流程高度契合。 |
| **Continual Learning** | Cursor Marketplace | 从对话中学习偏好并**增量更新 AGENTS.md** 要点，减轻「每次架构/路径变更必须手动改 AGENTS.md」的负担。 |

---

## 二、建议配置的 **MCP Servers**

| MCP | 用途 | 说明 |
|-----|------|------|
| **Git 类 MCP** | 查 staged 文件、diff、blame、log | 便于 Agent 执行 `check_doc_sync --staged` 前理解变更、写交接摘要；可从 cursor.store / 官方示例选一个兼容的 Git MCP。 |
| **Filesystem / File Reader** | 多目录、大仓库下的读文件/列目录 | 项目 `converter/`、`gui/`、`docs/` 分散，Agent 需要跨路径读文件时有用；若 Cursor 已提供基础文件能力可不必重复安装。 |
| **PDF / Markdown 提取（可选）** | 与 PDF 合并、Markdown 导出 相关时查内容 | 例如 [PDF Extraction MCP](https://cursor.directory/mcp/pdf-extraction)；仅在希望 Agent 直接分析 PDF/ Markdown 内容时安装。 |

**配置位置**：Cursor 设置 → Features → Model Context Protocol；或用户级 `~/.cursor/mcp.json`（Windows 为 `%USERPROFILE%\.cursor\mcp.json`）。

---

## 三、**Hooks** 建议

| 类型 | 建议 | 说明 |
|------|------|------|
| **Git pre-commit** | ✅ 已具备 | 已通过 `scripts/install_git_hook.py` 安装，提交前自动跑 `check_doc_sync.py --staged`，无需再装同类 hook。 |
| **Cursor 插件 Hooks** | 可选 | 若插件支持「保存时 / Agent 完成后」触发：可配置「当修改 `converter/`、`gui/`、`tests/` 时提示运行全量测试或 `check_doc_sync`」；具体取决于所用插件的 hook 能力。 |

---

## 四、本仓库已具备、无需重复的

- **Rules**：`.cursor/rules/project-structure.mdc` 已 alwaysApply，覆盖目录与命名。
- **Pre-commit**：`scripts/install_git_hook.py` + `check_doc_sync.py --staged`。
- **CI**：`.github/workflows/quality-gate.yml`（文档同步校验 + 全量 unittest）。

可选增强：在 `.cursor/rules/` 下新增一条 **rule**，内容为：修改 `office_converter.py`、`office_gui.py`、`task_manager.py`、`converter/**`、`gui/**`、`tests/**` 时，必须同步更新交接文档与 `TEST_REPORT_SUMMARY.md`（与 AGENTS.md §3 一致），避免 AI 遗漏。

---

## 五、安装步骤摘要

1. **Skills**：Cursor 设置 → Plugins / Marketplace，搜索并安装 **Superpowers**、**Cursor Team Kit**、**Continual Learning**。
2. **MCP**：设置 → Features → Model Context Protocol，添加 Git、必要时 File/PDF 类 MCP（按各 MCP 的 install snippet 填写）。
3. **Hooks**：保持现有 pre-commit；若使用带 Hooks 的插件，在插件配置中按需开启「文档/测试提醒」类自动化。
4. **Rule 增强**：在 `.cursor/rules/` 新增 `doc-sync-reminder.mdc`（可选），内容见上文。

---

*文档生成日期：2026-02-26。随 Cursor 市场与 MCP 生态更新可再修订。*
