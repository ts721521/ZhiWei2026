# Cursor / Codex 规则与约定模板（可复制到其他项目复用）

本目录存放与 **AGENTS.md 文档同步规则** 配套的模板，便于在**任意项目**（含非 Cursor 项目，如用 Codex、VS Code、其他 AI 助手）中保持「改代码必更交接/测试汇总」的同一套约定。

## 文件说明

| 文件 | 用途 | 适用环境 |
|------|------|----------|
| `doc-sync-reminder-content.md` | **与编辑器无关**的约定正文：触发路径、必须同步的文档、测试命令、自检命令。 | 任意；可单独给 AI 读，或作为 Cursor/Codex 模板的正文来源。 |
| `doc-sync-reminder.mdc.template` | Cursor Rule 模板：在上述正文基础上加上 Cursor 用 frontmatter，放入 `.cursor/rules/`。 | **Cursor** |
| `codex-skill-doc-sync/SKILL.md` | Codex Skill 模板：同一套约定写成 Skill，放入项目或用户 skill 目录。 | **Codex** |

---

## 其他项目如何复用

### 方式一：只用 AGENTS.md（任意环境）

若项目已采用本仓库的 `docs/AGENTS_TEMPLATE.md` 作为 AGENTS.md，其中 §3（文档目录）、§4（强制文档同步）、§5（测试记录）已包含相同约定。**不装 Cursor / Codex 时**，只要 AI 先读 AGENTS.md，按「接手必读」执行即可复用，无需本目录任何文件。

### 方式二：Cursor 项目

1. 将 `doc-sync-reminder.mdc.template` 复制到目标项目的 `.cursor/rules/doc-sync-reminder.mdc`（无则先建 `.cursor/rules/`）。
2. 在 `.mdc` 中按项目填写：**触发路径**、**必须同步文档路径**、**测试命令**、**自检命令**（可选），与该项目 AGENTS.md 的「本项目的文档目录」「代码路径约定」「测试命令」一致。

### 方式三：Codex 项目（或用 Codex 开发的其他仓库）

1. 将 `codex-skill-doc-sync` 整个目录复制到目标项目中，例如：  
   `目标项目/.agents/skills/doc-sync-reminder/`  
   （Codex 会从仓库 `.agents/skills` 或用户 `$HOME/.agents/skills` 等位置加载 skills。）
2. 打开复制后的 `SKILL.md`，将「触发路径」「必须同步的文档」「测试命令」等示例替换为该项目的实际约定，与 AGENTS.md 一致。
3. 开发时可通过 `/skills` 或 `$doc-sync-reminder` 显式调用该 skill，或依赖 Codex 按任务匹配 skill 描述自动带入上下文。

### 方式四：仅复用约定内容（任意编辑器 / 任意 AI）

复制 `doc-sync-reminder-content.md` 到目标项目（如 `docs/doc-sync-reminder.md`），按项目填写占位后，在 AGENTS.md 或「接手必读」中注明：**修改约定代码路径后须按 `docs/doc-sync-reminder.md` 更新文档与测试汇总**。任何能读该文件的 AI 或人都可复用同一约定。

---

## 与本仓库的关系

- 本仓库（知喂 2026）中：**Cursor** 实际生效规则在 `2026/.cursor/rules/doc-sync-reminder.mdc`；**Codex** 可将 `codex-skill-doc-sync` 拷到 `2026/.agents/skills/doc-sync-reminder` 并按项目改路径。
- 本目录仅提供**通用模板**，供 Cursor、Codex 或纯 AGENTS.md 项目按需复制与修改。
