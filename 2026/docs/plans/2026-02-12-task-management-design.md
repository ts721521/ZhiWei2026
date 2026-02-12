# 任务管理功能设计（2026-02-12，基于当前代码的优化版）

## 1. 文档目标
本设计以当前代码实现为基线，定义任务管理能力的后续优化路线，确保：
- 规划与现状一致，不空转；
- 每一步可实施、可验证、可回滚；
- 用户能明确“任务到底用什么配置在运行”。

适用代码范围：
- `office_gui.py`
- `task_manager.py`
- `office_converter.py`
- `ui_translations.py`

---

## 2. 当前实现快照（已完成）

### 2.1 任务主流程（已具备）
- 任务 CRUD、运行、停止、续传已具备。
- 任务页已接入独立 Tab，并有运行态按钮互斥控制。
- 断点续传机制已落地：`planned_files` + `completed_files`。

代码锚点：
- `office_gui.py:_build_task_tab_content`
- `office_gui.py:_on_click_task_run`
- `office_gui.py:_set_running_ui_state`
- `task_manager.py:TaskStore`

### 2.2 Converter 联动（已具备）
- Converter 支持 `run(resume_file_list=None)`，续传可跳过全量扫描。
- 提供 `file_plan_callback` / `file_done_callback`，任务层可持续更新 checkpoint。

代码锚点：
- `office_converter.py:run`
- `office_converter.py:_emit_file_plan`
- `office_converter.py:_emit_file_done`

### 2.3 配置冲突消解（已具备）
- `convert_then_merge` 下强制 `merge_source=target`。
- `output_enable_md` 与 `enable_markdown` 已做一致性对齐。
- 每任务独立增量账本路径已生效。

代码锚点：
- `task_manager.py:build_task_runtime_config`

### 2.4 质量状态
- 当前相关测试通过（`python -m unittest discover -s tests`）。

---

## 3. 当前主要缺口（按优先级）

### P0（必须先做）
1. 任务配置来源不稳定
- 当前任务运行仍依赖“当前激活配置文件”作为基线；切 profile 后同一任务行为可能变化。

2. 任务配置编辑入口不闭环
- 任务配置主要来自主界面当前值快照，不是任务页内完整编辑模型。

3. 生效配置不可见
- 用户无法在运行前看到任务最终生效配置（来源、覆盖项、强制项）。

### P1（强烈建议）
4. 缺少正式的任务配置域白名单
- 目前由 `_build_task_overrides_from_ui` 隐式决定，缺少统一常量与校验。

5. 续传一致性策略缺失
- 配置关键项变化后仍可续传，存在结果不一致风险。

### P2（后续增强）
6. 任务配置与全局默认模板的双向复制能力未落地。

---

## 4. 设计原则（保持与当前代码兼容）

1. 增量演进，不推翻当前可用能力。
2. 先补“可解释性”和“稳定性”，再扩展高级交互。
3. 新结构必须向后兼容旧任务文件（`config_overrides`）。

---

## 5. 目标配置模型（兼容迁移版）

### 5.1 运行时三层合成（固定）
1. `global_task_defaults`（全局任务默认模板）
2. `task_config`（任务专属配置）
3. 运行期强制项（路径、增量开关、账本路径、模式强制规则）

### 5.2 与旧结构兼容
- 旧任务字段：`config_overrides`
- 新任务字段：`task_config`
- 兼容策略：
  - 读取时优先 `task_config`，否则回退 `config_overrides`；
  - 保存时统一写 `task_config`，保留 `config_overrides` 一段过渡期；
  - 提供一次性迁移脚本（可选）。

### 5.3 强制规则（保持）
- `run_incremental` 覆盖 `enable_incremental_mode`
- `convert_then_merge => merge_source=target`
- Markdown 双键一致性
- 每任务独立 `incremental_registry_path`

---

## 6. 分阶段实施计划（优化后）

## Phase 1：稳定与可解释（先做）
目标：不改大结构，先让用户看得懂、跑得稳。

改动：
- 抽出任务配置白名单常量（单一来源）。
- 任务详情增加：
  - 配置来源（base + task + forced）
  - 覆盖项数量
  - 关键开关摘要
- 增加“查看生效配置”按钮（只读弹窗/导出 JSON）。

验收：
- 运行前可看到最终生效配置。
- 关键配置项与实际运行一致。

## Phase 2：结构收敛
目标：建立稳定配置来源，消除任务行为漂移。

改动：
- 在 `config.json` 引入 `global_task_defaults`。
- 任务文件引入 `task_config` 与 `task_config_version`。
- 运行时改为固定三层合成，不再依赖主界面临时值。
- 保留旧字段兼容读取。

验收：
- 切换 profile 不影响已配置任务行为（除显式复制/编辑）。
- 新旧任务都能运行。

## Phase 3：一致性与效率增强
目标：降低误操作风险，提升管理效率。

改动：
- 增加续传一致性签名（关键项 hash）。
- 若关键项变化，续传前提示“建议全量重跑”。
- 增加双向复制：
  - Global -> Task
  - Task -> Global
  - 差异预览 + 二次确认。

验收：
- 配置变化时续传策略有明确提示。
- 双向复制仅作用于白名单键。

---

## 7. 关键实现点（按文件）

### `task_manager.py`
- 增加：任务配置白名单、三层合成函数。
- 增加：旧字段到新字段的兼容读取。
- 增加：关键配置签名计算函数（Phase 3）。

### `office_gui.py`
- 任务页新增“查看生效配置”。
- 任务详情扩展配置摘要。
- 新建/编辑任务改为仅编辑任务配置域。
- Phase 3 增加双向复制入口与差异确认。

### `office_converter.py`
- 保持 `resume_file_list` 机制。
- 继续以 callback 向任务层回传计划与完成事件。

### `ui_translations.py`
- 补齐新增按钮与提示语键。

---

## 8. 测试与验收清单

单元测试（建议新增/扩充）：
1. 任务三层合成正确性（含强制规则）。
2. 旧字段兼容读取（`config_overrides`）。
3. 续传列表计算正确性。
4. 关键配置签名变化检测（Phase 3）。
5. 白名单复制策略（Global->Task / Task->Global）。

回归验证：
- 任务新建->运行->停止->续传->完成全链路。
- 全量重跑会清理任务账本并重建 checkpoint。
- 运行互斥（同一时刻仅一个任务运行）。

---

## 9. 非目标（本轮不做）
- 定时调度（cron/间隔）
- 多任务并行队列
- 分布式执行

---

## 10. 结论
当前代码的任务主链路已经可用，下一步不应推翻重来；应按“Phase 1 -> Phase 2 -> Phase 3”渐进优化。优先解决“配置来源不清”和“结果不可解释”，再推进结构收敛与高级管理能力。
