---
name: 任务管理功能规划
overview: 将任务能力升级为“任务配置 / 非任务配置”双域模型，明确配置归属、编辑边界、运行优先级与双向复制策略，避免用户混乱。
todos: []
isProject: false
---

# 任务管理功能规划（优化版）

## 0. 评审结论

当前规划方向是对的，现有代码也已经完成了任务运行、停止、断点续传、增量账本隔离的主干能力。

但在“配置认知”上存在一个关键缺口：
- 用户无法明确知道“任务到底用哪套配置在跑”。
- 任务配置与非任务配置在 UI/语义上没有完全隔离。

你提出的思路（任务配置和非任务配置分治，并支持双向复制）是合理且必要的，应作为下一阶段核心优化。

---

## 1. 与当前代码的对齐（现状）

### 1.1 已完成能力
- 任务的创建、编辑、删除、运行、停止、续传。
- checkpoint 记录计划文件和完成文件，支持续传仅跑剩余文件。
- 每任务独立增量账本路径。
- 任务运行时会做项目配置 + 任务覆盖项合并，并包含关键冲突修正：
  - `convert_then_merge` 强制 `merge_source=target`
  - `output_enable_md` / `enable_markdown` 对齐

### 1.2 当前不足（需要优化）
- 任务没有绑定“稳定配置源”，当前取的是“当前激活 config/profile”，存在漂移风险。
- 任务配置项仍主要来自“主运行界面当前值”，不是“任务界面内专属配置编辑”。
- 用户在任务详情里看不到完整生效配置、配置来源、覆盖项数量。
- 缺少任务配置与全局配置之间的显式同步机制（仅能间接影响）。

---

## 2. 目标模型：配置双域分治

## 2.1 域定义
- `非任务配置域`（Global/App Domain）
  - 系统与应用层配置。
  - 仅可在“非任务配置界面”维护。
- `任务配置域`（Task Runtime Domain）
  - 任务执行行为配置（run_mode、输出、merge、增量等）。
  - 仅可在“任务界面”维护。

## 2.2 设计原则
- 同一个配置键只能归属一个域，禁止双归属。
- 非任务界面不能直接改任务配置；任务界面不能直接改非任务配置。
- 手动临时运行（非任务）使用“全局任务默认配置”；任务运行使用“任务专属配置”。

---

## 3. 键归属建议（白名单）

## 3.1 非任务配置域（示例）
- `ui.*`（主题、语言、tooltip、窗口状态）
- 日志/目录管理：`log_folder`
- 外部工具与平台能力：`everything.*`、`listary.*`、`privacy.*`
- profile 管理相关元信息

## 3.2 任务配置域（示例）
- 模式相关：`run_mode`、`collect_mode`、`content_strategy`
- 转换策略：`default_engine`、`kill_process_mode`
- 输出控制：`output_enable_pdf`、`output_enable_md`、`output_enable_merged`、`output_enable_independent`
- merge：`merge_convert_submode`、`merge_mode`、`merge_source`、`enable_merge_index`、`enable_merge_excel`
- 增量：`enable_incremental_mode`、`incremental_verify_hash`、`incremental_reprocess_renamed`、`source_priority_skip_same_name_pdf`、`global_md5_dedup`
- AI导出：`enable_markdown_quality_report`、`enable_excel_json`、`enable_chromadb_export`
- 任务路径相关：`source_folder`、`target_folder`（任务内固定）

注：最终以代码常量白名单为准，文档先定义方向。

---

## 4. 运行时优先级（必须固化）

任务运行时生效配置推荐固定为：
1. 非任务配置中的“任务默认配置模板”（Global Task Defaults）
2. 任务专属配置（Task Config）
3. 运行期强制项（最高优先级）

运行期强制项包括：
- 任务绑定 `source_folder` / `target_folder`
- 任务级增量开关 `run_incremental`
- 任务独立 `incremental_registry_path`
- `convert_then_merge => merge_source=target`
- Markdown 键一致性约束

---

## 5. 双向复制设计（你提出的核心）

## 5.1 Global -> Task
- 用于“从全局任务默认配置初始化/覆盖任务配置”。
- 只复制任务域白名单键。
- 弹窗展示差异预览（新增/修改键）。

## 5.2 Task -> Global
- 用于“将某任务配置沉淀为新的全局任务默认模板”。
- 只回写任务域白名单键。
- 建议写入 `global_task_defaults`（或等价字段）而不是污染非任务域。

## 5.3 安全约束
- 禁止复制运行态字段：`status`、`last_run_at`、checkpoint 相关字段。
- 禁止跨域键写入。
- 复制动作需二次确认，并记录日志。

---

## 6. 任务数据模型优化

在现有 `tasks/<task_id>.json` 基础上，补充：
- `task_config`：任务域完整配置（建议存完整值，不仅差异值，便于可读性与稳定复现）
- `task_config_version`：配置结构版本
- `config_last_copied_from`：最近一次从全局复制的时间/来源（可选）

保留：
- `source_folder`、`target_folder`
- `run_incremental`
- `status`、`last_run_at`

checkpoint 结构保持不变。

---

## 7. GUI 优化要求

## 7.1 任务界面
- 提供任务配置编辑区（仅任务域键）。
- 提供按钮：
  - `从全局默认复制到任务`
  - `将任务配置设为全局默认`
  - `查看生效配置`

## 7.2 非任务配置界面
- 明确分区：
  - 非任务配置
  - 全局任务默认配置模板
- 不展示任务实例配置。

## 7.3 可解释性
任务详情至少展示：
- 配置来源（任务配置 + 全局默认 + 运行期强制）
- 任务域关键开关摘要
- 覆盖项数量

---

## 8. 关键风险与防错

- **风险1：用户误以为复制是双向实时同步**
  - 解决：明确“复制是一次性动作，不是绑定关系”。
- **风险2：续传前配置被改导致结果不一致**
  - 解决：对关键键做签名；签名变化时提示“需全量重跑或确认继续”。
- **风险3：跨域键漂移**
  - 解决：白名单校验 + 单元测试保护。

---

## 9. 分阶段实施建议

### Phase 1（最小可用，先降混乱）
- 定义 `TASK_CONFIG_KEYS` / `GLOBAL_CONFIG_KEYS` 白名单常量。
- 任务界面显示“生效配置摘要 + 来源”。
- 新增双向复制按钮与差异确认弹窗。

### Phase 2（结构收敛）
- 在配置文件引入 `global_task_defaults`。
- 任务文件引入 `task_config`（完整任务域配置）。
- 运行时改为“三层合并模型”。

### Phase 3（稳态增强）
- 续传一致性签名校验。
- 生效配置导出（用于审计/问题复现）。

---

## 10. 验收标准（更新）

- 用户在任务界面可以独立编辑任务配置，不依赖主运行界面临时状态。
- 非任务配置只能在非任务配置界面维护。
- 可执行 Global->Task 与 Task->Global 双向复制，且有差异预览和白名单保护。
- 任意任务运行前，用户可看到清晰的“最终生效配置”。
- 续传流程在关键配置变化时有显式提示与策略。

---

## 11. Obsidian 同步说明

当前工作区未发现本机可写的 `Obsidian Vault/任务管理功能规划.md` 实体路径，故先更新仓库文档：
- `docs/任务管理功能规划_bcfc715c.plan.md`

如果你给我 Obsidian 实际磁盘路径（例如 `D:\Obsidian Vault\任务管理功能规划.md`），我可以再同步写入同名文档。
