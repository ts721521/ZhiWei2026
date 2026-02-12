# 任务管理功能设计（2026-02-12）

## 1. 目标与范围
- 为 Office 转换工具新增“任务管理”能力。
- 任务绑定固定的源目录、目标目录与任务级配置覆盖项。
- 支持任务的手动运行、停止、断点续传。
- 本期不包含定时调度与任务队列。

## 2. 数据模型与存储
任务数据存储在 `tasks/` 目录：
- `tasks/tasks_index.json`：任务摘要索引。
- `tasks/<task_id>.json`：任务完整定义。
- `tasks/<task_id>_checkpoint.json`：任务断点信息（运行计划 + 已完成列表）。

任务核心字段：
- `id`
- `name`
- `source_folder`
- `target_folder`
- `run_incremental`
- `config_overrides`
- `status`
- `created_at` / `updated_at` / `last_run_at`

## 3. 断点续传机制
- 运行开始时记录本次 `planned_files`。
- 每处理完成一个文件，写入 `completed_files`。
- 停止后保留 checkpoint，状态标记为 `paused`。
- 点击“断点续传”时，计算 `planned - completed` 作为本次输入列表，仅处理剩余文件。
- 全部完成后自动清理 checkpoint。

## 4. Converter 扩展
在 `office_converter.py` 增加：
- `run(resume_file_list=None)`：支持直接处理给定文件列表，跳过全量扫描。
- `file_plan_callback`：回调本次计划文件列表。
- `file_done_callback`：每个文件处理完成后回调结果记录。

任务层利用上述回调实现 checkpoint 持续更新。

## 5. 配置覆盖与冲突消解
配置合并规则：
- `runtime = deep_merge(project_config, task.config_overrides)`
- 然后强制注入任务绑定路径：`source_folder` / `target_folder`

冲突消解规则：
- 任务只保存“差异覆盖项”，不保存与项目配置一致的冗余键。
- `run_incremental` 是任务级开关，运行时覆盖 `enable_incremental_mode`。
- `convert_then_merge` 模式下，强制 `merge_source = target`。
- 统一 `output_enable_md` 与 `enable_markdown`，避免双键语义漂移。

## 6. 增量账本隔离
每个任务使用独立账本路径：
- `<target>/_AI/registry/task_<task_id>_incremental_registry.json`

全量重跑时：
- 清理该任务专属账本。
- 清理 checkpoint。
- 本次按全量模式执行。

## 7. GUI 行为
在主 Notebook 新增“任务管理”页，提供：
- 新建任务
- 编辑任务
- 删除任务
- 刷新任务列表
- 运行任务
- 停止任务
- 断点续传
- 本次强制全量重跑开关

运行互斥：
- 同一时刻只允许一个任务运行。
- 运行期间禁用其他任务操作按钮。

## 8. 国际化
在 `ui_translations.py` 增加任务管理相关中英文键：
- Tab 标题、按钮文案、提示消息、对话框标题、详情模板等。

## 9. 后续可扩展
- 定时调度（cron / 间隔触发）。
- 任务历史（每次运行统计、耗时、失败记录）。
- 多任务队列与优先级。
- 任务模板与从当前 UI 快速保存任务。
