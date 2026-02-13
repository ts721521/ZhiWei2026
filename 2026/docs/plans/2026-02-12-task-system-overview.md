# 任务运行体系总设计（任务模式 + 传统模式共存）

> 面向 2026 版 Office 批量转换工具的运行体系重构说明。  
> 目标：以“任务模式”为主线，保留“传统模式”兼容路径，统一规划数据结构、运行流程和 GUI 行为，方便后续多人接力开发。

---

## 0. 运行模式总览

- **任务模式（推荐主线）**
  - 以“任务”为单位运行；
  - 任务绑定：源目录列表、目标目录、任务级配置覆盖项；
  - 支持：独立增量账本、断点续传、任务管理页、新建任务向导等；
  - 适合：长期维护固定目录、增量更新、可重复运行的场景。

- **传统模式（非任务模式，兼容用）**
  - 沿用现有用法：根据当前 UI 参数（运行参数 Tab + config.json）直接运行；
  - 不强制先创建任务；
  - 高级能力（“按任务隔离的增量账本 / checkpoint”）不作为保证前提；
  - 适合：一次性/临时任务、老用户习惯、调试配置。

> 约定：除非特别说明，本文“运行流程/数据结构/GUI 设计”等内容默认指 **任务模式**。  
> 传统模式的保留方式和差异，集中在第 7 章说明。

---

## 1. 背景与总体目标

### 1.1 现状

- 使用方式：
  - 通过 `config.json` + GUI 各运行参数 Tab 配置源/目标/模式/输出；
  - 点击“开始”后按当前界面参数直接跑；
  - 没有“任务”这个抽象，配置与运行行为是一次性的。
- 已有能力：
  - 增量账本（`incremental_registry.json`）和增量包；
  - MSHelp 模式、收集模式、LLM 输出等；
  - GUI 结构已经有多个功能 Tab（转换、合并、收集、MSHelp、定位、输出、配置中心）。

### 1.2 目标

1. **引入“任务模式”作为主线运行方式**
   - 用“任务”描述固定目录 + 输出策略的组合；
   - 复用已有增量账本与增量包能力；
   - 新增任务管理页、新建任务向导、断点续传。
2. **保留“传统模式”以兼容旧用法**
   - 不破坏现有用户的操作习惯；
   - 逐步引导用户迁移到任务模式。
3. **统一运行入口与配置合并规则**
   - 只有一个真正的“执行入口”（开始按钮）；
   - 模式只决定：**配置从哪里来**、**是否走任务相关能力**。

---

## 2. 核心概念与运行时配置来源

### 2.1 任务（Task）

- 描述一类“固定源目录 + 目标目录 + 一组运行策略”的抽象。
- 典型字段：
  - `id`、`name`、`description`；
  - `source_folders[]`、`target_folder`；
  - `run_incremental_by_default`；
  - `task_overrides`：任务级配置覆盖项；
  - 状态与时间：`status`（idle/running/paused）、`created_at`、`updated_at`、`last_run_at`。

### 2.2 全局配置 vs 任务覆盖

- 全局配置（`config.json`）：
  - 日志、隐私、沙箱、Office 引擎、一部分默认运行策略、UI 设置等；
  - 所有模式和所有任务共享。
- 任务覆盖（`task_overrides`）：
  - 与具体任务强相关的运行/输出/增量策略；
  - 只保存“与全局配置不同”的差异键。

### 2.3 有效配置（effective_config）

统一规则：

```text
effective_config = deep_merge(global_config, task_overrides)
强制覆盖:
  effective_config["source_folders"] = task.source_folders
  effective_config["target_folder"]  = task.target_folder
  effective_config["enable_incremental_mode"] = task.run_incremental_by_default
```

- 在 **任务模式** 下，运行入口使用 `effective_config`。
- 在 **传统模式** 下，运行入口直接使用“当前 UI 参数 + 全局配置合并”的结果，不依赖 Task。

---

## 3. 数据模型与存储结构

### 3.1 任务数据存储（tasks/）

目录建议：

- `tasks/tasks_index.json`：任务索引与摘要；
- `tasks/<task_id>.json`：任务详细定义；
- `tasks/checkpoints/<task_id>_checkpoint.json`：任务断点续传信息。

任务 JSON（示意）：

```jsonc
{
  "id": "task_2026_demo",
  "name": "2026年度招投标转换",
  "description": "E盘招投标→C盘PDF，启用增量与合并",
  "created_at": "2026-02-12T10:00:00",
  "updated_at": "2026-02-12T10:00:00",
  "source_folders": [
    "E:\\\\21_SE_Doc\\\\BaiduNetdiskWorkspace\\\\SynologyDrive\\\\5_投标\\\\2026"
  ],
  "target_folder": "C:\\\\PDFs",
  "run_incremental_by_default": true,
  "task_overrides": {
    "run_mode": "convert_then_merge",
    "output_enable_pdf": true,
    "output_enable_md": true,
    "output_enable_merged": true,
    "output_enable_independent": true,
    "merge_mode": "all_in_one",
    "merge_source": "source",
    "enable_update_package": true,
    "incremental_registry_path": ""
  },
  "status": "idle",
  "last_run_at": null
}
```

### 3.2 checkpoint 结构（断点续传）

```jsonc
{
  "task_id": "task_2026_demo",
  "run_id": "20260212T120000",
  "planned_files": [
    "E:/.../file1.docx",
    "E:/.../file2.pptx"
  ],
  "completed_files": [
    "E:/.../file1.docx"
  ],
  "created_at": "2026-02-12T12:00:00",
  "updated_at": "2026-02-12T12:10:00"
}
```

### 3.3 增量账本按任务隔离

- 每个任务拥有独立增量账本路径：

```text
<target_folder>/_AI/registry/task_<task_id>_incremental_registry.json
```

- 好处：
  - 不同任务间互不影响（即便源目录有交集）；
  - 便于根据任务删除/归档其增量历史。

---

## 4. 任务模式下的运行流程

### 4.1 用户视角

1. 在 Header 中选择 **任务模式**；
2. 打开“任务管理”Tab：
   - 若没有任务 → 使用新建任务向导创建；
   - 若已有任务 → 选中一个任务；
3. 点击顶部 `开始` 按钮；
4. 运行中可查看进度、日志；
5. 可在运行中选择“停止”，之后通过“续传”继续本轮；
6. 运行完成后，任务列表中更新“最后运行时间/状态”等。

### 4.2 内部流程（含增量、checkpoint）

1. **选择任务**
   - 从 `TaskStore` 读取 `task` 和对应 checkpoint（如存在）。

2. **决定本轮入口：新一轮 vs 续传**
   - 若存在未完成 checkpoint：
     - `planned_files` 与 `completed_files` 已存在；
     - 本轮待处理列表 = `planned_files - completed_files`；
     - 不再重新扫描源目录。
   - 若不存在 checkpoint：
     - 合并配置得到 `effective_config`；
     - 按 `run_mode` 和过滤条件扫描源目录；
     - 若 `run_incremental_by_default == true`：
       - 使用该任务专属增量账本过滤 added/modified 文件；
     - 将得到的待处理文件列表记为 `planned_files`，写入 checkpoint。

3. **执行批处理**
   - 将待处理文件列表传给 `OfficeConverter`（可通过现有 `run()` + 扫描，也可扩展新增 `run_with_file_list`）；
   - 每处理完成一个文件：
     - 添加到 `completed_files`；
     - 更新 checkpoint（耐心程度视性能权衡，可定期批量写入）。

4. **结束与清理**
   - 若用户未中途停止且所有文件处理完：
     - 删除 checkpoint；
     - 更新任务状态与 `last_run_at`；
     - 刷新增量账本与增量包。
   - 若用户中途停止：
     - 保留 checkpoint（`planned_files` + `completed_files`）；
     - 标记任务状态为 `paused`；
     - 下次运行同一任务时优先提供“续传”选项。

---

## 5. GUI 设计要点

### 5.1 顶部 Header：运行模式开关

- 位置：`office_gui.py` 中 `_build_ui` 的 `header_frame` / `ctrl_frame`。
- 控件：
  - `self.var_app_mode = tk.StringVar(value="classic")`；
  - 两个选项：
    - 传统模式（classic）；
    - 任务模式（task）。
- 文案：
  - `app_mode_classic`：传统模式；
  - `app_mode_task`：任务模式；
  - `tip_app_mode`：解释使用场景。
- 持久化：
  - 在 `config.json` 中新增 `app_mode` 字段；
  - `_load_config_to_ui` 时读取，保存配置时写回。

### 5.2 任务管理 Tab

- Notebook 中已有 `tab_run_tasks` + `_scroll_tasks` 容器。
- 推荐布局：
  - 左侧：任务列表（Treeview）：名称、源/目标、状态、最后运行时间；
  - 右侧：任务详情/编辑区域 或 任务新建向导入口。
- 工具栏按钮：
  - 新建任务（打开向导）；
  - 编辑任务；
  - 删除任务；
  - 运行任务；
  - 停止任务；
  - 断点续传；
  - 全量重跑本任务。

### 5.3 新建任务向导（4 步）

> 详见已有文档 `2026-02-12-task-management-design.md` 第 10 章，这里仅列关键点。

1. **基础信息**：名称、描述（校验名称唯一且非空）。
2. **目录与增量策略**：
   - 源目录列表（多选）；
   - 目标目录；
   - 是否默认使用增量、Hash 校验、重命名处理等。
3. **运行模式与输出**：
   - run_mode、输出格式开关、merge 模式/来源；
   - LLM 输出、增量包等高阶选项在高级折叠区。
4. **确认与保存**：
   - 汇总展示前三步关键信息；
   - 保存任务 JSON & 更新索引；
   - 可选“完成后立即运行任务”。

---

## 6. 实现步骤建议（任务模式）

1. **数据层**
   - 实现 `TaskStore`：负责读写任务、索引、checkpoint；
   - 规范 `tasks/` 目录结构与 JSON 字段名。
2. **运行层**
   - 封装一个“运行任务”的入口函数：
     - 输入：`task_id`；
     - 步骤：合并配置 → 读取/生成 checkpoint → 调用 `OfficeConverter` → 更新状态。
   - 视需要扩展 `OfficeConverter`（例如支持 `run_with_file_list` 或文件计划回调）。
3. **GUI 层**
   - 完成任务管理 Tab（列表 + 工具栏）；
   - 实现新建任务向导 UI；
   - 在 `_on_click_start` 里根据 `app_mode` 分支调用：
     - `task`：运行当前选中任务；
     - `classic`：执行原有逻辑。

---

## 7. 传统模式（非任务模式）的保留策略

### 7.1 行为定义

- 当 `app_mode == "classic"` 时：
  - `开始` 按钮：
    - 使用当前界面参数（共享参数/转换/合并等 Tab）与 `config.json` 合并后的结果；
    - 不依赖 `TaskStore` 和任务选择；
    - 增量逻辑沿用原有全局实现（不做 per-task 隔离保证）。
  - 任务管理页：
    - 可以选择保留为“只读/编辑”界面，禁用“运行/续传”按钮；
    - 或允许从任务中“加载配置到界面”，作为创建配置模板的手段。

### 7.2 文档提示

- 在“运行模式总览”和“传统模式”小节中明确写出：
  - 若需要长期维护某个目录的增量更新与断点续传，应使用任务模式；
  - 传统模式仅保证“像旧版本一样能跑”，不保证所有新特性。

---

## 8. 已有实现与设计的一致性核对（供接手者参考）

- `office_gui.py`：
  - 已有 `tab_run_tasks`、`_scroll_tasks`、`TaskStore(self.script_dir)` 等挂载点；
  - **传统模式**：`_update_task_tab_for_app_mode()` 在 `app_mode==classic` 时隐藏任务 Tab；`_on_run_mode_change` 末尾调用该函数，保证切换运行模式后仍按应用模式显隐任务页；
  - **任务模式**：任务 Tab 显示，`btn_start` 走 `_on_click_start` → 选中任务后 `_on_click_task_run`；
  - **传统模式**：`_on_click_start` 使用当前 UI 参数与 config 合并后运行，不依赖任务选择；
  - `app_mode` 在 `_load_config_to_ui` 中读取、保存配置时写回；
  - 新建任务向导 `_open_task_wizard()` 含 4 步（基础信息、目录与增量、运行模式与输出、确认与保存）。
- `office_converter.py`：
  - 已有增量账本、增量包实现；
  - 需要在上层调用时，按任务维度指定独立的 `incremental_registry_path`。
- **自动化测试**：`tests/test_gui_task_mode.py` 覆盖 app_mode 配置回合、传统模式隐藏任务 Tab、任务模式显示 Tab（需 ttkbootstrap）、向导 4 步 key、Windows 下步骤标签、开始按钮按 app_mode 分支；`tests/test_task_manager.py` 覆盖 TaskStore、checkpoint、build_task_runtime_config。运行：`python -m unittest discover -s tests -p "test_*.py" -v`。
- **测试报告**：各功能模块的用例清单与验证结果见 `docs/test-reports/`，含 GUI 任务模式、任务管理、转换器续传、合并/转换流水线、输出控制等分报告及索引 [README](../test-reports/README.md)。

---

## 9. 小结

通过引入“任务模式 + 传统模式共存”的运行体系，本方案实现：

- 不破坏现有使用方式的前提下，引入了适合长期维护文档库的任务化工作流；
- 用统一的 effective_config 和运行入口，避免两套完全割裂的逻辑；
- 把增量账本、断点续传、LLM 输出等高级能力，集中托管在“任务模式”下实现与文档约定。

后续若需要扩展（定时调度、多任务队列、任务模板等），均可直接在任务模式的这一套结构上迭代，无需再做大规模重构。

