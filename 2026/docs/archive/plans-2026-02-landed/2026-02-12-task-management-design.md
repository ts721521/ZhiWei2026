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
    "E:\\21_SE_Doc\\BaiduNetdiskWorkspace\\SynologyDrive\\5_投标\\2026"
  ],
  "target_folder": "C:\\PDFs",
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
3. **运行模式与输出（引导式）**：
   - 第一层只展示“我要什么结果”：
     - 仅独立文件；
     - 仅合并文件；
     - 独立 + 合并都要。
   - 第二层（仅当选择“需要合并”时）展示“怎么合并”：
     - 合成一个总文件（对应 `merge_mode=all_in_one`）；
     - 按大小分卷（推荐，默认 80MB，对应 `merge_mode=category_split` + `max_merge_size_mb`）。
   - 第三层才展示高级参数（`merge_source`、LLM 输出、增量包等）。
4. **确认与保存**：
   - 汇总展示前三步关键信息；
   - 保存任务 JSON & 更新索引；
   - 可选“完成后立即运行任务”。

---

### 5.4 输出控制引导化（新增）

#### 5.4.1 目标

- 解决用户对“合并输出”和“合并模式”的概念混淆；
- 避免“设置了 80MB 但未生效”的预期落差；
- 保持底层配置兼容（不破坏既有 `config.json` 和任务 JSON）。

#### 5.4.2 交互原则

- 先问业务目标，再映射技术参数；
- 默认值可直接运行，减少必须理解的配置项；
- 每个选择都给出“将产生什么文件”的即时预览。

#### 5.4.3 引导流程（任务向导与运行页共用）

1. **你要哪些产物？**
   - `仅独立文件`
   - `仅合并文件`
   - `独立 + 合并`
2. **如果要合并 PDF，希望怎么出包？**
   - `单个总文件（不按大小分卷）`
   - `按大小分卷（推荐）`
3. **如果按大小分卷，单卷上限是多少？**
   - 默认 `80 MB`，可改为正整数。
4. **确认页展示“用户语义 + 技术映射”**
   - 示例：`按大小分卷(80MB)` -> `output_enable_merged=true, merge_mode=category_split, max_merge_size_mb=80`

#### 5.4.4 配置映射规则（兼容旧字段）

- `需要合并 = false` -> `output_enable_merged=false`（忽略 `merge_mode` 与 `max_merge_size_mb`）。
- `需要合并 = true` 且选择 `单个总文件` -> `merge_mode=all_in_one`。
- `需要合并 = true` 且选择 `按大小分卷` -> `merge_mode=category_split`，并启用 `max_merge_size_mb`。
- 反向加载（旧配置回填 UI）：
  - `merge_mode=all_in_one` -> 显示“单个总文件”；
  - `merge_mode=category_split` -> 显示“按大小分卷”并回填阈值。

#### 5.4.5 防呆与提示文案

- 当选择 `all_in_one` 时：
  - 禁用 `max_merge_size_mb` 输入框；
  - 显示提示：`当前模式不会按大小分卷。`
- 当选择 `category_split` 时：
  - `max_merge_size_mb` 必填且必须为正整数；
  - 预览文本显示：`将按分类并按大小分卷输出。`
- 汇总栏始终显示一行“本次合并策略摘要”，避免用户只看单个控件误判。

#### 5.4.6 验收标准（引导化）

1. 用户仅勾选“合并输出”，未选分卷策略时，UI 能明确展示当前默认策略。
2. 用户选择“按大小分卷 80MB”后，确认页和保存结果均包含 `max_merge_size_mb=80`。
3. 用户切换到“单个总文件”后，大小输入框自动禁用，且提示“80MB 不生效”。
4. 旧任务（仅有 `merge_mode`）加载到新 UI 时，策略显示正确且可无损保存。
5. 任务模式与传统模式在该引导逻辑上的交互行为一致（仅配置来源不同）。

---

### 5.5 用户操作模拟测试（脑内走查）

> 方法：基于当前 `office_gui.py` 交互逻辑做“用户路径模拟”（cognitive walkthrough），用于发现理解成本高的节点。  
> 参考实现点：`_on_run_mode_change`、`_build_task_overrides_from_ui`、`_save_config_sections_to_file`。

#### 5.5.1 场景 A：用户想要“转换并合并，且按 80MB 分包”

1. 用户选择 `run_mode=convert_then_merge`。  
2. 用户勾选“合并输出”，看到“最大MB=80”，认为“已启用按 80MB 分包”。  
3. 用户未注意“合并模式”当前为 `all_in_one`（常见于从历史配置加载）。  
4. 运行后得到单个大 PDF，用户预期落空。

结论：当前 UI 中“大小阈值输入框”与“是否生效”缺少联动提示，易造成误判。

#### 5.5.2 场景 B：用户切换到不相关模式，仍看到可编辑配置

1. 用户在 `merge_only` 配完合并参数。  
2. 切换到 `collect_only` 或 `mshelp_only`。  
3. 部分合并参数虽逻辑上无效，但视觉上仍可能可见或不够明确。  
4. 用户不确定“当前改动是否会影响本次运行”。

结论：需要“模式相关配置矩阵 + 灰显 + 说明文案”三件套，避免无效编辑。

#### 5.5.3 场景 C：加载旧配置与当前模式冲突

1. 用户加载旧配置（包含 `merge_source=source`、`merge_mode=all_in_one`）。  
2. 当前运行模式为 `convert_then_merge`。  
3. 系统内部会强制 `merge_source=target`，但用户界面未必明确看到“已自动纠偏”。

结论：必须有“自动校正提示条”，告诉用户哪些字段已被模式规则覆盖。

### 5.6 引导提示与功能说明增强（新增）

#### 5.6.1 文案策略

- 每个关键选项采用“名称 + 一句话结果描述”：
  - `合并输出`：是否生成合并产物文件。
  - `单个总文件`：合并成一个 PDF，不按大小分卷。
  - `按大小分卷`：按阈值拆分多个合并包（默认 80MB）。
- 在参数区下方固定展示“本次输出摘要”：
  - 示例：`将输出：独立PDF + 合并PDF（按分类，80MB分卷）`。

#### 5.6.2 即时提示（Inline Help）

- 当 `merge_mode=all_in_one`：
  - 展示黄条提示：`当前为单文件模式，大小阈值不参与分包。`
- 当 `merge_mode=category_split`：
  - 展示绿条提示：`当前将按分类并按大小分卷输出。`
- 当 `output_enable_merged=false`：
  - 展示灰条提示：`已关闭合并输出，以下合并参数仅作预设，不参与本次运行。`

#### 5.6.3 “为什么灰色”可解释

- 所有被禁用控件支持统一 tooltip：
  - `因当前运行模式为 {mode}，此项不参与本次执行。`
- 减少“灰了但不知道为什么”的挫败感。

### 5.7 模式驱动的配置可编辑矩阵（新增）

#### 5.7.1 规则总则

- 只要“当前模式不会使用该配置”，该控件必须灰显不可编辑。
- 灰显不等于丢值：保留历史值用于模式切回后的恢复。
- 运行时只取“当前模式相关字段”，其余字段进入 `inactive_config`（仅展示，不执行）。

#### 5.7.2 模式-配置分组矩阵

1. `convert_only`
   - 启用：转换相关、增量相关、输出格式（PDF/MD）、独立输出。
   - 禁用：合并策略（merge_mode/max_merge_size_mb/merge_source/merge_index/merge_excel）、collect、mshelp 专属。
2. `convert_then_merge`
   - 启用：转换相关、增量相关、输出格式、合并策略。
   - 强制：`merge_source=target`（控件只读显示“目标目录”）。
3. `merge_only`
   - 启用：合并策略、输出格式、merge 子模式。
   - 禁用：转换引擎、沙箱、增量扫描、日期过滤等转换专属项。
4. `collect_only`
   - 启用：collect 相关配置。
   - 禁用：转换与合并参数（含 `max_merge_size_mb`）。
5. `mshelp_only`
   - 启用：MSHelp 相关配置。
   - 禁用：常规合并参数（与 MSHelp 合并分开管理时）。

#### 5.7.3 合并区内部联动

- `output_enable_merged=false`：
  - 灰显：`merge_mode`、`max_merge_size_mb`、`merge_source`、`enable_merge_index`、`enable_merge_excel`。
- `output_enable_merged=true` 且 `merge_mode=all_in_one`：
  - 灰显：`max_merge_size_mb`（显示“仅分卷模式生效”）。
- `output_enable_merged=true` 且 `merge_mode=category_split`：
  - 启用：`max_merge_size_mb`（正整数校验）。

### 5.8 后端配置与当前模式强一致（新增）

#### 5.8.1 统一收敛函数

- 在运行前增加统一收敛步骤（建议函数名：`sanitize_runtime_config_for_mode`）：
  - 输入：UI 当前配置 + 当前 `run_mode`；
  - 输出：`effective_runtime_config` + `inactive_config` + `coercion_messages[]`。

#### 5.8.2 必须执行的纠偏规则

- `run_mode=convert_then_merge` -> 强制 `merge_source=target`。
- `output_enable_merged=false` -> 合并字段不参与执行（保留但不生效）。
- `merge_mode=all_in_one` -> `max_merge_size_mb` 标记为 inactive。
- `merge_mode=category_split` -> 校验 `max_merge_size_mb>=1`，非法时回退 `80` 并提示。

#### 5.8.3 用户可见反馈

- 启动任务前弹出一次“配置校正摘要”或写入日志首段：
  - 示例：`已按当前模式自动调整：merge_source source->target；max_merge_size_mb 已忽略（all_in_one）。`
- 用户点击“查看详情”可展开完整映射表，减少“系统偷偷改我配置”的不信任感。

#### 5.8.4 验收测试（模拟测试用例）

1. 切换到 `collect_only` 时，合并相关控件全部灰显，且 tooltip 说明原因。  
2. `convert_then_merge + all_in_one` 时，`max_merge_size_mb` 灰显并显示“不生效”提示。  
3. `convert_then_merge + category_split + 80` 时，运行摘要明确显示“80MB 分卷”。  
4. 加载冲突配置后，系统显示“已自动校正”明细，并能继续运行。  
5. 任务模式与传统模式使用同一套 `sanitize` 规则，结果一致。

---

### 5.9 配套文档交付（新增）

为降低个人使用时的认知负担，交付两份长期维护文档：

1. 操作手册：`docs/操作手册-引导模式.md`
   - 按“用户点击顺序”写，不按代码结构写；
   - 每个区域都说明“改这个会影响什么”。
2. 流程文件：`docs/流程文件-引导模式.md`
   - 明确“模式 -> 可编辑控件 -> 收敛规则 -> 日志反馈”链路；
   - 用于开发、测试、排障三方统一口径。

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
   - 在“运行模式与输出”中加入引导式流程（产物选择 -> 合并策略 -> 分卷阈值 -> 映射预览）；
   - 将“合并输出 / 合并模式 / 大小阈值”的状态联动收敛到统一函数，避免多处散落判断；
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
  - `_on_run_mode_change` 中 `self._set_run_tab_state(self.tab_run_tasks, "normal")`，保证任务页常驻；
  - `btn_start` 统一绑定 `_on_click_start`，是唯一运行入口。
- `office_converter.py`：
  - 已有增量账本、增量包实现；
  - 需要在上层调用时，按任务维度指定独立的 `incremental_registry_path`。

---

## 9. 小结

通过引入“任务模式 + 传统模式共存”的运行体系，本方案实现：

- 不破坏现有使用方式的前提下，引入了适合长期维护文档库的任务化工作流；
- 用统一的 effective_config 和运行入口，避免两套完全割裂的逻辑；
- 把增量账本、断点续传、LLM 输出等高级能力，集中托管在“任务模式”下实现与文档约定。

后续若需要扩展（定时调度、多任务队列、任务模板等），均可直接在任务模式的这一套结构上迭代，无需再做大规模重构。
