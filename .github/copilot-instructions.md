**Repository Overview**

- **Purpose**: 这是一个 Windows 下的 Office 文件批量转换与梳理工具，支持 CLI 与 GUI 两种运行方式（WPS / Microsoft Office）。主要逻辑在 [office_converter.py](office_converter.py#L1-L40)，GUI 在 [office_gui.py](office_gui.py#L1-L40)。

**How This Project Runs**

- **CLI**: 主入口位于 [office_converter.py](office_converter.py#L1800-L2050)。典型运行：`python office_converter.py --config config.json`（脚本会在首次运行时生成默认 `config.json`）。
- **GUI**: 在 [office_gui.py](office_gui.py#L800-L835) 启动 Tk 窗口，界面会覆盖但不写回 `config.json`（界面有“保存配置”按钮写入）。

**Key Config & Behaviour**

- **Main config file**: [config.json](config.json#L1-L40). 重要字段：
  - **source_folder / target_folder**: 源与目标目录，必须为绝对路径。
  - **enable_sandbox / temp_sandbox_root**: 是否使用临时目录转换（默认启用）。
  - **default_engine**: `wps` / `ms` / `ask`（GUI 会覆盖此值）。
  - **kill_process_mode**: `ask` / `auto` / `keep` — 控制是否杀 Office 进程（CLI 的 `ask` 会交互，GUI 不应使用 `ask`）。
  - **enable_merge / merge_mode / max_merge_size_mb**: 合并行为与大小阈值（见合并逻辑在 office_converter.py）。

**Patterns & Conventions for Code Changes**

- **Dual-mode design (CLI vs GUI)**: `OfficeConverter` 在 [office_converter.py](office_converter.py#L1-L40) 为核心类；GUI 使用 `GUIOfficeConverter`（在 [office_gui.py](office_gui.py#L1-L120)）覆写所有会触发 `input()` 的方法。若修改运行参数，优先考虑两者的交互差异。
- **Config-first**: 大部分运行时参数从 `config.json` 加载并可被 UI 覆盖但不自动写回（除非用户点击“保存配置”）。代码里经常通过 `converter.config` 在运行时覆盖。
- **Platform / dependency assumptions**: 只在 Windows 上运行（依赖 `win32com.client` / `pythoncom`），对 Office COM 自动化有显式进程管理（`tasklist`/`taskkill`）。注意：更改进程管理路径或名称需兼顾 WPS 与 MS Office 两套映射。

**Dependencies & Optional Features**

- Required (Windows): `pywin32` (win32com)，`pythoncom`。
- Optional: `pypdf`（PDF 合并，检查 HAS_PYPDF in [office_converter.py](office_converter.py#L1-L40)），`openpyxl`（collect_only 时写 Excel 索引）。代码会根据导入是否成功启用/禁用对应功能。

**Common Change Patterns / Examples**

- 如果添加新的转换后处理步骤，请：
  1. 在 `run_batch` / `merge_pdfs` 周边调整（参见 [office_converter.py](office_converter.py#L1700-L1850) 的 run/merge 流程）。
  2. 在 GUI 中暴露对应参数时，更新 `_load_config_to_ui` 和 `save_config_from_ui`（参见 [office_gui.py](office_gui.py#L320-L420) / [office_gui.py](office_gui.py#L640-L740)）。
- 若更改文件扫描或排除逻辑，更新 `config.json` 字段 `excluded_folders` 的读取方式（见 [office_converter.py](office_converter.py#L1820-L1840)）。

**Testing & Debugging Tips**

- 在修改与 Office 互操作或进程管理相关代码时，先在一个隔离目录测试并把 `enable_sandbox` 设为 `True`，以避免直接破坏源文件。
- 日志会写入 `log_folder`（默认 `./logs`），GUI 会把 stdout/stderr 投到界面日志区；用日志来追踪长流程问题（参见 `setup_logging` in [office_converter.py](office_converter.py#L300-L360)）。

**What Not to Change Without Caution**

- 不要在没有考虑 GUI 的情况下直接调用 `input()`；如果需要交互，应像 `GUIOfficeConverter` 那样提供无交互替代。
- 不要默认在 GUI 模式下把 `kill_process_mode` 设为 `ask`，因为 GUI 无法响应 CLI 交互（代码里已有判断）。

如果有任何部分需要更详细的说明（例如具体函数注释或常见故障堆栈），请告诉我想要扩展的部分或提供你希望突出的问题清单，我会迭代更新本文件。