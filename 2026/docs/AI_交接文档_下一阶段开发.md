# 知喂 (ZhiWei) — 下一阶段开发交接文档

供下一位开发者或 AI 接续开发时使用。文档日期：2026-02-13。

---

## 1. 项目概览

| 项目 | 说明 |
|------|------|
| **产品名** | 知喂 (ZhiWei)，副标题：知识投喂工具 |
| **版本** | v5.17.0（唯一定义在 `office_converter.py` 的 `__version__`） |
| **仓库** | 私有仓库，主分支 `main`；项目代码在 `2026/` 目录下 |
| **界面** | 仅中文，无英文模式 |

**核心定位**：为知识库与 AI 服务准备语料——Office 批量转 PDF、合并、梳理，以及面向 NotebookLM 的溯源与 LLM 上传目录。

---

## 2. 代码结构与入口

### 2.1 主要文件

| 文件 | 作用 |
|------|------|
| `office_converter.py` | 核心：转换 / 合并 / 梳理 / MSHelp / 增量 / LLM 归集；`OfficeConverter` 类，`run()` 主流程，版本号 `__version__` |
| `office_gui.py` | GUI：Tk + ttkbootstrap，单层 7 Tab（模式与路径、转换选项、合并梳理、MSHelp、快速定位、成果文件、高级设置）；经典模式 + 任务模式 |
| `ui_translations.py` | 界面文案（当前仅中文），`TRANSLATIONS["zh"]`，`tr(key)` 取文案 |
| `task_manager.py` | 任务存储与 checkpoint，`tasks/` 目录与 `tasks_index.json` |
| `build_exe.py` | 一键打包：先清空 dist/build，再 PyInstaller，产出 `ZhiWei_v<版本>.exe` |

### 2.2 运行方式

```bash
# 图形界面（推荐）
python office_gui.py

# 命令行
python office_converter.py --source "C:\源路径" --target "D:\目标路径" --run-mode convert_then_merge
python office_converter.py --help   # 查看参数
```

### 2.3 配置与敏感文件（勿提交）

- `config.json`：运行时配置，首次运行自动生成；已加入 `.gitignore`，仓库内仅保留 `config.example.json`。
- `2026/tasks/*.json`、`2026/config_profiles/*.json`：用户数据，已忽略，勿提交。

---

## 3. 已实现能力摘要

- **运行模式**：`convert_only` / `merge_only` / `convert_then_merge` / `collect_only` / `mshelp_only`
- **多源目录**：支持多选源文件夹，转换/合并/归集均会扫描并处理所有选中目录
- **合并输出**：`merge_filename_pattern` 可配置，占位符 `{category}`, `{timestamp}`, `{date}`, `{time}`, `{idx}`
- **任务模式**：多组「源+目标+参数」保存与一键运行，断点续传
- **增量同步**：账本、Added/Modified/Renamed/Deleted、增量包、MD5 去重、同名优先 Office
- **LLM 上传目录**：`_LLM_UPLOAD` 集中输出，manifest、扁平化、去重策略
- **沙盒与空间**：`temp_sandbox_root`、`sandbox_min_free_gb`、低空间策略（block/confirm/warn）
- **MSHelp**：扫描 MSHelpViewer/CAB，转 Markdown，索引与合并包
- **NotebookLM 溯源**：合并 PDF + 页码或短 ID 定位，Everything/Listary 集成
- **产物目录**：`_MERGED/`、`_AI/Markdown/`、`_AI/ExcelJSON/`、`_AI/Records/`、`_AI/ChromaDB/`、`_AI/Update_Package/`、`_LLM_UPLOAD/` 等

---

## 4. 建议的下一阶段功能（优先级供参考）

以下来自现有需求与任务清单，接续开发时可从中选取：

### 4.1 高优先级

1. **大批量压测与文档**  
   - 对 ≥10k 文件场景做压测，记录磁盘占用与耗时；将 V1.1 验收用例与文档补全（见 `docs/TASK_LIST.md` Phase 6 未勾选项）。

2. **Google Drive 上传（桌面 OAuth）**  
   - 配置项：`enable_gdrive_upload`、`gdrive_folder_id`、`gdrive_client_secrets_path`。  
   - GUI：「上传 _LLM_UPLOAD 到 Google Drive」按钮；首次浏览器 OAuth，本地缓存 token。  
   - 详见 `docs/AI_HANDOVER_20260211.md` 第 11 节。

### 4.2 中低优先级

- 非 Windows 下沙盒可用空间检测的兼容与降级策略（保守 warn）。
- LLM 归集可选硬链接/符号链接模式（当前为复制，减少重复占用）。
- 其它见 `docs/TASK_LIST.md` Phase 7 及 `docs/PRODUCT_REQUIREMENTS.md`。

---

## 5. 实现时需注意的约束

- **兼容性**：不删除、不破坏现有 `_AI/*` 输出结构；新增能力以“叠加”为主（如 LLM 目录、新配置项）。
- **配置**：新功能尽量有对应 `config.json` 键与 GUI 控件，并在 `office_converter.py` 的默认配置/加载处补齐。
- **文案**：界面新增文案放入 `ui_translations.py` 的 `zh` 字典，键名语义清晰。
- **版本**：仅改 `office_converter.py` 的 `__version__`；打包与文档中的版本号均由此带出。
- **安全**：不要提交 `config.json`、用户任务与配置 profile；路径与示例用占位符（如 `config.example.json`）。

---

## 6. 关键文档索引

| 文档 | 用途 |
|------|------|
| `docs/AI_HANDOVER_20260211.md` | 英文交接：设计要点、LLM hub、沙盒、验收标准、实现顺序 |
| `docs/TASK_LIST.md` | 任务清单与 Phase 划分，未勾选为待办 |
| `docs/PRODUCT_REQUIREMENTS.md` | 产品需求与 P0/P1/P2 能力说明 |
| `使用说明书.md` | 用户向：界面、流程、配置、常见问题 |
| `CHANGELOG.md` | 版本更新记录 |
| `README.md` | 仓库/项目说明、安装、运行、打包 |

---

## 7. 建议的实现顺序（新功能）

1. 在 `office_converter.py` 中增加或扩展能力（新配置键、默认值、主流程调用点）。
2. 在 `office_gui.py` 中增加对应控件与运行时参数传递。
3. 在 `ui_translations.py` 中增加文案键。
4. 更新 `使用说明书.md`、`CHANGELOG.md`，必要时更新 `docs/TASK_LIST.md` 与 `docs/PRODUCT_REQUIREMENTS.md`。
5. 运行一次完整转换+合并+归集流程做回归；若有新依赖，在 `requirements.txt` 与 `build_exe.py` 中补齐。

---

## 8. 快速自检命令

```bash
# 语法检查
python -m py_compile office_converter.py office_gui.py ui_translations.py

# 单元测试（若有）
python -m unittest discover -s tests -p "test_*.py" -v
```

---

*本文档随版本与需求更新，当前对应 v5.17.0。*
