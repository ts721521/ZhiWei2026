# 产品需求文档（PRD）与技术规格

## 1. 背景与目标
- 当前工具已具备：Office 转 PDF、PDF 合并、索引映射、反向定位。
- 新阶段目标：把工具升级为“可供 NotebookLM / RAG 稳定消费的数据生产器”。
- 核心原则：先做高价值、低风险、可落地能力，再扩展到复杂语义能力。

## 2. 能力分层
### 2.1 P0（优先上线）
- `Corpus Manifest`：生成 `corpus.json`，统一记录本次产物、路径、哈希、时间、来源映射。
- `Excel 结构化导出`：为 Excel 增加 JSON Records 导出（供 Code Interpreter/程序处理）。
- `Markdown 导出基础版`：先实现稳定文本导出，再逐步增强清洗质量。

### 2.1.1 当前落地状态（2026-02-10）
- 已完成：`Corpus Manifest`（`enable_corpus_manifest`，默认开启）。
- 已完成：GUI 开关（运行页共用配置中可开启/关闭 `corpus.json` 生成）。
- 已完成：GUI 产物摘要（每步执行后输出转换/合并/map数量及关键索引路径）。
- 已完成：基础 Markdown 导出（基于转换后 PDF 文本抽取）。
- 已完成：Markdown 清洗增强（页眉/页脚与页码噪声清理、标题/段落结构优化）。
- 已完成：Markdown 质量报告（按批次汇总清洗统计，并输出抽样记录）。
- 已完成：索引 Records JSON 导出（convert/merge）。
- 已完成：Excel 专项 JSON 导出（语义增强 v2，含表头识别、空行剔除、records 预览、列类型统计、公式样本、跨表引用统计、合并单元格范围信息；输出 `_AI/ExcelJSON`，无 `openpyxl` 时写入降级状态）。
- 已完成：Excel JSON 深度语义补充（工作簿命名区域 `defined_names`、工作表图表元数据、透视表元数据）。
- 已完成：增量同步 v2（`FileRegistry` 账本、Added/Modified/Renamed/Unchanged/Deleted 统计、同目录同名 Office 优先、同类型 MD5 去重、`Update_Package` 产物、GUI 开关）。
- 已完成：核心 CLI 输出乱码清理（merge/collect 主流程提示、日志文案与索引表头）。
- 已完成：ChromaDB 适配基础版（可选开关、Markdown 分块、PersistentClient 入库、JSONL 回退清单）。
- 已完成：增量设计文档校正（`INCREMENTAL_SYNC_DESIGN.md` 乱码修复并与当前实现字段/路径对齐）。
- 已完成：运行页转换配置分区细化（核心转换 / 过滤策略 / AI 导出 / 增量同步）。
- 已完成：配置中心分层补齐（新增 `AI默认配置` 与 `增量默认配置` Tab）。
- 已完成：配置中心二次拆分（AI 默认项与增量默认项进一步分组，降低参数混放与误配风险）。
- 已完成：`Shared Defaults` 再拆分（配置路径 / 进程策略 / 日志输出）。
- 已完成：`Merge Defaults` 再拆分（合并行为 / 合并产物），并在配置中心新增“持久化默认值”提示文案。
- 已完成：`Rules & Keywords` 再拆分（排除规则 / 关键词策略）。
- 已完成：配置中心分区级重置（每个配置页支持“重置本分区”，用于快速回到内置默认值，需手动保存后生效到配置文件）。
- 已完成：配置中心未保存状态提示（默认值被修改后显示“尚未保存”，执行保存全部后恢复为“已保存”）。
- 已完成：保存后状态精确回算（`Save Current Mode` 也会按当前 UI 与 `config.json` 差异自动判断 dirty 状态，避免误清空）。
- 已完成：配置分区级 dirty 提示（对应配置 Tab 在存在未保存改动时追加 `*`，便于快速定位改动范围）。
- 已完成：未保存分区列表提示（配置中心顶部展示具体未保存分区名称，辅助保存前复核）。
- 已完成：未保存分区快速跳转（“跳转未保存分区”按钮可直接定位到首个未保存配置 Tab）。
- 已完成：未保存分区点击跳转（顶部未保存分区名称支持逐项点击，直达对应配置 Tab）。
- 已完成：仅保存未保存分区（新增“保存未保存分区”动作，仅写回 dirty 分区，降低配置误覆盖风险）。
- 已完成：分区保存结果反馈（“保存未保存分区”后明确展示已保存分区名称；无改动时给出“无未保存分区”提示）。
- 已完成：分区级单独保存（每个配置分区新增“保存本分区”，支持按需落盘，降低跨分区误覆盖风险）。
- 已完成：分区回滚能力（新增“恢复未保存分区”动作，可将 dirty 分区回滚到 `config.json` 基线，并提示已恢复分区）。
- 已完成：未保存分区计数提示（保存/恢复按钮动态展示待处理分区数量，提升操作可见性）。
- 已完成：分区回滚安全增强（恢复前二次确认回滚范围，且执行前先刷新 `config.json` 基线以避免过期回滚）。
- 已完成：回滚确认可配置（UI 默认配置新增开关，可关闭“恢复未保存分区”前的确认弹窗，并持久化到 `config.json`）。
- 已完成：转换运行页配置二次分层（核心转换 / 过滤策略 / AI 导出 / 增量同步），降低配置混杂度。

### 2.2 P1（增强）
- Markdown 清洗增强：页眉页脚噪声识别、段落结构优化。
- 主索引增强：统一索引 PDF / Excel / JSON 关联关系，提升可追溯性。
- GUI 增强：提供 AI 导出开关（Markdown/Excel JSON）与任务结果摘要。

### 2.3 P2（后置）
- 向量库增强能力（检索策略、召回评估、在线服务化）。
- 说明：基础 ChromaDB 适配已上线；生产级检索质量与服务化仍属于后置阶段。

### 2.4 新增需求校正（2026-02-10）
- `增量同步与去重` 价值高，但风险也高（误判会直接影响数据完整性），建议按“账本 -> 检测 -> 去重 -> 打包”四步分阶段上线。
- `Source Priority` 规则需收敛为“同目录同名优先 Office，跨目录不做强跳过”，避免误伤同名不同文档。
- `Global MD5` 建议先用于“同类型输出去重”（例如 PDF 对 PDF），不跨类型（Office 原件 vs PDF）直接判重。
- `Update_Package` 不建议只输出单一合并 PDF，应至少包含：增量 PDF、增量索引、增量 manifest。

## 3. 与当前程序匹配性评估
### 3.1 已匹配（当前已有）
- 合并映射与定位链路：`*.map.json` + `locate_source.py`（按页码/短 ID）。
- 合并索引导出：支持 Excel 索引与定位辅助。
- 多模式运行：转换、合并、转换后合并、归集索引。

### 3.2 部分匹配（建议先补）
- AI 数据导出开关：已具备 `Markdown` / `Excel JSON` / `Records JSON` 开关。

### 3.3 已补齐（基础版）
- 向量库直连（ChromaDB）已提供基础适配：从 Markdown 分块写入向量库，失败时输出 JSONL/manifest 便于离线追溯。

## 4. 数据契约（Corpus Manifest）
- 输出文件：`<target_folder>/corpus.json`
- 开关配置：`enable_corpus_manifest`（默认 `true`）
- 建议字段：
  - `version`
  - `generated_at`
  - `run_mode`
  - `collect_mode`
  - `merge_mode`
  - `content_strategy`
  - `source_folder`
  - `target_folder`
  - `artifacts[]`
  - `conversion_records[]`
  - `merge_records[]`
  - `summary`
- `artifacts` 单项建议字段：
  - `kind`（如 `converted_pdf` / `merged_pdf` / `merge_map_json` / `index_excel`）
  - `path_abs`
  - `path_rel_to_target`
  - `size_bytes`
  - `mtime`
  - `md5`
  - `sha256`

## 4.1 数据契约（Excel JSON，语义增强 v2）
- 输出目录：`<target_folder>/_AI/ExcelJSON/`
- 关键字段：
  - `parse_status`（`ok` / `openpyxl_missing` / `unsupported_format_xls` / `parse_failed`）
  - `limits`（行列上限与语义提取开关）
  - `sheets[]`
    - `header_detected`、`header`
    - `rows`（清洗后的二维数据）
    - `records_preview`（按表头展开的记录预览）
    - `column_profiles`（列类型统计与样本）
    - `formula_stats`（公式样本、跨表引用统计）
    - `merged_ranges`（合并单元格范围与左上值）
- `workbook_links[]`（工作簿级跨表引用边）
- `workbook_defined_names[]`（命名区域语义）
- `chart_count_total`、`pivot_table_count_total`（工作簿图表/透视表总量）
- `sheets[].charts[]`（图表类型/标题/锚点/系列引用）
- `sheets[].pivot_tables[]`（透视表名称/缓存ID/位置）

## 4.2 数据契约（Update_Package，增量包）
- 输出目录：`<target_folder>/_AI/Update_Package/Update_Package_YYYYMMDD_HHMMSS/`
- 关键文件：
  - `incremental_manifest.json`
  - `incremental_index.json`
  - `incremental_index.csv`（若环境具备 `openpyxl`，额外输出 `incremental_index.xlsx`）
  - `PDF/`（本次增量成功产出的 PDF 子集）

## 4.3 数据契约（ChromaDB 导出，基础版）
- 输出目录：`<target_folder>/_AI/ChromaDB/`
- 开关配置：`enable_chromadb_export`（默认 `false`）
- 关键文件：
  - `chroma_export_YYYYMMDD_HHMMSS.json`（导出状态与统计）
  - `chroma_docs_YYYYMMDD_HHMMSS.jsonl`（文档分块回退清单）
- 关键配置：
  - `chromadb_persist_dir`（留空时默认 `<target_folder>/_AI/ChromaDB/db`）
  - `chromadb_collection_name`
  - `chromadb_max_chars_per_chunk`
  - `chromadb_chunk_overlap`
  - `chromadb_write_jsonl_fallback`

## 5. 实施策略（当前阶段）
1. 已完成 `corpus.json`，先保持稳定运行（作为所有后续能力的数据基座）。
2. 先落地“AI 导出基础版”：配置开关 + GUI 开关 + 基础 Markdown + 索引 Records JSON。
3. 再推进增量同步与去重（按“账本 -> 检测 -> 去重 -> 打包”分步上线）。
4. 在基础适配稳定后，再评估检索质量优化与服务化集成，避免过早提升系统复杂度。

### 5.1 增量同步当前边界（v2）
- 已实现：账本落盘、新增/修改检测、删除统计、同目录同名优先规则、同类型 MD5 去重、GUI 配置。
- 已实现：Rename 检测（优先 hash 匹配，降级到 mtime/大小匹配），并在增量统计与增量包中输出 `renamed` 记录。

## 6. 验收标准（本阶段）
- 每次运行结束后稳定生成 `corpus.json`。
- Manifest 能覆盖关键产物（转换、合并、索引、映射）。
- 失败不会中断主流程；Manifest 生成失败仅记录日志不崩溃。

## 7. 本轮已交付（2026-02-10）
- `office_converter.py`
  - 新增产物采集与 `corpus.json` 写入。
  - 在主流程末尾接入 manifest 输出（失败不阻断主流程）。
  - 新增基础 Markdown 导出（从转换后 PDF 提取文本）。
  - 新增 Markdown 清洗与结构化（页眉/页脚/页码清理、标题/段落优化）。
  - 新增 Markdown 质量报告导出（`_AI/MarkdownQuality/`）。
  - 新增 Excel 专项 JSON 导出（语义增强 v2）。
  - 新增 Excel JSON 深度语义（命名区域、图表、透视表元数据）。
  - 新增索引 Records JSON 导出（convert/merge）。
  - 新增增量同步 v2：`FileRegistry`、变更检测、同目录同名 Office 优先、同类型 MD5 去重、`Update_Package`。
  - 修复 `ask_retry_failed_files` 语法损坏导致的启动失败问题。
  - 修复 `_compute_md5` 绑定问题，确保增量索引中的 `source_md5` 可稳定写出。
  - 清理 merge/collect 主流程终端输出与索引表头乱码。
  - 新增 ChromaDB 适配基础版：Markdown 分块、可选入库、JSONL 回退、manifest 记录。
- `office_gui.py`
  - 新增 `corpus.json` 开关。
  - 新增“本次产物摘要”日志输出。
  - 新增 AI 导出开关（Markdown / Records JSON）。
  - 新增增量/去重开关（增量扫描、哈希校验、重命名是否重转、同名优先、MD5 去重）。
  - 新增 ChromaDB 导出开关与产物摘要展示。
  - 修复跨页面滚动区在 Windows 下鼠标滚轮易失效的问题。
- `ui_translations.py`
  - 新增中英文文案（开关、提示词、产物摘要文案）。
