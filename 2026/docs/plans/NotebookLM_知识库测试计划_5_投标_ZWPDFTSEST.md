# NotebookLM 知识库测试计划

**目的**：用知喂（ZhiWei）将投标源目录转为适合 NotebookLM 的语料，输出到指定目标目录，并验证配置与产物是否符合 NotebookLM 要求。

**测试源目录**：`Z:\Schneider\5_投标`  
**测试目标目录**：`D:\ZWPDFTSEST`

---

## 一、NotebookLM 知识库要求（摘要）

根据公开说明与帮助文档：

| 项目 | 限制/要求 |
|------|-----------|
| **支持的文档类型** | PDF、Google Docs、Google Slides、文本/ Markdown、网页 URL、YouTube、音频（会转写为文本） |
| **单文件大小** | 不超过 200MB |
| **单来源字数** | 不超过 50 万词（约 50 万英文词 / 对应中文篇幅） |
| **每个 Notebook 来源数** | 免费版最多 50 个来源；Plus 最多 300 个 |
| **建议** | 以 PDF 或文本/Markdown 为主；内容清晰、结构分明更利于引用与问答 |

因此，若源文件很多，需要在本工具侧通过「合并」或「只输出合并件」控制最终文件数量与单文件大小，以便上传后不超过 NotebookLM 的来源数与单文件限制。

---

## 二、推荐知喂（ZhiWei）配置（面向 NotebookLM）

目标：在 `D:\ZWPDFTSEST` 下得到可直接用于 NotebookLM 的 PDF 和/或 Markdown，并集中到 `_LLM_UPLOAD` 便于一次性上传或拷贝。

### 2.1 路径与模式

- **源目录**：`Z:\Schneider\5_投标`（可多选，此处先单目录测试）
- **目标目录**：`D:\ZWPDFTSEST`
- **运行模式**：**合并与转换**（先转 Office→PDF/MD，再按类别或规则合并），便于控制最终文件数量与单文件大小；若仅需「只转不合并」可选用「仅转换」

### 2.2 输出格式与策略（关键）

- **输出格式**：建议 **PDF = 开、MD = 开**（NotebookLM 均支持；MD 便于后续若需再处理）
- **输出策略**：
  - 若希望**减少来源数、适合 NotebookLM 50 来源限制**：**合并 = 开、独立 = 关**，主要使用合并后的 PDF/MD；
  - 若希望**同时保留独立文件与合并文件**：**合并 = 开、独立 = 开**（来源数会增多，需自行控制上传数量）
- **合并与转换子功能**：按需选「合并与转换」下的「先转换再合并」或「仅合并（已有 PDF/MD）」（若已有现成 PDF/MD 只做合并）

### 2.3 LLM 交付中心（_LLM_UPLOAD）

- **启用 LLM 交付中心**：开（保证产出写入 `D:\ZWPDFTSEST\_LLM_UPLOAD`）
- **LLM 交付目录**：使用默认即可（即目标目录下的 `_LLM_UPLOAD`）
- **扁平化**：建议开（`llm_delivery_flatten = true`），所有文件在 `_LLM_UPLOAD` 根下，便于在 NotebookLM 中按文件名识别
- **合并去重**：建议开（已并入合并文档的源不再单独放入 _LLM_UPLOAD，减少重复与数量）
- **上传清单**：建议同时生成 `llm_upload_manifest.json` 与 `README_UPLOAD_LIST.txt`，便于核对来源与追溯

### 2.4 其他建议

- **溯源与短 ID**：若需在 NotebookLM 中通过「合并 PDF + 页码」或短 ID 反查源文件，可开启「溯源锚点与 merge map」（`enable_traceability_anchor_and_map`），便于后续用定位工具查回 `Z:\Schneider\5_投标` 中的原始文件。
- **单文件大小**：若合并后单文件接近或超过 200MB，需在「配置中心」调整合并策略（如按类别/按大小拆分），或关闭部分合并，使单文件 &lt; 200MB。
- **数量**：若最终 _LLM_UPLOAD 内文件数 &gt; 50（免费版），可先只上传部分（如只上传合并后的 PDF/MD），或升级 Plus 后上传更多。

---

## 三、测试步骤

### 3.1 前置检查

- [ ] 确认 `Z:\Schneider\5_投标` 可读且包含待转换的 Office 文件（如 .docx、.xlsx、.pptx）。
- [ ] 确认 `D:\ZWPDFTSEST` 已创建且可写；若不存在，由本工具或用户先创建。
- [ ] 确认本机已安装知喂（ZhiWei）运行环境（Python、依赖、可选 Office/WPS），可正常启动 GUI 或 CLI。

### 3.2 配置并执行一次完整运行

1. **打开知喂**（如 `python office_gui.py`），在「运行参数」或对应 Tab 中：
   - 源目录：添加 `Z:\Schneider\5_投标`
   - 目标目录：`D:\ZWPDFTSEST`
   - 运行模式：**合并与转换**（或按 2.1 选择）
   - 输出格式：PDF=开、MD=开；输出策略：合并=开、独立=关（或按 2.2 调整）
2. 在「成果文件」或「配置中心」中确认：
   - LLM 交付中心：开
   - 目标为默认（即 `D:\ZWPDFTSEST\_LLM_UPLOAD`）
   - 扁平化、合并去重、上传清单：按 2.3 设置
3. 执行一次完整「开始」运行，等待结束（含转换、合并、写入 _LLM_UPLOAD）。

### 3.3 产物检查

- [ ] 检查 `D:\ZWPDFTSEST\_LLM_UPLOAD` 是否存在且非空。
- [ ] 检查 `llm_upload_manifest.json`、`README_UPLOAD_LIST.txt` 是否生成，内容是否与目录内文件一致。
- [ ] 统计 _LLM_UPLOAD 内文件数量与总大小；抽查单文件大小是否均 &lt; 200MB。
- [ ] 若启用溯源：检查 merge map / 合并索引等是否生成，便于后续定位。

### 3.4 与 NotebookLM 的对接验证（可选）

- [ ] 将 `_LLM_UPLOAD` 内部分或全部文件上传至 Google Drive（可用知喂内置「上传到 Google Drive」），或直接在本机用 NotebookLM 的「上传 PDF/文档」添加本地文件。
- [ ] 在 NotebookLM 中新建 Notebook，添加上述来源，确认能正常识别、生成摘要并可基于来源回答。
- [ ] 若使用合并 PDF + 页码溯源：用知喂「定位工具」用页码或短 ID 反查，确认能定位回 `Z:\Schneider\5_投标` 下对应源文件。

### 3.5 记录与问题

- 记录：源文件数量、转换/合并耗时、_LLM_UPLOAD 文件数、单文件最大体积、任何报错或跳过文件。
- 若出现单文件 &gt; 200MB 或来源数 &gt; 50：记录当前配置（合并策略、是否独立输出等），并调整配置后重跑一轮，在本文档或 TEST_REPORT 中补充结果。

---

## 四、推荐配置小结（复制参考）

| 配置项 | 推荐值 | 说明 |
|--------|--------|------|
| 源目录 | `Z:\Schneider\5_投标` | 测试源 |
| 目标目录 | `D:\ZWPDFTSEST` | 测试目标 |
| 运行模式 | 合并与转换 | 便于控制文件数与大小 |
| 输出 PDF | 开 | NotebookLM 支持 |
| 输出 MD | 开 | NotebookLM 支持 |
| 合并输出 | 开 | 减少文件数 |
| 独立输出 | 关（或按需开） | 控制来源数 ≤50 时建议关 |
| LLM 交付中心 | 开 | 产出进 _LLM_UPLOAD |
| 扁平化 | 开 | 便于识别 |
| 合并去重 | 开 | 避免重复 |
| 溯源/merge map | 按需开 | 用于页码/短 ID 反查 |

---

## 五、文档与后续

- 本计划可放在 `docs/plans/` 下，与「NotebookLM 溯源」「LLM 交付」等现有文档并列。
- 执行完一轮后，建议在 `docs/test-reports/` 或交接文档中记录：执行时间、配置快照、_LLM_UPLOAD 统计结果、NotebookLM 验证结论（若已做），以及后续可优化点（如合并粒度、单文件上限策略等）。

---

## 六、AI 自主执行与修复决策

本节约定：AI（Cursor Agent）如何无交互跑完 E2E、如何判定通过/失败、失败时如何按决策表修复并重跑直至通过或需人工介入。详见 [docs/plans/NotebookLM_E2E_计划_第一性审阅与修改版.md](NotebookLM_E2E_计划_第一性审阅与修改版.md)。  
各程序/脚本的职责分工见 [docs/dev/程序与脚本清单.md](../dev/程序与脚本清单.md)（包括被测主程序 `OfficeConverter` 与 E2E 驱动脚本 `scripts/run_notebooklm_e2e.py` 的关系）。

### 6.1 AI 执行策略

执行 NotebookLM E2E 时**须循环**：在 **2026** 目录下运行：

```text
python scripts/run_notebooklm_e2e.py
```

可选参数：`--config ...`、`--output ...`、`--repair-prompt ...`、`--source ...`、`--target ...`；环境变量 `ZW_E2E_SOURCE`、`ZW_E2E_TARGET` 可覆盖源/目标路径。

- 若**退出码为 0**：通过，结束或写 TEST_REPORT。
- 若**退出码非 0**：根据 `docs/test-reports/notebooklm_e2e_repair_prompt.txt` 与其中 `log_path` 指向的 log、按本节「结果→动作表」修复配置或代码 → **再次运行同一命令**。
- 直至**退出码 0**，或遇到**需人工介入**的 `error_category`（见下表），或**同一 category 连续 2 次修复后仍失败**。**不得在未达到上述结束条件时结束。**

### 6.2 可执行规格

- **工作目录**：以 **2026** 为当前工作目录。
- **config**：默认 `configs/scenarios/notebooklm/config.notebooklm_test.json`（相对 2026）；若不存在则由脚本按第四节推荐配置生成（含 `auto_open_output_dir: false`；大目录时保留 enable_checkpoint、checkpoint_auto_resume、office_restart_every_n_files、sandbox_min_free_gb）。
- **通过（退出码 0）**：无异常；log 无 Traceback/关键错误；`<target>/_LLM_UPLOAD` 存在且非空或存在 manifest；`llm_upload_manifest.json` 存在且合法。
- **失败（退出码 1）**：异常或 log 含 Traceback/关键错误；或 _LLM_UPLOAD 应为非空但为空。
- **成功但需优化（退出码 2）**：满足通过条件但 _LLM_UPLOAD 文件数>50 或单文件>200MB；视为通过可结束，或按决策表调整后再跑。

### 6.3 结果→动作表（与脚本 error_category 一致）

| error_category        | AI 动作 |
|-----------------------|--------|
| path_not_found        | 创建缺失目录或修正 config 中 source/target，再跑 E2E |
| permission_denied     | 检查目标目录权限或换可写路径；若无法解决则停止并提示人工介入 |
| office_not_installed  | 检查 Office/WPS 安装与 COM；若无法解决则停止并提示人工介入 |
| office_com            | 检查 Office/WPS/COM 或切 engine，再跑 |
| config                | 检查 config 键与值，对照文档修正，再跑 |
| llm_hub_empty         | 检查 enable_llm_delivery_hub、run_mode、output_enable_* 及 corpus_manifest 调用，修正后再跑 |
| file_too_large_or_many| 记录到 TEST_REPORT，调整 max_merge_size_mb 或合并策略后再跑，或标记需优化后结束 |

当同一 category 连续 2 次修复后仍失败，Agent 应停止并提示人工介入。

### 6.4 修复提示文件

失败或退出码 2 时，脚本会写入 **`docs/test-reports/notebooklm_e2e_repair_prompt.txt`**（或 `--repair-prompt` 指定路径），内容含 `log_path`、`error_category`、`result_json` 路径及本节决策表引用。Agent 读该文件与 log 后按表执行修复并重跑。

### 6.5 长时间运行与「所有情况都测到」

- 单次 E2E 可能运行很久（目录内文档多时）；脚本不设总超时。转换器支持 checkpoint 断点续跑：若运行中途中断，用**同一 config** 再跑同一命令会只处理未完成文件。
- 单次运行只会得到一个结果（0/1/2）。要覆盖多种结果与 error_category，需多轮运行（例如全量跑→得 0 或 2；用错误路径再跑→得 1 path_not_found；或从 2 调配置再跑→得 0）。详见审阅文档「长时间运行与所有情况都测到」小节。
