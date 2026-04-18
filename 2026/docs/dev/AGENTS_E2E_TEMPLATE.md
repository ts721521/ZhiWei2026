# AGENTS.md（通用协作规范 · 带自测 E2E 版本模板）

**用途**：复制到项目根目录并重命名为 `AGENTS.md`。  
**适用场景**：希望项目具备「程序自测、自读 log、自生成修复提示」能力的仓库。

> 若只需要最基本的“改代码要记文档 + 记录测试结果”，可改用基础版模板 `AGENTS_TEMPLATE.md`。

---

## 0. 接手必读（新 AI / 新开发者入仓后按顺序执行）

1. **读本文档**（AGENTS.md），了解规则与文档路径约定。
2. **读项目现状与下一步**（路径见「本项目的文档目录」）  
   - 交接/现状文档：项目概览、入口、已实现能力、当前状态、**下一步建议**。  
   - 测试汇总：最新全量回归结果（用例数、OK/FAILED）。  
3. **读 E2E 场景计划文档**（若已定义）：  
   - `docs/plans/<SCENARIO>_E2E_计划.md`（一个或多个），了解本项目为哪些关键业务路径定义了端到端自测场景。  
4. **开始工作前**（建议）：若你对本轮改动负责，在记录中注明**执行者/会话标识**（如 Agent 名或会话 ID + 日期），便于多协作者追溯。

---

## 1. 本项目的文档目录（请按实际填写）

**规则**：若仓库中已存在下列用途的目录或文件，直接使用；若不存在，默认在当前项目下创建（若创建失败或用户不想在当前项目创建，则请用户指定路径）。

| 用途       | 路径（示例）                               |
|------------|--------------------------------------------|
| 计划/交接  | `docs/plans/`                              |
| 测试汇总   | `docs/test-reports/TEST_REPORT_SUMMARY.md` |
| 交接总览   | `docs/HANDOVER.md`（可选，可并入首份计划文档） |
| 任务清单   | `docs/TASK_LIST.md`（可选）                |
| 归档       | `docs/archive/`（可选）                    |

**请在本项目中使用时填写实际路径**（沿用已有的、当前项目下新建的或用户指定的）：

- 计划/交接：`docs/plans/`
- 测试汇总：`docs/test-reports/TEST_REPORT_SUMMARY.md`
- 交接总览：`docs/HANDOVER.md`
- 任务清单：`docs/TASK_LIST.md`

---

## 2. 初始化清单（AI 读完 AGENTS 后若发现缺失应执行）

- [ ] **检测是否已有对应目录**  
  仅检查：`docs/plans/`、`docs/test-reports/`，或「本项目的文档目录」中已填写的路径。**若已存在，则直接沿用，不创建、不修改**。
- [ ] **创建目录与文件**（仅当上一步未找到时）  
  默认在**当前项目目录**下创建 `docs/plans/`、`docs/test-reports/` 及首份文档。若**创建失败**（如无写权限）或用户**明确不想在当前项目创建**，则请用户指定路径后再创建。
- [ ] **创建计划/交接首份文档**（仅当缺失时）  
  例如 `docs/plans/handover.md`，内容含：项目一句话说明、当前状态（可填「待初始化」）、**下一步建议**（可填「待补充」）。
- [ ] **创建测试汇总**（仅当缺失时）  
  例如 `docs/test-reports/TEST_REPORT_SUMMARY.md`，内容含：验收命令、最新结果（可填「待首次运行」或「无自动化测试」）、历史轮次占位。
- [ ] **（可选）交接总览或任务清单**  
  若需单独入口可创建 `docs/HANDOVER.md` 或 `docs/TASK_LIST.md`。
- [ ] **更新 AGENTS.md**  
  将「本项目的文档目录」中路径替换为**实际使用的路径**。

完成后在交接/计划文档中记一条：「初始化完成，执行者：<标识>，日期：<日期>」。

---

## 3. 基本原则

- 代码改动必须有可追溯记录（提交信息或文档中的变更摘要）。
- 文档是交付物的一部分，不是补充。
- **任何导致本文档所列路径或目录失效的变更**（如模块搬迁、目录重命名、导入规则变化），须同步更新 AGENTS.md。

---

## 4. 强制文档同步规则（必须遵守）

当**约定的代码路径**（见下）发生变更时，必须更新**计划/交接文档**（记录变更摘要）；若本轮执行了测试，必须更新**测试汇总**（记录命令、用例数、结果）。**若未填写下方「代码路径约定」**，则任意代码变更均视为需更新文档。

若变更**导致本文档所列路径或目录失效**，还必须更新 **AGENTS.md**。

**本项目的代码路径约定**（请按实际填写；不填则任意代码变更均触发文档更新）：

- 核心入口：`src/main.py`、`app/`
- 测试：`tests/`
- （可选）SUT 主要模块：`core/`、`service/` 等

---

## 5. 强制测试记录规则（必须遵守）

完成一轮开发后，若项目有约定测试命令，则执行并在**测试汇总**与**计划/交接文档**中记录：执行命令、用例总数、结果（OK/FAILED）、本轮变更摘要。**执行者/会话**（建议）：可注明本轮执行者或会话 ID + 日期。

**若项目暂时无自动化测试**：在测试汇总中注明「无」或「待补充」，并仍建议在计划/交接中记录本轮摘要。

**本项目的测试命令**（请按实际填写）：

```bash
python -m pytest tests/ -v
# 或
python -m unittest discover -s tests -p "test_*.py" -v
```

---

## 6. 通用 E2E 自举方法论（可选但推荐）

若本项目希望具备「程序自测、自读 log、自生成修复提示」能力，建议为**若干关键场景**按以下约定实现 E2E 自举：

### 6.1 角色与边界

- **SUT（System Under Test，被测系统）**
  - 定义：项目中“真正提供业务功能”的入口，例如 `App.run(config)`、`main_pipeline(config_path)` 等。
  - 要求：调用一次即可完成一轮完整业务（或一个清晰的业务场景）。

- **Config 工厂**
  - 定义：根据默认配置和场景覆盖项，产出一次可运行的配置文件/对象。
  - 要求：若配置文件不存在，E2E Runner 能够通过 Config 工厂自动生成一个最小可行版本，并写入磁盘。

- **E2E Runner 脚本**
  - 命名与位置约定：`scripts/run_<scenario>_e2e.py`。
  - 职责（仅三件事）：
    1. **准备配置**：若无配置则创建默认配置，之后应用场景推荐覆盖项（路径、禁用交互、超时策略等）；
    2. **调用 SUT**：导入项目内真实入口，调用 SUT 执行一轮（例如 `MyApp(config_path).run()`）；
    3. **检查结果与输出结构化信息**：读取 log 和约定输出位置，判断是否“通过 / 失败 / 成功但需优化”，生成统一结构的结果 JSON 与修复提示文件，并用退出码表达状态。

- **决策表（Result → Action Table）**
  - 定义：将 `error_category` 映射为“下一步动作”，例如：
    - `path_not_found` → 修复路径 → 再跑；
    - `config` → 修复配置 → 再跑；
    - `output_empty` → 检查 SUT 是否正确写出产物 → 再跑；
    - `sut_exception` → 检查异常信息，必要时标记“需人工介入”。
  - 要求：每个 E2E 场景的计划文档必须以表格形式列出该场景采用的 `error_category` 枚举及含义，并保证与 Runner 中输出的值**完全一致**。

- **文档锚点**
  - 每个 E2E 场景的计划文档 `docs/plans/<SCENARIO>_E2E_计划.md` 必须指定：
    - 工作目录（cwd）；
    - 执行命令（例如 `python scripts/run_<scenario>_e2e.py`）；
    - 默认配置路径、结果 JSON 路径、修复提示路径；
    - “结果→动作表”所在章节位置。

### 6.2 Runner 通用结构（伪代码）

任意场景的 `scripts/run_<scenario>_e2e.py` 推荐遵循以下结构（示意）：

```python
def ensure_config(config_path: str, scene_overrides: dict) -> bool:
    """
    若 config 文件不存在：
      - 调用项目内 default_config 工厂创建默认配置；
      - 应用 scene_overrides（源/目标路径、禁用 auto_open、禁用交互等）；
      - 写回磁盘。
    若存在：
      - 读取配置（JSON/YAML 等）；
      - 应用 scene_overrides；
      - 写回。
    失败返回 False。
    """

def run_sut_once(config_path: str, extra_args: dict) -> tuple[bool, str]:
    """
    导入并调用本项目的 SUT：
      - from <project_root> import <SUT>
      - sut = <SUT>(config_path, **extra_args)
      - sut.run()
    返回 (执行是否抛异常, log_path)。
    """

def analyze_result(outputs_root: str, log_path: str) -> dict:
    """
    按该场景的业务约定检查输出：
      - 检查关键输出目录/文件是否存在；
      - 统计关键指标（如文件数、大小、延迟、error 条数等）；
      - 判断是否 success / need_optimization / fail；
      - 推导 error_category。
    返回统一结构，例如：
      {
        "success": bool,
        "need_optimization": bool,
        "error_category": "config" | "path_not_found" | "output_empty" | "sut_exception" | "",
        "errors": [...],
        "message": "...",
        "...其他指标..."
      }
    """

def write_result_and_prompt(result_json_path: str, repair_prompt_path: str, result: dict) -> None:
    """
    - 将 result 写入 result_json_path；
    - 写入 repair_prompt_path，内容包括：
        - result_json: <绝对或相对路径>
        - log_path: <日志路径>
        - error_category: <枚举值>
        - 简要提示：引用 docs/plans/<SCENARIO>_E2E_计划.md 某节的结果→动作表，并给出再次运行的命令。
    """

def main() -> int:
    # 1. 解析命令行参数（config 路径、输出路径、源/目标目录、max-files 等）
    # 2. 通过环境变量或参数决定 source/target/cwd
    # 3. ensure_config(...)
    # 4. run_sut_once(...)
    # 5. analyze_result(...)
    # 6. write_result_and_prompt(...)
    # 7. 根据 success / need_optimization 决定退出码：
    #    - 0 = 通过
    #    - 1 = 失败
    #    - 2 = 成功但需优化（例如“输出过多”“单文件过大”等）
    ...
```

> **约束**：Runner **不得在脚本中实现替代的业务逻辑**，只能调用项目内真实 SUT 并对结果做分析与断言。

### 6.3 E2E 场景文档与执行协议

对于每个 E2E 场景 `<SCENARIO>`，项目应：

1. 创建 `docs/plans/<SCENARIO>_E2E_计划.md`，内容包括但不限于：
   - 场景目标与前置条件；
   - 推荐配置（可在表格中列出关键配置项）；
   - 测试步骤；
   - 「可执行规格」小节，明确：
     - cwd；
     - 执行命令：`python scripts/run_<scenario>_e2e.py`；
     - config 默认路径；
     - 结果 JSON 路径与修复提示路径；
     - 成功 / 失败 / 成功但需优化 的判定条件；
   - 「结果→动作表」小节，列出 `error_category` → AI/开发者应执行的修复动作。

2. 在本 AGENTS 文档中，为每个 E2E 场景增加一条执行规则（示例）：

```markdown
**<SCENARIO> E2E 场景**：执行该场景 E2E 测试时须**循环运行直至通过或需人工介入**。  
- 场景规格与决策表：见 `docs/plans/<SCENARIO>_E2E_计划.md` 第六节；  
- 执行命令（在项目根或文档指定 cwd 下）：`python scripts/run_<scenario>_e2e.py`；  
- 结果 JSON：`docs/test-reports/<scenario>_e2e_result.json`；  
- 修复提示文件：`docs/test-reports/<scenario>_e2e_repair_prompt.txt`。
```

3. AI 或开发者在执行 `<SCENARIO>` E2E 时，遵循以下循环：

- 运行 `python scripts/run_<scenario>_e2e.py`；  
- 若退出码为 0：视为通过，可在测试汇总与计划文档中记录结果；  
- 若退出码为 1 或 2：
  - 打开 `*_e2e_result.json` 与 `*_e2e_repair_prompt.txt`；
  - 按计划文档「结果→动作表」中的规则修复配置或代码；
  - 再次运行同一命令；  
- 循环上述过程，直到：
  - 退出码为 0；或
  - 连续多次同类致命错误且决策表要求“需人工介入”。

---

## 7. 禁止项

- 只改代码不记文档。
- 只跑测试不写回归结果到约定文档。
- 做了导致 AGENTS 所列路径失效的变更但不更新 AGENTS.md。
- 在项目根目录散落临时脚本或临时文档（应放入约定目录或 archive）。
- 在 E2E Runner 中实现替代业务逻辑而不是调用项目真实 SUT。

---

## 8. 执行口径

若规则与临时口头要求冲突，以「可追溯、可回放、可验收」为优先。  
任何例外必须在交接文档中注明：例外原因、影响范围、后续补齐计划。

---

## 9. 可选：自动校验与 CI

若项目已具备或计划引入：文档同步校验脚本、pre-commit 或 CI 执行校验与全量测试，可在此记录约定。无则可不提供。

- 建议：在 CI 中为关键 `<SCENARIO>` 场景执行 `python scripts/run_<scenario>_e2e.py`，并根据退出码 0/1/2 决定工作流通过/失败或标记“需人工介入”。

---

**模板版本**：E2E-extended v1.0 · 基于通用 AGENTS 模板扩展的自测方法论版本。

