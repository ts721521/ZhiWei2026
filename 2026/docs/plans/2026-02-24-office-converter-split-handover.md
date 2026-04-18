# office_converter 拆分交接文档（同步副本）

完整历史与早期章节见：[docs/archive/plans-2026-02-landed/2026-02-24-office-converter-split-handover.md](../archive/plans-2026-02-landed/2026-02-24-office-converter-split-handover.md)。

下文仅追加**门禁所需**的最新轮次摘要（与 `docs/test-reports/TEST_REPORT_SUMMARY.md` 保持一致）。

---

## 66. 2026-04-18 v5.21.0 Collect「复制 + 索引」复制布局选项

### 66.1 本轮改动（执行者：Cloud Agent）

- **配置键** `collect_copy_layout`：`preserve_tree`（默认，保持源目录相对结构）与 `flat`（全部落在目标根目录；同名 basename 依次重命名为 `name__1.ext`、`name__2.ext`…）。
- **核心逻辑**：[converter/collect_index.py](../../converter/collect_index.py) 在 `copy_and_index` 子模式下按选项重写目标路径；`index_only` 不受影响。
- **配置链路**：`default_config.py`、`config_defaults.py`、`config_validation.py`、GUI 运行参数 Tab 与任务向导（`gui_run_tab_mixin.py`、`gui_run_mode_state_mixin.py`、`gui_config_compose_mixin.py`、`gui_config_io_mixin.py`、`gui_task_workflow_mixin.py`）、`ui_translations.py`。
- **版本**：[office_converter.py](../../office_converter.py) `__version__` → `5.21.0`。

### 66.2 定向测试（本轮）

```bash
cd 2026
python3 -m unittest \
  tests.test_converter_collect_index_module \
  tests.test_converter_config_validation_module \
  tests.test_converter_config_load_module \
  tests.test_default_config_schema \
  tests.test_converter_constants_module \
  tests.test_converter_config_defaults_module \
  -v
```

结果：`Ran 17 tests in …` · **OK**（环境中若未安装 openpyxl，collect 测试会打印预期错误信息后仍通过）。

### 66.3 全量套件说明

在无图形界面（无 `tkinter`）或路径断言依赖 Windows 分隔符的环境中，`python3 -m unittest discover -s tests -p "test_*.py" -v` 可能出现与本轮无关的 ERROR/FAIL；CI 与本地完整验证请在标准桌面 Python 环境中执行。

### 66.4 当前状态

- Collect 复制布局可由配置与 GUI 控制；默认保持目录树，与旧行为一致。
