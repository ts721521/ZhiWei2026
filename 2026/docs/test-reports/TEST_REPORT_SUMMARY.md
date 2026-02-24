# 全量测试验证总览

## 验证命令

```bash
python -m unittest discover -s tests -p "test_*.py" -v
```

## 最新结果（2026-02-24）

| 项目 | 结果 |
|------|------|
| 总用例数 | 71 |
| 通过 | 71 |
| 跳过 | 0 |
| 失败 | 0 |
| 退出码 | 0 |
| 结论 | OK |

## 控制台摘要

```text
Ran 71 tests in 7.762s
OK
```

## 覆盖模块概览

- 任务系统：任务管理、绑定关系、列表筛选/排序、选中详情渲染、任务切换状态保持。
- GUI 运行态：Run Mode 状态切换、关键 Mixin 的依赖绑定与入口保障。
- 核心流水线：转换恢复、合并/转换路径、输出策略、失败重试与扫描跳过。
- 默认配置与 i18n：默认 schema 结构、翻译 key 覆盖一致性。
- 三方兼容：`pypdf` API 兼容性测试。

## 关键新增验证点（相对旧报告）

- `test_task_list_filter_sort.py`：任务列表关键字/状态过滤、排序、当前配置范围过滤。
- `test_default_config_schema.py`：`ui.task_current_config_only` 默认值与 schema 存在性。
- `test_nonfatal_ui_error_reporting.py`：UI 非致命错误记录窗口与日志输出行为。
- `test_gui_run_mode_state_behavior.py`：run mode 切换异常上报及无裸 `except` 约束。

## 分报告链接

- [TEST_REPORT_GUI_TASK_MODE.md](TEST_REPORT_GUI_TASK_MODE.md)
- [TEST_REPORT_TASK_MANAGER.md](TEST_REPORT_TASK_MANAGER.md)
- [TEST_REPORT_CONVERTER_RESUME.md](TEST_REPORT_CONVERTER_RESUME.md)
- [TEST_REPORT_MERGE_CONVERT_PIPELINE.md](TEST_REPORT_MERGE_CONVERT_PIPELINE.md)
- [TEST_REPORT_OUTPUT_CONTROLS.md](TEST_REPORT_OUTPUT_CONTROLS.md)

