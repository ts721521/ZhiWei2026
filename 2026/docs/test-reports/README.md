# 测试报告索引

本目录存放自动化测试的分报告与全量汇总。

## 全量回归命令

```bash
python -m unittest discover -s tests -p "test_*.py" -v
```

## 最新全量结果

- 执行日期：2026-02-24
- 结果：`Ran 71 tests ... OK`
- 汇总报告：[TEST_REPORT_SUMMARY.md](TEST_REPORT_SUMMARY.md)

## 自动化用例矩阵（当前仓库）

| 测试文件 | 用例数 |
|------|------|
| `test_converter_resume.py` | 1 |
| `test_default_config_schema.py` | 1 |
| `test_failed_file_trace_log.py` | 2 |
| `test_gui_config_logic_mixin.py` | 2 |
| `test_gui_locator_mixin_globals.py` | 1 |
| `test_gui_misc_ui_mixin.py` | 1 |
| `test_gui_profile_mixin.py` | 1 |
| `test_gui_run_mode_state_behavior.py` | 2 |
| `test_gui_run_tab_mixin_globals.py` | 1 |
| `test_gui_task_mode.py` | 10 |
| `test_gui_ui_shell_mixin_globals.py` | 3 |
| `test_locator.py` | 3 |
| `test_merge_convert_pipeline.py` | 4 |
| `test_nonfatal_ui_error_reporting.py` | 2 |
| `test_office_gui_entrypoint.py` | 2 |
| `test_output_controls.py` | 5 |
| `test_pypdf.py` | 2 |
| `test_pypdf_dir.py` | 1 |
| `test_pypdf_ver.py` | 2 |
| `test_pypdf_ver2.py` | 1 |
| `test_retry_and_scan_skip.py` | 2 |
| `test_task_binding_summary.py` | 3 |
| `test_task_detail_render.py` | 1 |
| `test_task_list_filter_sort.py` | 3 |
| `test_task_manager.py` | 13 |
| `test_task_selection_tab_preserve.py` | 1 |
| `test_ui_translation_coverage.py` | 1 |
| **总计** | **71** |

## 分报告

- [TEST_REPORT_GUI_TASK_MODE.md](TEST_REPORT_GUI_TASK_MODE.md)
- [TEST_REPORT_TASK_MANAGER.md](TEST_REPORT_TASK_MANAGER.md)
- [TEST_REPORT_CONVERTER_RESUME.md](TEST_REPORT_CONVERTER_RESUME.md)
- [TEST_REPORT_MERGE_CONVERT_PIPELINE.md](TEST_REPORT_MERGE_CONVERT_PIPELINE.md)
- [TEST_REPORT_OUTPUT_CONTROLS.md](TEST_REPORT_OUTPUT_CONTROLS.md)

说明：旧分报告主要覆盖核心模块；新增测试（如任务筛选、非致命 UI 错误上报、默认配置 schema 等）以 `TEST_REPORT_SUMMARY.md` 与测试文件为准。

