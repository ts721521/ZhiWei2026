# 全量测试验证总览

## 验证命令

```bash
# 项目根目录执行
python -m unittest discover -s tests -p "test_*.py" -v
```

## 最新验证结果

| 项目 | 结果 |
|------|------|
| **执行日期** | 2026-02-13 |
| **总用例数** | 25 |
| **通过** | 23 |
| **跳过** | 2（GUI 任务模式：需 ttkbootstrap 的 2 个 Tab 显示用例） |
| **失败** | 0 |
| **退出码** | 0 |
| **耗时** | ~1.0s |

## 输出摘要

```
Ran 25 tests in 0.993s
OK (skipped=2)
```

## 各模块结果

| 模块 | 文件 | 用例数 | 通过 | 跳过 |
|------|------|--------|------|------|
| GUI 任务/传统模式 | test_gui_task_mode.py | 10 | 8 | 2 |
| 任务管理 | test_task_manager.py | 5 | 5 | 0 |
| 转换器断点续传 | test_converter_resume.py | 1 | 1 | 0 |
| 合并/转换流水线 | test_merge_convert_pipeline.py | 4 | 4 | 0 |
| 输出控制 | test_output_controls.py | 5 | 5 | 0 |

## 分报告链接

- [GUI 任务模式](TEST_REPORT_GUI_TASK_MODE.md)
- [任务管理](TEST_REPORT_TASK_MANAGER.md)
- [转换器断点续传](TEST_REPORT_CONVERTER_RESUME.md)
- [合并/转换流水线](TEST_REPORT_MERGE_CONVERT_PIPELINE.md)
- [输出控制](TEST_REPORT_OUTPUT_CONTROLS.md)
