# 测试报告索引

本目录存放各功能模块的测试报告，每次完整回归后更新。

## 运行全部测试

```bash
# 项目根目录
python -m unittest discover -s tests -p "test_*.py" -v
```

## 报告列表

| 模块 | 测试文件 | 报告文档 | 用例数 |
|------|----------|----------|--------|
| GUI 任务/传统模式 | tests/test_gui_task_mode.py | [TEST_REPORT_GUI_TASK_MODE.md](TEST_REPORT_GUI_TASK_MODE.md) | 10 |
| 任务管理（TaskStore/checkpoint） | tests/test_task_manager.py | [TEST_REPORT_TASK_MANAGER.md](TEST_REPORT_TASK_MANAGER.md) | 5 |
| 转换器断点续传 | tests/test_converter_resume.py | [TEST_REPORT_CONVERTER_RESUME.md](TEST_REPORT_CONVERTER_RESUME.md) | 1 |
| 合并/转换流水线 | tests/test_merge_convert_pipeline.py | [TEST_REPORT_MERGE_CONVERT_PIPELINE.md](TEST_REPORT_MERGE_CONVERT_PIPELINE.md) | 4 |
| 输出控制（输出计划） | tests/test_output_controls.py | [TEST_REPORT_OUTPUT_CONTROLS.md](TEST_REPORT_OUTPUT_CONTROLS.md) | 5 |

**全量**：共 25 个用例，`python -m unittest discover -s tests -p "test_*.py" -v` 运行结果：OK（skipped=2）。

最后全量运行时间与结果见各报告中的「验证命令与结果」。最新一次全量验证见 [TEST_REPORT_SUMMARY.md](TEST_REPORT_SUMMARY.md)。
