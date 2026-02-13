# 输出控制（输出计划）— 测试报告

## 功能模块说明

- **功能**：根据 run_mode 与 output_* 配置计算 need_final_pdf / need_markdown 等输出计划
- **实现位置**：`office_converter.py`（OfficeConverter.compute_convert_output_plan）

## 测试文件

- **路径**：`tests/test_output_controls.py`
- **运行**：`python -m unittest tests.test_output_controls -v`

## 测试用例清单

| 序号 | 测试用例 | 说明 | 结果 |
|------|----------|------|------|
| 1 | OutputPlanTests.test_convert_only_pdf_independent | convert_only + 仅 PDF 独立 | ✅ ok |
| 2 | OutputPlanTests.test_convert_only_md_independent | convert_only + 仅 MD 独立 | ✅ ok |
| 3 | OutputPlanTests.test_convert_then_merge_pdf_merged_only | convert_then_merge + 仅 PDF 合并 | ✅ ok |
| 4 | OutputPlanTests.test_convert_then_merge_md_merged_only | convert_then_merge + 仅 MD 合并 | ✅ ok |
| 5 | OutputPlanTests.test_all_formats_disabled | 全部格式关闭 | ✅ ok |

**统计**：共 5 个用例，全部通过。

## 验证命令与结果

```bash
python -m unittest tests.test_output_controls -v
```

- **执行日期**：2026-02-13（报告生成日）
- **结果**：Ran 5 tests，OK
- **退出码**：0

## 结论

- compute_convert_output_plan 在多种 run_mode 与 output 组合下行为正确，当前全部通过。
