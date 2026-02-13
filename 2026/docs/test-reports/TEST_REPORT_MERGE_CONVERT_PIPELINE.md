# 合并/转换流水线 — 测试报告

## 功能模块说明

- **功能**：merge_only 模式调用 PDF/MD 合并、pdf_to_md 子模式与合并逻辑
- **实现位置**：`office_converter.py`（OfficeConverter 合并与转换流水线）

## 测试文件

- **路径**：`tests/test_merge_convert_pipeline.py`
- **运行**：`python -m unittest tests.test_merge_convert_pipeline -v`

## 测试用例清单

| 序号 | 测试用例 | 说明 | 结果 |
|------|----------|------|------|
| 1 | MergeConvertPipelineTests.test_merge_only_calls_pdf_and_md_merge | merge_only 调用 PDF 与 MD 合并 | ✅ ok |
| 2 | MergeConvertPipelineTests.test_merge_only_missing_md_can_abort | merge_only 缺少 MD 时可中止 | ✅ ok |
| 3 | MergeConvertPipelineTests.test_pdf_to_md_with_merged_md_merges_generated_md | pdf_to_md 子模式与已合并 MD 时合并生成 MD | ✅ ok |
| 4 | MergeConvertPipelineTests.test_pdf_to_md_without_merged_only_converts | 无 merged 时仅转换 | ✅ ok |

**统计**：共 4 个用例，全部通过。

## 验证命令与结果

```bash
python -m unittest tests.test_merge_convert_pipeline -v
```

- **执行日期**：2026-02-13（报告生成日）
- **结果**：Ran 4 tests，OK
- **退出码**：0

## 结论

- 合并模式与 pdf_to_md 子模式行为已覆盖，当前全部通过。
