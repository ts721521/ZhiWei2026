# 转换器断点续传 — 测试报告

## 功能模块说明

- **功能**：按文件列表续传时调用 process_single_file 回调、不触发全量扫描
- **实现位置**：`office_converter.py`（OfficeConverter 续传逻辑）

## 测试文件

- **路径**：`tests/test_converter_resume.py`
- **运行**：`python -m unittest tests.test_converter_resume -v`

## 测试用例清单

| 序号 | 测试用例 | 说明 | 结果 |
|------|----------|------|------|
| 1 | ConverterResumeTests.test_run_resume_file_list_emits_callbacks | 续传文件列表触发单文件处理回调、不触发扫描 | ✅ ok |

**统计**：共 1 个用例，通过。

## 验证命令与结果

```bash
python -m unittest tests.test_converter_resume -v
```

- **执行日期**：2026-02-13（报告生成日）
- **结果**：Ran 1 test，OK
- **退出码**：0

## 结论

- 续传模式下的文件列表驱动与回调行为已覆盖，测试通过。
