# office_converter 拆分交接文档（2026-02-24）

## 1. 建议的拆分停止标准

建议采用“双阈值”来决定是否停止：

- 硬阈值：`office_converter.py <= 2000` 行。
- 质量阈值：核心业务函数（除 `__init__` 外）单函数尽量 `<= 80` 行；超过 `120` 行必须继续拆。

当前状态（本机最新）：

- `office_converter.py`：约 `2897` 行。
- 说明：还未达到硬阈值，需要继续拆分。

## 2. 本轮前后进度摘要

本轮及上一轮已完成的关键拆分（均已接回 `office_converter.py` 委托）：

- `converter/process_single.py`
- `converter/batch_sequential.py`
- `converter/batch_parallel.py`
- `converter/run_workflow.py`
- `converter/update_package_export.py`
- `converter/excel_json_export.py`
- `converter/corpus_manifest.py`
- `converter/collect_index.py`
- `converter/merge_pdfs.py`

并已配套新增模块测试（委托测试 + 核心行为测试），全量回归通过。

## 3. 当前剩余大函数（优先继续拆）

按体量排序：

1. `convert_logic_in_thread`：149 行（约在 `office_converter.py:1167`）
2. `_render_html_to_markdown`：111 行（约在 `office_converter.py:1549`）
3. `_get_merge_tasks`：109 行（约在 `office_converter.py:2097`）
4. `_export_pdf_markdown`：108 行（约在 `office_converter.py:2729`）
5. `__init__`：100 行（约在 `office_converter.py:316`）
6. `_convert_cab_to_markdown`：80 行（约在 `office_converter.py:1661`）
7. `_create_index_doc_and_convert`：76 行（约在 `office_converter.py:2018`）
8. `_write_markdown_quality_report`：72 行（约在 `office_converter.py:2838`）

## 4. 推荐继续拆分顺序

建议按“低耦合优先、风险可控优先”执行：

1. `_render_html_to_markdown` -> `converter/markdown_render.py`
2. `_get_merge_tasks` -> `converter/merge_tasks.py`
3. `_write_markdown_quality_report` -> `converter/markdown_quality_report.py`
4. `_export_pdf_markdown` -> `converter/pdf_markdown_export.py`
5. `convert_logic_in_thread` -> `converter/convert_thread.py`
6. `_convert_cab_to_markdown` + `_create_index_doc_and_convert`（视耦合拆到 `converter/cab_convert.py` / `converter/merge_index_doc.py`）

## 5. 续拆统一模式（必须遵守）

每个函数都按下面流程：

1. 新建模块函数（参数注入依赖，不直接用全局）。
2. `office_converter.py` 仅保留薄委托方法。
3. 新增测试：
   - 模块核心行为测试（至少 1 个）。
   - 委托测试（monkeypatch `office_converter` 中 `*_impl`）。
4. 先跑定向测试，再跑全量测试。

## 6. 回归命令与基线

- 定向测试：按本次变更相关测试模块执行。
- 全量基线：

```bash
python -m unittest discover -s tests -p "test_*.py" -v
```

当前已验证基线：

- `Ran 193 tests ... OK`

若续拆后低于 193，不一定是失败（可能新增/删除测试），但必须保证全绿。

## 7. 已知注意事项（避免踩坑）

1. 编码问题：新增文件务必 `UTF-8`。曾出现过因错误编码导致 `UnicodeDecodeError` 导入失败。
2. 不要改行为：当前目标是“结构拆分”，不是业务改写。
3. 保留委托命名：延续 `xxx_impl` 导入别名模式，便于写委托测试。
4. 不要清理用户既有脏改动：仓库当前有较多未提交改动，避免误回滚。

## 8. 验收定义（交给下一位 AI）

满足以下条件即可认为“拆分一期完成”：

- `office_converter.py <= 2000` 行。
- 除 `__init__` 外不存在 >120 行函数。
- 全量测试绿（`unittest discover`）。
- 每个新拆模块都有对应测试文件。

## 9. 建议交付物

继续拆分时，每完成 1~2 个大函数，更新：

- 本文档（进度与剩余函数列表）。
- `docs/test-reports/TEST_REPORT_SUMMARY.md`（记录最新测试结果）。

