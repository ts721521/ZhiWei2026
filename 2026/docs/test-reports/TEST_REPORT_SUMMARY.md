# 测试报告总览

## 执行命令

```bash
# 项目根目录
python -m unittest discover -s tests -p "test_*.py" -v
```

## 当前状态（v5.20.0 · 2026-04-18）

最近一次按需回归（与 v5.20.0 后的 UI 改动相关）：

| 套件 | 结果 |
|------|------|
| `tests.test_task_list_filter_sort` | OK |
| `tests.test_task_workflow_exception_narrowing` | OK |
| `tests.test_ui_translation_coverage` | OK |
| `tests.test_gui_task_mode` | OK |
| `tests.test_task_binding_summary` | OK |
| `tests.test_converter_incremental_filters_module` | OK |
| **合计** | **24 tests, OK in 3.4s** |

完整套件（394+ 用例）按需运行：

```bash
python -m unittest discover -s tests -p "test_*.py" -v
```

## 各报告索引

| 报告 | 范围 |
|------|------|
| [TEST_REPORT_GUI_TASK_MODE.md](TEST_REPORT_GUI_TASK_MODE.md) | GUI 任务模式可见性、向导、按钮分支（v5.19 历史快照） |
| [TEST_REPORT_CONVERTER_RESUME.md](TEST_REPORT_CONVERTER_RESUME.md) | 转换断点续传 |
| [README.md](README.md) | 测试报告总目录 |

## 历史快照

| 日期 | 版本 | 套件 | 备注 |
|------|------|------|------|
| 2026-02-13 | v5.18 | 25 (23 PASS / 2 SKIP) | 早期最小套件 |
| 2026-02-24 | v5.19.1 | 71 | 全量回归 |
| 2026-02-26 | v5.19.1 | 394 | Google Drive 上传模块加入 |
| 2026-04-18 | v5.20.0 | 按需 24（核心任务/向导/i18n） | classic 模式移除后回归 |

## 历史遗留事项

- v5.19.1 时期 `test_classic_mode_hides_task_tab` 曾断言失败 — v5.20.0 起 classic 模式整体移除，相关测试已删除，问题不再存在。
- v5.19.1 时期 4 个 pypdf 相关测试因未装 pypdf 报错 — 改用 `unittest.skipIf` + 显式依赖检查。
