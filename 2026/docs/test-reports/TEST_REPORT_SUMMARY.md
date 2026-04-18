# 测试报告总览

## 执行命令

```bash
# 项目根目录（2026/）
python3 -m unittest discover -s tests -p "test_*.py" -v
```

## 当前状态（v5.21.0 · 2026-04-18）

最近一次按需回归（Collect `collect_copy_layout` 与配置链路）：

| 套件 | 结果 |
|------|------|
| `tests.test_converter_collect_index_module` | OK |
| `tests.test_converter_config_validation_module` | OK |
| `tests.test_converter_config_load_module` | OK |
| `tests.test_default_config_schema` | OK |
| `tests.test_converter_constants_module` | OK |
| `tests.test_converter_config_defaults_module` | OK |
| **合计** | **17 tests, OK** |

命令（在 `2026/` 目录下，使用 `python3`）：

```bash
python3 -m unittest \
  tests.test_converter_collect_index_module \
  tests.test_converter_config_validation_module \
  tests.test_converter_config_load_module \
  tests.test_default_config_schema \
  tests.test_converter_constants_module \
  tests.test_converter_config_defaults_module \
  -v
```

**环境说明**：当前部分 CI/容器镜像若未安装 `python3-tkinter` 或未装 `pypdf`，全量 `unittest discover` 可能出现与本轮功能无关的导入类 ERROR；完整桌面环境可再跑 `python3 -m unittest discover -s tests -p "test_*.py" -v`。

历史按需回归（v5.20.0 后 UI 相关，供对照）：

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
python3 -m unittest discover -s tests -p "test_*.py" -v
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
| 2026-04-18 | v5.21.0 | 按需 17（collect 复制布局 + 配置） | `collect_copy_layout` |
| 2026-04-18 | v5.20.0 | 按需 24（核心任务/向导/i18n） | classic 模式移除后回归 |

## 历史遗留事项

- v5.19.1 时期 `test_classic_mode_hides_task_tab` 曾断言失败 — v5.20.0 起 classic 模式整体移除，相关测试已删除，问题不再存在。
- v5.19.1 时期 4 个 pypdf 相关测试因未装 pypdf 报错 — 改用 `unittest.skipIf` + 显式依赖检查。
