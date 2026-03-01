---
name: doc-sync-reminder
description: 修改约定代码路径时，必须同步更新计划/交接文档与测试汇总；与 AGENTS.md 强制文档同步规则一致。在编辑或提交前触发本 skill 可避免遗漏。
---

# 文档同步提醒（Skill）

当你在本仓库中**修改或新增**以下约定路径时，**必须**同步更新对应文档，否则视为未完成交付。

## 触发路径（请按本项目填写）

- 示例：`src/main.py`、`app/**`、`tests/**`
- 或：`office_converter.py`、`converter/**`、`gui/**`、`tests/**`

## 必须同步的文档（请按本项目填写）

- 计划/交接文档：`docs/plans/handover.md` — 当前状态与变更摘要
- 测试汇总：`docs/test-reports/TEST_REPORT_SUMMARY.md` — 执行命令、用例数、OK/FAILED、本轮变更摘要

若涉及**架构/目录迁移**，还必须更新 `AGENTS.md`。

## 流程

1. 完成代码改动后先执行**本项目的测试命令**（见 AGENTS.md 或下方）。
2. 将测试结果与变更摘要写入上述两个文档。
3. 若有自检脚本，提交前运行（如 `python scripts/check_doc_sync.py --staged`）。

## 本项目的测试命令（示例，请按实际填写）

```bash
python -m unittest discover -s tests -p "test_*.py" -v
```

详细约定以仓库根目录 **AGENTS.md** 为准。
