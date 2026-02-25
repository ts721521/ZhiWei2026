# 文档同步门禁落地说明

## 1. 目标

保证“代码变更 -> 文档记录 -> 测试回归”形成强制闭环，避免漏记。

## 2. 本地启用

1. 安装 hook：

```bash
python scripts/install_git_hook.py
```

2. 提交前会自动执行：

```bash
python scripts/check_doc_sync.py --staged
```

## 3. CI 启用

已提供：`.github/workflows/quality-gate.yml`

门禁步骤：

1. 计算本次变更文件
2. 运行文档同步校验 `scripts/check_doc_sync.py`
3. 运行全量单测 `python -m unittest discover -s tests -p "test_*.py" -v`

## 4. 失败处理

- 若提示缺失 `docs/plans/2026-02-24-office-converter-split-handover.md` 或 `docs/test-reports/TEST_REPORT_SUMMARY.md`：补记本轮改动与测试结果。
- 若提示缺失 `AGENTS.md`：说明你做了架构/目录迁移，但未同步规范；请补充规则说明。

## 5. 建议团队流程

- 开发分支每次提交都通过 pre-commit。
- PR 合并前必须通过 quality-gate。
- 若有规则例外，必须在交接文档记录“例外原因 + 影响范围 + 补齐计划”。
