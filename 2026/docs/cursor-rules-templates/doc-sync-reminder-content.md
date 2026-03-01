# 文档同步约定（与 AGENTS.md 强制文档同步规则一致）

**用途**：本文件为**与编辑器/IDE 无关**的约定内容。可单独复用到任何项目；在 Cursor 中可转为 Rule（.mdc），在 Codex 中可转为 Skill（SKILL.md），或仅依赖 AGENTS.md 时由 AI 按 §3/§4 执行。

---

当**修改或新增**以下任意路径时，**必须**同步更新对应文档，否则视为未完成交付：

**本项目的触发路径**（请按实际填写，与 AGENTS.md 中「代码路径约定」一致）：
- 示例：`src/main.py`、`app/**`、`tests/**`
- 或：`office_converter.py`、`converter/**`、`gui/**`、`tests/**`

**本项目的必须同步文档**（请按实际填写，与 AGENTS.md 中「本项目的文档目录」一致）：
- 计划/交接文档：`docs/plans/handover.md` — 当前状态与变更摘要
- 测试汇总：`docs/test-reports/TEST_REPORT_SUMMARY.md` — 最新全量回归结果（执行命令、用例数、OK/FAILED、本轮变更摘要）

**若涉及架构/目录迁移**（模块搬迁、导入规则变化、兼容层新增或移除）**还必须更新：**
- `AGENTS.md`

完成代码改动后，须先执行**本项目的测试命令**（见下），再将结果与变更摘要写入上述文档。若项目有文档同步自检脚本，提交前运行自检。

**本项目的测试命令**（请按实际填写）：
```bash
python -m unittest discover -s tests -p "test_*.py" -v
# 或 python -m pytest tests/ -v
```

**本项目的自检命令**（可选，无则删除本段）：
```bash
python scripts/check_doc_sync.py --staged
```
