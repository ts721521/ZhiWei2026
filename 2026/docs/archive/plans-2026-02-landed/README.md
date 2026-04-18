# 已落地的 2026-02 规划文档

本目录归档 2026-02 期间编写、当前已落地或被取代的规划文档。仅供历史参考，不再维护。

## 索引

| 文件 | 当时主题 | 落地状态（截至 v5.20.0） |
|------|----------|--------------------------|
| `2026-02-12-task-management-design.md` | 任务模式 + 经典模式并存的设计 | classic 模式已移除（v5.20.0），任务中心成为唯一入口 |
| `2026-02-12-task-system-overview.md` | 同上的概览版（与 management-design 重复） | 同上 |
| `2026-02-12-unified-output-controls-implementation.md` | 统一输出控件计划 | 已落地：PDF/MD 切换、合并/独立切换、Merge & Convert 子选项均在向导/配置页 |
| `任务管理功能规划.md` | 任务/非任务双域配置规划（中文） | 已演化为"每任务独立 profile"，详见 `config_profiles/task_<id>.json` |
| `2026-02-24-office-converter-split-plan.md` | `office_converter.py` 拆分规划 | 已落地：`converter/` 子模块化（ai_paths / batch_helpers / chromadb / merge / 等 90+ 文件），`office_converter.py` 由 7k+ 行降至 ~1.7k 行 |
| `2026-02-24-office-converter-split-handover.md` | 拆分交接细节 | 同上 |
| `任务管理功能规划_bcfc715c.plan.md` | 早期任务管理 plan 快照 | 已被实现取代 |

## 当前规划文档去向

- 仍在用的规划：`docs/plans/`（如 V6.0 评估、code-review 优化建议）。
- AI 交接：`docs/dev/AI_交接文档_下一阶段开发.md`。
- 变更日志：`CHANGELOG.md`。
