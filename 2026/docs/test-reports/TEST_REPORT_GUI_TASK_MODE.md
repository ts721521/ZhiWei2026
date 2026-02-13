# GUI 任务模式 / 传统模式 — 测试报告

## 功能模块说明

- **对应设计文档**：`docs/plans/2026-02-12-task-system-overview.md`（5.1～5.3、7.1、7.2）
- **实现位置**：`office_gui.py`（Header 模式开关、任务管理 Tab、新建任务向导、开始按钮分支、任务 Tab 显隐）

## 测试文件

- **路径**：`tests/test_gui_task_mode.py`
- **运行**：`python -m unittest tests.test_gui_task_mode -v`

## 测试用例清单

| 序号 | 测试用例 | 说明 | 结果 |
|------|----------|------|------|
| 1 | TestAppModeConfig.test_app_mode_roundtrip_classic | app_mode=classic 配置回合 | ✅ ok |
| 2 | TestAppModeConfig.test_app_mode_roundtrip_task | app_mode=task 配置回合 | ✅ ok |
| 3 | TestGuiTaskTabVisibility.test_classic_mode_hides_task_tab | 传统模式下任务 Tab 被隐藏 | ✅ ok |
| 4 | TestGuiTaskTabVisibility.test_switch_to_task_mode_shows_task_tab | 切换到任务模式后任务 Tab 显示 | ⏭ skipped（需 ttkbootstrap） |
| 5 | TestGuiTaskTabVisibility.test_task_mode_config_shows_task_tab | 配置为 task 时启动后任务 Tab 显示 | ⏭ skipped（需 ttkbootstrap） |
| 6 | TestWizardExists.test_open_task_wizard_method_exists | 新建任务向导方法存在 | ✅ ok |
| 7 | TestWizardExists.test_start_button_branches_on_app_mode | 开始按钮按 app_mode 分支（6.3/7.1） | ✅ ok |
| 8 | TestWizardExists.test_update_task_tab_for_app_mode_exists | 任务 Tab 显隐更新方法存在 | ✅ ok |
| 9 | TestWizardExists.test_wizard_step_keys_defined | 向导 4 步 tr key 存在（5.3） | ✅ ok |
| 10 | TestWizardStepLabelsOnWindows.test_wizard_four_step_labels_return_strings | Windows 下 4 步标签可解析 | ✅ ok |

**统计**：共 10 个用例，通过 8，跳过 2（依赖 ttkbootstrap 的 Tab 恢复显示）。

## 验证命令与结果

```bash
python -m unittest tests.test_gui_task_mode -v
```

- **执行日期**：2026-02-13（报告生成日）
- **结果**：Ran 10 tests，OK (skipped=2)
- **退出码**：0

## 结论

- 文档 5.1（Header + app_mode 持久化）、5.2（任务 Tab）、5.3（向导 4 步）、7.1（传统模式隐藏 Tab、开始分支）均有实现并有对应测试。
- 跳过项：在未安装 ttkbootstrap 的环境下，任务模式“恢复显示”任务 Tab 的两项用例不执行；安装 ttkbootstrap 后应全部通过。
