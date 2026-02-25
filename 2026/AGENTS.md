# 项目规范（AI 通用）

本文件面向所有在此仓库工作的 AI 助手（Cursor、Claude、Copilot 等），请务必遵守以下规范，以保持项目整洁并避免将无关文件提交到 GitHub。

---

## 一、命名规范

### 1.1 产品与品牌

| 用途 | 规范 | 示例 |
|------|------|------|
| 产品中文名 | 知喂 | 文档、界面文案 |
| 产品英文/代号 | ZhiWei | 可执行文件名、配置中的程序名、代码注释 |
| 可执行文件 | `ZhiWei_v<版本号>.exe` | 版本号来自 `office_converter.py` 的 `__version__` |
| 用户目录/文件夹 | 知喂 或 ZhiWei | 如 `%APPDATA%/知喂/`、网盘「知喂上传」 |

### 1.2 Python 模块与脚本

| 类型 | 规范 | 示例 |
|------|------|------|
| 根目录主模块 | **snake_case** | `office_converter.py`, `office_gui.py`, `task_manager.py`, `gdrive_upload.py`, `locate_source.py`, `cab_to_pdf.py` |
| `scripts/` 下脚本 | **snake_case** | `sync_docs_to_obsidian.py` |
| 临时/内部脚本 | 下划线前缀 + snake_case | `_update_docs_pack.py`（一般不提交或按项目约定） |
| **新增模块** | 一律使用 **snake_case**，禁止全大写或驼峰 | 新文件：`my_feature.py` ✓；避免：`MyFeature.py`、`MYFEATURE.py` |

### 1.3 测试与配置

| 类型 | 规范 | 示例 |
|------|------|------|
| 测试文件 | `test_<模块或功能>.py`，全部放在 `tests/` | `test_task_manager.py`, `test_converter_resume.py` |
| 配置示例 | `config.example.json`（入库） | 用户复制为 `config.json` 后修改 |
| 本地配置 | `config.json`、`config_profiles/`（不提交） | 见 `.gitignore` |

### 1.4 文档与目录

| 类型 | 规范 | 示例 |
|------|------|------|
| 文档目录 | 小写，多词用连字符或下划线 | `docs/plans/`, `docs/test-reports/` |
| 文档文件名 | 英文或拼音，可含数字与连字符 | `PRODUCT_REQUIREMENTS.md`, `AI_交接文档_下一阶段开发.md` |
| 打包产物目录 | 固定名，不提交 | `build/`, `dist/` |

---

## 二、项目结构

| 路径 | 用途 | 是否提交到 Git |
|------|------|-----------------|
| `office_gui.py`, `office_converter.py` 等（见命名规范） | 主程序与核心模块 | 是 |
| `config.example.json`, `requirements.txt` | 配置示例与依赖 | 是 |
| `README.md`, `CHANGELOG.md`, `AGENTS.md` | 项目说明与规范 | 是 |
| `docs/` | 开发文档、计划、测试报告 | 是 |
| `docs/plans/` | 设计/计划文档 | 是 |
| `docs/test-reports/` | 测试报告 | 是 |
| `tests/` | 自动化测试 | 是 |
| `scripts/` | 工具脚本（打包、同步等） | 是 |
| `.cursor/rules/` | Cursor 规则文件 | 是 |
| `logs/` | 运行日志 | **否** |
| `output/` | 转换/合并输出 | **否** |
| `build/`, `dist/` | 打包产物 | **否** |
| `config.json`, `config_profiles/` | 本地配置 | **否** |
| `tasks/` | 本地任务数据 | **否** |
| `.venv/`, `__pycache__/`, `.ruff_cache/` | 虚拟环境与缓存 | **否** |
| `.tmp_test_work/` | 临时测试工作目录 | **否** |
| `skills/` | 第三方 skill 仓库（若存在） | **否** |

---

## 三、文件放置规则

1. **与程序直接相关的代码与配置**
   - 放在项目根目录或 `scripts/`、`tests/` 中。
   - 仅提交：源代码、`config.example.json`、`requirements.txt`。

2. **与程序无关或临时性文档**
   - 一律放在 `docs/` 下相应子目录（如 `docs/plans/`、`docs/notes/` 等）。
   - 禁止在项目根目录随意新建与程序无关的 `.md`、`.txt`、`.docx` 等文档。

3. **新增文档的默认位置**
   - 开发/设计/计划类：`docs/plans/`
   - 测试报告：`docs/test-reports/`
   - 其他说明、笔记：`docs/` 或 `docs/notes/`（可自建子目录）。

4. **测试代码**
   - 所有 `test_*.py` 统一放在 `tests/` 目录，根目录不保留零散测试脚本。

---

## 四、禁止提交到 GitHub 的内容

- 虚拟环境：`.venv/`, `venv/`, `env/`
- 本地配置与任务数据：`config.json`, `config_profiles/`, `tasks/`
- 日志与输出：`logs/`, `output/`, `.tmp_test_work/`
- 打包与缓存：`build/`, `dist/`, `*.spec`, `__pycache__/`, `.ruff_cache/`, `.pytest_cache/`, `.mypy_cache/`
- 敏感与凭证：`client_secrets*.json`, `*gdrive_token*.json`, `credentials.json`, `.env`, `.env.*`
- 临时文件：`out.txt`, `*.tmp`, `*.temp`, `*.bak`, `~$*.docx` 等
- 第三方 clone 仓库：`skills/`（若存在）

提交前请确认未包含上述路径或文件类型。

---

## 五、代码与提交习惯

- 新增与程序无关的文档时，先放到 `docs/` 指定子目录，再编辑或引用。
- 不在根目录堆积临时文件；临时输出使用 `output/` 或系统临时目录。
- 不将个人笔记、草稿等与程序无关内容放在根目录或未在 `.gitignore` 中忽略的路径。

---

## 六、参考

- 仓库根目录 `.gitignore` 已包含本项目忽略规则，请勿提交其中列出的路径。
- 项目说明与使用方式见 `README.md`，开发交接见 `docs/` 下相关文档。

---

## Cursor Cloud specific instructions

### 项目概述

知喂 (ZhiWei) 是 Windows 桌面 Python 应用（Office 文档批量转 PDF），主代码位于 `2026/` 目录。

### 运行测试

```bash
cd /workspace/2026
python3 -m unittest discover -s tests -p "test_*.py" -v
```

- 总计约 193 个测试。在 Linux 上会有 ~5 个因 Windows 路径差异导致的 FAIL，属于预期行为（项目仅面向 Windows）。
- 需要 `python3-tk` 系统包，否则 GUI 相关测试会 ERROR。

### Lint

项目无正式 lint 配置。可用 `ruff check .` 做基础检查（已有大量历史告警，非本项目 CI 阻断项）。

### 启动 GUI

```bash
cd /workspace/2026
python3 office_gui.py
```

- GUI 基于 tkinter + ttkbootstrap，需要 X display（Cloud VM 已自带 `:1`）。
- 实际文档转换功能依赖 Windows COM（`win32com`），在 Linux 上不可用；但 GUI 可正常启动、配置和保存。

### 关键注意事项

- 代码入口在 `2026/` 子目录，不是仓库根目录。所有 Python 命令需从 `2026/` 运行。
- `requirements.txt` 中注释掉的可选依赖（`pypdf`、`openpyxl`、`python-docx`、`beautifulsoup4`）在测试中会被用到，建议安装。
- `config.json` 和 `config_profiles/` 在 `.gitignore` 中，不要提交。
