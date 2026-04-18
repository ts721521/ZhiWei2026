# Google Drive 上传功能 — 实现规划

基于既有规划（AI_HANDOVER §11、PRODUCT_REQUIREMENTS §8.5、TASK_LIST Phase 7）展开为可执行开发任务。目标：单用户本机将 `_LLM_UPLOAD` 一键上传到自己的 Google Drive，供 NotebookLM 等使用。

---

## 1. 范围与约束

- **场景**：用户在本机运行知喂，生成 `_LLM_UPLOAD` 后，通过 GUI 手动点击「上传到 Google Drive」，将当前 LLM 上传目录内容同步到 Drive。
- **认证**：桌面 OAuth 2.0（用户本人账号），不引入 Service Account；首次打开浏览器授权，token 本地缓存，后续静默刷新。
- **约束**：不改变现有转换/合并/LLM 归集流程；上传为独立动作，可选功能，未配置或未启用时不影响主流程。

---

## 2. 配置项设计

| 键 | 类型 | 默认值 | 说明 |
|----|------|--------|------|
| `enable_gdrive_upload` | bool | `false` | 是否启用 Google Drive 上传能力（显示/启用上传按钮与配置区） |
| `gdrive_client_secrets_path` | string | `""` | 本机 OAuth 客户端密钥 JSON 路径（Google Cloud 控制台下载） |
| `gdrive_folder_id` | string | `""` | 可选。上传到此文件夹下；为空则在用户 Drive 根下创建 `知喂上传/Run_YYYYMMDD_HHMMSS` |
| `gdrive_token_path` | string | `""` | 可选。token 缓存路径；为空则用 `<script_dir>/.gdrive_token.json` |

**说明**：
- 使用前用户需在 Google Cloud 创建 OAuth 2.0 客户端（桌面应用），下载 `client_secrets.json` 并填入 `gdrive_client_secrets_path`。
- `gdrive_folder_id` 为空时，在 Drive 根目录创建 `知喂上传/Run_YYYYMMDD_HHMMSS` 并上传到该子文件夹。

---

## 2.1 敏感数据与安全（必读）

`client_secrets.json` 与 OAuth **token** 属于敏感数据，必须做到：

| 要求 | 说明 |
|------|------|
| **禁止提交到 Git** | 已将 `**/client_secrets*.json`、`**/.gdrive_token*.json`、`**/gdrive_token*.json`、`**/credentials.json` 加入 `.gitignore`，切勿强制 add。 |
| **存放位置** | 建议将 `client_secrets.json` 放在**项目外**（如用户目录、仅本机可读的目录），config 里只存**路径**，不存文件内容。 |
| **Token 默认路径** | token 缓存建议默认放在用户目录（如 `%APPDATA%/知喂/.gdrive_token.json` 或 `~/.config/知喂/gdrive_token.json`），避免放在仓库目录下。 |
| **不记录敏感内容** | 日志、错误提示、调试输出中**不得**打印 client_secrets 或 token 的完整内容；仅可提示“未配置密钥路径”或“token 已过期”等。 |
| **配置示例** | `configs/templates/config.example.json` 中**不要**出现真实路径或占位路径指向敏感文件；文档中说明“请将密钥文件放在安全位置并填写路径”。 |

实现时：读取密钥/ token 仅用于 OAuth 与 API 调用，用毕不缓存到内存以外；若需写入 token 文件，权限设为仅当前用户可读（如 Windows 仅当前用户、Linux `chmod 600`）。

---

## 3. 依赖

- **新增 Python 包**（建议加入 `requirements.txt`）：
  - `google-auth-oauthlib`：OAuth 桌面流程、token 刷新
  - `google-api-python-client`：Drive API 调用（上传、创建文件夹）
- **可选**：`google-auth-httplib2`（若 google-api-python-client 未自带适配）

```text
# Google Drive 上传（可选）
google-auth-oauthlib>=1.0.0
google-api-python-client>=2.0.0
```

---

## 4. 模块与职责划分

### 4.1 新建模块 `gdrive_upload.py`（推荐）

- **职责**：OAuth 流程、token 读写、调用 Drive API 上传、创建远程文件夹、更新 manifest。
- **接口建议**：
  - `ensure_credentials(client_secrets_path, token_path) -> credentials`：若未授权则弹浏览器，否则从 token_path 加载并刷新。
  - `upload_llm_folder_to_drive(local_llm_root, credentials, parent_folder_id=None) -> dict`：  
    在 Drive 创建 `Run_YYYYMMDD_HHMMSS`（或使用 parent_folder_id），遍历 local_llm_root 下文件并上传；返回 `{ "folder_id": str, "uploaded_at": str, "file_count": int }`。
  - `update_manifest_gdrive_section(manifest_path, gdrive_info)`：在 `llm_upload_manifest.json` 中追加或更新 `gdrive` 区段（remote_folder_id、uploaded_at 等）。
- **错误处理**：网络异常、配额、无权限等应抛明确异常或返回错误结构，由 GUI 捕获并提示。

### 4.2 在 `office_converter.py` 中的改动

- **默认配置**：在 `default_config` 与 `cfg.setdefault` 中增加上述 4 个键。
- **不在此处调用上传**：上传仅由 GUI 按钮触发，converter 不自动调 Drive API，保持“核心只做本地生成”的边界。

### 4.3 在 `office_gui.py` 中的改动

- **成果文件 Tab**：在 LLM 上传相关区域下方新增一块「Google Drive 上传」：
  - 勾选框：`启用 Google Drive 上传`（绑定 `enable_gdrive_upload`）。
  - 输入框：`客户端密钥 JSON 路径`（`gdrive_client_secrets_path`），可选「浏览」选文件。
  - 输入框：`上传目标文件夹 ID`（`gdrive_folder_id`），留空则自动创建 `知喂上传/Run_*`。
  - 按钮：`上传 _LLM_UPLOAD 到 Google Drive`。  
    点击时：读取当前 config 中的 `llm_delivery_root`（或当前运行的目标目录下的 `_LLM_UPLOAD`），若目录存在则调用 `gdrive_upload.ensure_credentials` + `upload_llm_folder_to_drive`，进度/结果在日志区或弹窗提示；若已生成 `llm_upload_manifest.json` 则调用 `update_manifest_gdrive_section`。
- **配置加载/保存**：在 `_load_config_to_ui` / `_save_settings_to_file`（或等价逻辑）中读写上述 4 个键；高级设置中如需统一管理，可增加对应项。

### 4.4 在 `ui_translations.py` 中

- 新增文案键，例如：`chk_enable_gdrive_upload`、`lbl_gdrive_client_secrets_path`、`lbl_gdrive_folder_id`、`btn_upload_llm_to_gdrive`、`tip_*`、`msg_gdrive_upload_success`、`msg_gdrive_upload_failed`、`msg_gdrive_no_secrets` 等（中英文可仅维护中文）。

---

## 5. OAuth 与 Token 流程（简要）

1. 用户填写 `gdrive_client_secrets_path` 并点击上传。
2. 若本地无有效 token（或 `gdrive_token_path` 指向的文件不存在/过期）：
   - 使用 `google_auth_oauthlib.flow.InstalledAppFlow`，scope 包含 `https://www.googleapis.com/auth/drive.file`（仅访问通过本应用创建/打开的文件，更安全）或 `drive` 全量（若需写入任意位置）。
   - 运行 `run_local_server()` 打开浏览器，用户登录并授权。
   - 将返回的 credentials 序列化写入 `gdrive_token_path`（或默认 `.gdrive_token.json`）。
3. 若已有 token：用 `google.auth.load_credentials_from_file` 或等价方式加载，按需 refresh。
4. 使用 credentials 构建 Drive API 的 `Resource`，执行创建文件夹、上传文件。

---

## 6. 上传策略（细化）

- **目标文件夹**：  
  - 若 `gdrive_folder_id` 非空：在该文件夹下创建子文件夹 `Run_YYYYMMDD_HHMMSS`。  
  - 若为空：在 Drive 根下先查找/创建「知喂上传」，再在其下创建 `Run_YYYYMMDD_HHMMSS`。
- **本地路径**：以当前配置的 `llm_delivery_root` 为准；若为空则用 `<target_folder>/_LLM_UPLOAD`（与 `_maybe_build_llm_delivery_hub` 一致）。仅上传该目录下的文件（可递归或仅一层，与现有 flatten 行为一致），跳过子目录若用户选的是扁平化输出。
- **manifest**：上传成功后，若存在 `llm_upload_manifest.json`，则追加/更新 `gdrive` 段：`remote_folder_id`、`uploaded_at`（ISO8601）、可选 `run_folder_name`。

---

## 7. 实现顺序（建议分步）

| 步骤 | 内容 | 产出 |
|------|------|------|
| 1 | 依赖与配置 | `requirements.txt` 增加包；`office_converter.py` 默认配置 + setdefault 增加 4 个键 |
| 2 | `gdrive_upload.py` 骨架 | 模块存在，`ensure_credentials`、`upload_llm_folder_to_drive`、`update_manifest_gdrive_section` 接口定义；OAuth 与 Drive 创建文件夹 + 上传单文件先跑通 |
| 3 | OAuth 与 token | 完整首次授权与 token 缓存/刷新逻辑，命令行或小脚本可单独验证 |
| 4 | 批量上传与 manifest | 遍历 _LLM_UPLOAD、批量上传、写 manifest 的 gdrive 段 |
| 5 | GUI 控件与绑定 | 成果文件 Tab 增加开关、路径、按钮；加载/保存 config；按钮回调调用 gdrive_upload |
| 6 | 文案与错误提示 | ui_translations 全部键；GUI 中上传成功/失败/未配置密钥等提示 |
| 7 | 文档与打包 | 使用说明书、CHANGELOG、交接文档补充；`build_exe.py` 若需打包可选依赖则注明或条件导入 |

---

## 8. 测试要点

- 未配置 `gdrive_client_secrets_path` 时点击上传：提示先配置密钥路径。
- 首次上传：弹出浏览器完成 OAuth，上传成功后本地有 token 文件，再次上传不再弹窗。
- 上传后检查 Drive 上是否存在 `知喂上传/Run_*`（或指定父文件夹下），文件数与本地 _LLM_UPLOAD 一致。
- 存在 `llm_upload_manifest.json` 时，上传后其内包含 `gdrive` 段且 `remote_folder_id`、`uploaded_at` 正确。
- 关闭「启用 Google Drive 上传」时，上传区域可隐藏或禁用，保存/加载 config 正常。

---

## 9. 风险与注意

- **敏感数据**：`client_secrets.json` 与 token 已通过 `.gitignore` 排除，实现时默认 token 存用户目录、不写日志/不暴露内容，详见上文 §2.1。
- **配额与限流**：Drive API 有配额，大批量文件时考虑分批与简单重试；首次实现可先保证「小规模验证通过」。
- **可选依赖**：若希望未安装 google 包时 GUI 不报错，可在导入 `gdrive_upload` 时 try/except，上传按钮禁用并提示「请安装 google-auth-oauthlib 与 google-api-python-client」。

---

## 10. 文件改动清单（汇总）

| 文件 | 改动 |
|------|------|
| `requirements.txt` | 增加 google-auth-oauthlib、google-api-python-client |
| `office_converter.py` | default_config 与 setdefault 增加 4 个 gdrive 键 |
| `gdrive_upload.py` | 新建：OAuth、上传、manifest 更新 |
| `office_gui.py` | 成果文件 Tab 增加 GDrive 区与按钮，config 读写，按钮回调 |
| `ui_translations.py` | 新增 GDrive 相关 zh 文案 |
| `.gitignore` | 增加 `.gdrive_token.json` 或通用 token 文件名 |
| `docs/notes/使用说明书.md` | 增加「Google Drive 上传」小节 |
| `CHANGELOG.md` | 新版本条目标注 GDrive 上传功能 |
| `docs/dev/AI_交接文档_下一阶段开发.md` | 可选：将 GDrive 从「建议」改为「已实现」并简述用法 |

按上述顺序实现即可分步落地，且不破坏现有流程。
