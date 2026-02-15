# -*- coding: utf-8 -*-
"""
Google Drive 上传：OAuth 桌面流程 + 将 _LLM_UPLOAD 目录上传到 Drive。
敏感：client_secrets.json 与 token 文件不得提交到 Git，勿写入日志。
"""
import os
import json
import logging
from datetime import datetime

# 可选依赖：未安装时本模块对外返回明确错误，不抛 ImportError 给上层
try:
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaFileUpload
    from googleapiclient.errors import HttpError
    HAS_GDEPEND = True
except ImportError:
    HAS_GDEPEND = False
    HttpError = None

# Drive API 仅访问本应用创建/打开的文件
SCOPES = ["https://www.googleapis.com/auth/drive.file"]
# 远程根目录名（在用户 Drive 根下创建）
REMOTE_ROOT_FOLDER_NAME = "知喂上传"
# 未启用 Drive API 时提示用户打开的页面
DRIVE_API_ENABLE_URL = "https://console.developers.google.com/apis/api/drive.googleapis.com/overview"


def _format_drive_api_error(exc):
    """将 Drive API 的 403 accessNotConfigured 转为用户可读提示，含启用链接。"""
    if HttpError is None or not isinstance(exc, HttpError):
        return str(exc)
    try:
        status = getattr(getattr(exc, "resp", None), "status", None)
        if status == 403:
            content = getattr(exc, "content", b"") or b""
            text = content.decode("utf-8", errors="replace") if isinstance(content, bytes) else str(content)
            if "accessNotConfigured" in text or "has not been used" in text or "is disabled" in text:
                return "当前 Google Cloud 项目未启用 Drive API。请先启用后再重试：\n\n" + DRIVE_API_ENABLE_URL
    except Exception:
        pass
    return str(exc)


def _default_token_path():
    """Token 缓存默认路径：用户目录，避免放在仓库内。"""
    if os.name == "nt":
        base = os.environ.get("APPDATA") or os.path.expanduser("~")
        dir_path = os.path.join(base, "知喂")
    else:
        base = os.path.expanduser("~/.config")
        dir_path = os.path.join(base, "知喂")
    try:
        os.makedirs(dir_path, exist_ok=True)
    except OSError:
        pass
    return os.path.join(dir_path, "gdrive_token.json")


def ensure_credentials(client_secrets_path, token_path=None):
    """
    获取有效凭证：若 token_path 存在且有效则加载并刷新，否则走浏览器 OAuth 并写入 token_path。
    :param client_secrets_path: client_secrets.json 的路径
    :param token_path: 保存/读取 token 的路径，为空则用用户目录下的默认路径
    :return: (credentials, None) 或 (None, error_message)
    """
    if not HAS_GDEPEND:
        return None, "未安装 Google 依赖，请执行: pip install google-auth-oauthlib google-api-python-client"
    if not client_secrets_path or not os.path.isfile(client_secrets_path):
        return None, "未配置或找不到客户端密钥文件，请在设置中填写 client_secrets.json 路径"

    token_path = token_path or _default_token_path()
    creds = None
    if os.path.exists(token_path):
        try:
            creds = Credentials.from_authorized_user_file(token_path, SCOPES)
        except Exception:
            creds = None
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except Exception:
                creds = None
        if not creds:
            try:
                flow = InstalledAppFlow.from_client_secrets_file(
                    client_secrets_path, SCOPES
                )
                creds = flow.run_local_server(port=0)
            except Exception as e:
                return None, "OAuth 授权失败: " + str(e)
        try:
            with open(token_path, "w", encoding="utf-8") as f:
                f.write(creds.to_json())
            if os.name != "nt":
                try:
                    os.chmod(token_path, 0o600)
                except OSError:
                    pass
        except Exception as e:
            return None, "保存 token 失败: " + str(e)
    return creds, None


def upload_llm_folder_to_drive(local_llm_root, credentials, parent_folder_id=None):
    """
    将本地 _LLM_UPLOAD 目录内容上传到 Google Drive。
    :param local_llm_root: 本地 _LLM_UPLOAD 目录绝对路径
    :param credentials: google.oauth2.credentials.Credentials
    :param parent_folder_id: 可选，上传到此文件夹下；为空则在用户根下创建 知喂上传/Run_YYYYMMDD_HHMMSS
    :return: (result_dict, None) 或 (None, error_message)。result_dict 含 folder_id, uploaded_at, file_count
    """
    if not HAS_GDEPEND:
        return None, "未安装 Google 依赖"
    if not os.path.isdir(local_llm_root):
        return None, "本地目录不存在: " + local_llm_root

    # 仅统计一层文件，与上传逻辑一致
    top_level_files = [n for n in os.listdir(local_llm_root) if os.path.isfile(os.path.join(local_llm_root, n))]
    if not top_level_files:
        return None, "目录下没有可上传的文件（仅上传一层文件）"

    try:
        service = build("drive", "v3", credentials=credentials)
    except Exception as e:
        return None, "创建 Drive 服务失败: " + str(e)

    run_name = "Run_" + datetime.now().strftime("%Y%m%d_%H%M%S")
    parent_id = parent_folder_id

    if not parent_id:
        # 在根下查找或创建「知喂上传」
        q = "name='%s' and mimeType='application/vnd.google-apps.folder' and trashed=false" % (
            REMOTE_ROOT_FOLDER_NAME.replace("'", "\\'"),
        )
        try:
            results = (
                service.files()
                .list(q=q, spaces="drive", fields="files(id, name)", pageSize=1)
                .execute()
            )
            files = results.get("files", [])
            if files:
                parent_id = files[0]["id"]
            else:
                meta = {
                    "name": REMOTE_ROOT_FOLDER_NAME,
                    "mimeType": "application/vnd.google-apps.folder",
                }
                folder = service.files().create(body=meta, fields="id").execute()
                parent_id = folder.get("id")
        except Exception as e:
            return None, "创建/查找远程根目录失败: " + (_format_drive_api_error(e) if HttpError and isinstance(e, HttpError) else str(e))

    # 创建本次运行子文件夹
    try:
        meta = {
            "name": run_name,
            "mimeType": "application/vnd.google-apps.folder",
            "parents": [parent_id],
        }
        folder = service.files().create(body=meta, fields="id").execute()
        run_folder_id = folder.get("id")
    except Exception as e:
        return None, "创建运行目录失败: " + (_format_drive_api_error(e) if HttpError and isinstance(e, HttpError) else str(e))

    # 遍历本地上传文件（仅一层，与扁平化输出一致）
    uploaded = 0
    try:
        for name in os.listdir(local_llm_root):
            path = os.path.join(local_llm_root, name)
            if not os.path.isfile(path):
                continue
            mime = "application/octet-stream"
            if name.endswith(".pdf"):
                mime = "application/pdf"
            elif name.endswith(".md"):
                mime = "text/markdown"
            elif name.endswith(".json"):
                mime = "application/json"
            elif name.endswith(".txt"):
                mime = "text/plain"
            media = MediaFileUpload(path, mimetype=mime, resumable=True)
            file_meta = {"name": name, "parents": [run_folder_id]}
            service.files().create(body=file_meta, media_body=media, fields="id").execute()
            uploaded += 1
    except Exception as e:
        msg = _format_drive_api_error(e) if HttpError and isinstance(e, HttpError) else str(e)
        return (
            {"folder_id": run_folder_id, "uploaded_at": datetime.utcnow().isoformat() + "Z", "file_count": uploaded},
            "部分上传后出错: " + msg,
        )

    return (
        {
            "folder_id": run_folder_id,
            "uploaded_at": datetime.utcnow().isoformat() + "Z",
            "file_count": uploaded,
        },
        None,
    )


def list_remote_folder_structure(credentials, parent_folder_id=None):
    """
    列出远程目录结构，便于测试：若 parent_folder_id 为空则列出「知喂上传」及其子项，否则列出该文件夹下内容。
    :return: (structure_text, None) 或 (None, error_message)
    """
    if not HAS_GDEPEND:
        return None, "未安装 Google 依赖"
    try:
        service = build("drive", "v3", credentials=credentials)
    except Exception as e:
        return None, "创建 Drive 服务失败: " + str(e)

    root_id = parent_folder_id
    root_name = "指定文件夹"
    if not root_id:
        q = "name='%s' and mimeType='application/vnd.google-apps.folder' and trashed=false" % (
            REMOTE_ROOT_FOLDER_NAME.replace("'", "\\'"),
        )
        try:
            results = (
                service.files()
                .list(q=q, spaces="drive", fields="files(id, name)", pageSize=1)
                .execute()
            )
            files = results.get("files", [])
            if not files:
                return None, "根目录下未找到「%s」，请先执行一次上传。" % REMOTE_ROOT_FOLDER_NAME
            root_id = files[0]["id"]
            root_name = REMOTE_ROOT_FOLDER_NAME
        except Exception as e:
            return None, "查找远程根目录失败: " + (_format_drive_api_error(e) if HttpError and isinstance(e, HttpError) else str(e))

    lines = []
    folder_mime = "application/vnd.google-apps.folder"

    def list_children(folder_id, prefix=""):
        q = "'%s' in parents and trashed=false" % folder_id
        try:
            results = (
                service.files()
                .list(q=q, spaces="drive", fields="files(id, name, mimeType)", pageSize=200)
                .execute()
            )
        except Exception as e:
            lines.append(prefix + "[ 列出失败: %s ]" % str(e))
            return
        items = results.get("files", [])
        for f in sorted(items, key=lambda x: (x.get("mimeType") != folder_mime, (x.get("name") or "").lower())):
            name = f.get("name") or "(无名称)"
            mid = f.get("id", "")
            is_dir = f.get("mimeType") == folder_mime
            lines.append("%s%s %s [%s]" % (prefix, "📁 " if is_dir else "📄 ", name, mid))
            if is_dir:
                list_children(mid, prefix + "    ")

    lines.append("%s [%s]" % (root_name, root_id))
    list_children(root_id, "  ")
    return "\n".join(lines), None


def update_manifest_gdrive_section(manifest_path, gdrive_info):
    """
    在 llm_upload_manifest.json 中追加或更新 gdrive 区段。
    :param manifest_path: 清单文件路径
    :param gdrive_info: dict，含 folder_id, uploaded_at, file_count 等；同时写入 remote_folder_id 与规划一致
    """
    if not os.path.isfile(manifest_path):
        return
    try:
        with open(manifest_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        # 兼容规划文档中的 remote_folder_id 键名
        section = dict(gdrive_info)
        if "folder_id" in section and "remote_folder_id" not in section:
            section["remote_folder_id"] = section["folder_id"]
        data["gdrive"] = section
        with open(manifest_path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
    except Exception as e:
        logging.warning("更新 manifest gdrive 段失败: %s", e)
