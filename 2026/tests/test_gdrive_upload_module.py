# -*- coding: utf-8 -*-
"""Unit tests for gdrive_upload module (Google Drive 上传)."""
import json
import os
import shutil
import tempfile
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

# 被测模块在仓库根目录，tests 在 tests/，需能导入 gdrive_upload
import sys

_REPO_ROOT = Path(__file__).resolve().parent.parent
if str(_REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(_REPO_ROOT))


class TestGDriveUploadModule(unittest.TestCase):
    """gdrive_upload 模块单元测试（不调用真实 Google API）。"""

    def test_format_drive_api_error_non_httperror(self):
        import gdrive_upload as gd

        self.assertIn("hello", gd._format_drive_api_error(ValueError("hello")))

    def test_format_drive_api_error_403_access_not_configured(self):
        import gdrive_upload as gd

        if gd.HttpError is None:
            self.skipTest("Google API not installed")
        err = gd.HttpError(MagicMock(status=403), b'{"error":{"message":"accessNotConfigured"}}')
        out = gd._format_drive_api_error(err)
        self.assertIn("Drive API", out)
        self.assertIn(gd.DRIVE_API_ENABLE_URL, out)

    def test_default_token_path_returns_path_under_user_dir(self):
        import gdrive_upload as gd

        path = gd._default_token_path()
        self.assertTrue(path.endswith("gdrive_token.json"), path)
        self.assertIn("知喂", path)

    def test_ensure_credentials_no_deps(self):
        import gdrive_upload as gd

        with patch.object(gd, "HAS_GDEPEND", False):
            creds, err = gd.ensure_credentials("/any/path.json")
        self.assertIsNone(creds)
        self.assertIsNotNone(err)
        self.assertTrue("google" in (err or "").lower() or "pip" in (err or "").lower())

    def test_ensure_credentials_missing_client_secrets_file(self):
        import gdrive_upload as gd

        if not gd.HAS_GDEPEND:
            self.skipTest("Google deps not installed")
        creds, err = gd.ensure_credentials("")
        self.assertIsNone(creds)
        self.assertIsNotNone(err)
        self.assertIn("密钥", err)
        creds2, err2 = gd.ensure_credentials(os.path.join(tempfile.gettempdir(), "nonexistent_secrets.json"))
        self.assertIsNone(creds2)
        self.assertIsNotNone(err2)

    def test_upload_llm_folder_no_deps(self):
        import gdrive_upload as gd

        with patch.object(gd, "HAS_GDEPEND", False):
            result, err = gd.upload_llm_folder_to_drive("/any/path", None)
        self.assertIsNone(result)
        self.assertIn("依赖", err or "")

    def test_upload_llm_folder_not_a_directory(self):
        import gdrive_upload as gd

        if not gd.HAS_GDEPEND:
            self.skipTest("Google deps not installed")
        result, err = gd.upload_llm_folder_to_drive("/nonexistent_dir_xyz", MagicMock())
        self.assertIsNone(result)
        self.assertIn("不存在", err or "")

    def test_upload_llm_folder_empty_directory(self):
        import gdrive_upload as gd

        if not gd.HAS_GDEPEND:
            self.skipTest("Google deps not installed")
        with tempfile.TemporaryDirectory() as d:
            result, err = gd.upload_llm_folder_to_drive(d, MagicMock())
        self.assertIsNone(result)
        self.assertIn("没有可上传", err or "")

    def test_upload_llm_folder_success_with_mocked_service(self):
        import gdrive_upload as gd

        if not gd.HAS_GDEPEND:
            self.skipTest("Google deps not installed")
        d = tempfile.mkdtemp()
        try:
            (Path(d) / "a.pdf").write_bytes(b"pdf")
            (Path(d) / "b.txt").write_bytes(b"text")
            mock_creds = MagicMock()
            mock_list = MagicMock()
            mock_list.execute.return_value = {"files": []}
            mock_create_folder = MagicMock()
            mock_create_folder.execute.return_value = {"id": "run_folder_id_123"}
            mock_create_file = MagicMock()
            mock_create_file.execute.return_value = {"id": "file_id"}
            mock_files = MagicMock()
            mock_files.list.return_value = mock_list
            mock_files.create.side_effect = [mock_create_folder, mock_create_file, mock_create_file]
            mock_service = MagicMock()
            mock_service.files.return_value = mock_files
            with patch.object(gd, "build", return_value=mock_service):
                result, err = gd.upload_llm_folder_to_drive(d, mock_creds, parent_folder_id="parent_abc")
            self.assertIsNone(err, err)
            self.assertIsNotNone(result)
            self.assertEqual(result["folder_id"], "run_folder_id_123")
            self.assertEqual(result["file_count"], 2)
            self.assertIn("uploaded_at", result)
        finally:
            shutil.rmtree(d, ignore_errors=True)

    def test_list_remote_folder_structure_no_deps(self):
        import gdrive_upload as gd

        with patch.object(gd, "HAS_GDEPEND", False):
            text, err = gd.list_remote_folder_structure(MagicMock())
        self.assertIsNone(text)
        self.assertIn("依赖", err or "")

    def test_update_manifest_gdrive_section_manifest_missing(self):
        import gdrive_upload as gd

        gd.update_manifest_gdrive_section(
            os.path.join(tempfile.gettempdir(), "nonexistent_manifest.json"), {"folder_id": "x"}
        )
        # 无异常即可

    def test_update_manifest_gdrive_section_writes_gdrive_section(self):
        import gdrive_upload as gd

        with tempfile.NamedTemporaryFile(mode="w", suffix=".json", delete=False, encoding="utf-8") as f:
            json.dump({"sources": []}, f, ensure_ascii=False)
            path = f.name
        try:
            gd.update_manifest_gdrive_section(
                path, {"folder_id": "fid1", "uploaded_at": "2026-01-01T00:00:00Z", "file_count": 3}
            )
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            self.assertIn("gdrive", data)
            self.assertEqual(data["gdrive"]["folder_id"], "fid1")
            self.assertEqual(data["gdrive"]["remote_folder_id"], "fid1")
            self.assertEqual(data["gdrive"]["file_count"], 3)
        finally:
            os.unlink(path)

    def test_constants(self):
        import gdrive_upload as gd

        self.assertEqual(gd.REMOTE_ROOT_FOLDER_NAME, "知喂上传")
        self.assertIn("drive.file", str(gd.SCOPES))


if __name__ == "__main__":
    unittest.main()
