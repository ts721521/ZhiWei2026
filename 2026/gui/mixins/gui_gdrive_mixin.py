# -*- coding: utf-8 -*-
"""Google Drive integration methods extracted from OfficeGUI."""

import importlib
import os
import re
import subprocess
import sys
import threading
import tkinter as tk
import webbrowser
from tkinter import filedialog, messagebox

try:
    import ttkbootstrap as tb
except ModuleNotFoundError:
    from tkinter import ttk

    class _BootstyleMixin:
        def __init__(self, *args, **kwargs):
            kwargs.pop("bootstyle", None)
            super().__init__(*args, **kwargs)

        def configure(self, cnf=None, **kwargs):
            kwargs.pop("bootstyle", None)
            if isinstance(cnf, dict) and "bootstyle" in cnf:
                cnf = dict(cnf)
                cnf.pop("bootstyle", None)
            return super().configure(cnf, **kwargs)

        config = configure

    class _FallbackFrame(_BootstyleMixin, ttk.Frame):
        pass

    class _FallbackButton(_BootstyleMixin, ttk.Button):
        pass

    class _TBNamespace:
        Frame = _FallbackFrame
        Button = _FallbackButton

    tb = _TBNamespace()


class GDriveMixin:
    def _on_open_gdrive_console(self):
        """在浏览器中打开 Google Cloud 凭据页面，方便用户获取 client_secrets.json。"""
        webbrowser.open("https://console.cloud.google.com/apis/credentials")

    def _on_open_gdrive_enable_api(self):
        """在浏览器中打开 Drive API 启用页面，解决 403 accessNotConfigured。"""
        try:
            import gdrive_upload as gd

            url = getattr(
                gd,
                "DRIVE_API_ENABLE_URL",
                "https://console.developers.google.com/apis/api/drive.googleapis.com/overview",
            )
        except Exception:
            url = "https://console.developers.google.com/apis/api/drive.googleapis.com/overview"
        webbrowser.open(url)

    def _on_browse_gdrive_secrets(self):
        path = filedialog.askopenfilename(
            title=self.tr("lbl_gdrive_client_secrets_path"),
            filetypes=[("JSON", "*.json"), ("All", "*.*")],
        )
        if path:
            self.var_gdrive_client_secrets_path.set(path)

    def _on_install_gdrive_deps(self):
        """后台执行 pip install google-auth-oauthlib google-api-python-client，完成后在主线程弹窗。"""
        pkgs = ["google-auth-oauthlib", "google-api-python-client"]
        running_msg = self.tr("msg_gdrive_install_running")

        def run_pip():
            try:
                r = subprocess.run(
                    [sys.executable, "-m", "pip", "install", "-q"] + pkgs,
                    capture_output=True,
                    text=True,
                    timeout=120,
                )
                err = (r.stderr or "").strip() if r.stderr else ""
                out = (r.stdout or "").strip() if r.stdout else ""
                if r.returncode != 0:
                    detail = err or out or str(r.returncode)
                    self.after(
                        0,
                        lambda: messagebox.showerror(
                            self.tr("msg_gdrive_install_failed"),
                            self.tr("msg_gdrive_install_failed_detail").format(
                                detail=detail
                            ),
                            parent=self,
                        ),
                    )
                else:

                    def on_ok():
                        messagebox.showinfo(
                            self.tr("msg_gdrive_install_ok"),
                            self.tr("msg_gdrive_install_ok_detail"),
                            parent=self,
                        )
                        # 安装成功后尝试重新启用 GDrive 控件（无需重启）
                        try:
                            import importlib
                            import gdrive_upload as _gd

                            importlib.reload(_gd)
                            if getattr(_gd, "HAS_GDEPEND", False):
                                for w in (
                                    self.chk_enable_gdrive_upload,
                                    self.entry_gdrive_client_secrets_path,
                                    self.btn_gdrive_open_console,
                                    self.btn_gdrive_enable_api,
                                    self.btn_browse_gdrive_secrets,
                                    self.entry_gdrive_folder_id,
                                    self.btn_upload_to_gdrive,
                                    self.btn_fetch_gdrive_structure,
                                ):
                                    try:
                                        w.configure(state="normal")
                                    except Exception:
                                        pass
                        except Exception:
                            pass

                    self.after(0, on_ok)
            except subprocess.TimeoutExpired:
                self.after(
                    0,
                    lambda: messagebox.showerror(
                        self.tr("msg_gdrive_install_failed"),
                        self.tr("msg_gdrive_install_timeout"),
                        parent=self,
                    ),
                )
            except Exception as e:
                self.after(
                    0,
                    lambda: messagebox.showerror(
                        self.tr("msg_gdrive_install_failed"),
                        str(e),
                        parent=self,
                    ),
                )

        messagebox.showinfo(
            self.tr("lbl_gdrive_install_title"),
            running_msg,
            parent=self,
        )
        t = threading.Thread(target=run_pip, daemon=True)
        t.start()

    def _on_fetch_gdrive_structure(self):
        """后台上拉远程目录结构，在独立窗口显示，便于测试。"""
        try:
            import gdrive_upload as gd
        except ImportError:
            messagebox.showerror(
                self.tr("msg_gdrive_upload_failed"),
                self.tr("msg_gdrive_no_deps"),
                parent=self,
            )
            return
        if not getattr(gd, "HAS_GDEPEND", True):
            messagebox.showerror(
                self.tr("msg_gdrive_upload_failed"),
                self.tr("msg_gdrive_no_deps"),
                parent=self,
            )
            return
        client_secrets = (self.var_gdrive_client_secrets_path.get() or "").strip()
        if not client_secrets:
            messagebox.showwarning(
                self.tr("msg_gdrive_upload_failed"),
                self.tr("msg_gdrive_no_secrets"),
                parent=self,
            )
            return
        folder_id = (self.var_gdrive_folder_id.get() or "").strip() or None

        def do_fetch():
            creds, err = gd.ensure_credentials(client_secrets, token_path=None)
            if err:
                self.after(0, lambda: self._gdrive_structure_done(None, err))
                return
            text, err = gd.list_remote_folder_structure(creds, folder_id)
            self.after(0, lambda: self._gdrive_structure_done(text, err))

        self.btn_fetch_gdrive_structure.configure(state="disabled")
        threading.Thread(target=do_fetch, daemon=True).start()

    def _gdrive_structure_done(self, text, err):
        """显示远程目录结构结果：错误弹窗或独立窗口；支持设为目标文件夹、在浏览器中打开。"""
        try:
            self.btn_fetch_gdrive_structure.configure(state="normal")
        except Exception:
            pass
        if err:
            messagebox.showerror(
                self.tr("title_gdrive_structure"),
                err,
                parent=self,
            )
            return
        if not text:
            return
        try:
            win = tk.Toplevel(self)
            win.title(self.tr("title_gdrive_structure"))
            win.geometry("720x560")
            try:
                from ttkbootstrap.widgets.scrolled import ScrolledText as ST
            except Exception:
                from tkinter.scrolledtext import ScrolledText as ST
            txt = ST(win, height=26, wrap=tk.WORD, font=("Consolas", 9))
            txt.pack(fill=tk.BOTH, expand=True, padx=8, pady=(8, 4))
            txt.insert(tk.END, text)
            # 保持可选以便「设为目标文件夹」/「在浏览器中打开」从选中行取 ID
            frm_actions = tb.Frame(win)
            frm_actions.pack(fill=tk.X, padx=8, pady=(0, 8))

            def _get_id_from_selection():
                try:
                    sel = txt.get(tk.SEL_FIRST, tk.SEL_LAST)
                except tk.TclError:
                    return None
                if not sel or not sel.strip():
                    return None
                # 从选中文本中取最后一个 [id]（Drive ID 为字母数字与 -_）
                matches = re.findall(r"\[([a-zA-Z0-9_-]+)\]", sel)
                return matches[-1] if matches else None

            def _on_set_as_target():
                fid = _get_id_from_selection()
                if not fid:
                    messagebox.showwarning(
                        self.tr("title_gdrive_structure"),
                        self.tr("msg_gdrive_select_line"),
                        parent=win,
                    )
                    return
                self.var_gdrive_folder_id.set(fid)
                messagebox.showinfo(
                    self.tr("title_gdrive_structure"),
                    self.tr("msg_gdrive_target_set").format(folder_id=fid),
                    parent=win,
                )

            def _on_open_in_browser():
                fid = _get_id_from_selection()
                if not fid:
                    messagebox.showwarning(
                        self.tr("title_gdrive_structure"),
                        self.tr("msg_gdrive_select_line"),
                        parent=win,
                    )
                    return
                webbrowser.open("https://drive.google.com/drive/folders/" + fid)

            tb.Button(
                frm_actions,
                text=self.tr("btn_gdrive_set_target"),
                bootstyle="primary-outline",
                command=_on_set_as_target,
            ).pack(side=tk.LEFT, padx=(0, 8))
            tb.Button(
                frm_actions,
                text=self.tr("btn_gdrive_open_in_browser"),
                bootstyle="secondary-outline",
                command=_on_open_in_browser,
            ).pack(side=tk.LEFT)
        except Exception as e:
            messagebox.showerror(
                self.tr("title_gdrive_structure"),
                "显示结果失败: " + str(e),
                parent=self,
            )

    def _on_upload_llm_to_gdrive(self):
        """将 _LLM_UPLOAD 目录上传到 Google Drive（后台上传，避免界面卡顿）。"""
        try:
            import gdrive_upload as gd
        except ImportError:
            messagebox.showerror(
                self.tr("msg_gdrive_upload_failed"),
                self.tr("msg_gdrive_no_deps"),
                parent=self,
            )
            return
        if not getattr(gd, "HAS_GDEPEND", True):
            messagebox.showerror(
                self.tr("msg_gdrive_upload_failed"),
                self.tr("msg_gdrive_no_deps"),
                parent=self,
            )
            return
        llm_root = (self.var_llm_delivery_root.get() or "").strip()
        if not llm_root:
            tgt = (self.var_target_folder.get() or "").strip()
            llm_root = os.path.join(tgt, "_LLM_UPLOAD") if tgt else ""
        if not llm_root or not os.path.isdir(llm_root):
            messagebox.showwarning(
                self.tr("msg_gdrive_upload_failed"),
                self.tr("msg_gdrive_no_llm_folder"),
                parent=self,
            )
            return
        client_secrets = (self.var_gdrive_client_secrets_path.get() or "").strip()
        if not client_secrets:
            messagebox.showwarning(
                self.tr("msg_gdrive_upload_failed"),
                self.tr("msg_gdrive_no_secrets"),
                parent=self,
            )
            return
        folder_id = (self.var_gdrive_folder_id.get() or "").strip() or None

        # 后台上传，避免 OAuth 与上传时界面卡死
        def do_upload():
            creds, err = gd.ensure_credentials(client_secrets, token_path=None)
            if err:
                self.after(0, lambda: self._gdrive_upload_done(None, err))
                return
            result, err = gd.upload_llm_folder_to_drive(llm_root, creds, folder_id)
            if result:
                manifest_path = os.path.join(llm_root, "llm_upload_manifest.json")
                gd.update_manifest_gdrive_section(manifest_path, result)
            self.after(0, lambda: self._gdrive_upload_done(result, err))

        self.btn_upload_to_gdrive.configure(state="disabled")
        # 不阻塞：后台上传，OAuth 会自行打开浏览器；完成后 _gdrive_upload_done 弹窗并恢复按钮
        threading.Thread(target=do_upload, daemon=True).start()

    def _gdrive_upload_done(self, result, err):
        """上传完成回调：恢复按钮并弹窗结果。"""
        try:
            self.btn_upload_to_gdrive.configure(state="normal")
        except Exception:
            pass
        if err and not result:
            messagebox.showerror(self.tr("msg_gdrive_upload_failed"), err, parent=self)
            return
        if result:
            fid = result.get("folder_id", "")
            detail = self.tr("msg_gdrive_upload_success_detail").format(
                file_count=result.get("file_count", 0),
                folder_id=fid,
            )
            if fid:
                detail += "\n\n" + self.tr("msg_gdrive_folder_link").format(
                    folder_id=fid
                )
            messagebox.showinfo(
                self.tr("msg_gdrive_upload_success"),
                detail,
                parent=self,
            )
        if err:
            messagebox.showwarning(
                self.tr("msg_gdrive_upload_failed"), err, parent=self
            )

