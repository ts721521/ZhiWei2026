
import os
import sys
import time
import json
import shutil
import logging
import argparse
import ctypes
import uuid
import tempfile
import subprocess
import threading
import signal
import random
import win32com.client
import pythoncom
import pywintypes
from datetime import datetime
from pathlib import Path

# 全局版本号
__version__ = "5.5.0"

# 常量定义
wdFormatPDF = 17
xlTypePDF = 0
ppSaveAsPDF = 32
ppFixedFormatTypePDF = 2
xlPDF_SaveAs = 57
xlRepairFile = 1

# 引擎类型常量
ENGINE_WPS = "wps"
ENGINE_MS = "ms"
ENGINE_ASK = "ask"

# 进程处理模式
KILL_MODE_ASK = "ask"
KILL_MODE_AUTO = "auto"
KILL_MODE_KEEP = "keep"

# COM 繁忙错误码
ERR_RPC_SERVER_BUSY = -2147417846

def get_app_path():
    """获取程序的运行目录"""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))


class OfficeConverter:
    def __init__(self, config_path):
        self.temp_sandbox = None
        self.failed_dir = None
        self.config_path = config_path
        self.engine_type = None
        self.is_running = True
        self.reuse_process = False

        signal.signal(signal.SIGINT, self.signal_handler)
        signal.signal(signal.SIGTERM, self.signal_handler)

        self.load_config(config_path)

        self.select_engine_mode()
        self.check_and_handle_running_processes()
        self.confirm_config_with_dialog()
        self.ask_for_subfolder()

        self.temp_sandbox = os.path.join(tempfile.gettempdir(), "OfficeToPDF_Sandbox")
        if not os.path.exists(self.temp_sandbox):
            os.makedirs(self.temp_sandbox)

        self.failed_dir = os.path.join(self.config['target_folder'], "_FAILED_FILES")
        if not os.path.exists(self.failed_dir):
            os.makedirs(self.failed_dir)

        self.setup_logging()

        self.stats = {
            "total": 0, "success": 0, "failed": 0,
            "skipped": 0, "conflicts": 0, "timeout": 0
        }
        self.error_records = []

    def signal_handler(self, signum, frame):
        print("\n\n" + "!" * 60)
        print("  接收到终止信号 (Ctrl+C)！正在紧急停止...")
        self.is_running = False
        if not self.reuse_process:
            print("  正在清理后台进程...")
            self.cleanup_all_processes()

        if self.temp_sandbox and os.path.exists(self.temp_sandbox):
            try:
                shutil.rmtree(self.temp_sandbox, ignore_errors=True)
            except:
                pass

        print("--> 程序已退出。")
        sys.exit(0)

    def get_process_list(self, process_names):
        running_list = []
        try:
            cmd = 'tasklist /FO CSV /NH'
            output = subprocess.check_output(cmd, shell=True).decode('gbk', errors='ignore')
            for line in output.splitlines():
                parts = line.split(',')
                if len(parts) > 1:
                    p_name = parts[0].strip('"').lower()
                    p_pid = parts[1].strip('"')
                    base_name = p_name.replace('.exe', '')
                    if base_name in process_names:
                        running_list.append(f"{p_name} (PID: {p_pid})")
        except Exception:
            pass
        return running_list

    def check_and_handle_running_processes(self):
        mode = self.config.get('kill_process_mode', KILL_MODE_ASK)
        if mode == KILL_MODE_KEEP:
            self.reuse_process = True
            print(f"--> 根据配置，已启用[进程复用]模式。")
            return

        target_apps = ["wps", "et", "wpp", "wpscenter"] if self.engine_type == ENGINE_WPS else ["winword", "excel", "powerpnt"]
        app_label = "WPS Office" if self.engine_type == ENGINE_WPS else "Microsoft Office"

        print(f"\n正在扫描系统中的 {app_label} 进程...")
        running = self.get_process_list(target_apps)

        if not running:
            print("--> 未发现相关残留进程，环境干净。")
            return

        if mode == KILL_MODE_AUTO:
            print(f"--> [自动模式] 检测到 {len(running)} 个进程，正在清理...")
            self.cleanup_all_processes()
            self.reuse_process = False
            return

        print("\n" + "!" * 60)
        print(f" [警告] 检测到以下 {app_label} 程序正在运行：")
        for p in running:
            print(f"  - {p}")
        print("!" * 60)
        print("请选择操作：")
        print("  [1] 强力清理：自动杀掉上述进程（推荐）")
        print("  [2] 尝试复用：不杀进程（可能因弹窗导致失败）")
        print("  [3] 退出程序")
        print("-" * 60)

        while True:
            choice = input("请输入序号 (1-3): ").strip()
            if choice == '1':
                self.cleanup_all_processes()
                self.reuse_process = False
                break
            elif choice == '2':
                self.reuse_process = True
                break
            elif choice == '3':
                sys.exit(0)
            else:
                print("输入无效。")
        print("\n")

    def cleanup_all_processes(self):
        apps = ["wps", "et", "wpp", "wpscenter", "wpscloudsvr"] if self.engine_type == ENGINE_WPS or self.engine_type is None else []
        if self.engine_type == ENGINE_MS or self.engine_type is None:
            apps.extend(["winword", "excel", "powerpnt"])

        for app in apps:
            self._kill_process_by_name(app)

    def _kill_process_by_name(self, app_name):
        if not app_name:
            return
        try:
            cmd = f"taskkill /F /IM {app_name}.exe"
            subprocess.run(cmd, shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        except:
            pass

    def load_config(self, path):
        try:
            with open(path, 'r', encoding='utf-8') as f:
                content = f.read().replace('\\', '/')
                self.config = json.loads(content)
        except Exception as e:
            print(f"[错误] 无法加载配置文件: {e}")
            sys.exit(1)

        self.config['source_folder'] = os.path.abspath(self.config['source_folder'])
        self.config['target_folder'] = os.path.abspath(self.config['target_folder'])

        # 通用超时、沙盒等配置
        if 'timeout_seconds' not in self.config:
            self.config['timeout_seconds'] = 60
        if 'enable_sandbox' not in self.config:
            self.config['enable_sandbox'] = True
        if 'default_engine' not in self.config:
            self.config['default_engine'] = ENGINE_ASK
        if 'kill_process_mode' not in self.config:
            self.config['kill_process_mode'] = KILL_MODE_ASK
        if 'auto_retry_failed' not in self.config:
            self.config['auto_retry_failed'] = False
        if 'pdf_wait_seconds' not in self.config:
            self.config['pdf_wait_seconds'] = 15

        # PPT 专属配置（如果没配，就走通用值）
        if 'ppt_timeout_seconds' not in self.config:
            self.config['ppt_timeout_seconds'] = self.config['timeout_seconds']
        if 'ppt_pdf_wait_seconds' not in self.config:
            self.config['ppt_pdf_wait_seconds'] = self.config['pdf_wait_seconds']

        exts = self.config.setdefault('allowed_extensions', {})
        exts.setdefault('word', [])
        exts.setdefault('excel', [])
        exts.setdefault('powerpoint', [])
        if 'pdf' not in exts:
            exts['pdf'] = ['.pdf']

    def save_config(self):
        try:
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=4, ensure_ascii=False)
            print("--> 配置文件已更新。")
        except Exception:
            pass

    def select_engine_mode(self):
        default = self.config.get('default_engine', ENGINE_ASK)
        if default == ENGINE_WPS:
            self.engine_type = ENGINE_WPS
            print(f"--> [自动选择] 使用引擎: WPS Office")
            return
        elif default == ENGINE_MS:
            self.engine_type = ENGINE_MS
            print(f"--> [自动选择] 使用引擎: Microsoft Office")
            return

        print("=" * 60)
        print(f"Office转换工具 v{__version__} - 引擎选择")
        print("=" * 60)
        print("  [1] WPS Office (兼容性好)")
        print("  [2] Microsoft Office (速度快)")
        print("-" * 60)
        while True:
            choice = input("请输入序号 (1/2): ").strip()
            if choice == '1' or choice == '':
                self.engine_type = ENGINE_WPS
                break
            elif choice == '2':
                self.engine_type = ENGINE_MS
                break
            else:
                print("输入无效，请输入 1 或 2。")
        print(f"--> 已选择: {self.engine_type.upper()} 模式\n")

    def confirm_config_with_dialog(self):
        msg = f"源目录:\n{self.config['source_folder']}\n\n目标目录:\n{self.config['target_folder']}\n\n是否需要修改？"
        if ctypes.windll.user32.MessageBoxW(0, msg, "配置确认", 4 | 0x40) == 6:
            print("\n=== 配置修改模式 ===")
            print(f"当前源目录: {self.config['source_folder']}")
            new_s = input(f"请输入新源目录 [回车保持不变]: ").strip().replace('"', '').replace("'", "")
            if new_s:
                self.config['source_folder'] = os.path.abspath(new_s)

            print(f"\n当前目标目录: {self.config['target_folder']}")
            new_t = input(f"请输入新目标目录 [回车保持不变]: ").strip().replace('"', '').replace("'", "")
            if new_t:
                self.config['target_folder'] = os.path.abspath(new_t)

            self.save_config()
            print("配置已保存。\n")

    def ask_for_subfolder(self):
        print("-" * 60)
        sub = input("请输入本次输出子文件夹名称 (直接回车不创建): ").strip()
        if sub:
            for char in '<>:"/\\|?*':
                sub = sub.replace(char, '')
            self.config['target_folder'] = os.path.abspath(os.path.join(self.config['target_folder'], sub))
            print(f"--> 最终位置: {self.config['target_folder']}")

    def setup_logging(self):
        log_dir = self.config.get('log_folder', './logs')
        if not os.path.isabs(log_dir):
            log_dir = os.path.join(get_app_path(), log_dir)
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
        self.log_path = os.path.join(log_dir, f"conversion_log_{datetime.now().strftime('%Y%m%d_%H%M')}.txt")

        logging.basicConfig(
            filename=self.log_path,
            level=logging.INFO,
            format='%(message)s',
            encoding='utf-8',
            force=True
        )
        console = logging.StreamHandler()
        console.setLevel(logging.INFO)
        console.setFormatter(logging.Formatter('%(message)s'))
        logging.getLogger('').addHandler(console)

        with open(self.log_path, 'a', encoding='utf-8') as f:
            f.write(f"{datetime.now()} === 转换任务开始 (v{__version__}) ===\n")
            f.write(f"使用引擎: {self.engine_type.upper()}\n")
            f.write(f"源目录: {self.config['source_folder']}\n")
            f.write(f"目标目录: {self.config['target_folder']}\n")
            f.write(f"沙盒模式: {'开启' if self.config['enable_sandbox'] else '关闭'}\n")
            f.write(f"PPT线程超时: {self.config.get('ppt_timeout_seconds')}s, PPT I/O等待: {self.config.get('ppt_pdf_wait_seconds')}s\n")

    def _kill_current_app(self, app_type):
        if self.reuse_process:
            return
        name_map = {
            ENGINE_WPS: {'word': 'wps', 'excel': 'et', 'ppt': 'wpp'},
            ENGINE_MS: {'word': 'winword', 'excel': 'excel', 'ppt': 'powerpnt'}
        }
        app_name = name_map[self.engine_type].get(app_type, '')
        self._kill_process_by_name(app_name)

    def _get_local_app(self, app_type):
        pythoncom.CoInitialize()

        prog_id = ""
        if self.engine_type == ENGINE_WPS:
            prog_id = {"word": "Kwps.Application", "excel": "Ket.Application", "ppt": "Kwpp.Application"}.get(app_type)
        else:
            prog_id = {"word": "Word.Application", "excel": "Excel.Application", "ppt": "PowerPoint.Application"}.get(app_type)

        app = None
        try:
            app = win32com.client.Dispatch(prog_id)
        except:
            try:
                app = win32com.client.DispatchEx(prog_id)
            except Exception as e:
                raise Exception(f"无法启动 {prog_id}: {e}")

        try:
            app.Visible = False
            if app_type != 'ppt':
                app.DisplayAlerts = False
        except:
            pass

        if self.engine_type == ENGINE_MS and app_type == 'excel':
            try:
                app.AskToUpdateLinks = False
            except:
                pass

        return app

    def close_office_apps(self):
        if not self.reuse_process:
            self.cleanup_all_processes()

    def get_target_path(self, source_file_path):
        filename = os.path.basename(source_file_path)
        base_name = os.path.splitext(filename)[0]
        return os.path.join(self.config['target_folder'], base_name + ".pdf")

    def handle_file_conflict(self, temp_pdf_path, target_pdf_path):
        if not os.path.exists(target_pdf_path):
            os.makedirs(os.path.dirname(target_pdf_path), exist_ok=True)
            shutil.move(temp_pdf_path, target_pdf_path)
            return "成功", target_pdf_path

        if os.path.getsize(temp_pdf_path) == os.path.getsize(target_pdf_path):
            try:
                os.remove(target_pdf_path)
                shutil.move(temp_pdf_path, target_pdf_path)
                return "覆盖", target_pdf_path
            except:
                return "覆盖失败", target_pdf_path
        else:
            self.stats["conflicts"] += 1
            conflict_dir = os.path.join(os.path.dirname(target_pdf_path), "conflicts")
            os.makedirs(conflict_dir, exist_ok=True)
            fname = os.path.splitext(os.path.basename(target_pdf_path))[0]
            ts = datetime.now().strftime("%Y%m%d%H%M%S")
            new_path = os.path.join(conflict_dir, f"{fname}_{ts}.pdf")
            shutil.move(temp_pdf_path, new_path)
            return "冲突备份", new_path

    def _safe_exec(self, func, *args, retries=3, **kwargs):
        for attempt in range(retries + 1):
            if not self.is_running:
                raise Exception("程序已终止")
            try:
                return func(*args, **kwargs)
            except pywintypes.com_error as e:
                error_code = e.hresult
                if error_code == ERR_RPC_SERVER_BUSY:
                    wait_time = random.randint(2, 5)
                    time.sleep(wait_time)
                    continue
                if attempt < retries:
                    time.sleep(1)
                    continue
                raise Exception(f"COM错误 ({error_code}): {e}")
            except Exception:
                if attempt < retries:
                    time.sleep(1)
                    continue
                raise

    # [核心] 尝试解除文件锁定（删除 Zone.Identifier）
    def _unblock_file(self, file_path):
        try:
            zone_path = file_path + ":Zone.Identifier"
            try:
                os.remove(zone_path)
                logging.info(f"[UNBLOCK] 已尝试解除锁定: {file_path}")
            except FileNotFoundError:
                # 本地文件 / 同步目录一般没有 ADS 流
                logging.debug(f"[UNBLOCK] 未发现 Zone.Identifier: {file_path}")
            except Exception as e:
                logging.warning(f"[UNBLOCK] 删除 Zone.Identifier 失败: {file_path} | {e}")
        except Exception as e:
            logging.warning(f"[UNBLOCK] 处理解除锁定时出现异常: {file_path} | {e}")

    def convert_logic_in_thread(self, file_source, sandbox_target_pdf, ext):
        app = None
        doc = None

        try:
            if ext in self.config['allowed_extensions']['word']:
                app = self._get_local_app('word')
                try:
                    if self.engine_type == ENGINE_WPS:
                        try:
                            doc = self._safe_exec(app.Documents.Open, file_source, ReadOnly=True)
                        except:
                            doc = self._safe_exec(app.Documents.Open, file_source)
                    else:
                        doc = self._safe_exec(
                            app.Documents.Open,
                            file_source,
                            ReadOnly=True,
                            Visible=False,
                            OpenAndRepair=True
                        )
                    self._safe_exec(doc.ExportAsFixedFormat, sandbox_target_pdf, wdFormatPDF)
                finally:
                    if doc:
                        try:
                            doc.Close(SaveChanges=False)
                        except:
                            pass

            elif ext in self.config['allowed_extensions']['excel']:
                app = self._get_local_app('excel')
                try:
                    if self.engine_type == ENGINE_WPS:
                        try:
                            doc = self._safe_exec(app.Workbooks.Open, file_source, ReadOnly=True)
                        except:
                            doc = self._safe_exec(app.Workbooks.Open, file_source)
                        try:
                            self._safe_exec(doc.ExportAsFixedFormat, xlTypePDF, sandbox_target_pdf)
                        except:
                            if os.path.exists(sandbox_target_pdf):
                                os.remove(sandbox_target_pdf)
                            self._safe_exec(doc.SaveAs, sandbox_target_pdf, FileFormat=xlPDF_SaveAs)
                    else:
                        doc = self._safe_exec(
                            app.Workbooks.Open,
                            file_source,
                            UpdateLinks=0,
                            ReadOnly=True,
                            IgnoreReadOnlyRecommended=True,
                            CorruptLoad=xlRepairFile
                        )
                        self._safe_exec(doc.ExportAsFixedFormat, xlTypePDF, sandbox_target_pdf)
                finally:
                    if doc:
                        try:
                            doc.Close(SaveChanges=False)
                        except:
                            pass

            elif ext in self.config['allowed_extensions']['powerpoint']:
                app = self._get_local_app('ppt')
                try:
                    if self.engine_type == ENGINE_WPS:
                        try:
                            doc = self._safe_exec(app.Presentations.Open, file_source, WithWindow=False)
                        except:
                            doc = self._safe_exec(app.Presentations.Open, file_source)
                        self._safe_exec(doc.SaveCopyAs, sandbox_target_pdf, ppSaveAsPDF)
                    else:
                        doc = self._safe_exec(
                            app.Presentations.Open,
                            file_source,
                            WithWindow=False,
                            ReadOnly=True
                        )
                        try:
                            self._safe_exec(doc.ExportAsFixedFormat, sandbox_target_pdf, ppFixedFormatTypePDF)
                        except:
                            if os.path.exists(sandbox_target_pdf):
                                os.remove(sandbox_target_pdf)
                            self._safe_exec(doc.SaveCopyAs, sandbox_target_pdf, ppSaveAsPDF)
                finally:
                    if doc:
                        try:
                            doc.Close()
                        except:
                            pass

        finally:
            if app:
                try:
                    app.Quit()
                except:
                    pass
            pythoncom.CoUninitialize()

    def copy_pdf_direct(self, source, temp_target):
        try:
            shutil.copy2(source, temp_target)
        except Exception as e:
            raise Exception(f"[PDF复制失败] {e}")

    def quarantine_failed_file(self, source_path, should_copy=True):
        if not should_copy:
            return
        try:
            fname = os.path.basename(source_path)
            target = os.path.join(self.failed_dir, fname)
            if os.path.exists(target):
                name, ext = os.path.splitext(fname)
                target = os.path.join(self.failed_dir, f"{name}_{datetime.now().strftime('%H%M%S')}{ext}")
            shutil.copy2(source_path, target)
        except Exception:
            pass

    def process_single_file(self, file_path, target_path, ext, progress_str, is_retry=False):
        sandbox_pdf = os.path.join(self.temp_sandbox, f"{uuid.uuid4()}.pdf")
        filename = os.path.basename(file_path)

        use_sandbox = self.config.get('enable_sandbox', True)
        working_src = file_path
        sandbox_src_path = None

        # 是否 PPT，用于专门的超时和等待
        is_ppt = ext in self.config['allowed_extensions'].get('powerpoint', [])

        # 线程执行超时配置
        base_timeout = self.config.get('timeout_seconds', 60)
        ppt_timeout = self.config.get('ppt_timeout_seconds', base_timeout)
        current_timeout = ppt_timeout if is_ppt else base_timeout

        # I/O 等待 PDF 生成时间
        base_wait = self.config.get('pdf_wait_seconds', 15)
        ppt_wait = self.config.get('ppt_pdf_wait_seconds', base_wait)
        current_pdf_wait = ppt_wait if is_ppt else base_wait

        try:
            if use_sandbox:
                sandbox_src_path = os.path.join(self.temp_sandbox, f"{uuid.uuid4()}{ext}")
                shutil.copy2(file_path, sandbox_src_path)
                # 解除文件锁定（如果有 Zone.Identifier）
                self._unblock_file(sandbox_src_path)
                working_src = sandbox_src_path

            if ext == '.pdf':
                self.copy_pdf_direct(working_src, sandbox_pdf)
            else:
                convert_thread = threading.Thread(
                    target=self.convert_logic_in_thread,
                    args=(working_src, sandbox_pdf, ext),
                    daemon=True
                )
                convert_thread.start()

                wait_start = time.time()
                while convert_thread.is_alive():
                    elapsed = time.time() - wait_start
                    if elapsed > current_timeout:
                        break

                    print(f"\r{progress_str} 正在处理: {filename} ({elapsed:.1f}s)   ", end="", flush=True)
                    time.sleep(0.1)

                convert_thread.join(timeout=0.1)

                if convert_thread.is_alive():
                    self.stats["timeout"] += 1
                    logging.error(f"超时跳过 (>{current_timeout}s) - {file_path}")

                    if not self.reuse_process:
                        if ext in self.config['allowed_extensions']['word']:
                            self._kill_current_app('word')
                        elif ext in self.config['allowed_extensions']['excel']:
                            self._kill_current_app('excel')
                        elif ext in self.config['allowed_extensions']['powerpoint']:
                            self._kill_current_app('ppt')

                    raise Exception("超时")

            # 使用配置的等待时间
            wait_pdf_start = time.time()
            while time.time() - wait_pdf_start < current_pdf_wait:
                if os.path.exists(sandbox_pdf):
                    time.sleep(0.5)
                    return self.handle_file_conflict(sandbox_pdf, target_path)
                time.sleep(0.5)

            raise Exception(f"转换指令已发送但未生成PDF ({current_pdf_wait}s内未检测到文件)")

        finally:
            try:
                if sandbox_src_path and os.path.exists(sandbox_src_path):
                    os.remove(sandbox_src_path)
                if os.path.exists(sandbox_pdf):
                    os.remove(sandbox_pdf)
            except Exception:
                pass

    def get_progress_prefix(self, current, total):
        width = len(str(total))
        percent = current / total if total > 0 else 0
        bar_len = 20
        filled = int(bar_len * percent)
        bar = '█' * filled + '░' * (bar_len - filled)
        return f"[{int(percent * 100):>3}%]{bar} [{str(current).rjust(width)}/{total}]"

    def run_batch(self, file_list, is_retry=False):
        total = len(file_list)
        for i, fpath in enumerate(file_list, 1):
            if not self.is_running:
                break

            fname = os.path.basename(fpath)
            ext = os.path.splitext(fpath)[1].lower()
            target_path = os.path.join(self.config['target_folder'], os.path.splitext(fname)[0] + ".pdf")

            progress_prefix = self.get_progress_prefix(i, total)
            label = "[重试]" if is_retry else "正在处理"
            print(f"\r{progress_prefix} {label}: {fname}" + " " * 20, end="", flush=True)

            start = time.time()
            try:
                status, final_path = self.process_single_file(fpath, target_path, ext, progress_prefix, is_retry)
                self.stats["success"] += 1
                elapsed = time.time() - start
                print(f"\r{progress_prefix} {status}: {fname} (耗时: {elapsed:.2f}s)   ")
                logging.info(f"{status}: {fpath} -> {final_path}")
            except Exception as e:
                elapsed = time.time() - start
                err_msg = str(e)
                if "超时" in err_msg:
                    print(f"\r{progress_prefix} 超时: {fname} (耗时: {elapsed:.2f}s)   ")
                else:
                    self.stats["failed"] += 1
                    print(f"\r{progress_prefix} 失败: {fname} (耗时: {elapsed:.2f}s)   ")

                logging.error(f"失败: {fpath} | 原因: {e}")

                if not is_retry:
                    self.quarantine_failed_file(fpath)
                    self.error_records.append(fpath)

    def run(self):
        files = []
        logging.info("正在扫描文件...")

        if not os.path.exists(self.config['source_folder']):
            print(f"\n[警告] 源目录不存在: {self.config['source_folder']}")
            logging.error(f"源目录不存在: {self.config['source_folder']}")
            return

        for root, _, filenames in os.walk(self.config['source_folder']):
            for f in filenames:
                if not f.startswith("~$"):
                    ext = os.path.splitext(f)[1].lower()
                    valid_exts = [e for sub in self.config['allowed_extensions'].values() for e in sub]
                    if ext in valid_exts:
                        files.append(os.path.join(root, f))

        logging.info(f"开始处理 {len(files)} 个文件...")
        self.stats["total"] = len(files)

        if len(files) > 0:
            self.run_batch(files)
        else:
            print("\n[提示] 源目录中没有发现可转换的 Office 文件。")

        self.close_office_apps()

        failed_count = self.stats["failed"] + self.stats["timeout"]
        should_retry = False
        if failed_count > 0:
            if self.config.get('auto_retry_failed', False):
                should_retry = True
                print(f"\n[配置] 自动重试失败文件 ({failed_count}个)...")
            else:
                msg = f"任务完成，但有 {failed_count} 个文件失败。\n是否尝试重新处理这些失败的文件？"
                if ctypes.windll.user32.MessageBoxW(0, msg, "失败重试", 4 | 0x20) == 6:
                    should_retry = True

        if should_retry:
            print("\n" + "=" * 60)
            print("  开始重试失败文件...")
            print("  正在重新检查并清理进程...")
            print("=" * 60)

            if not self.reuse_process:
                self.cleanup_all_processes()

            retry_files = []
            if os.path.exists(self.failed_dir):
                if self.config.get('enable_sandbox', True) and not os.path.exists(self.temp_sandbox):
                    os.makedirs(self.temp_sandbox)

                for f in os.listdir(self.failed_dir):
                    if not f.startswith("~$"):
                        ext = os.path.splitext(f)[1].lower()
                        valid_exts = [e for sub in self.config['allowed_extensions'].values() for e in sub]
                        if ext in valid_exts:
                            retry_files.append(os.path.join(self.failed_dir, f))

            if retry_files:
                self.run_batch(retry_files, is_retry=True)
            else:
                print("未在失败目录找到可重试的文件。")

            self.close_office_apps()

        summary = (
            f"\n=== 最终统计 (v{__version__}) ===\n"
            f"总处理: {self.stats['total']}\n"
            f"成功: {self.stats['success']}\n"
            f"失败: {self.stats['failed']}\n"
            f"超时: {self.stats['timeout']}\n"
        )
        logging.info(summary)
        print(summary)

        try:
            os.startfile(self.config['target_folder'])
        except Exception:
            pass

        if self.temp_sandbox and os.path.exists(self.temp_sandbox):
            try:
                shutil.rmtree(self.temp_sandbox, ignore_errors=True)
            except Exception:
                pass


if __name__ == "__main__":
    script_dir = get_app_path()
    default_config_path = os.path.join(script_dir, "config.json")
    parser = argparse.ArgumentParser()
    parser.add_argument("--config", default=default_config_path)
    args = parser.parse_args()

    if not os.path.exists(args.config):
        try:
            default_config = {
                "source_folder": "C:\\Docs",
                "target_folder": "C:\\PDFs",
                "log_folder": "./logs",
                "enable_sandbox": True,
                "default_engine": "ask",
                "kill_process_mode": "ask",
                "auto_retry_failed": False,
                "timeout_seconds": 60,
                "pdf_wait_seconds": 15,
                "ppt_timeout_seconds": 180,      # PPT 单独线程超时时间（可以按需改）
                "ppt_pdf_wait_seconds": 30,      # PPT 单独 I/O 等待时间（可以按需改）
                "allowed_extensions": {
                    "word": [".doc", ".docx"],
                    "excel": [".xls", ".xlsx"],
                    "powerpoint": [".ppt", ".pptx"],
                    "pdf": [".pdf"]
                },
                "overwrite_same_size": True
            }
            with open(args.config, "w", encoding='utf-8') as f:
                json.dump(default_config, f, indent=4)
            print(f"已生成默认配置文件: {args.config}")
        except Exception:
            pass

    converter = OfficeConverter(args.config)
    converter.run()
