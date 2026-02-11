#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
生产版 V9：路径超长优化
- 所有工作目录和输出目录都使用 C:\ctmp 下极短路径，最大程度规避 Win 路径上限。
- 解压 zip 时自动加 \\?\ 前缀。
- 兼容你的全部处理统计、日志和容错流程。
"""

import os
import subprocess
import zipfile
import shutil
import sys
from datetime import datetime

try:
    from bs4 import BeautifulSoup
    from tqdm import tqdm
except ImportError as e:
    sys.exit(f"错误：缺少必要的 Python 库 - {e.name}。请运行 'pip install beautifulsoup4 lxml tqdm' 来安装依赖。")

# ========================================================================
#                             路径配置
# ========================================================================
BASE_DIR = r"C:\ctmp"
PROCESSING_STAGE_DIR = os.path.join(BASE_DIR, "stage")
WORK_DIR = os.path.join(BASE_DIR, "work")
OUTPUT_PDF_DIR = os.path.join(BASE_DIR, "pdfs")
ROOT_SEARCH_DIR = os.path.join(BASE_DIR, "Program")  # 你可以把原始 Program 拷贝到此目录
WKHTMLTOPDF_EXECUTABLE_PATH = r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe"
SEVEN_ZIP_EXECUTABLE_PATH = r"C:\Program Files\7-Zip\7z.exe"
# ========================================================================

def to_win_long_path(path):
    """加长路径前缀，Win10+ 支持长路径，py3.6+生效。"""
    if os.name == 'nt' and not path.startswith('\\\\?\\'):
        path = os.path.abspath(path)
        return '\\\\?\\' + path
    return path

def find_files_with_exts(root_dir, exts):
    """查找指定目录下指定扩展名文件（非递归）"""
    results = []
    if not os.path.isdir(root_dir):
        return results
    with os.scandir(root_dir) as it:
        for entry in it:
            if entry.is_file() and entry.name.lower().endswith(exts):
                results.append(entry.path)
    return results

def extract_cab(cab_path, extract_dir):
    """优先 expand，失败再用 7-Zip"""
    cab_path_abs = os.path.abspath(cab_path)
    extract_dir_abs = os.path.abspath(extract_dir)
    original_cwd = os.getcwd()
    os.makedirs(extract_dir_abs, exist_ok=True)
    try:
        os.chdir(extract_dir_abs)
        cmd = ["expand", cab_path_abs, "-F:*", "."]
        subprocess.run(cmd, capture_output=True, text=True, encoding='gbk', errors='ignore', check=True)
    except Exception as e:
        tqdm.write(f"  [调试] expand.exe 尝试失败: {e}")
    finally:
        os.chdir(original_cwd)

    if find_files_with_exts(extract_dir_abs, ('.mshc', '.htm', '.html')):
        tqdm.write("  [信息] expand.exe 解压成功。")
        return

    tqdm.write(f"  [信息] expand.exe 未能有效解压，切换到 7-Zip 引擎重试...")
    try:
        os.chdir(extract_dir_abs)
        cmd = [SEVEN_ZIP_EXECUTABLE_PATH, "x", cab_path_abs, "-y"]
        subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8', errors='ignore', check=True)
    except subprocess.CalledProcessError as e:
        raise RuntimeError(f"所有解压引擎均失败！最终 7-Zip 错误: {e.stderr}")
    finally:
        os.chdir(original_cwd)

def extract_mshc(mshc_path, content_dir):
    """解压 .mshc 文件，自动加长路径前缀"""
    content_dir = to_win_long_path(content_dir)
    mshc_path = to_win_long_path(mshc_path)
    os.makedirs(content_dir, exist_ok=True)
    with zipfile.ZipFile(mshc_path, 'r') as z:
        z.extractall(content_dir)

def parse_toc(content_dir):
    nodes, root_ids = {}, []
    with os.scandir(content_dir) as it:
        html_files = [entry.name for entry in it if entry.is_file() and entry.name.lower().endswith((".htm", ".html"))]
    for fname in html_files:
        fpath = os.path.join(content_dir, fname)
        try:
            with open(fpath, 'r', encoding='utf-8', errors='ignore') as f:
                soup = BeautifulSoup(f, 'lxml')
            head = soup.find('head')
            if not head: continue
            id_meta, parent_meta, title_meta = (head.find('meta', {"name": name}) for name in ["Microsoft.Help.Id", "Microsoft.Help.TocParent", "Title"])
            title_tag = head.find('title')
            topic_id = id_meta['content'] if id_meta and 'content' in id_meta.attrs else fname
            parent_id = parent_meta['content'] if parent_meta and 'content' in parent_meta.attrs else None
            title = title_meta['content'] if title_meta and 'content' in title_meta.attrs else (title_tag.text if title_tag else fname)
            nodes[topic_id] = {"title": title, "file": fpath, "parent": parent_id, "children": []}
        except Exception as e:
            print(f"警告：解析文件 {fpath} 失败，已跳过。错误: {e}")
            continue
    for tid, node in nodes.items():
        pid = node.get("parent")
        if pid is None or pid == "-1" or pid not in nodes: root_ids.append(tid)
        elif pid in nodes: nodes[pid].setdefault("children", []).append(tid)
    for node in nodes.values():
        if "children" in node: node["children"].sort(key=lambda cid: nodes[cid]["title"])
    root_ids.sort(key=lambda rid: nodes[rid]["title"])
    return nodes, root_ids

def convert_to_pdf(nodes, root_ids, output_pdf):
    html_files = []
    def traverse(node_ids):
        for tid in node_ids:
            html_files.append(to_win_long_path(nodes[tid]["file"]))
            if "children" in nodes[tid] and nodes[tid]["children"]:
                traverse(nodes[tid]["children"])
    traverse(root_ids)
    if not html_files: return
    os.makedirs(os.path.dirname(output_pdf), exist_ok=True)
    output_pdf = to_win_long_path(output_pdf)
    cmd = [WKHTMLTOPDF_EXECUTABLE_PATH, "--enable-local-file-access", "--load-error-handling", "skip", "toc", "--toc-header-text", "目录"] + html_files + [output_pdf]
    result = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8', errors='ignore')
    if result.returncode != 0 and "Done" not in result.stderr:
        raise RuntimeError(f"PDF 转换失败: {output_pdf}\nwkhtmltopdf 错误信息: {result.stderr}")

def find_mshelpviewer_dirs(root_dir):
    result = []
    for dirpath, dirnames, _ in os.walk(root_dir):
        if "MSHelpViewer" in dirnames: result.append(os.path.join(dirpath, "MSHelpViewer"))
    return result

def process_mshv_directory(processing_dir, original_dir_for_reporting, work_dir_base, output_dir, stats):
    cab_files = find_files_with_exts(processing_dir, '.cab')
    if not cab_files:
        print(f"信息: 目录 {processing_dir} 中未找到 .cab 文件，跳过。")
        return
    stats["cab_found"] += len(cab_files)
    for cab_file in tqdm(cab_files, desc=f"处理 {os.path.basename(processing_dir)}"):
        try:
            base_name = os.path.splitext(os.path.basename(cab_file))[0]
            extract_dir = os.path.join(work_dir_base, base_name + "_cab_extract")
            extract_cab(cab_file, extract_dir)
            mshc_files = find_files_with_exts(extract_dir, '.mshc')
            html_source_dir = ""
            if mshc_files:
                tqdm.write(f"  [信息] 格式 'CAB -> MSHC'，找到 .mshc 文件，正在解压...")
                content_dir = os.path.join(work_dir_base, base_name + "_content")
                extract_mshc(mshc_files[0], content_dir)
                tqdm.write(f"  [验证] 检查 .mshc 解压结果...")
                if not find_files_with_exts(content_dir, ('.htm', '.html')):
                    raise RuntimeError(f"解压 .mshc 文件后，未能找到任何 .html 或 .htm 文件。文件可能为空或已损坏: {mshc_files[0]}")
                tqdm.write(f"  [信息] 验证成功，找到HTML文件。")
                html_source_dir = content_dir
            else:
                tqdm.write(f"  [信息] 格式 'CAB -> HTML'，未找到 .mshc 文件，直接使用当前内容。")
                html_source_dir = extract_dir
            nodes, root_ids = parse_toc(html_source_dir)
            if not nodes:
                raise RuntimeError("未能解析出任何有效的HTML页面主题，跳过PDF生成。")
            output_pdf = os.path.join(output_dir, base_name + ".pdf")
            convert_to_pdf(nodes, root_ids, output_pdf)
            stats["pdf_success"] += 1
        except Exception as e:
            stats["problematic_folders"].add(original_dir_for_reporting)
            stats["pdf_failed"] += 1
            tqdm.write(f"\n[错误] 处理文件 {cab_file} 时失败！", file=sys.stderr)
            tqdm.write(f"  - 原始文件夹: {original_dir_for_reporting}", file=sys.stderr)
            tqdm.write(f"  - 错误详情: {e}", file=sys.stderr)
            continue

def print_summary(stats, start_time):
    end_time = datetime.now()
    duration = end_time - start_time
    print("\n" + "="*60 + "\n" + " " * 22 + "运维结果报告" + "\n" + "="*60)
    print(f"任务开始时间: {start_time.strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"任务结束时间: {end_time.strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"总计耗时: {str(duration).split('.')[0]}")
    print("-" * 60 + "\n【文件夹处理统计】")
    print(f"  - 共找到 MSHelpViewer 文件夹: {stats['mshv_found']} 个")
    print(f"  - 已尝试处理 MSHelpViewer 文件夹: {stats['mshv_processed']} 个")
    print("\n【文件处理统计】")
    print(f"  - 共找到 .cab 文件总数: {stats['cab_found']}")
    print(f"  - 成功生成 PDF 文件数:  {stats['pdf_success']}")
    print(f"  - 处理失败文件数:      {stats['pdf_failed']}")
    print("-" * 60)
    if stats["problematic_folders"]:
        print("【有问题的文件夹列表 (处理过程中发生错误)】")
        for i, folder_path in enumerate(sorted(list(stats["problematic_folders"]))):
            print(f"  {i+1}. {folder_path}")
    elif stats['cab_found'] > 0:
        print("【状态】: 所有文件均已成功处理，无失败记录。")
    print("="*60)

if __name__ == "__main__":
    # =========== 文件夹准备 ===========
    for temp_dir in [PROCESSING_STAGE_DIR, WORK_DIR, OUTPUT_PDF_DIR]:
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir, ignore_errors=True)
        os.makedirs(temp_dir)

    if not os.path.isfile(WKHTMLTOPDF_EXECUTABLE_PATH):
         sys.exit("致命错误: wkhtmltopdf 路径配置不正确。")
    if not os.path.isdir(ROOT_SEARCH_DIR):
        sys.exit(f"错误: 未找到 Program 文件夹: '{ROOT_SEARCH_DIR}'（请将原始程序包拷贝到 {ROOT_SEARCH_DIR}）")

    stats = {"mshv_found": 0, "mshv_processed": 0, "cab_found": 0, "pdf_success": 0, "pdf_failed": 0, "problematic_folders": set()}
    start_time = datetime.now()

    try:
        print("[日志] 正在初始化临时工作区...")
        print(f"[日志] Phase 1: 搜索 'MSHelpViewer' 文件夹...")
        original_mshv_dirs = find_mshelpviewer_dirs(ROOT_SEARCH_DIR)
        stats["mshv_found"] = len(original_mshv_dirs)
        if not original_mshv_dirs:
            print("[日志] 未找到任何 'MSHelpViewer' 文件夹。")
        else:
            print(f"[日志] 找到 {stats['mshv_found']} 个文件夹。")
            copied_dirs_map = {}
            print(f"\n[日志] Phase 2: 将源文件夹复制到临时处理区 '{PROCESSING_STAGE_DIR}'...")
            for i, original_dir in enumerate(tqdm(original_mshv_dirs, desc="复制进度")):
                dest_dir = os.path.join(PROCESSING_STAGE_DIR, f"MSHelpViewer_{i}")
                try:
                    if os.path.exists(dest_dir): shutil.rmtree(dest_dir)
                    shutil.copytree(original_dir, dest_dir)
                    copied_dirs_map[dest_dir] = original_dir
                except Exception as e:
                    stats["problematic_folders"].add(original_dir)
            print(f"\n[日志] Phase 3: 开始处理 {len(copied_dirs_map)} 个已复制的文件夹...")
            if copied_dirs_map:
                for copied_dir, original_dir in copied_dirs_map.items():
                    stats["mshv_processed"] += 1
                    print(f"\n[日志] ==> 正在处理源文件夹: {original_dir}")
                    process_mshv_directory(copied_dir, original_dir, WORK_DIR, OUTPUT_PDF_DIR, stats)
        print_summary(stats, start_time)

    except Exception as e:
        print(f"\n[致命错误] 主流程发生严重错误: {e}", file=sys.stderr)
        print_summary(stats, start_time)
    finally:
        print("\n[日志] 正在进行最终清理...")
        for temp_path in [PROCESSING_STAGE_DIR, WORK_DIR]:
            if os.path.isdir(temp_path):
                try:
                    shutil.rmtree(temp_path, ignore_errors=True)
                    print(f"[日志] 已清理临时目录: {temp_path}")
                except OSError as e:
                    print(f"错误: 清理临时目录 {temp_path} 失败: {e}", file=sys.stderr)
