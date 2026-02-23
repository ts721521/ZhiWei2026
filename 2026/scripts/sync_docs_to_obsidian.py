import argparse
import json
import os
import shutil
from datetime import datetime


def _load_config(config_path):
    with open(config_path, "r", encoding="utf-8") as f:
        return json.load(f)


def _resolve_obsidian_root(cfg):
    if os.name == "nt":
        root = cfg.get("obsidian_root_win")
        if root:
            return os.path.abspath(root)
    root = cfg.get("obsidian_root") or ""
    if root:
        return os.path.abspath(root)
    return ""


def _project_root():
    """Project root (2026 directory). When script is in scripts/, use parent dir."""
    _script_dir = os.path.dirname(os.path.abspath(__file__))
    if os.path.basename(_script_dir) == "scripts":
        return os.path.dirname(_script_dir)
    return _script_dir


def _iter_markdown_files(project_root):
    root_entries = sorted(os.listdir(project_root))
    for name in root_entries:
        p = os.path.join(project_root, name)
        if os.path.isfile(p) and name.lower().endswith(".md"):
            yield p

    docs_root = os.path.join(project_root, "docs")
    if os.path.isdir(docs_root):
        for cur, _, files in os.walk(docs_root):
            for fn in sorted(files):
                if fn.lower().endswith(".md"):
                    yield os.path.join(cur, fn)


def _determine_category_and_dest(rel_path):
    """
    Determine the category and destination path within Obsidian.

    Args:
        rel_path: path relative to project root (e.g. 'README.md' or 'docs/foo.md')

    Returns:
        (category_name, dest_sub_path)
    """
    path_lower = rel_path.lower()
    filename = os.path.basename(rel_path)

    # 1. Product / Project Level
    if filename in ["README.md", "README_zh.md", "readme.md"]:
        return "00_项目主页", filename
    if "使用说明" in filename or "user_guide" in path_lower:
        return "01_使用手册", filename
    if "打包" in filename or "deployment" in path_lower:
        return "02_部署指南", filename
    if "changelog" in path_lower or "更新" in filename or "history" in path_lower:
        return "03_更新日志", filename

    # 2. Technical / Dev Docs
    if (
        rel_path.startswith("docs")
        or "dev" in path_lower
        or "开发" in filename
        or "design" in path_lower
    ):
        # Keep nested structure for docs folder, but strip 'docs/' prefix if present
        if rel_path.startswith("docs"):
            sub = rel_path[5:]  # strip 'docs/' or 'docs\\'
            return "10_技术开发", sub
        return "10_技术开发", filename

    # 3. Default
    return "99_其他文档", filename


def sync_docs(config_path):
    project_root = _project_root()
    cfg = _load_config(config_path)
    obsidian_root = _resolve_obsidian_root(cfg)
    if not obsidian_root:
        raise RuntimeError("obsidian_root is empty in config")

    program_name = str(cfg.get("obsidian_program_name") or "ZhiWei").strip() or "ZhiWei"
    # Root for this program in Obsidian
    prog_root_in_obsidian = os.path.join(obsidian_root, program_name)

    # We will clear the target folders we manage to ensure clean sync (optional but good for '分级' changes)
    # But safety first: let's just overwrite.

    copied = []

    # Categorized list for MOC
    # Structure: { "Category": [ (filename, relative_path_in_obsidian) ] }
    moc_data = {}

    for src in _iter_markdown_files(project_root):
        rel = os.path.relpath(src, project_root)
        category, sub_path = _determine_category_and_dest(rel)

        # Destination: <ObsidianRoot>/<Program>/<Category>/<SubPath>
        dst = os.path.join(prog_root_in_obsidian, category, sub_path)

        os.makedirs(os.path.dirname(dst), exist_ok=True)
        shutil.copy2(src, dst)

        copied.append(dst)

        if category not in moc_data:
            moc_data[category] = []

        # For MOC link
        name_no_ext = os.path.splitext(os.path.basename(src))[0]
        moc_data[category].append(name_no_ext)

    # Generate MOC (Map of Content) at Program Root
    index_path = os.path.join(prog_root_in_obsidian, f"00_{program_name}_索引.md")

    lines = [
        f"# {program_name} 文档索引",
        "",
        f"> Generated at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        "",
    ]

    # Sort categories
    sorted_cats = sorted(moc_data.keys())

    for cat in sorted_cats:
        lines.append(f"## {cat}")
        for name in sorted(moc_data[cat]):
            lines.append(f"- [[{name}]]")
        lines.append("")

    with open(index_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")

    return prog_root_in_obsidian, copied, index_path


def main():
    parser = argparse.ArgumentParser()
    default_config = os.path.join(_project_root(), "config.json")
    parser.add_argument("--config", default=default_config)
    args = parser.parse_args()

    dst_root, copied, index_path = sync_docs(args.config)
    print(f"Synced docs to: {dst_root}")
    print(f"Copied markdown files: {len(copied)}")
    print(f"Index file: {index_path}")


if __name__ == "__main__":
    main()
