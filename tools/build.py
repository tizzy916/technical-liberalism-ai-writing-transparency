#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build.py
========
把 chapters/*.md 合并 → 用 pandoc 以 reference.docx 为样式模板，
导出一份快速预览用 .docx 文件。

用法：
    python3 build.py
        默认输出：../_deliverables/word_exports/current/技术自由主义_latest_快速预览.docx

    python3 build.py -o ../my_export.docx
        自定义输出路径

    python3 build.py --keep-merged
        保留合并后的 markdown 文件（_export/_merged.md），便于排查

依赖：
    pandoc 2.x+    （macOS: brew install pandoc）

工作目录约定：本脚本应放在 _export/ 目录下，与 reference.docx 同目录。
"""

import argparse
import re
import shutil
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent          # _export/
PAPER_DIR = HERE.parent                         # 技术自由主义/
CHAPTERS_DIR = PAPER_DIR / "chapters"
REFERENCE_DOCX = HERE / "reference.docx"
DEFAULT_OUTPUT = PAPER_DIR / "_deliverables" / "word_exports" / "current" / "技术自由主义_latest_快速预览.docx"
MERGED_MD = HERE / "_merged.md"

# 章节顺序（与 chapters/ 中的文件名严格对应）
CHAPTER_ORDER = [
    "00_前置.md",      # 封面 / 评阅人
    "00_摘要.md",      # 摘要 + Abstract
    "00_导言.md",      # 导言
    "01_信息何以自由.md",
    "02_审美革命中的技术问题.md",
    "03_技艺的分裂与重逢.md",
    "04_技艺自由主义.md",
    "05_选择走出去.md",
    "06_参考文献.md",
    "07_后记.md",      # 致谢/声明/个人简历/评语/决议书
]


def strip_yaml_frontmatter(text: str) -> str:
    """移除文件顶部的 YAML frontmatter（--- ... ---）"""
    if text.startswith("---\n"):
        end = text.find("\n---\n", 4)
        if end != -1:
            return text[end + 5:]
    return text


def merge_chapters() -> str:
    """按 CHAPTER_ORDER 顺序拼接所有章节，去掉 frontmatter"""
    pieces = []
    for name in CHAPTER_ORDER:
        path = CHAPTERS_DIR / name
        if not path.exists():
            print(f"  [警告] 缺少章节：{name}", file=sys.stderr)
            continue
        body = strip_yaml_frontmatter(path.read_text(encoding="utf-8"))
        # 章节之间留两个空行，确保 pandoc 正确识别块级元素
        pieces.append(body.rstrip() + "\n\n")
        print(f"  ✓ {name}  ({len(body):>6} chars)")
    return "".join(pieces)


def run_pandoc(input_md: Path, output_docx: Path):
    """调用 pandoc 转换"""
    cmd = [
        "pandoc",
        "-f", "gfm",
        "-t", "docx",
        "--reference-doc", str(REFERENCE_DOCX),
        str(input_md),
        "-o", str(output_docx),
    ]
    print("\n[pandoc] " + " ".join(repr(c) if " " in c else c for c in cmd))
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        print("\n[pandoc 失败]", file=sys.stderr)
        print(result.stderr, file=sys.stderr)
        sys.exit(result.returncode)
    if result.stderr.strip():
        # pandoc 偶尔有非致命警告
        print("[pandoc 提示]\n" + result.stderr.strip())


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("-o", "--output", type=Path, default=DEFAULT_OUTPUT,
                    help=f"输出 docx 路径（默认：{DEFAULT_OUTPUT.name}）")
    ap.add_argument("--keep-merged", action="store_true",
                    help="保留 _export/_merged.md 中间文件")
    args = ap.parse_args()

    # 检查 pandoc
    if shutil.which("pandoc") is None:
        sys.exit("错误：未找到 pandoc。请先安装：brew install pandoc")

    # 检查 reference.docx
    if not REFERENCE_DOCX.exists():
        sys.exit(f"错误：缺少样式模板 {REFERENCE_DOCX}\n"
                 "请把当前格式基线 docx 复制为 _export/reference.docx")

    print(f"[1/3] 合并 {len(CHAPTER_ORDER)} 个章节文件 ...")
    merged = merge_chapters()
    MERGED_MD.write_text(merged, encoding="utf-8")
    print(f"      合并稿写入 {MERGED_MD.name}（{len(merged):,} chars）")

    print(f"\n[2/3] 通过 pandoc 套用 reference.docx 样式导出 ...")
    args.output.parent.mkdir(parents=True, exist_ok=True)
    run_pandoc(MERGED_MD, args.output)

    print(f"\n[3/3] 完成 ✓")
    print(f"      输出：{args.output}")
    print(f"      大小：{args.output.stat().st_size / 1024:.1f} KB")

    if not args.keep_merged:
        try:
            MERGED_MD.unlink()
        except (PermissionError, OSError) as e:
            # 某些只读 FUSE mount 下 unlink 会失败；不致命，下次会被覆盖
            print(f"      （提示：未能删除 {MERGED_MD.name}：{e.__class__.__name__}）")
    else:
        print(f"      （已保留 {MERGED_MD.name}）")


if __name__ == "__main__":
    main()
