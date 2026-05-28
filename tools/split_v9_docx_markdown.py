#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Split the final v9 thesis Markdown exported from DOCX into repo chapters."""

from __future__ import annotations

import argparse
import re
from pathlib import Path


BROKEN_LINK_FIXES = {
    "[claude-skill-academic-writer](footer10.xml)": (
        "[claude-skill-humanities-writing-companion]"
        "(https://github.com/tizzy916/claude-skill-humanities-writing-companion)"
    ),
    "[scholar-wendao-skill](header10.xml)": (
        "[scholar-wendao-skill](https://github.com/tizzy916/scholar-wendao-skill)"
    ),
    "视角sikll": "视角 skill",
}


OUTPUTS = [
    ("00_前置.md", "# 摘　要", None),
    ("00_摘要.md", "# 摘　要", "# **目　录**"),
    ("01_技术自由主义的当代危机.md", "# 第1章 技术自由主义的当代危机", "# 第2章 "),
    ("02_信息何以自由.md", "# 第2章 信息何以自由？——控制论、信息革命与技术自由主义的思想源流", "# 第3章 "),
    ("03_审美革命及其技术问题.md", "# 第3章 审美革命及其技术问题——从朗西埃到斯蒂格勒", "# 第4章 "),
    ("04_技艺的分裂与重逢.md", "# 第4章 技－艺的分裂与重逢", "# 第5章 "),
    ("05_技艺自由主义.md", "# 第5章 “技艺自由主义”：回归技术自由的本义", "# 第6章 "),
    ("06_结语_选择走出去.md", "# 第6章 结语——“选择走出去”", "# 参考文献"),
    ("07_参考文献.md", "# 参考文献", "# 致　谢"),
    ("08_后记与声明.md", "# 致　谢", None),
]


def apply_known_fixes(text: str) -> str:
    for old, new in BROKEN_LINK_FIXES.items():
        text = text.replace(old, new)
    return text


def find_line(lines: list[str], marker: str, start: int = 0) -> int:
    for index in range(start, len(lines)):
        if lines[index].startswith(marker):
            return index
    raise ValueError(f"Cannot find marker: {marker}")


def parse_footnotes(lines: list[str]) -> tuple[int, dict[str, list[str]]]:
    first = None
    for index, line in enumerate(lines):
        if re.match(r"^\[\^([0-9]+)\]:", line):
            first = index
            break
    if first is None:
        return len(lines), {}

    notes: dict[str, list[str]] = {}
    current_id: str | None = None
    current_lines: list[str] = []

    for line in lines[first:]:
        match = re.match(r"^\[\^([0-9]+)\]:", line)
        if match:
            if current_id is not None:
                notes[current_id] = current_lines
            current_id = match.group(1)
            current_lines = [line]
        elif current_id is not None:
            current_lines.append(line)

    if current_id is not None:
        notes[current_id] = current_lines

    return first, notes


def append_used_footnotes(section: str, notes: dict[str, list[str]]) -> str:
    used = sorted({int(note_id) for note_id in re.findall(r"\[\^([0-9]+)\]", section)})
    if not used:
        return section.rstrip() + "\n"

    pieces = [section.rstrip(), ""]
    for number in used:
        note = notes.get(str(number))
        if note:
            pieces.extend(note)
            pieces.append("")
    return "\n".join(pieces).rstrip() + "\n"


def slice_section(lines: list[str], start_marker: str, end_marker: str | None, stop_at: int) -> str:
    if start_marker == "# 摘　要" and end_marker is None:
        start = 0
        end = find_line(lines, start_marker)
    else:
        start = find_line(lines, start_marker)
        end = stop_at
        if end_marker is not None:
            end = find_line(lines, end_marker, start + 1)
    return "\n".join(lines[start:end]).strip() + "\n"


def write_chapters(input_markdown: Path, chapters_dir: Path) -> None:
    raw = input_markdown.read_text(encoding="utf-8")
    raw = apply_known_fixes(raw)
    lines = raw.splitlines()
    footnote_start, notes = parse_footnotes(lines)
    body_lines = lines[:footnote_start]

    chapters_dir.mkdir(parents=True, exist_ok=True)
    for old_file in chapters_dir.glob("*.md"):
        old_file.unlink()

    for filename, start_marker, end_marker in OUTPUTS:
        section = slice_section(body_lines, start_marker, end_marker, footnote_start)
        section = append_used_footnotes(section, notes)
        (chapters_dir / filename).write_text(section, encoding="utf-8")
        print(f"wrote {filename}")


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input_markdown", type=Path, help="Markdown exported from the v9 DOCX")
    parser.add_argument(
        "--chapters-dir",
        type=Path,
        default=Path("paper/chapters"),
        help="Destination chapter directory",
    )
    args = parser.parse_args()
    write_chapters(args.input_markdown, args.chapters_dir)


if __name__ == "__main__":
    main()
